"""Stateless PPTX -> per-slide thumbnails/extracts conversion sidecar.

No database access, no persistent volume mount, and no published port
outside the Docker Compose network (see ../docker-compose.yml's `render`
service) -- this container has nothing to protect and nothing to exfiltrate
even if compromised. Every request works entirely inside a fresh temp
directory that's deleted before the response is returned; nothing survives
past one request.

POST /convert with a multipart 'file' field (a .pptx/.potx) returns a zip:
  manifest.json  -- {"slides": [{index, title, thumbnail, insertMode,
                     reconstructSpec, pptx}, ...]}
  slide-N.png    -- whitespace-trimmed PNG thumbnail for slide N (1-based)
  slide-N.pptx   -- present only for insertMode "file": that slide alone,
                     as its own single-slide presentation

Pipeline verified directly against a real multi-shape .pptx before this was
written: LibreOffice (--convert-to pdf) -> pdftoppm (per-slide PNG) ->
ImageMagick (-trim +repage) -> python-pptx (title hint, single-slide
extraction, and reconstruct/file classification -- the same technique and
the same classify_shape_tree/extract_reconstruct_spec logic
scripts/slice-catalog-source.py already uses for the offline bulk-seed
pipeline, ported here so admin bulk uploads get the same one-click
"reconstruct" fidelity for simple shapes instead of always falling back to
the heavier temp-slide/copy-paste "file" flow).
"""

import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from lxml import etree
from flask import Flask, jsonify, request, send_file
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # real templates carry embedded fonts/images

SUBPROCESS_TIMEOUT_SECONDS = 120
THUMBNAIL_DPI = 150
THINK_CELL_MARKER = "think-cell"


class ConversionError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.message = message
        self.status = status


@app.errorhandler(ConversionError)
def handle_conversion_error(err):
    return jsonify({"error": err.message}), err.status


@app.errorhandler(413)
def handle_too_large(_err):
    return jsonify({"error": "File too large."}), 413


@app.get("/health")
def health():
    return jsonify({"ok": True})


def run(cmd):
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=SUBPROCESS_TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired:
        raise ConversionError(f"{cmd[0]} timed out.", status=422)
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", "replace")[:500]
        raise ConversionError(f"{cmd[0]} failed: {stderr}", status=422)
    return result


def extract_title_hint(slide):
    """Prefers a real Title/Center-Title placeholder; falls back to the
    first non-empty text run on the slide. Same preference order already
    established in scripts/slice-catalog-source.py."""
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type in (
            PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.CENTER_TITLE,
        ):
            text = shape.text_frame.text.strip() if shape.has_text_frame else ""
            if text:
                return text
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                return text[:80]
    return None


def slice_single_slide(source_path, slide_index, out_path):
    """Reloads the source presentation fresh and deletes every slide except
    slide_index -- python-pptx has no clone operation, so mutating one
    shared Presentation object across iterations would progressively delete
    slides out from under later iterations."""
    prs = Presentation(source_path)
    slide_id_list = prs.slides._sldIdLst
    slide_ids = list(slide_id_list)
    for i, slide_id in enumerate(slide_ids):
        if i != slide_index:
            slide_id_list.remove(slide_id)
    prs.save(out_path)


# ---------------------------------------------------------------------------
# reconstruct-vs-file classification, ported directly from
# scripts/slice-catalog-source.py -- that script's own module docstring and
# per-function comments carry the full "why" (which OOXML shape features
# have no PowerPoint JS API equivalent, why an unmapped preset falls back to
# 'file' instead of guessing, etc.); this is a straight port, not a
# reinterpretation, so consult that file for the reasoning if it's not
# repeated verbatim here.
# ---------------------------------------------------------------------------


def is_think_cell_placeholder(shape) -> bool:
    return THINK_CELL_MARKER in shape._element.xml


def emu_to_pt(value):
    return round(Emu(value).pt, 2) if value is not None else None


def has_cust_geom(element) -> bool:
    return element.find(".//" + qn("a:custGeom")) is not None


def has_picture(element) -> bool:
    return element.tag == qn("p:pic") or element.find(".//" + qn("p:pic")) is not None


def has_other_graphic_frame(element) -> bool:
    return element.tag == qn("p:graphicFrame") or element.find(".//" + qn("p:graphicFrame")) is not None


def has_custom_bullet_char(element) -> bool:
    return element.find(".//" + qn("a:buChar")) is not None


# Maps an OOXML <a:prstGeom prst="..."> value to the exact string
# PowerPoint.GeometricShapeType's TypeScript enum uses. NOT a mechanical
# capitalization -- a deliberately conservative, hand-checked subset against
# the real enum list (see slice-catalog-source.py). A prst not in this table
# routes to 'file' mode rather than risk addGeometricShape silently
# accepting a wrong-but-validly-spelled value.
PRST_TO_GEOMETRIC_SHAPE_TYPE = {
    "rect": "Rectangle",
    "ellipse": "Ellipse",
    "roundRect": "RoundRectangle",
    "triangle": "Triangle",
    "rtTriangle": "RightTriangle",
    "diamond": "Diamond",
    "parallelogram": "Parallelogram",
    "trapezoid": "Trapezoid",
    "pentagon": "Pentagon",
    "hexagon": "Hexagon",
    "heptagon": "Heptagon",
    "octagon": "Octagon",
    "decagon": "Decagon",
    "dodecagon": "Dodecagon",
    "star4": "Star4",
    "star5": "Star5",
    "star6": "Star6",
    "star7": "Star7",
    "star8": "Star8",
    "star10": "Star10",
    "star12": "Star12",
    "star16": "Star16",
    "star24": "Star24",
    "star32": "Star32",
    "chevron": "Chevron",
    "donut": "Donut",
    "heart": "Heart",
    "sun": "Sun",
    "moon": "Moon",
    "cube": "Cube",
    "can": "Can",
    "cloud": "Cloud",
    "wave": "Wave",
    "plus": "Plus",
    "arc": "Arc",
    "chord": "Chord",
    "bevel": "Bevel",
    "frame": "Frame",
    "corner": "Corner",
    "pie": "Pie",
    "blockArc": "BlockArc",
    "smileyFace": "SmileyFace",
    "lightningBolt": "LightningBolt",
    "plaque": "Plaque",
    "teardrop": "Teardrop",
    "homePlate": "HomePlate",
    "rightArrow": "RightArrow",
    "leftArrow": "LeftArrow",
    "upArrow": "UpArrow",
    "downArrow": "DownArrow",
    "leftRightArrow": "LeftRightArrow",
    "upDownArrow": "UpDownArrow",
    "leftBracket": "LeftBracket",
    "rightBracket": "RightBracket",
    "leftBrace": "LeftBrace",
    "rightBrace": "RightBrace",
    "wedgeRectCallout": "WedgeRectCallout",
    # Deliberately NOT mapped: prst="line" has no GeometricShapeType
    # equivalent -- a straight line needs PowerPoint.Slide.addLine, a
    # different API the reconstruct path doesn't call.
}


# a:pPr algn -> PowerPoint.ParagraphHorizontalAlignment's exact string
# values (confirmed against @types/office-js: Left/Center/Right/Justify/
# JustifyLow/Distributed/ThaiDistributed). PP_ALIGN.to_xml() round-trips a
# python-pptx alignment member straight back to its raw OOXML value -- same
# technique already used for MSO_THEME_COLOR below -- so this table only
# maps that raw value to Office.js's naming, it doesn't re-derive it.
ALGN_TO_OFFICEJS_ALIGNMENT = {
    "l": "Left",
    "ctr": "Center",
    "r": "Right",
    "just": "Justify",
    "justLow": "JustifyLow",
    "dist": "Distributed",
    "thaiDist": "ThaiDistributed",
}

# a:bodyPr anchor (+ anchorCtr) -> PowerPoint.TextVerticalAlignment's exact
# string values (Top/Middle/Bottom/TopCentered/MiddleCentered/
# BottomCentered). The "Centered" variants correspond to anchorCtr="1" (the
# whole text block is also centered horizontally as a group) -- a separate
# OOXML attribute from paragraph-level algn, combined here since Office.js
# models both as one enum.
ANCHOR_TO_OFFICEJS_VERTICAL_ALIGNMENT = {"t": "Top", "ctr": "Middle", "b": "Bottom"}


def classify_shape_tree(shape) -> str:
    """Returns 'file' if anything under this shape can't be reconstructed, else 'reconstruct'."""
    el = shape._element
    if has_cust_geom(el) or has_picture(el) or has_other_graphic_frame(el) or has_custom_bullet_char(el):
        return "file"
    if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
        return "reconstruct"
    prst_el = el.find(".//" + qn("a:prstGeom"))
    prst = prst_el.get("prst") if prst_el is not None else None
    if prst not in PRST_TO_GEOMETRIC_SHAPE_TYPE:
        return "file"
    return "reconstruct"


# ---------------------------------------------------------------------------
# Theme-color resolution. A color set via this app's own Fill/Line/Text
# color picker (src/features/fillLineColors.ts) -- or via PowerPoint's own
# theme-based color swatches, which is the common case for any shape not
# explicitly given a custom RGB color -- is stored as <a:schemeClr val="..">
# (e.g. "accent1"), not a literal RGB value. python-pptx's ColorFormat.rgb
# raises AttributeError on this color type ("no .rgb property on color type
# '_SchemeColor'") -- confirmed directly, and without handling it here every
# theme-colored fill/line/font color was silently dropped (caught by a
# blanket except and turned into None), which is why "recreate" mode was
# losing exactly the styles reported: colors are almost always theme-based
# in real content, explicit RGB is the unusual case.
#
# Resolving requires walking the real OOXML relationship chain: the shape's
# slide -> slide layout -> slide master -> theme part (for the actual
# <a:clrScheme> RGB values, keyed by raw slot names dk1/lt1/dk2/lt2/
# accent1-6/hlink/folHlink) and the slide master's own <p:clrMap> (which
# remaps the four semantic bg1/tx1/bg2/tx2 names PowerPoint's UI/API uses to
# whichever raw slot the current master actually points them at -- a "dark"
# master can map bg1 to dk1 instead of lt1. accent1-6/hlink/folHlink map to
# themselves, but only by convention; only the actual clrMap is Definitive.
# Verified against a real theme-colored test file before writing this: the
# resolved hex values matched the theme's own <a:clrScheme> exactly.
# ---------------------------------------------------------------------------


def _load_master_theme(master):
    """Returns (clr_map, clr_scheme) for a slide master, both plain dicts of
    XML slot name -> value. Cached per (per-request) theme_cache dict, since
    a whole category's worth of slides typically share one master and
    re-walking the relationship chain/re-parsing the theme XML per shape
    would be wasteful."""
    clr_map_el = master.element.find(qn("p:clrMap"))
    clr_map = dict(clr_map_el.attrib) if clr_map_el is not None else {}

    theme_part = None
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            theme_part = rel.target_part
            break

    scheme = {}
    if theme_part is not None:
        theme_el = etree.fromstring(theme_part.blob)
        clr_scheme_el = theme_el.find(".//" + qn("a:clrScheme"))
        if clr_scheme_el is not None:
            for child in clr_scheme_el:
                tag = etree.QName(child).localname
                srgb = child.find(qn("a:srgbClr"))
                sys_clr = child.find(qn("a:sysClr"))
                if srgb is not None:
                    scheme[tag] = srgb.get("val")
                elif sys_clr is not None:
                    scheme[tag] = sys_clr.get("lastClr")

    return clr_map, scheme


# Raw OOXML schemeClr value (pre-clrMap -- MSO_THEME_COLOR.to_xml() returns
# exactly this, the semantic role name PowerPoint's own UI/API uses, not the
# raw theme-XML slot it happens to map to on this particular master) ->
# Office.js's ThemeColorScheme.getThemeColor() role names, confirmed
# against src/features/fillLineColors.ts's own THEME_COLOR_ROLES list.
# hlink/folHlink are deliberately absent: that file's getThemeColors() never
# exposes them either, so there's no live-lookup path for a hyperlink color
# at this app's PowerPointApi floor -- those colors fall back to the static
# hex snapshot below and nothing else.
#
# MSO_THEME_COLOR.to_xml() can also return the four RAW scheme slot names
# (dk1/lt1/dk2/lt2) instead of the semantic ones, when a shape's
# <a:schemeClr val="..."> happens to reference a slot directly rather than
# through clrMap (confirmed directly: python-pptx round-trips whichever
# form was actually written -- DARK_1/LIGHT_1/etc. produce "dk1"/"lt1"/...,
# while TEXT_1/BACKGROUND_1/etc. produce "tx1"/"bg1"/...). Mapped here to
# the same Office.js role their semantic counterpart would typically
# resolve to, since tx1->dk1 and bg1->lt1 is clrMap's standard/default
# configuration -- an approximation for the (uncommon) direct-slot case,
# not a guess for the common semantic case, which is exact.
XML_NAME_TO_OFFICEJS_ROLE = {
    "bg1": "Light1",
    "tx1": "Dark1",
    "bg2": "Light2",
    "tx2": "Dark2",
    "lt1": "Light1",
    "dk1": "Dark1",
    "lt2": "Light2",
    "dk2": "Dark2",
    "accent1": "Accent1",
    "accent2": "Accent2",
    "accent3": "Accent3",
    "accent4": "Accent4",
    "accent5": "Accent5",
    "accent6": "Accent6",
}


def resolve_theme_color(shape, mso_theme_color, theme_cache):
    """Returns (hex, officejs_role). hex is a static snapshot of the SOURCE
    presentation's theme at extraction time -- kept as a fallback. When
    officejs_role is non-None, the insert-time client can instead resolve
    the color live against the TARGET presentation's own theme (see
    src/features/libraryInsert.ts's two-phase theme-role resolve), so a
    shape colored e.g. "Accent1" in the source lands as whatever Accent1
    actually is in the presentation it's inserted into, rather than a
    baked-in hex that's wrong (commonly white-on-white) the moment the two
    presentations' themes differ."""
    try:
        xml_name = MSO_THEME_COLOR.to_xml(mso_theme_color)
    except Exception:
        return None, None
    officejs_role = XML_NAME_TO_OFFICEJS_ROLE.get(xml_name)
    try:
        master = shape.part.slide_layout.slide_master
    except Exception:
        return None, officejs_role
    cache_key = id(master.part)
    if cache_key not in theme_cache:
        theme_cache[cache_key] = _load_master_theme(master)
    clr_map, scheme = theme_cache[cache_key]
    slot = clr_map.get(xml_name, xml_name)
    hexval = scheme.get(slot)
    return (f"#{hexval}" if hexval else None), officejs_role


def resolve_color_spec(color_format, shape, theme_cache):
    """Resolves a python-pptx ColorFormat to {"hex", "themeRole"}: hex a
    real #RRGGBB string, themeRole a live Office.js theme-color role name
    when this color is theme-derived and mappable, else None. Returns None
    entirely if no color is set, or for the rare color types (HSL/preset/
    system/scRGB) this doesn't handle -- safer to omit than to guess."""
    try:
        color_type = color_format.type
    except Exception:
        return None
    if color_type == MSO_COLOR_TYPE.RGB:
        return {"hex": f"#{color_format.rgb}", "themeRole": None}
    if color_type == MSO_COLOR_TYPE.SCHEME:
        hexval, role = resolve_theme_color(shape, color_format.theme_color, theme_cache)
        return {"hex": hexval, "themeRole": role} if hexval else None
    return None


def extract_fill(shape, theme_cache):
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        if str(fill.type) == "MSO_FILL_TYPE.SOLID (1)" or fill.type == 1:
            color = resolve_color_spec(fill.fore_color, shape, theme_cache)
            return {"type": "solid", "color": color} if color else None
        return None
    except Exception:
        return None


def extract_line(shape, theme_cache):
    try:
        line = shape.line
        if line.fill.type is None:
            return None
        width_pt = emu_to_pt(line.width) if line.width else None
        color = resolve_color_spec(line.color, shape, theme_cache)
        return {"color": color, "widthPt": width_pt} if color else None
    except Exception:
        return None


def extract_text_spec(shape, theme_cache):
    if not shape.has_text_frame:
        return None
    paragraphs = []
    for p in shape.text_frame.paragraphs:
        runs = []
        for r in p.runs:
            runs.append(
                {
                    "text": r.text,
                    "bold": r.font.bold,
                    "italic": r.font.italic,
                    "size": r.font.size.pt if r.font.size else None,
                    "fontName": r.font.name,
                    "color": resolve_color_spec(r.font.color, shape, theme_cache) if r.font.color else None,
                }
            )
        align = None
        try:
            if p.alignment is not None:
                align = ALGN_TO_OFFICEJS_ALIGNMENT.get(PP_ALIGN.to_xml(p.alignment))
        except Exception:
            align = None
        paragraphs.append({"level": p.level, "bullet": None, "align": align, "runs": runs})
    return paragraphs


def extract_vertical_alignment(shape):
    """Maps a:bodyPr anchor/anchorCtr to Office.js's TextVerticalAlignment
    string. Returns None (not a guessed default) when the shape has no text
    frame or no anchor attribute at all -- the caller only sets
    TextFrame.verticalAlignment when this is non-null, leaving PowerPoint's
    own default (Top) untouched otherwise."""
    if not shape.has_text_frame:
        return None
    body_pr = shape._element.find(".//" + qn("a:bodyPr"))
    if body_pr is None:
        return None
    base = ANCHOR_TO_OFFICEJS_VERTICAL_ALIGNMENT.get(body_pr.get("anchor"))
    if base is None:
        return None
    return f"{base}Centered" if body_pr.get("anchorCtr") == "1" else base


def extract_reconstruct_spec(shape, theme_cache):
    is_text_box = shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX
    spec = {
        "kind": "textBox" if is_text_box else "geometricShape",
        "left": emu_to_pt(shape.left),
        "top": emu_to_pt(shape.top),
        "width": emu_to_pt(shape.width),
        "height": emu_to_pt(shape.height),
        "rotation": shape.rotation,
        "fill": extract_fill(shape, theme_cache),
        "line": extract_line(shape, theme_cache),
        "verticalAlignment": extract_vertical_alignment(shape),
        "paragraphs": extract_text_spec(shape, theme_cache),
    }
    if not is_text_box:
        prst_el = shape._element.find(".//" + qn("a:prstGeom"))
        prst = prst_el.get("prst") if prst_el is not None else None
        spec["presetGeometry"] = PRST_TO_GEOMETRIC_SHAPE_TYPE[prst]
    return spec


@app.post("/convert")
def convert():
    upload = request.files.get("file")
    if upload is None or upload.filename == "":
        raise ConversionError("No file uploaded (expected multipart field 'file').")
    if not upload.filename.lower().endswith((".pptx", ".potx")):
        raise ConversionError("File must be a .pptx or .potx.")

    work_dir = tempfile.mkdtemp(prefix="render-")
    try:
        source_path = os.path.join(work_dir, "source.pptx")
        upload.save(source_path)

        try:
            prs = Presentation(source_path)
            slide_count = len(prs.slides)
        except Exception as exc:
            raise ConversionError(f"Could not read the file as a presentation: {exc}")
        if slide_count == 0:
            raise ConversionError("Presentation has no slides.")

        # Isolated LibreOffice profile per request -- a profile dir shared
        # across concurrent requests can lock/conflict.
        profile_dir = os.path.join(work_dir, "loprofile")
        run([
            "soffice", "--headless", "--invisible", "--nocrashreport",
            "--nodefault", "--nofirststartwizard", "--nolockcheck", "--nologo",
            "--norestore", f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to", "pdf", "--outdir", work_dir, source_path,
        ])
        pdf_path = os.path.join(work_dir, "source.pdf")
        if not os.path.exists(pdf_path):
            raise ConversionError("LibreOffice did not produce a PDF.", status=422)

        run(["pdftoppm", "-png", "-r", str(THUMBNAIL_DPI), pdf_path, os.path.join(work_dir, "slide")])

        # pdftoppm zero-pads its page-number suffix based on total page
        # count (slide-1.png for <10 pages, slide-01.png for 10-99, etc.),
        # so a sorted glob -- not a guessed filename -- is the only correct
        # way to recover slide order for an arbitrary-sized deck.
        produced_pngs = sorted(
            name for name in os.listdir(work_dir) if re.fullmatch(r"slide-\d+\.png", name)
        )
        if len(produced_pngs) != slide_count:
            raise ConversionError(
                f"Expected {slide_count} rendered slide(s), got {len(produced_pngs)}.",
                status=422,
            )

        # One per request (not module-level) so a slide master's cached
        # theme never leaks across unrelated uploads -- see
        # _load_master_theme's docstring for why caching happens at all.
        theme_cache = {}

        manifest = {"slides": []}
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for index, raw_png_name in enumerate(produced_pngs):
                page_num = index + 1
                raw_png = os.path.join(work_dir, raw_png_name)
                trimmed_png = os.path.join(work_dir, f"trimmed-{page_num}.png")
                run(["convert", raw_png, "-trim", "+repage", trimmed_png])
                thumb_name = f"slide-{page_num}.png"
                zf.write(trimmed_png, thumb_name)

                slide = prs.slides[index]
                real_shapes = [s for s in slide.shapes if not is_think_cell_placeholder(s)]

                entry = {
                    "index": index,
                    "title": extract_title_hint(slide),
                    "thumbnail": thumb_name,
                }

                if real_shapes and all(classify_shape_tree(s) == "reconstruct" for s in real_shapes):
                    if len(real_shapes) == 1:
                        spec = extract_reconstruct_spec(real_shapes[0], theme_cache)
                    else:
                        spec = {
                            "kind": "group",
                            "shapes": [extract_reconstruct_spec(s, theme_cache) for s in real_shapes],
                        }
                    entry["insertMode"] = "reconstruct"
                    entry["reconstructSpec"] = spec
                    entry["pptx"] = None
                else:
                    slide_pptx = os.path.join(work_dir, f"slide-{page_num}.pptx")
                    slice_single_slide(source_path, index, slide_pptx)
                    pptx_name = f"slide-{page_num}.pptx"
                    zf.write(slide_pptx, pptx_name)
                    entry["insertMode"] = "file"
                    entry["reconstructSpec"] = None
                    entry["pptx"] = pptx_name

                manifest["slides"].append(entry)
            zf.writestr("manifest.json", json.dumps(manifest))

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype="application/zip", download_name="converted.zip")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8090)
