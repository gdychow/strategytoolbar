#!/usr/bin/env python3
"""
One-time content-prep step for Tier 3's shared catalog. For a given
boilerplate library .pptx (e.g. Text.pptx), classifies each slide's real
content (the inert think-cell placeholder frame is always stripped first,
regardless of mode) as either:

  reconstruct - every real shape is plain preset geometry (<a:prstGeom>) or
    a plain text box, so it can be rebuilt with full fidelity via
    addGeometricShape/addTextBox/addGroup at insert time (see
    src/features/libraryInsert.ts). No .pptx file is written for these
    items - a draft reconstruct_spec JSON fragment is printed instead, for
    a human to check and fold into db/seed/catalog-<category>.json.

  file - anything else (custom <a:custGeom> geometry, embedded pictures,
    tables/other graphicFrames) - no PowerPoint JS API can reconstruct
    these, so the slide is sliced out into its own minimal single-slide
    .pptx under <output-dir>/<category>/, to be inserted via
    insertSlidesFromBase64 at runtime.

Usage:
    python3 scripts/slice-catalog-source.py <source.pptx> <category-slug> [output-dir]

Example:
    python3 scripts/slice-catalog-source.py \
        "../Package Files/StrategyToolbar/BoilerPlates/16x9/Text.pptx" \
        text data/catalog
"""
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

THINK_CELL_MARKER = "think-cell"
EMU_PER_POINT = 12700


def is_think_cell_placeholder(shape) -> bool:
    return THINK_CELL_MARKER in shape._element.xml


def emu_to_pt(value) -> float:
    return round(Emu(value).pt, 2) if value is not None else None


def has_cust_geom(element) -> bool:
    return element.find(".//" + qn("a:custGeom")) is not None


def has_picture(element) -> bool:
    return element.find(".//" + qn("p:pic")) is not None


def has_other_graphic_frame(element) -> bool:
    # Any graphicFrame other than the (already-excluded) think-cell one -
    # tables, charts, SmartArt. Not handled by the reconstruct path.
    return element.find(".//" + qn("p:graphicFrame")) is not None


def has_custom_bullet_char(element) -> bool:
    # PowerPoint.BulletFormat (confirmed in src/features/otherTweaks.ts,
    # while porting L_Other_Tweaks.bas) only exposes type
    # (None/Numbered/Unnumbered), a fixed numbering-style enum, and
    # visible - no way to set a specific bullet character, font, or
    # colour. A <a:buChar> (an explicit fixed bullet glyph, as opposed to
    # <a:buAutoNum>'s auto-incrementing numbering) can't be reproduced, so
    # route it to the file-based path instead of guessing at a lossy
    # approximation.
    return element.find(".//" + qn("a:buChar")) is not None


def classify_shape_tree(shape) -> str:
    """Returns 'file' if anything under this shape can't be reconstructed, else 'reconstruct'."""
    el = shape._element
    if has_cust_geom(el) or has_picture(el) or has_other_graphic_frame(el) or has_custom_bullet_char(el):
        return "file"
    return "reconstruct"


def extract_fill(shape):
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        if str(fill.type) == "MSO_FILL_TYPE.SOLID (1)" or fill.type == 1:
            return {"type": "solid", "color": f"#{fill.fore_color.rgb}"}
        return {"type": str(fill.type), "note": "non-solid fill - verify manually against ShapeFill API"}
    except Exception:
        return None


def extract_line(shape):
    try:
        line = shape.line
        if line.fill.type is None:
            return None
        width_pt = emu_to_pt(line.width) if line.width else None
        try:
            color = f"#{line.color.rgb}"
        except Exception:
            color = None
        return {"color": color, "widthPt": width_pt}
    except Exception:
        return None


def extract_bullet(paragraph_xml):
    bu_char = paragraph_xml.find(".//" + qn("a:buChar"))
    if bu_char is not None:
        return bu_char.get("char")
    if paragraph_xml.find(".//" + qn("a:buNone")) is not None:
        return None
    return "(inherited - verify against shape's lstStyle or slide layout)"


def extract_text_spec(shape):
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
                    "color": (f"#{r.font.color.rgb}" if r.font.color and r.font.color.type is not None else None),
                }
            )
        paragraphs.append({"level": p.level, "bullet": extract_bullet(p._p), "runs": runs})
    return paragraphs


def extract_reconstruct_spec(shape):
    is_text_box = shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX
    spec = {
        "kind": "textBox" if is_text_box else "geometricShape",
        "left": emu_to_pt(shape.left),
        "top": emu_to_pt(shape.top),
        "width": emu_to_pt(shape.width),
        "height": emu_to_pt(shape.height),
        "rotation": shape.rotation,
        "fill": extract_fill(shape),
        "line": extract_line(shape),
        "text": extract_text_spec(shape),
    }
    if not is_text_box:
        prst_el = shape._element.find(".//" + qn("a:prstGeom"))
        spec["presetGeometry"] = prst_el.get("prst") if prst_el is not None else None
        spec["_todo"] = "map presetGeometry (raw OOXML prst value) to a PowerPoint.GeometricShapeType enum member by hand"
    return spec


def extract_text_hint(shape) -> str:
    if not shape.has_text_frame:
        return ""
    return " / ".join(p.text for p in shape.text_frame.paragraphs if p.text.strip())[:80]


def slice_single_slide(source_path: Path, slide_index: int, dest_path: Path) -> None:
    """Duplicates the source presentation, keeps only slide_index, strips the think-cell placeholder, saves to dest_path."""
    prs = Presentation(str(source_path))
    slide_id_list = prs.slides._sldIdLst
    slide_id_elements = list(slide_id_list)
    for i, slide_id_elm in enumerate(slide_id_elements):
        if i == slide_index:
            continue
        rId = slide_id_elm.get(qn("r:id"))
        prs.part.drop_rel(rId)
        slide_id_list.remove(slide_id_elm)

    remaining_slide = prs.slides[0]
    for shape in list(remaining_slide.shapes):
        if is_think_cell_placeholder(shape):
            shape._element.getparent().remove(shape._element)

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dest_path))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    source_path = Path(sys.argv[1])
    category = sys.argv[2]
    output_root = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("data/catalog")
    category_dir = output_root / category

    prs = Presentation(str(source_path))
    results = []

    for index, slide in enumerate(prs.slides):
        real_shapes = [s for s in slide.shapes if not is_think_cell_placeholder(s)]
        if not real_shapes:
            print(f"slide {index + 1}: no real content shape found, skipping", file=sys.stderr)
            continue

        mode = "reconstruct" if all(classify_shape_tree(s) == "reconstruct" for s in real_shapes) else "file"
        text_hint = " | ".join(filter(None, (extract_text_hint(s) for s in real_shapes)))

        if mode == "file":
            filename = f"{category}-{index + 1:03d}.pptx"
            dest = category_dir / filename
            slice_single_slide(source_path, index, dest)
            results.append(
                {
                    "slideIndex": index + 1,
                    "insertMode": "file",
                    "sourceFile": f"{category}/{filename}",
                    "textHint": text_hint,
                }
            )
        else:
            if len(real_shapes) == 1:
                spec = extract_reconstruct_spec(real_shapes[0])
            else:
                spec = {"kind": "group", "shapes": [extract_reconstruct_spec(s) for s in real_shapes]}
            results.append(
                {
                    "slideIndex": index + 1,
                    "insertMode": "reconstruct",
                    "textHint": text_hint,
                    "reconstructSpec": spec,
                }
            )

    print(json.dumps(results, indent=2))
    if any(r["insertMode"] == "file" for r in results):
        print(f"\nSliced 'file'-mode items into: {category_dir}", file=sys.stderr)


if __name__ == "__main__":
    main()
