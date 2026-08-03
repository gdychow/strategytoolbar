"""Stateless PPTX -> per-slide thumbnails/extracts conversion sidecar.

No database access, no persistent volume mount, and no published port
outside the Docker Compose network (see ../docker-compose.yml's `render`
service) -- this container has nothing to protect and nothing to exfiltrate
even if compromised. Every request works entirely inside a fresh temp
directory that's deleted before the response is returned; nothing survives
past one request.

POST /convert with a multipart 'file' field (a .pptx/.potx) returns a zip:
  manifest.json          -- {"slides": [{"index", "title", "thumbnail", "pptx"}, ...]}
  slide-N.png            -- whitespace-trimmed PNG thumbnail for slide N (1-based)
  slide-N.pptx            -- that slide alone, as its own single-slide presentation

Pipeline verified directly against a real multi-shape .pptx before this was
written: LibreOffice (--convert-to pdf) -> pdftoppm (per-slide PNG) ->
ImageMagick (-trim +repage) -> python-pptx (title hint + single-slide
extraction, the same technique scripts/slice-catalog-source.py already uses
for the bulk-seed pipeline).
"""

import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from flask import Flask, jsonify, request, send_file
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # real templates carry embedded fonts/images

SUBPROCESS_TIMEOUT_SECONDS = 120
THUMBNAIL_DPI = 150


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

        manifest = {"slides": []}
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for index, raw_png_name in enumerate(produced_pngs):
                page_num = index + 1
                raw_png = os.path.join(work_dir, raw_png_name)
                trimmed_png = os.path.join(work_dir, f"trimmed-{page_num}.png")
                run(["convert", raw_png, "-trim", "+repage", trimmed_png])

                slide_pptx = os.path.join(work_dir, f"slide-{page_num}.pptx")
                slice_single_slide(source_path, index, slide_pptx)

                thumb_name = f"slide-{page_num}.png"
                pptx_name = f"slide-{page_num}.pptx"
                zf.write(trimmed_png, thumb_name)
                zf.write(slide_pptx, pptx_name)
                manifest["slides"].append({
                    "index": index,
                    "title": extract_title_hint(prs.slides[index]),
                    "thumbnail": thumb_name,
                    "pptx": pptx_name,
                })
            zf.writestr("manifest.json", json.dumps(manifest))

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype="application/zip", download_name="converted.zip")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8090)
