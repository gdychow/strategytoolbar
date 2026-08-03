const { unzipSync, strFromU8 } = require("fflate");

// http://render:8090 matches the `render` service name/port in
// docker-compose.yml — reachable only over Compose's internal network,
// never published to the host. Node 20 ships fetch/FormData/Blob as
// globals (undici-backed), so no new dependency is needed for this call.
const RENDER_SERVICE_URL = process.env.RENDER_SERVICE_URL || "http://render:8090";

/**
 * Sends one .pptx to the render sidecar and returns its per-slide
 * thumbnail/single-slide-pptx/title-hint results, unzipped in memory —
 * nothing touches disk here. See render-sidecar/app.py for the actual
 * conversion pipeline (LibreOffice -> pdftoppm -> ImageMagick trim ->
 * python-pptx slice+title-hint).
 */
async function convertPptxToSlides(buffer, filename) {
  const form = new FormData();
  form.append("file", new Blob([buffer]), filename);

  let res;
  try {
    res = await fetch(`${RENDER_SERVICE_URL}/convert`, { method: "POST", body: form });
  } catch (err) {
    throw new Error(`Could not reach the render service: ${err instanceof Error ? err.message : String(err)}`);
  }
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Render service returned HTTP ${res.status}`);
  }

  const zipBytes = new Uint8Array(await res.arrayBuffer());
  const files = unzipSync(zipBytes);
  const manifest = JSON.parse(strFromU8(files["manifest.json"]));
  // Each item is either insertMode "reconstruct" (a JSON spec, pptx null —
  // the one-click native insert path, built via the same
  // classify_shape_tree/extract_reconstruct_spec logic
  // scripts/slice-catalog-source.py uses for the offline bulk-seed
  // pipeline) or "file" (a single-slide pptx Buffer, reconstructSpec null
  // — the temp-slide/copy-paste path), never both.
  return manifest.slides.map((slide) => ({
    index: slide.index,
    title: slide.title,
    thumbnail: Buffer.from(files[slide.thumbnail]),
    insertMode: slide.insertMode,
    reconstructSpec: slide.reconstructSpec ?? null,
    pptx: slide.pptx ? Buffer.from(files[slide.pptx]) : null,
  }));
}

module.exports = { convertPptxToSlides };
