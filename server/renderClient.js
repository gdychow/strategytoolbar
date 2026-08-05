const { unzipSync, strFromU8 } = require("fflate");

// http://render:8090 matches the `render` service name/port in
// docker-compose.yml — reachable only over Compose's internal network,
// never published to the host. Node 20 ships fetch/FormData/Blob as
// globals (undici-backed), so no new dependency is needed for this call.
const RENDER_SERVICE_URL = process.env.RENDER_SERVICE_URL || "http://render:8090";

// Node's default fetch (undici) times out around 5 minutes with no way to
// distinguish "server is slow" from "server is unreachable" beyond a bare
// "fetch failed". This call is always made from server.js's background
// upload-job processor (see the async job queue around convertPptxToSlides's
// call site), never from inside a live browser request — so there's no
// user-facing HTTP deadline forcing a short timeout here. A large,
// picture-heavy, many-slide source file (e.g. one of the Maps category
// decks) can legitimately take several minutes in aggregate even though no
// single step inside app.py exceeds its own SUBPROCESS_TIMEOUT_SECONDS —
// confirmed directly against a real report: the render container stayed
// healthy the whole time (no crash, no OOM, no error logged) while Node's
// fetch gave up and reported a bare connection failure. 15 minutes is a
// generous ceiling for the largest realistic single upload.
const CONVERT_TIMEOUT_MS = 15 * 60 * 1000;

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
    res = await fetch(`${RENDER_SERVICE_URL}/convert`, {
      method: "POST",
      body: form,
      signal: AbortSignal.timeout(CONVERT_TIMEOUT_MS),
    });
  } catch (err) {
    // err.cause carries undici's real underlying reason (a connection
    // refusal, a reset, this abort's own timeout, etc.) — err.message alone
    // is just the generic "fetch failed" wrapper, which is what made this
    // failure mode hard to diagnose the first time around.
    const cause = err instanceof Error && err.cause ? `: ${err.cause.message || err.cause}` : "";
    throw new Error(`Could not reach the render service: ${err instanceof Error ? err.message : String(err)}${cause}`);
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
