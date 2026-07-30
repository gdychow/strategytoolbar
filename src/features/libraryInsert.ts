/**
 * Tier 3: the shared content library's insert engine. Two paths, chosen
 * per item at content-prep time by scripts/slice-catalog-source.py (see
 * db/init.sql for the exact classification rule):
 *
 * - 'reconstruct': built directly on the current slide via addTextBox/
 *   addGeometricShape/addGroup, from a structured spec. True one-click
 *   insert, matching the VBA original's actual behaviour.
 * - 'file': PowerPoint JS has no API for custom vector geometry or custom
 *   bullet characters (confirmed while porting L_Other_Tweaks.bas — see
 *   otherTweaks.ts), so these items are pre-sliced into their own
 *   single-slide .pptx and brought in via insertSlidesFromBase64, which
 *   can only insert a new slide, never merge into the current one (no
 *   shape-level copy/import or clipboard API exists anywhere in Office.js
 *   or the Common API). insertFileItem() automates everything around that
 *   gap except the final copy/paste keystrokes themselves: it selects the
 *   new content and hands off to the user, then finishFileInsert() cleans
 *   up the temporary slide once they're done.
 */

import { extractErrorMessage } from "../core/ui";

export interface TextRunSpec {
  text: string;
  bold: boolean | null;
  italic: boolean | null;
  size: number | null;
  fontName: string | null;
  color: string | null;
}

export interface ParagraphSpec {
  level: number;
  bullet: { char: string; font: string } | null;
  runs: TextRunSpec[];
}

export interface ShapeSpec {
  kind: "textBox" | "geometricShape";
  presetGeometry?: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  fill: { type: "solid"; color: string } | null;
  line: { color: string; widthPt: number } | null;
  paragraphs: ParagraphSpec[] | null;
}

export type ReconstructSpec = ShapeSpec | { kind: "group"; shapes: ShapeSpec[] };

export interface CatalogItem {
  id: number;
  title: string;
  insertMode: "reconstruct" | "file" | "unicode-char";
  reconstructSpec: ReconstructSpec | null;
  unicodeChar: string | null;
  thumbnailUrl: string | null;
  groupId: number | null;
  groupName: string | null;
  tags: string[];
  // Task Pane Phase 13: non-null only for items returned from
  // GET /api/catalog/personal — lets the gallery/task pane tell "this is
  // mine" without a second round-trip. Always null for shared/global items.
  ownerOid: string | null;
  ownerTid: string | null;
  // Task Pane Phase 14: non-null only for items returned from
  // GET /api/catalog/company — lets the gallery/task pane tell "this
  // belongs to my company" without a second round-trip. Always null for
  // global and personal items.
  companyDomain: string | null;
}

export interface CatalogGroup {
  id: number;
  name: string;
  sortOrder: number;
}

export interface CatalogResponse {
  groups: CatalogGroup[];
  items: CatalogItem[];
}

export interface FileInsertHandle {
  tempSlideId: string;
  originalSlideId: string;
}

/** insertSlidesFromBase64 and its options are gated at PowerPointApi 1.2 — far below every other requirement-set this app already checks. */
export function isLibraryInsertSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
}

/**
 * Response is { groups, items }, not a bare item array (Phase 5) — the
 * gallery dialog needs the category's admin-defined groups in their own
 * order to render group headers correctly.
 */
export async function fetchCatalog(category: string): Promise<CatalogResponse> {
  const res = await fetch(`/api/catalog/${category}`);
  if (!res.ok) throw new Error(await extractErrorMessage(res, `Failed to load the "${category}" library (${res.status}).`));
  return res.json();
}

// ---------------------------------------------------------------------------
// 'file' mode — insert as a temp slide, hand off for a native copy/paste,
// clean up once the user's done.
// ---------------------------------------------------------------------------

const fileCache = new Map<number, string>();

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function fetchFileBase64(itemId: number): Promise<string> {
  const cached = fileCache.get(itemId);
  if (cached) return cached;

  const res = await fetch(`/api/catalog/file/${itemId}`);
  if (!res.ok) throw new Error(await extractErrorMessage(res, `Failed to fetch that item (${res.status}).`));
  const base64 = arrayBufferToBase64(await res.arrayBuffer());
  fileCache.set(itemId, base64);
  return base64;
}

async function getTargetSlide(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide> {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items");
  await context.sync();
  if (selected.items.length > 0) return selected.items[0];

  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  if (slides.items.length === 0) throw new Error("This presentation has no slides to insert next to.");
  return slides.items[slides.items.length - 1];
}

/** Shared by insertFileItem and insertFileItemAsNewSlide — inserts the item's slide immediately after the current one and locates it, without deciding what happens to it afterward. */
async function insertFileSlideAfterCurrent(
  context: PowerPoint.RequestContext,
  base64: string
): Promise<{ originalSlideId: string; newSlide: PowerPoint.Slide }> {
  const originalSlide = await getTargetSlide(context);
  originalSlide.load("id");
  await context.sync();
  const originalSlideId = originalSlide.id;

  context.presentation.insertSlidesFromBase64(base64, {
    formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
    targetSlideId: originalSlideId,
  });
  await context.sync();

  // insertSlidesFromBase64 always places the new slide immediately after
  // targetSlideId — there's no way to learn the new slide's ID ahead of
  // time, so re-locate it by position once the collection reflects it.
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const originalIndex = slides.items.findIndex((s) => s.id === originalSlideId);
  if (originalIndex === -1 || originalIndex + 1 >= slides.items.length) {
    throw new Error("Couldn't locate the newly inserted slide.");
  }
  return { originalSlideId, newSlide: slides.items[originalIndex + 1] };
}

export async function insertFileItem(itemId: number): Promise<FileInsertHandle> {
  const base64 = await fetchFileBase64(itemId);

  return PowerPoint.run(async (context) => {
    const { originalSlideId, newSlide: tempSlide } = await insertFileSlideAfterCurrent(context, base64);

    tempSlide.shapes.load("items/id");
    await context.sync();
    tempSlide.setSelectedShapes(tempSlide.shapes.items.map((s) => s.id));
    context.presentation.setSelectedSlides([tempSlide.id]);
    await context.sync();

    return { tempSlideId: tempSlide.id, originalSlideId };
  });
}

/**
 * The direct-insert route for 'file'-mode items: skips the temp-slide/
 * copy-paste/Finish dance entirely — the inserted slide is the permanent
 * result, exactly like a plain "insert new slide" action. No handle is
 * returned since there's nothing left to clean up.
 */
export async function insertFileItemAsNewSlide(itemId: number): Promise<void> {
  const base64 = await fetchFileBase64(itemId);

  await PowerPoint.run(async (context) => {
    const { newSlide } = await insertFileSlideAfterCurrent(context, base64);
    context.presentation.setSelectedSlides([newSlide.id]);
    await context.sync();
  });
}

/**
 * `keep: true` ("Finish and Keep") leaves the temp slide in the deck as a
 * permanent one instead of deleting it — for when the copy/paste has
 * already been done and the user decides the extra slide is worth keeping
 * too, not just scratch space. Either way, selection returns to the
 * slide the user started on.
 */
export async function finishFileInsert(handle: FileInsertHandle, options: { keep?: boolean } = {}): Promise<void> {
  await PowerPoint.run(async (context) => {
    if (!options.keep) {
      const tempSlide = context.presentation.slides.getItemOrNullObject(handle.tempSlideId);
      await context.sync();
      if (!tempSlide.isNullObject) {
        tempSlide.delete();
      }
    }
    context.presentation.setSelectedSlides([handle.originalSlideId]);
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// 'reconstruct' mode — build directly on the current slide, one click, no
// temporary slide involved.
// ---------------------------------------------------------------------------

function applyParagraphs(shape: PowerPoint.Shape, paragraphs: ParagraphSpec[]): void {
  const textRange = shape.textFrame.textRange;
  textRange.text = paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\r");

  // TextRange has no .paragraphs collection — per-paragraph formatting goes
  // through getSubstring(start, length) instead. Run-level formatting is
  // similarly re-applied per paragraph here (one style per paragraph,
  // matching every item currently seeded — none mixes styles within a
  // single paragraph). BulletFormat can't set a custom character/font
  // (confirmed in otherTweaks.ts) — items needing that are routed to
  // 'file' mode at content-prep time, so paragraphs reaching here are
  // expected to have no bullet.
  let charOffset = 0;
  for (const paragraph of paragraphs) {
    const paragraphText = paragraph.runs.map((r) => r.text).join("");
    const run = paragraph.runs[0];
    if (run && paragraphText.length > 0) {
      const range = textRange.getSubstring(charOffset, paragraphText.length);
      if (run.bold !== null) range.font.bold = run.bold;
      if (run.italic !== null) range.font.italic = run.italic;
      if (run.size !== null) range.font.size = run.size;
      if (run.fontName !== null) range.font.name = run.fontName;
      if (run.color !== null) range.font.color = run.color;
      // indentLevel needs PowerPointApi 1.10 (this feature otherwise only
      // needs 1.2) — only touch it when a paragraph actually needs
      // indenting, so the common flat case works on older builds too.
      if (paragraph.level > 0) range.paragraphFormat.indentLevel = paragraph.level;
    }
    charOffset += paragraphText.length + 1; // +1 for the "\r" paragraph break
  }
}

export function buildShape(slide: PowerPoint.Slide, spec: ShapeSpec): PowerPoint.Shape {
  const options: PowerPoint.ShapeAddOptions = {
    left: spec.left,
    top: spec.top,
    width: spec.width,
    height: spec.height,
  };

  const shape =
    spec.kind === "textBox"
      ? slide.shapes.addTextBox("", options)
      : slide.shapes.addGeometricShape(spec.presetGeometry as PowerPoint.GeometricShapeType, options);

  // rotation needs PowerPointApi 1.10 (this feature otherwise only needs
  // 1.2) — only touch it when actually rotated, so the common unrotated
  // case (every item currently seeded) works on older builds too.
  if (spec.rotation !== 0) shape.rotation = spec.rotation;

  if (spec.fill) {
    shape.fill.setSolidColor(spec.fill.color);
  } else {
    shape.fill.clear();
  }

  if (spec.line) {
    shape.lineFormat.visible = true;
    shape.lineFormat.color = spec.line.color;
    shape.lineFormat.weight = spec.line.widthPt;
  } else {
    shape.lineFormat.visible = false;
  }

  if (spec.paragraphs && spec.paragraphs.length > 0) {
    applyParagraphs(shape, spec.paragraphs);
  }

  return shape;
}

export async function insertReconstructedItem(spec: ReconstructSpec): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getTargetSlide(context);

    if (spec.kind === "group") {
      const children = spec.shapes.map((shapeSpec) => buildShape(slide, shapeSpec));
      await context.sync();
      if (children.length > 1) {
        slide.shapes.addGroup(children);
      }
    } else {
      buildShape(slide, spec);
    }

    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Admin-only: edit an existing item's graphic natively in PowerPoint, or add
// a new one from the current slide (Task Pane Phase 12). Two Office.js
// primitives make this possible with no server-side rendering
// infrastructure at all — both confirmed directly against @types/office-js,
// both gated at PowerPointApi 1.8 (already an implicit baseline elsewhere in
// this app, e.g. insertReconstructedItem's addGroup call above):
//   - Slide.exportAsBase64() — the slide as its own base64-encoded .pptx.
//   - Slide.getImageAsBase64({ width, height }) — a PNG rendered by
//     PowerPoint itself, exactly as faithful as the qlmanage-generated
//     thumbnails already in use, with nothing new to install anywhere.
// ---------------------------------------------------------------------------

/** Slide.exportAsBase64/getImageAsBase64 are both PowerPointApi 1.8. */
export function isAdminLibraryEditSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
}

/**
 * The 'reconstruct'-mode equivalent of insertFileItem: builds the item onto
 * a fresh temporary slide (not the current one) so an admin can edit it
 * natively without disturbing whatever they're actually working on, and
 * returns the same FileInsertHandle shape so the caller can treat both
 * insert modes identically from here on (export the result, then
 * finishFileInsert to clean up).
 */
export async function insertReconstructedItemOnTempSlide(spec: ReconstructSpec): Promise<FileInsertHandle> {
  return PowerPoint.run(async (context) => {
    const originalSlide = await getTargetSlide(context);
    originalSlide.load("id");
    await context.sync();
    const originalSlideId = originalSlide.id;

    context.presentation.slides.add(); // PowerPointApi 1.3 — appends at the end
    await context.sync();
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const tempSlide = slides.items[slides.items.length - 1];

    if (spec.kind === "group") {
      const children = spec.shapes.map((shapeSpec) => buildShape(tempSlide, shapeSpec));
      await context.sync();
      if (children.length > 1) {
        tempSlide.shapes.addGroup(children);
      }
    } else {
      buildShape(tempSlide, spec);
    }
    await context.sync();

    tempSlide.shapes.load("items/id");
    await context.sync();
    tempSlide.setSelectedShapes(tempSlide.shapes.items.map((s) => s.id));
    context.presentation.setSelectedSlides([tempSlide.id]);
    await context.sync();

    return { tempSlideId: tempSlide.id, originalSlideId };
  });
}

export interface AdminExport {
  pptxBase64: string;
  thumbnailBase64: string;
}

const THUMBNAIL_MAX_DIMENSION = 480;
const THUMBNAIL_PADDING_PX = 12;

/**
 * Crops a captured slide image down to its actual content's bounding box
 * — library source content is a single shape/graphic on an otherwise
 * blank slide, so a raw, unsized getImageAsBase64() capture is mostly
 * white space around whatever the admin actually drew. Downscales the
 * trimmed result afterward if it's still larger than a sane thumbnail
 * size (this only ever shrinks; a small crop stays small).
 *
 * Assumes a white/near-white background, matching every item already in
 * the library (all built from blank-background source content) — a
 * slide with a genuinely colored background just won't find anything to
 * trim around and falls back to the untrimmed capture, not a wrong one.
 *
 * Pure Canvas/Image work, nothing Office-specific — runs outside
 * PowerPoint.run on purpose, since it has nothing to do with it.
 */
async function trimThumbnail(base64Png: string): Promise<string> {
  const img = new Image();
  const loaded = new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("Couldn't process the captured thumbnail."));
  });
  img.src = `data:image/png;base64,${base64Png}`;
  await loaded;

  const sourceCanvas = document.createElement("canvas");
  sourceCanvas.width = img.naturalWidth;
  sourceCanvas.height = img.naturalHeight;
  const sourceCtx = sourceCanvas.getContext("2d");
  if (!sourceCtx) return base64Png;
  sourceCtx.drawImage(img, 0, 0);

  const { data, width, height } = sourceCtx.getImageData(0, 0, sourceCanvas.width, sourceCanvas.height);
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const isBackground = data[i + 3] === 0 || (data[i] > 250 && data[i + 1] > 250 && data[i + 2] > 250);
      if (!isBackground) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) return base64Png; // nothing but background found — leave it as-is rather than crop to nothing

  minX = Math.max(0, minX - THUMBNAIL_PADDING_PX);
  minY = Math.max(0, minY - THUMBNAIL_PADDING_PX);
  maxX = Math.min(width - 1, maxX + THUMBNAIL_PADDING_PX);
  maxY = Math.min(height - 1, maxY + THUMBNAIL_PADDING_PX);
  const cropWidth = maxX - minX + 1;
  const cropHeight = maxY - minY + 1;

  const scale = Math.min(1, THUMBNAIL_MAX_DIMENSION / Math.max(cropWidth, cropHeight));
  const outWidth = Math.round(cropWidth * scale);
  const outHeight = Math.round(cropHeight * scale);

  const outCanvas = document.createElement("canvas");
  outCanvas.width = outWidth;
  outCanvas.height = outHeight;
  const outCtx = outCanvas.getContext("2d");
  if (!outCtx) return base64Png;
  outCtx.drawImage(sourceCanvas, minX, minY, cropWidth, cropHeight, 0, 0, outWidth, outHeight);

  return outCanvas.toDataURL("image/png").split(",")[1];
}

/** Captures a specific slide (by ID) for saving back to the library — used once an admin has finished editing it. */
export async function exportSlideForAdmin(slideId: string): Promise<AdminExport> {
  const { pptxBase64, rawThumbnailBase64 } = await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    const pptxResult = slide.exportAsBase64();
    // No width/height — "the true size of the slide is used" per the
    // API's own docs, which is exactly what trimThumbnail needs to work
    // with (a scaled-down capture would just mean trimming fewer, blurrier
    // pixels).
    const imgResult = slide.getImageAsBase64();
    await context.sync();
    return { pptxBase64: pptxResult.value, rawThumbnailBase64: imgResult.value };
  });
  return { pptxBase64, thumbnailBase64: await trimThumbnail(rawThumbnailBase64) };
}

/** Captures whatever slide is currently selected (or the last slide, via the same fallback insertFileItem/insertReconstructedItem already use) — used to add a brand-new item. */
export async function exportCurrentSlideForAdmin(): Promise<AdminExport> {
  const { pptxBase64, rawThumbnailBase64 } = await PowerPoint.run(async (context) => {
    const slide = await getTargetSlide(context);
    const pptxResult = slide.exportAsBase64();
    const imgResult = slide.getImageAsBase64();
    await context.sync();
    return { pptxBase64: pptxResult.value, rawThumbnailBase64: imgResult.value };
  });
  return { pptxBase64, thumbnailBase64: await trimThumbnail(rawThumbnailBase64) };
}

// ---------------------------------------------------------------------------
// 'unicode-char' mode — insert a single character into the current text
// selection/cursor position. Ported from the VBA original's InsertChar,
// which called TextRange.InsertSymbol on the active selection; PowerPoint
// JS has no equivalent method (PowerPoint.TextRange's real surface is
// start/length/text/getSubstring/getParentTextFrame — confirmed against
// @types/office-js), so this uses the same
// presentation.getSelectedTextRangeOrNullObject() pattern already proven in
// otherTweaks.ts's setBulletsVisible. Setting .text on a zero-length
// selection (a bare cursor) inserts there; on a real selection it replaces
// it — both match ordinary typing behaviour.
//
// With nothing selected at all, the VBA original's ppSelectionNone branch
// dropped the character into its own small, centered text box rather than
// erroring — ported here the same way (a shape-selected-but-not-in-text-
// edit-mode branch also existed in VBA but isn't ported; that case still
// falls through to the same text-box fallback rather than a dedicated
// third path, since PowerPoint JS has no way to distinguish "a shape is
// selected" from "nothing is selected" via getSelectedTextRangeOrNullObject
// alone).
async function insertUnicodeCharAsTextBox(context: PowerPoint.RequestContext, char: string): Promise<void> {
  const slide = await getTargetSlide(context);
  const pageSetup = context.presentation.pageSetup;
  pageSetup.load("slideWidth,slideHeight");
  await context.sync();

  const width = 50;
  const height = 50;
  const shape = slide.shapes.addTextBox(char, {
    left: (pageSetup.slideWidth - width) / 2,
    top: (pageSetup.slideHeight - height) / 2,
    width,
    height,
  });
  shape.textFrame.wordWrap = false;
  shape.textFrame.leftMargin = 0;
  shape.textFrame.rightMargin = 0;
  shape.textFrame.topMargin = 0;
  shape.textFrame.bottomMargin = 0;
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
  shape.textFrame.textRange.font.size = 16;
  await context.sync();
}

export async function insertUnicodeChar(char: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const textRange = context.presentation.getSelectedTextRangeOrNullObject();
    await context.sync();
    if (textRange.isNullObject) {
      await insertUnicodeCharAsTextBox(context, char);
      return;
    }
    textRange.text = char;
    await context.sync();
  });
}

export async function insertCatalogItem(item: CatalogItem): Promise<FileInsertHandle | null> {
  if (item.insertMode === "reconstruct") {
    if (!item.reconstructSpec) throw new Error(`"${item.title}" is missing its reconstruction data.`);
    await insertReconstructedItem(item.reconstructSpec);
    return null;
  }
  if (item.insertMode === "unicode-char") {
    if (!item.unicodeChar) throw new Error(`"${item.title}" is missing its character data.`);
    await insertUnicodeChar(item.unicodeChar);
    return null;
  }
  return insertFileItem(item.id);
}
