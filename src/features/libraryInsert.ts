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
  if (!res.ok) throw new Error(`Failed to load the "${category}" library (${res.status}).`);
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
  if (!res.ok) throw new Error(`Failed to fetch that item (${res.status}).`);
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

export async function insertFileItem(itemId: number): Promise<FileInsertHandle> {
  const base64 = await fetchFileBase64(itemId);

  return PowerPoint.run(async (context) => {
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
    const tempSlide = slides.items[originalIndex + 1];

    tempSlide.shapes.load("items/id");
    await context.sync();
    tempSlide.setSelectedShapes(tempSlide.shapes.items.map((s) => s.id));
    context.presentation.setSelectedSlides([tempSlide.id]);
    await context.sync();

    return { tempSlideId: tempSlide.id, originalSlideId };
  });
}

export async function finishFileInsert(handle: FileInsertHandle): Promise<void> {
  await PowerPoint.run(async (context) => {
    const tempSlide = context.presentation.slides.getItemOrNullObject(handle.tempSlideId);
    await context.sync();
    if (!tempSlide.isNullObject) {
      tempSlide.delete();
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

function buildShape(slide: PowerPoint.Slide, spec: ShapeSpec): PowerPoint.Shape {
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
