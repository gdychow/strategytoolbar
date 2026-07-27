import { unzipSync } from "fflate";

/**
 * Reads the presentation's actual "Custom Colours" list — the same named
 * palette the original VBA tool's frmColourPicker exposed via
 * CustomColours.xml. No Office.js API exposes this: it lives as
 * <a:custClrLst> inside the applied theme's own theme XML part
 * (ppt/theme/themeN.xml), a sibling of <a:clrScheme> — confirmed by
 * unzipping the real "Wiley Template 2024 Berry.potx" and reading it
 * directly. PowerPoint.CustomXmlPart (checked against its own docs) only
 * reaches the package's separate customXml/itemN.xml parts — an unrelated
 * OOXML part type used for arbitrary developer-attached metadata, not
 * theme content. So the only way to reach it is to read the file's own raw
 * bytes and walk the same relationship chain PowerPoint itself uses to
 * resolve a slide's theme: slide -> layout -> master -> theme.
 */

export interface CustomColor {
  name: string;
  hex: string;
}

export function isCustomColorsSupported(): boolean {
  return typeof Office !== "undefined" && typeof Office.context?.document?.getFileAsync === "function";
}

function getFile(): Promise<Office.File> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
      else resolve(result.value);
    });
  });
}

function getSlice(file: Office.File, index: number): Promise<Office.Slice> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
      else resolve(result.value);
    });
  });
}

function closeFile(file: Office.File): Promise<void> {
  return new Promise((resolve) => file.closeAsync(() => resolve()));
}

/** Reads the whole current file as raw zip bytes, in ≤4MB slices (Common API, universal PowerPoint platform support). */
async function getFileBytes(): Promise<Uint8Array> {
  const file = await getFile();
  try {
    const sliceBuffers: number[][] = [];
    let total = 0;
    for (let i = 0; i < file.sliceCount; i++) {
      const slice = await getSlice(file, i);
      const data = slice.data as number[];
      sliceBuffers.push(data);
      total += data.length;
    }
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const buf of sliceBuffers) {
      bytes.set(buf, offset);
      offset += buf.length;
    }
    return bytes;
  } finally {
    await closeFile(file);
  }
}

/** The zero-based ordinal position (matching <p:sldIdLst>'s order) of the currently selected slide, or null if none is selected. */
async function getSelectedSlideIndex(): Promise<number | null> {
  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();
    if (selectedSlides.items.length === 0) return null;
    const slide = selectedSlides.items[0];
    slide.load("index");
    await context.sync();
    return slide.index;
  });
}

/** Resolves a relative OOXML Target (e.g. "../slideLayouts/slideLayout1.xml") against the directory a .rels file lives alongside. */
function resolvePath(baseDir: string, target: string): string {
  const parts = baseDir.split("/").filter(Boolean);
  for (const segment of target.split("/").filter(Boolean)) {
    if (segment === "..") parts.pop();
    else if (segment !== ".") parts.push(segment);
  }
  return parts.join("/");
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "application/xml");
}

/** Finds the Target of the first <Relationship> whose Type ends with typeSuffix (e.g. "/slideMaster"). */
function findRelationshipTarget(relsXml: string, typeSuffix: string): string | null {
  const doc = parseXml(relsXml);
  const rels = Array.from(doc.getElementsByTagName("Relationship"));
  const match = rels.find((r) => r.getAttribute("Type")?.endsWith(typeSuffix));
  return match?.getAttribute("Target") ?? null;
}

/** Follows a part's own .rels file to the target of a given relationship type, resolved to a full zip-entry path. */
function followRelationship(readText: (path: string) => string | null, partPath: string, typeSuffix: string): string | null {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = partPath.substring(0, lastSlash);
  const fileName = partPath.substring(lastSlash + 1);
  const relsXml = readText(`${dir}/_rels/${fileName}.rels`);
  if (!relsXml) return null;
  const target = findRelationshipTarget(relsXml, typeSuffix);
  return target ? resolvePath(dir, target) : null;
}

/** The r:id (e.g. "rId9") of the Nth <p:sldId> entry in <p:sldIdLst> — its display order, matching Slide.index. */
function findNthSlideRelationshipId(presentationXml: string, slideIndex: number): string | null {
  const doc = parseXml(presentationXml);
  const entry = doc.getElementsByTagName("p:sldId")[slideIndex];
  return entry?.getAttribute("r:id") ?? null;
}

/** Resolves an r:id from presentation.xml.rels to its ppt/slides/slideN.xml path. */
function resolveSlidePath(presentationRels: string, rId: string): string | null {
  const doc = parseXml(presentationRels);
  const rel = Array.from(doc.getElementsByTagName("Relationship")).find((r) => r.getAttribute("Id") === rId);
  const target = rel?.getAttribute("Target");
  return target ? resolvePath("ppt", target) : null;
}

function parseCustomColorList(themeXml: string): CustomColor[] {
  const doc = parseXml(themeXml);
  const colors: CustomColor[] = [];
  for (const entry of Array.from(doc.getElementsByTagName("a:custClr"))) {
    const name = entry.getAttribute("name");
    const hex = entry.getElementsByTagName("a:srgbClr")[0]?.getAttribute("val");
    if (name && hex) colors.push({ name, hex: `#${hex.toUpperCase()}` });
  }
  return colors;
}

// Only the parts needed to walk slide -> layout -> master -> theme are worth
// decompressing — media/embeddings can dwarf everything else in a real deck.
const RELEVANT_PART_PATTERN = /^ppt\/(presentation\.xml$|_rels\/presentation\.xml\.rels$|slides\/|slideLayouts\/|slideMasters\/|theme\/)/;

export async function getCustomColors(): Promise<CustomColor[]> {
  if (!isCustomColorsSupported()) return [];

  const slideIndex = await getSelectedSlideIndex();
  if (slideIndex === null) return [];

  const zipBytes = await getFileBytes();
  const decoder = new TextDecoder();
  const entries = unzipSync(zipBytes, { filter: (f) => RELEVANT_PART_PATTERN.test(f.name) });
  const readText = (path: string): string | null => {
    const data = entries[path];
    return data ? decoder.decode(data) : null;
  };

  const presentationXml = readText("ppt/presentation.xml");
  const presentationRels = readText("ppt/_rels/presentation.xml.rels");
  if (!presentationXml || !presentationRels) return [];

  const rId = findNthSlideRelationshipId(presentationXml, slideIndex);
  const slidePath = rId ? resolveSlidePath(presentationRels, rId) : null;
  if (!slidePath) return [];

  const layoutPath = followRelationship(readText, slidePath, "/slideLayout");
  const masterPath = layoutPath ? followRelationship(readText, layoutPath, "/slideMaster") : null;
  const themePath = masterPath ? followRelationship(readText, masterPath, "/theme") : null;
  const themeXml = themePath ? readText(themePath) : null;
  if (!themeXml) return [];

  return parseCustomColorList(themeXml);
}
