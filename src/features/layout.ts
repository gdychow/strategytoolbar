/**
 * Port of K_Layout.bas (align/distribute/position tools).
 *
 * Two known, unavoidable gaps vs. the VBA original, both because the JS API
 * has no way to read which specific preset shape (chevron vs. pentagon vs.
 * plain rectangle) an existing shape is — see the shape-copy prototype
 * findings. `alignAngles` and `chevronAlign` can no longer verify the
 * selection is actually chevrons/pentagons/arrows before adjusting them;
 * they trust the caller. Also, `shape.adjustments` requires PowerPointApi
 * 1.10 (Mac support only since Dec 2025) — call isAdjustmentsSupported()
 * before offering these two in the UI on older Mac builds.
 *
 * Bigger Mac-parity finding made while porting this module: slide dimensions
 * (`pageSetup.slideWidth`/`slideHeight`) also require API 1.10, not 1.4/1.5
 * as assumed earlier. That gates every "centre on slide" function here
 * (centreOnSlideHorizontal/Vertical/HV, twoUpTemplateHorizontalLeft/Right,
 * centreHalfSlide) — call isSlideDimensionsSupported() before offering them.
 * hvAlign/hvDistribute/switchPositions/get-set position/fixSize/EdgeJoin/
 * enhancedDistribute are unaffected (API 1.4/1.5/1.9 only).
 *
 * Not ported here (out of Tier 1 scope): InsertSlideGuides/InsertGuides
 * (library-copy, Tier 3), ResetTemplateCentrePointsMHFi (hardcoded offsets
 * tied to the old A4/Letter/OnScreen page-size model, not the current
 * 16:9-only templates), ResetTemplateCentrePointsCustom (opens frmMargins —
 * Tier 2 dialog, not yet built).
 */

import {
  getSavedPosition,
  setSavedPosition,
  getSavedBottomPosition,
  setSavedBottomPosition,
  getTemplateMargins,
  setTemplateMargins,
} from "../core/settings";

const DEG_TO_RAD = Math.PI / 180;

interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

function requireShapes(shapes: PowerPoint.Shape[], minCount = 1): void {
  if (shapes.length < minCount) {
    throw new Error(
      minCount > 1 ? `Select at least ${minCount} objects and try again.` : "Select one or more objects first."
    );
  }
}

export function isAdjustmentsSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
}

// ---------------------------------------------------------------------------
// hv_align / hv_distribute — grid-aware align and distribute
// ---------------------------------------------------------------------------

/** Detects rows/columns from shape positions (matches within `tolerance` points count as the same row/column), mirroring the VBA grid-detection logic. */
function detectGrid(rects: Rect[], tolerance = 15): number[][] {
  const indices = rects.map((_, i) => i);
  const rowOf = (i: number) => rects[i].top;
  const colOf = (i: number) => rects[i].left;

  const closeEnough = (a: number, b: number) => Math.abs(a - b) < tolerance;

  const rowGroups: number[][] = [];
  for (const i of indices) {
    const group = rowGroups.find((g) => closeEnough(rowOf(g[0]), rowOf(i)));
    if (group) group.push(i);
    else rowGroups.push([i]);
  }
  rowGroups.sort((a, b) => rowOf(a[0]) - rowOf(b[0]));
  for (const row of rowGroups) row.sort((a, b) => colOf(a) - colOf(b));
  return rowGroups;
}

/** Port of hv_align: aligns each row to a common height/top and each column to a common width/left. */
export async function hvAlign(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items, 2);

    shapes.items.forEach((s) => s.load("left, top, width, height"));
    await context.sync();

    const rects = shapes.items.map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height }));
    const grid = detectGrid(rects);
    const numColumns = grid[0]?.length ?? 0;
    if (numColumns === 0 || grid.some((row) => row.length !== numColumns)) {
      throw new Error("Please select a table with equal columns and rows and run again.");
    }

    // Align rows: every shape in a row takes the first shape's height/top.
    for (const row of grid) {
      const ref = rects[row[0]];
      for (const idx of row.slice(1)) {
        shapes.items[idx].height = ref.height;
        shapes.items[idx].top = ref.top;
      }
    }
    // Align columns: every shape in a column takes the top row's width/left.
    for (let c = 0; c < numColumns; c++) {
      const ref = rects[grid[0][c]];
      for (let r = 1; r < grid.length; r++) {
        shapes.items[grid[r][c]].width = ref.width;
        shapes.items[grid[r][c]].left = ref.left;
      }
    }

    await context.sync();
  });
}

/** Port of hv_distribute: packs rows/columns edge-to-edge starting from the top-left shape. */
export async function hvDistribute(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items, 2);

    shapes.items.forEach((s) => s.load("left, top, width, height"));
    await context.sync();

    const rects = shapes.items.map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height }));
    const grid = detectGrid(rects);
    const numColumns = grid[0]?.length ?? 0;
    if (numColumns === 0 || grid.some((row) => row.length !== numColumns)) {
      throw new Error("Please select a table with equal columns and rows and run again.");
    }

    for (const row of grid) {
      for (let c = 1; c < row.length; c++) {
        const prev = shapes.items[row[c - 1]];
        shapes.items[row[c]].left = prev.left + prev.width;
      }
    }
    for (let c = 0; c < numColumns; c++) {
      for (let r = 1; r < grid.length; r++) {
        const prev = shapes.items[grid[r - 1][c]];
        shapes.items[grid[r][c]].top = prev.top + prev.height;
      }
    }

    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// switch / get-set position / fix size
// ---------------------------------------------------------------------------

/** Port of switch(): swaps the position of exactly two selected shapes. */
export async function switchPositions(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length !== 2) throw new Error("Select two objects and run again.");

    const [a, b] = shapes.items;
    a.load("left, top");
    b.load("left, top");
    await context.sync();

    const aPos = { left: a.left, top: a.top };
    const bPos = { left: b.left, top: b.top };
    a.left = bPos.left;
    a.top = bPos.top;
    b.left = aPos.left;
    b.top = aPos.top;

    await context.sync();
  });
}

/** Port of get_pos(): remembers the selection's bounding position/size for later use by fix_size/fix_width/fix_height/set_pos. */
export async function getPosition(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    shapes.items.forEach((s) => s.load("left, top, width, height"));
    await context.sync();

    const lefts = shapes.items.map((s) => s.left);
    const tops = shapes.items.map((s) => s.top);
    const rights = shapes.items.map((s) => s.left + s.width);
    const bottoms = shapes.items.map((s) => s.top + s.height);
    const left = Math.min(...lefts);
    const top = Math.min(...tops);
    setSavedPosition({
      left,
      top,
      width: Math.max(...rights) - left,
      height: Math.max(...bottoms) - top,
    });
  });
}

/** Port of set_pos(): moves the selection to the position saved by getPosition(). */
export async function setPosition(): Promise<void> {
  const saved = getSavedPosition();
  if (!saved) throw new Error("Use 'Get Position' first, before 'Set Position'.");

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    for (const shape of shapes.items) {
      shape.left = saved.left;
      shape.top = saved.top;
    }
    await context.sync();
  });
}

async function fixDimension(dims: Array<"width" | "height">): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    if (shapes.items.length > 1) {
      shapes.items.forEach((s) => s.load("width, height"));
      await context.sync();
      const ref = shapes.items[0];
      for (const shape of shapes.items.slice(1)) {
        if (dims.includes("width")) shape.width = ref.width;
        if (dims.includes("height")) shape.height = ref.height;
      }
    } else {
      const saved = getSavedPosition();
      if (!saved) throw new Error("Select at least two objects, or use 'Get Position' first.");
      if (dims.includes("width")) shapes.items[0].width = saved.width;
      if (dims.includes("height")) shapes.items[0].height = saved.height;
    }
    await context.sync();
  });
}

/** Port of fix_size(): matches width and height across the selection (or from a saved position for a single shape). */
export const fixSize = () => fixDimension(["width", "height"]);
/** Port of fix_width(). */
export const fixWidth = () => fixDimension(["width"]);
/** Port of fix_height(). */
export const fixHeight = () => fixDimension(["height"]);

/** Port of GetPos_Bottom(): remembers the selection's bottom-anchored position. */
export async function getPositionBottom(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    shapes.items.forEach((s) => s.load("left, top, width, height"));
    await context.sync();

    const lefts = shapes.items.map((s) => s.left);
    const tops = shapes.items.map((s) => s.top);
    const rights = shapes.items.map((s) => s.left + s.width);
    const bottoms = shapes.items.map((s) => s.top + s.height);
    const left = Math.min(...lefts);
    const top = Math.min(...tops);
    setSavedBottomPosition({
      leftBottom: left,
      bottom: Math.max(...bottoms),
      width: Math.max(...rights) - left,
      height: Math.max(...bottoms) - top,
    });
  });
}

/** Port of SetPos_Bottom(): positions the selection so its bottom edge matches the saved bottom position. */
export async function setPositionBottom(): Promise<void> {
  const saved = getSavedBottomPosition();
  if (!saved) throw new Error("Use 'Get Position Bottom' first, before 'Set Position Bottom'.");

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    shapes.items.forEach((s) => s.load("height"));
    await context.sync();

    for (const shape of shapes.items) {
      shape.left = saved.leftBottom;
      shape.top = saved.bottom - shape.height;
    }
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Angle-adjustment matching (chevrons, pentagons, arrows)
// ---------------------------------------------------------------------------

/** Port of AlignAngles(): matches the angle-adjustment of every selected shape to the first one's, correcting for width differences. Caller is responsible for only selecting chevrons/pentagons/right-arrows — the API can't verify shape subtype (see module note). Requires PowerPointApi 1.10. */
export async function alignAngles(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items, 2);

    shapes.items.forEach((s) => s.load("width"));
    await context.sync();

    const first = shapes.items[0];
    const originalWidth = first.width;
    const originalAngle = first.adjustments.get(0);
    await context.sync();
    const angleDistance = originalWidth * (1 - originalAngle.value);

    for (const shape of shapes.items.slice(1)) {
      shape.adjustments.set(0, 1 - angleDistance / shape.width);
    }
    await context.sync();
  });
}

/** Port of chevronAlign(): lines up a sequence of chevrons/pentagons edge-to-edge based on each one's angle adjustment. Same subtype-verification caveat as alignAngles(). Requires PowerPointApi 1.10. */
export async function chevronAlign(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items, 2);

    shapes.items.forEach((s) => s.load("left, top, width, height, rotation"));
    await context.sync();

    for (let i = 1; i < shapes.items.length; i++) {
      const prev = shapes.items[i - 1];
      const prevAngle = prev.adjustments.get(0);
      await context.sync();
      const angleDistance = prev.height * prevAngle.value;

      if (prev.rotation === 0) {
        shapes.items[i].left = prev.left + prev.width - angleDistance;
      } else {
        shapes.items[i].top = prev.top + prev.width - angleDistance;
      }
      await context.sync();
    }
  });
}

// ---------------------------------------------------------------------------
// Centre-on-slide, with template margins
// ---------------------------------------------------------------------------

/** Slide dimensions require PowerPointApi 1.10 (Mac support only since Dec 2025) — every "centre on slide" function below inherits that requirement. */
export function isSlideDimensionsSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
}

async function getSlideSize(context: PowerPoint.RequestContext): Promise<{ width: number; height: number }> {
  context.presentation.pageSetup.load("slideWidth, slideHeight");
  await context.sync();
  return { width: context.presentation.pageSetup.slideWidth, height: context.presentation.pageSetup.slideHeight };
}

async function selectionBounds(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<Rect> {
  shapes.forEach((s) => s.load("left, top, width, height"));
  await context.sync();
  const left = Math.min(...shapes.map((s) => s.left));
  const top = Math.min(...shapes.map((s) => s.top));
  const right = Math.max(...shapes.map((s) => s.left + s.width));
  const bottom = Math.max(...shapes.map((s) => s.top + s.height));
  return { left, top, width: right - left, height: bottom - top };
}

/** Port of CentreOnSlideHorizontal(): centres the selection horizontally within the template margins. */
export async function centreOnSlideHorizontal(): Promise<void> {
  const margins = getTemplateMargins();
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    const bounds = await selectionBounds(context, shapes.items);
    const { width: pageWidth } = await getSlideSize(context);

    const target = (pageWidth - (margins.left + margins.right)) / 2 - bounds.width / 2 + margins.left;
    const delta = target - bounds.left;
    for (const shape of shapes.items) {
      shape.left = shape.left + delta;
    }
    await context.sync();
  });
}

/** Port of CentreOnSlideVertical(): centres the selection vertically within the template margins. */
export async function centreOnSlideVertical(): Promise<void> {
  const margins = getTemplateMargins();
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    const bounds = await selectionBounds(context, shapes.items);
    const { height: pageHeight } = await getSlideSize(context);

    const target = (pageHeight - (margins.top + margins.bottom)) / 2 - bounds.height / 2 + margins.top;
    const delta = target - bounds.top;
    for (const shape of shapes.items) {
      shape.top = shape.top + delta;
    }
    await context.sync();
  });
}

/** Port of CentreOnSlideHV(). */
export async function centreOnSlideHV(): Promise<void> {
  await centreOnSlideHorizontal();
  await centreOnSlideVertical();
}

/** Port of ResetTemplateCentrePointsNormal(): clears template margins back to zero. */
export function resetTemplateMarginsNormal(): void {
  setTemplateMargins({ left: 0, right: 0, top: 0, bottom: 0 });
}

/** Port of TwoUpTemplateHorizontalLeft(): positions the selection at the horizontal 1/4 mark. */
export async function twoUpTemplateHorizontalLeft(): Promise<void> {
  await positionAtQuarter(1);
}
/** Port of TwoUpTemplateHorizontalRight(): positions the selection at the horizontal 3/4 mark. */
export async function twoUpTemplateHorizontalRight(): Promise<void> {
  await positionAtQuarter(3);
}

async function positionAtQuarter(quarter: 1 | 3): Promise<void> {
  const margins = getTemplateMargins();
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items, 1);
    if (shapes.items.length > 1) throw new Error("Select a single object.");

    const shape = shapes.items[0];
    shape.load("width");
    const { width: pageWidth } = await getSlideSize(context); // also flushes the shape.load() above

    const usable = pageWidth - (margins.left + margins.right);
    shape.left = quarter * (usable / 4) - shape.width / 2 + margins.left;
    await context.sync();
  });
}

/** Port of CentreHalfSlide(): centres each selected shape within the left or right half of the slide. */
export async function centreHalfSlide(side: "Left" | "Right"): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    requireShapes(shapes.items);

    shapes.items.forEach((s) => s.load("width"));
    const { width: pageWidth } = await getSlideSize(context); // also flushes the per-shape loads above

    const halfWidth = pageWidth / 2;
    const leftCentre = pageWidth / 4;
    for (const shape of shapes.items) {
      shape.left = side === "Left" ? leftCentre - shape.width / 2 : leftCentre + halfWidth - shape.width / 2;
    }
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Edge join — move two shapes so their edges touch, accounting for rotation
// ---------------------------------------------------------------------------

function rotatedBoundingBox(rect: Rect, rotationDegrees: number) {
  const xLength = rect.width / 2;
  const yHeight = rect.height / 2;
  const xCentre = rect.left + xLength;
  const yCentre = rect.top + yHeight;
  const angle = rotationDegrees * DEG_TO_RAD;

  const dx = xLength * Math.cos(angle) + yHeight * Math.sin(angle);
  const dy = xLength * Math.sin(angle) + yHeight * Math.cos(angle);

  const leftA = xCentre - dx;
  const rightA = xCentre + dx;
  const topA = yCentre - dy;
  const bottomA = yCentre + dy;

  return {
    left: Math.min(leftA, rightA),
    right: Math.max(leftA, rightA),
    top: Math.min(topA, bottomA),
    bottom: Math.max(topA, bottomA),
    xLength,
    yHeight,
    angle,
  };
}

/** Port of EdgeJoin(): moves whichever of two shapes is further right/down so its edge touches the other, accounting for rotation. */
export async function edgeJoin(joinType: "Horizontal" | "Vertical"): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length !== 2) throw new Error("Select exactly two shapes to join.");

    const [one, two] = shapes.items;
    one.load("left, top, width, height, rotation");
    two.load("left, top, width, height, rotation");
    await context.sync();

    const boxOne = rotatedBoundingBox({ left: one.left, top: one.top, width: one.width, height: one.height }, one.rotation);
    const boxTwo = rotatedBoundingBox({ left: two.left, top: two.top, width: two.width, height: two.height }, two.rotation);

    if (joinType === "Horizontal") {
      if (boxOne.left <= boxTwo.left) {
        two.left = boxOne.right - boxTwo.xLength + (boxTwo.xLength * Math.cos(boxTwo.angle) + boxTwo.yHeight * Math.sin(boxTwo.angle));
      } else {
        one.left = boxTwo.right - boxOne.xLength + (boxOne.xLength * Math.cos(boxOne.angle) + boxOne.yHeight * Math.sin(boxOne.angle));
      }
    } else {
      if (boxOne.top <= boxTwo.top) {
        two.top = boxOne.bottom - boxTwo.yHeight + (boxTwo.xLength * Math.sin(boxTwo.angle) + boxTwo.yHeight * Math.cos(boxTwo.angle));
      } else {
        one.top = boxTwo.bottom - boxOne.yHeight + (boxOne.xLength * Math.sin(boxOne.angle) + boxOne.yHeight * Math.cos(boxOne.angle));
      }
    }
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Distribute (3+ shapes edge-to-edge, or table rows/columns)
// ---------------------------------------------------------------------------

/** Port of EnhancedDistributeHorizontal(): native distribute for 3+ shapes; splits a single table's columns evenly. */
export async function enhancedDistributeHorizontal(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length === 0) return;

    if (shapes.items.length > 2) {
      // No native "distribute" API — replicate by packing shapes edge-to-edge in left-to-right order.
      shapes.items.forEach((s) => s.load("left, width"));
      await context.sync();
      const ordered = [...shapes.items].sort((a, b) => a.left - b.left);
      for (let i = 1; i < ordered.length; i++) {
        ordered[i].left = ordered[i - 1].left + ordered[i - 1].width;
      }
      await context.sync();
      return;
    }

    const shape = shapes.items[0];
    shape.load("type, width");
    await context.sync();
    if (shape.type !== PowerPoint.ShapeType.table) return;

    const table = shape.getTable();
    table.load("columnCount");
    await context.sync();
    const targetWidth = shape.width / table.columnCount;
    for (let c = 0; c < table.columnCount; c++) {
      table.columns.getItemAt(c).width = targetWidth;
    }
    await context.sync();
  });
}

/** Port of EnhancedDistributeVertical(): native distribute for 3+ shapes; splits a single table's rows evenly. */
export async function enhancedDistributeVertical(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length === 0) return;

    if (shapes.items.length > 2) {
      shapes.items.forEach((s) => s.load("top, height"));
      await context.sync();
      const ordered = [...shapes.items].sort((a, b) => a.top - b.top);
      for (let i = 1; i < ordered.length; i++) {
        ordered[i].top = ordered[i - 1].top + ordered[i - 1].height;
      }
      await context.sync();
      return;
    }

    const shape = shapes.items[0];
    shape.load("type, height");
    await context.sync();
    if (shape.type !== PowerPoint.ShapeType.table) return;

    const table = shape.getTable();
    table.load("rowCount");
    await context.sync();
    const targetHeight = shape.height / table.rowCount;
    for (let r = 0; r < table.rowCount; r++) {
      table.rows.getItemAt(r).height = targetHeight;
    }
    await context.sync();
  });
}
