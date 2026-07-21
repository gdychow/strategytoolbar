/**
 * Port of V_FormatObject.bas — the shared formatting engine that Layout,
 * Fill/Line Colors, and Other Tweaks call into. Faithful to the original
 * dispatch: shapes/text-boxes get shape or text formatting, groups recurse
 * into their children, tables get per-cell handling.
 *
 * Deliberate simplification vs. the VBA original: VBA distinguishes "whole
 * group selected" from "some children of a group selected" via
 * Selection.HasChildShapeRange. Office.js's selection model doesn't expose
 * that distinction the same way — getSelectedShapes() returns whatever the
 * user has actually selected (including an individual child shape directly,
 * if that's what's selected). So a selected Group here is always treated as
 * "format every shape in the group."
 *
 * Uses shape.textFrame (API 1.4) rather than getTextFrameOrNullObject()
 * (API 1.10, only on Mac since Dec 2025) for broader Mac compatibility.
 */

export type FormatAction = "Fill" | "Line" | "Text";

export interface RGB {
  r: number;
  g: number;
  b: number;
}

export interface ShapeFormatOptions {
  color?: RGB;
  transparent?: boolean;
  lineWidth?: number;
  lineDashStyle?: PowerPoint.ShapeLineDashStyle;
  lineStyle?: PowerPoint.ShapeLineStyle;
}

export interface TextFormatOptions {
  color?: RGB;
  fontSize?: number;
  fontName?: string;
}

export function toHexColor(rgb: RGB): string {
  const h = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");
  return `#${h(rgb.r)}${h(rgb.g)}${h(rgb.b)}`;
}

const TEXT_CAPABLE_TYPES = new Set<PowerPoint.ShapeType | string>([
  PowerPoint.ShapeType.geometricShape,
  PowerPoint.ShapeType.callout,
  PowerPoint.ShapeType.diagram,
  PowerPoint.ShapeType.freeform,
  PowerPoint.ShapeType.line,
  PowerPoint.ShapeType.graphic,
  PowerPoint.ShapeType.image,
  PowerPoint.ShapeType.unsupported, // covers TextBox/TextEffect-equivalent shapes on some hosts
]);

/** Port of ApplyShapeFormatting */
function applyShapeFormatting(shape: PowerPoint.Shape, action: "Fill" | "Line", opts: ShapeFormatOptions): void {
  if (action === "Fill") {
    if (opts.color) {
      shape.fill.setSolidColor(toHexColor(opts.color));
    }
    if (opts.transparent) {
      shape.fill.clear();
    }
  } else if (action === "Line") {
    if (opts.color) {
      shape.lineFormat.visible = true;
      shape.lineFormat.color = toHexColor(opts.color);
    }
    if (opts.transparent) {
      shape.lineFormat.visible = false;
    }
    if (opts.lineWidth && opts.lineWidth > 0) {
      shape.lineFormat.weight = opts.lineWidth;
    }
    if (opts.lineDashStyle) {
      shape.lineFormat.dashStyle = opts.lineDashStyle;
    }
    if (opts.lineStyle) {
      shape.lineFormat.style = opts.lineStyle;
    }
  }
}

/** Port of ApplyTextFormatting */
function applyTextFormatting(textRange: PowerPoint.TextRange, opts: TextFormatOptions): void {
  if (opts.color) textRange.font.color = toHexColor(opts.color);
  if (opts.fontSize && opts.fontSize > 0) textRange.font.size = opts.fontSize;
  if (opts.fontName) textRange.font.name = opts.fontName;
}

/** Port of the table-cell branch of ObjectFormat: applies to every cell (no per-cell selection concept in Office.js, so this formats the whole table). */
async function applyTableFormatting(
  context: PowerPoint.RequestContext,
  table: PowerPoint.Table,
  action: FormatAction,
  shapeOpts: ShapeFormatOptions,
  textOpts: TextFormatOptions
): Promise<void> {
  table.load("rowCount, columnCount");
  await context.sync();

  const rowCount = table.rowCount;
  const columnCount = table.columnCount;
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < columnCount; c++) {
      const cell = table.getCellOrNullObject(r, c);
      if (action === "Text") {
        if (textOpts.color) cell.font.color = toHexColor(textOpts.color);
        if (textOpts.fontSize && textOpts.fontSize > 0) cell.font.size = textOpts.fontSize;
      } else if (action === "Fill") {
        if (shapeOpts.color) cell.fill.setSolidColor(toHexColor(shapeOpts.color));
        if (shapeOpts.transparent) cell.fill.clear();
      } else if (action === "Line") {
        const edges = [cell.borders.top, cell.borders.bottom, cell.borders.left, cell.borders.right];
        for (const edge of edges) {
          if (shapeOpts.transparent) {
            edge.transparency = 1;
            continue;
          }
          if (shapeOpts.color) {
            edge.transparency = 0;
            edge.color = toHexColor(shapeOpts.color);
          }
          if (shapeOpts.lineWidth && shapeOpts.lineWidth > 0) edge.weight = shapeOpts.lineWidth;
          if (shapeOpts.lineDashStyle) edge.dashStyle = shapeOpts.lineDashStyle;
        }
      }
    }
  }
}

async function formatShape(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  action: FormatAction,
  shapeOpts: ShapeFormatOptions,
  textOpts: TextFormatOptions
): Promise<void> {
  shape.load("type");
  await context.sync();

  if (shape.type === PowerPoint.ShapeType.table) {
    const table = shape.getTable();
    await applyTableFormatting(context, table, action, shapeOpts, textOpts);
    return;
  }

  if (shape.type === PowerPoint.ShapeType.group) {
    const group = shape.group;
    group.shapes.load("items");
    await context.sync();
    for (const child of group.shapes.items) {
      await formatShape(context, child, action, shapeOpts, textOpts);
    }
    return;
  }

  if (action === "Text") {
    if (TEXT_CAPABLE_TYPES.has(shape.type)) {
      applyTextFormatting(shape.textFrame.textRange, textOpts);
    }
  } else {
    applyShapeFormatting(shape, action, shapeOpts);
  }
}

export interface ObjectFormatOptions {
  shape?: ShapeFormatOptions;
  text?: TextFormatOptions;
}

/**
 * Port of ObjectFormat(TargetAction, ...). Applies formatting to the
 * current selection (shapes or in-place text selection), recursing into
 * groups and tables as needed.
 */
export async function objectFormat(action: FormatAction, options: ObjectFormatOptions = {}): Promise<void> {
  const shapeOpts = options.shape ?? {};
  const textOpts = options.text ?? {};

  await PowerPoint.run(async (context) => {
    // In-place text selection (cursor inside a text box, no shape selected): format that text
    // directly. Only relevant — and only called — for the Text action, so a problem with this
    // API can't take down Fill/Line as a side effect.
    if (action === "Text") {
      const selectedText = context.presentation.getSelectedTextRangeOrNullObject();
      await context.sync();

      if (!selectedText.isNullObject) {
        applyTextFormatting(selectedText, textOpts);
        await context.sync();
        return;
      }
    }

    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length === 0) {
      throw new Error("Select one or more shapes first.");
    }

    for (const shape of shapes.items) {
      await formatShape(context, shape, action, shapeOpts, textOpts);
    }

    await context.sync();
  });
}
