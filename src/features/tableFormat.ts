/**
 * Port of D_Tables.bas's TableAutoFormat/TableAutoFormat_2Axis: applies a
 * branded header style to a table (dark green header row, white text,
 * grey body text, brand-grey borders — colors match the custClrLst already
 * baked into the templates, see the custom-colours audit).
 *
 * One real behavioural difference from the VBA original: VBA only
 * restyled the cells the user had individually selected (`Cell(i,j).Selected`),
 * so you could select just a header row and style only that. Office.js has
 * no per-cell selection API — getSelectedShapes() only sees the table as a
 * whole — so this applies to every cell in the table, using each cell's
 * row/column position to decide its style. Same limitation already noted
 * in objectFormat.ts's table branch.
 */

/** cell.borders/fill/font/margins all require PowerPointApi 1.9. */
export function isTableCellFormattingSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
}

const BORDER_GREY = "#505050";
const WHITE = "#FFFFFF";
const BODY_TEXT = "#383838";
const HEADER_GREEN = "#115C36"; // Wiley "Dark Green" — matches custClrLst
const AXIS_GREY = "#D8D9DA"; // Wiley "Light Gray 1" — matches custClrLst

function getSelectedTable(shapes: PowerPoint.Shape[]): PowerPoint.Shape {
  if (shapes.length !== 1) {
    throw new Error("Select a single table first.");
  }
  return shapes[0];
}

function styleBordersGrey(cell: PowerPoint.TableCell): void {
  for (const edge of [cell.borders.top, cell.borders.bottom, cell.borders.left, cell.borders.right]) {
    edge.color = BORDER_GREY;
    edge.weight = 0.75;
    edge.transparency = 0;
  }
}

function whiteBorderEdges(cell: PowerPoint.TableCell, edges: Array<"top" | "bottom" | "left" | "right">): void {
  for (const edge of edges) {
    cell.borders[edge].color = WHITE;
  }
}

function baseCellStyle(cell: PowerPoint.TableCell): void {
  styleBordersGrey(cell);
  cell.margins.left = 2.8;
  cell.margins.right = 2.8;
  cell.margins.top = 2.8;
  cell.margins.bottom = 2.8;
  cell.fill.setSolidColor(WHITE);
  cell.font.name = "Open Sans";
  cell.font.bold = false;
  cell.font.color = BODY_TEXT;
}

/** Port of TableAutoFormat(): header row only. */
export async function applyHeaderRowStyle(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    const shape = getSelectedTable(shapes.items);

    const table = shape.getTable();
    table.load("rowCount, columnCount");
    await context.sync();
    if (table.rowCount <= 1) {
      throw new Error("This table needs more than one row to apply a header style.");
    }

    for (let r = 0; r < table.rowCount; r++) {
      for (let c = 0; c < table.columnCount; c++) {
        const cell = table.getCellOrNullObject(r, c);
        baseCellStyle(cell);

        if (r === 0) {
          cell.fill.setSolidColor(HEADER_GREEN);
          cell.font.name = "Open Sans ExtraBold";
          cell.font.color = WHITE;
          whiteBorderEdges(cell, ["top", "bottom", "left", "right"]);
        } else if (r === 1) {
          whiteBorderEdges(cell, ["top"]);
        } else if (r === table.rowCount - 1) {
          whiteBorderEdges(cell, ["bottom"]);
        }

        if (c === 0) {
          whiteBorderEdges(cell, ["left"]);
        } else if (c === table.columnCount - 1) {
          whiteBorderEdges(cell, ["right"]);
        }
      }
    }
    await context.sync();
  });
}

/** Port of TableAutoFormat_2Axis(): header row + header column (matrix-style table). */
export async function applyHeaderRowAndColumnStyle(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    const shape = getSelectedTable(shapes.items);

    const table = shape.getTable();
    table.load("rowCount, columnCount");
    await context.sync();
    if (table.rowCount <= 1) {
      throw new Error("This table needs more than one row to apply a header style.");
    }

    for (let r = 0; r < table.rowCount; r++) {
      for (let c = 0; c < table.columnCount; c++) {
        const cell = table.getCellOrNullObject(r, c);
        baseCellStyle(cell);

        if (r === 0) {
          cell.fill.setSolidColor(HEADER_GREEN);
          cell.font.name = "Open Sans ExtraBold";
          cell.font.color = WHITE;
          whiteBorderEdges(cell, ["top", "bottom", "left", "right"]);
          if (c === 0) {
            cell.fill.setSolidColor(WHITE);
          }
        } else if (r === 1) {
          whiteBorderEdges(cell, ["top"]);
        } else if (r === table.rowCount - 1) {
          whiteBorderEdges(cell, ["bottom"]);
        }

        if (c === 0 && r !== 0) {
          cell.fill.setSolidColor(AXIS_GREY);
          cell.font.name = "Open Sans";
          cell.font.bold = true;
          cell.font.color = BODY_TEXT;
          whiteBorderEdges(cell, ["left"]);
        } else if (c === 1) {
          whiteBorderEdges(cell, ["left"]);
        } else if (c === table.columnCount - 1) {
          whiteBorderEdges(cell, ["right"]);
        }
      }
    }
    await context.sync();
  });
}
