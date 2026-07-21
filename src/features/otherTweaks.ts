/**
 * Port of L_Other_Tweaks.bas — this module hit far more API gaps than
 * Layout or Fill/Line Colors. Roughly 37 of the ~47 subs in the source
 * file are NOT portable today because the PowerPoint JS API simply has no
 * equivalent surface, confirmed against the real @types/office-js
 * definitions (not assumed):
 *
 * - Paragraph/line spacing (NoLineBetweenParagraphs, Quarter/Half/OneLine-
 *   BetweenParagraphs, setLineSpacing, setLineSpacing_TextRange,
 *   show_lineSpacing — 7 subs): PowerPoint.ParagraphFormat only exposes
 *   bulletFormat, horizontalAlignment, and indentLevel. No spaceBefore/
 *   spaceAfter/spaceWithin/lineRule at all.
 * - Custom bullet character/font/colour (SetBulletFormatToStyle's real
 *   behaviour, setBullets, formatBulletsTextRange/ShapeRange/FrameRange —
 *   5 functions): PowerPoint.BulletFormat only exposes `type` (None/
 *   Numbered/Unnumbered), a fixed `style` enum of numbering schemes, and
 *   `visible` — no way to set an arbitrary bullet character, font, or
 *   colour. TextFrame has no `Ruler`/indent-margin-levels API either.
 * - Arrowheads (Arrow_Right/Left/Double/None — 4 subs): ShapeLineFormat
 *   has color/dashStyle/style/transparency/visible/weight — no begin/end
 *   arrowhead properties exist.
 * - Spell-check language switching (ChangeLanguage, CollectShapes, and 15
 *   LanguageXxx wrappers — 17 subs): no language property anywhere on
 *   ShapeFont or TextRange.
 * - Grayscale preview (grayscale — 1 sub): no blackWhiteMode equivalent.
 * - Footer text (UpdateFooter — 1 sub): no HeadersFooters/Footer API.
 * - Slide reorder (MoveSlideBack — 1 sub): SlideCollection has add/delete/
 *   getItem/getItemAt/exportAsBase64Presentation — no move/reorder.
 * - CleanUpTemplates (1 sub): depends on the already-skipped
 *   ResetTemplateCentrePointsMHFi and on ActivePresentation.Close, which
 *   has no add-in equivalent (and doesn't fit a per-use toolbar anyway —
 *   it's a one-off template-publishing step, not a formatting action).
 * - show_font_finder (1 sub): opens frmFindFonts, a Tier 2 dialog not yet
 *   built.
 *
 * What's actually portable and implemented below: word wrap toggle,
 * autosize toggle, text margins (shapes and table cells), clear text
 * (recursing into groups), paste-as-text (via the browser clipboard API,
 * the closest equivalent to PasteSpecial(ppPasteText)), and bullets
 * on/off — NOT the Dash-vs-Bullet character distinction the VBA had,
 * since BulletFormat can't set a custom character. Renamed accordingly
 * (`setBulletsVisible`, not `dash`/`bullet`) rather than shipping two
 * functions that claim to differ but would look identical.
 */

const TEXT_CAPABLE_TYPES = new Set<PowerPoint.ShapeType | string>([
  PowerPoint.ShapeType.geometricShape,
  PowerPoint.ShapeType.callout,
  PowerPoint.ShapeType.diagram,
  PowerPoint.ShapeType.freeform,
  PowerPoint.ShapeType.line,
  PowerPoint.ShapeType.graphic,
  PowerPoint.ShapeType.image,
  PowerPoint.ShapeType.unsupported,
]);

async function withSelectedShapes(fn: (context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]) => Promise<void>): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length === 0) {
      throw new Error("Select one or more shapes first.");
    }
    await fn(context, shapes.items);
  });
}

/** Port of word_wraptoggle(). */
export async function toggleWordWrap(): Promise<void> {
  await withSelectedShapes(async (context, shapes) => {
    shapes.forEach((s) => s.textFrame.load("wordWrap"));
    await context.sync();
    for (const shape of shapes) {
      shape.textFrame.wordWrap = !shape.textFrame.wordWrap;
    }
    await context.sync();
  });
}

/** Port of autosizeshapetoggle(). */
export async function toggleAutoSize(): Promise<void> {
  await withSelectedShapes(async (context, shapes) => {
    shapes.forEach((s) => s.textFrame.load("autoSizeSetting"));
    await context.sync();
    for (const shape of shapes) {
      shape.textFrame.autoSizeSetting =
        shape.textFrame.autoSizeSetting === PowerPoint.ShapeAutoSize.autoSizeShapeToFitText
          ? PowerPoint.ShapeAutoSize.autoSizeNone
          : PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    }
    await context.sync();
  });
}

/** Port of MarginObject(): sets text-frame margins on shapes, or every cell's margins on a selected table. */
export async function setTextMargins(margin: number): Promise<void> {
  await withSelectedShapes(async (context, shapes) => {
    shapes.forEach((s) => s.load("type"));
    await context.sync();

    for (const shape of shapes) {
      if (shape.type === PowerPoint.ShapeType.table) {
        const table = shape.getTable();
        table.load("rowCount, columnCount");
        await context.sync();
        for (let r = 0; r < table.rowCount; r++) {
          for (let c = 0; c < table.columnCount; c++) {
            const cell = table.getCellOrNullObject(r, c);
            cell.margins.left = margin;
            cell.margins.right = margin;
            cell.margins.top = margin;
            cell.margins.bottom = margin;
          }
        }
      } else {
        shape.textFrame.leftMargin = margin;
        shape.textFrame.rightMargin = margin;
        shape.textFrame.topMargin = margin;
        shape.textFrame.bottomMargin = margin;
      }
    }
    await context.sync();
  });
}

async function clearShapeText(context: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<void> {
  shape.load("type");
  await context.sync();

  if (shape.type === PowerPoint.ShapeType.group) {
    const group = shape.group;
    group.shapes.load("items");
    await context.sync();
    for (const child of group.shapes.items) {
      await clearShapeText(context, child);
    }
    return;
  }

  if (TEXT_CAPABLE_TYPES.has(shape.type)) {
    shape.textFrame.deleteText();
  }
}

/** Port of clear()/clear_group()/clear_shape(): deletes all text in the selection, recursing into groups. */
export async function clearText(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length === 0) {
      throw new Error("Select one or more shapes first.");
    }
    for (const shape of shapes.items) {
      await clearShapeText(context, shape);
    }
    await context.sync();
  });
}

/**
 * Port of paste_text()/PasteUnformattedText(): pastes the clipboard's
 * plain text into the selected shape's text, replacing existing content —
 * the closest equivalent to VBA's PasteSpecial(ppPasteText). Uses the
 * browser Clipboard API rather than PowerPoint's own clipboard, since
 * Office.js has no PasteSpecial equivalent; requires clipboard-read
 * permission, which may be blocked by browser/OS/MDM policy.
 */
export async function pasteAsText(): Promise<void> {
  let text: string;
  try {
    text = await navigator.clipboard.readText();
  } catch (err) {
    throw new Error(
      "Couldn't read the clipboard (permission blocked by the browser or your Mac's settings). Copy the text again and retry."
    );
  }

  await withSelectedShapes(async (context, shapes) => {
    for (const shape of shapes) {
      shape.textFrame.textRange.text = text;
    }
    await context.sync();
  });
}

/** Replaces dash()/bullet(): turns paragraph bullets on/off. Can't replicate the VBA originals' custom dash-vs-round bullet character — BulletFormat has no way to set one (see module note). */
export async function setBulletsVisible(visible: boolean): Promise<void> {
  await PowerPoint.run(async (context) => {
    const textRange = context.presentation.getSelectedTextRangeOrNullObject();
    await context.sync();
    if (textRange.isNullObject) {
      throw new Error("Place the cursor in some text, or select a text box, first.");
    }
    textRange.paragraphFormat.bulletFormat.visible = visible;
    await context.sync();
  });
}
