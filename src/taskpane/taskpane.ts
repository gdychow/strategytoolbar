import { bindStatusElement, notify, withErrorHandling } from "../core/ui";
import * as Layout from "../features/layout";
import * as FillLineColors from "../features/fillLineColors";
import * as OtherTweaks from "../features/otherTweaks";
import * as TableFormat from "../features/tableFormat";

function bindButton(id: string, handler: () => Promise<void>): void {
  const el = document.getElementById(id);
  if (!el) {
    console.warn(`Button #${id} not found in taskpane.html`);
    return;
  }
  el.addEventListener("click", withErrorHandling(handler));
}

/**
 * Applies a colour as the user picks it, with no separate Apply click.
 * Listens on `input` rather than `change`: macOS's native colour panel has
 * no explicit commit action, so `change` (which needs one) is unreliable in
 * WKWebView-hosted Mac Office task panes, while `input` fires live as the
 * user moves around the picker. Debounced so dragging around the picker
 * doesn't fire a PowerPoint.run call per pixel — it applies once movement
 * settles for 150ms.
 */
function bindColorAutoApply(inputId: string, handler: (hex: string) => Promise<void>): void {
  const input = document.getElementById(inputId) as HTMLInputElement | null;
  if (!input) {
    console.warn(`Color input #${inputId} not found in taskpane.html`);
    return;
  }
  let debounceTimer: number | undefined;
  input.addEventListener("input", () => {
    window.clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(withErrorHandling(() => handler(input.value)), 150);
  });
}

function setSectionEnabled(sectionId: string, enabled: boolean, reason?: string): void {
  const section = document.getElementById(sectionId);
  if (!section) return;
  section.querySelectorAll("button").forEach((btn) => {
    (btn as HTMLButtonElement).disabled = !enabled;
  });
  const note = section.querySelector(".unsupported-note");
  if (note) {
    (note as HTMLElement).style.display = enabled ? "none" : "block";
    if (reason) note.textContent = reason;
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  const statusEl = document.getElementById("status");
  if (statusEl) bindStatusElement(statusEl);

  // Fill, line & text color
  bindColorAutoApply("fillColorInput", async (hex) => {
    await FillLineColors.fillColor(FillLineColors.hexToRgb(hex));
    notify(`Fill set to ${hex}.`);
  });
  bindColorAutoApply("lineColorInput", async (hex) => {
    await FillLineColors.lineColor(FillLineColors.hexToRgb(hex));
    notify(`Line set to ${hex}.`);
  });
  bindColorAutoApply("textColorInput", async (hex) => {
    await FillLineColors.textColor(FillLineColors.hexToRgb(hex));
    notify(`Text color set to ${hex}.`);
  });
  bindButton("btnNoFill", async () => {
    await FillLineColors.noFill();
    notify("Fill removed.");
  });
  bindButton("btnNoLine", async () => {
    await FillLineColors.noLine();
    notify("Line removed.");
  });

  // Position & size
  bindButton("btnGetPosition", Layout.getPosition);
  bindButton("btnSetPosition", Layout.setPosition);
  bindButton("btnGetPositionBottom", Layout.getPositionBottom);
  bindButton("btnSetPositionBottom", Layout.setPositionBottom);
  bindButton("btnFixSize", Layout.fixSize);
  bindButton("btnFixWidth", Layout.fixWidth);
  bindButton("btnFixHeight", Layout.fixHeight);
  bindButton("btnSwitchPositions", Layout.switchPositions);

  // Align & distribute
  bindButton("btnHvAlign", Layout.hvAlign);
  bindButton("btnHvDistribute", Layout.hvDistribute);
  bindButton("btnDistributeH", Layout.enhancedDistributeHorizontal);
  bindButton("btnDistributeV", Layout.enhancedDistributeVertical);
  bindButton("btnEdgeJoinH", () => Layout.edgeJoin("Horizontal"));
  bindButton("btnEdgeJoinV", () => Layout.edgeJoin("Vertical"));

  // Centre on slide (requires PowerPointApi 1.10 — slide dimensions)
  bindButton("btnCentreH", Layout.centreOnSlideHorizontal);
  bindButton("btnCentreV", Layout.centreOnSlideVertical);
  bindButton("btnCentreHV", Layout.centreOnSlideHV);
  bindButton("btnTwoUpLeft", Layout.twoUpTemplateHorizontalLeft);
  bindButton("btnTwoUpRight", Layout.twoUpTemplateHorizontalRight);
  bindButton("btnHalfSlideLeft", () => Layout.centreHalfSlide("Left"));
  bindButton("btnHalfSlideRight", () => Layout.centreHalfSlide("Right"));
  bindButton("btnResetMargins", async () => {
    Layout.resetTemplateMarginsNormal();
    notify("Template margins reset to zero.");
  });
  setSectionEnabled(
    "sectionCentre",
    Layout.isSlideDimensionsSupported(),
    "Requires a newer PowerPoint build (PowerPointApi 1.10) than this one has."
  );

  // Angle tools (requires PowerPointApi 1.10 — shape.adjustments)
  bindButton("btnAlignAngles", Layout.alignAngles);
  bindButton("btnChevronAlign", Layout.chevronAlign);
  setSectionEnabled(
    "sectionAngles",
    Layout.isAdjustmentsSupported(),
    "Requires a newer PowerPoint build (PowerPointApi 1.10) than this one has."
  );

  // Other tweaks
  bindButton("btnToggleWordWrap", OtherTweaks.toggleWordWrap);
  bindButton("btnToggleAutoSize", OtherTweaks.toggleAutoSize);
  bindButton("btnSetMargins", async () => {
    const input = document.getElementById("marginInput") as HTMLInputElement;
    await OtherTweaks.setTextMargins(parseFloat(input.value));
    notify(`Margins set to ${input.value}pt.`);
  });
  bindButton("btnClearText", OtherTweaks.clearText);
  bindButton("btnPasteAsText", OtherTweaks.pasteAsText);
  bindButton("btnBulletsOn", () => OtherTweaks.setBulletsVisible(true));
  bindButton("btnBulletsOff", () => OtherTweaks.setBulletsVisible(false));

  // Table AutoFormat (requires PowerPointApi 1.9 — table cell borders/fill/font/margins)
  bindButton("btnTableHeaderRow", TableFormat.applyHeaderRowStyle);
  bindButton("btnTableHeaderRowCol", TableFormat.applyHeaderRowAndColumnStyle);
  setSectionEnabled(
    "sectionTables",
    TableFormat.isTableCellFormattingSupported(),
    "Requires a newer PowerPoint build (PowerPointApi 1.9) than this one has."
  );

  notify(`Ready. Host: ${info.host} / ${info.platform}`);
});
