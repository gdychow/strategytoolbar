import { bindStatusElement, notify, withErrorHandling } from "../core/ui";
import * as Layout from "../features/layout";
import * as FillLineColors from "../features/fillLineColors";
import * as OtherTweaks from "../features/otherTweaks";
import * as TableFormat from "../features/tableFormat";
import * as Library from "../features/libraryInsert";
import * as Auth from "../auth/msal";
import theme from "../config/theme.json";

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

interface SessionUser {
  oid: string;
  tid: string;
  email: string | null;
  displayName: string | null;
  isAdmin: boolean;
}

let currentFileInsertHandle: Library.FileInsertHandle | null = null;

function showLibraryFinishRow(show: boolean): void {
  const grid = document.getElementById("libraryGrid");
  const finishRow = document.getElementById("libraryFinishRow");
  const select = document.getElementById("librarySelect") as HTMLSelectElement | null;
  if (grid) (grid as HTMLElement).style.display = show ? "none" : "grid";
  if (finishRow) (finishRow as HTMLElement).style.display = show ? "block" : "none";
  if (select) select.disabled = show;
}

async function handleLibraryItemClick(item: Library.CatalogItem): Promise<void> {
  const handle = await Library.insertCatalogItem(item);
  if (handle) {
    currentFileInsertHandle = handle;
    showLibraryFinishRow(true);
    notify(`"${item.title}" added on a temporary slide — copy it across, then click Finish.`);
  } else {
    notify(`"${item.title}" inserted.`);
  }
}

function renderLibraryGrid(items: Library.CatalogItem[]): void {
  const grid = document.getElementById("libraryGrid");
  if (!grid) return;
  grid.innerHTML = "";

  for (const item of items) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "library-item";

    if (item.thumbnailUrl) {
      const img = document.createElement("img");
      img.src = item.thumbnailUrl;
      img.alt = item.title;
      button.appendChild(img);
    } else {
      const placeholder = document.createElement("span");
      placeholder.className = "library-item-placeholder";
      button.appendChild(placeholder);
    }

    const label = document.createElement("span");
    label.textContent = item.title;
    button.appendChild(label);

    button.addEventListener("click", withErrorHandling(() => handleLibraryItemClick(item)));
    grid.appendChild(button);
  }
}

async function loadLibrary(): Promise<void> {
  const select = document.getElementById("librarySelect") as HTMLSelectElement | null;
  const items = await Library.fetchCatalog(select?.value ?? "text");
  renderLibraryGrid(items);
}

/** Gates and (re)populates the Content Library section whenever sign-in state changes. */
function refreshLibrarySection(user: SessionUser | null): void {
  const signedIn = !!user;
  const supported = Library.isLibraryInsertSupported();
  setSectionEnabled(
    "sectionLibrary",
    signedIn && supported,
    signedIn ? "Requires a newer PowerPoint build (PowerPointApi 1.2) than this one has." : "Sign in above to browse the content library."
  );

  if (signedIn && supported) {
    loadLibrary().catch((err) => notify(`Error: ${err instanceof Error ? err.message : String(err)}`, "error"));
  } else {
    const grid = document.getElementById("libraryGrid");
    if (grid) grid.innerHTML = "";
    showLibraryFinishRow(false);
    currentFileInsertHandle = null;
  }
}

/** Exchanges a fresh Microsoft ID token for our own session cookie. */
async function establishSession(idToken: string): Promise<void> {
  const res = await fetch("/api/auth/session", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ idToken }),
  });
  if (!res.ok) throw new Error(`Failed to establish session (${res.status}).`);
}

/** Checks whether the session cookie from a previous sign-in is still valid, without forcing an interactive prompt. */
async function getSessionUser(): Promise<SessionUser | null> {
  const res = await fetch("/api/auth/me");
  if (!res.ok) return null;
  return res.json();
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  const statusEl = document.getElementById("status");
  if (statusEl) bindStatusElement(statusEl);

  let sessionUser = await getSessionUser().catch(() => null);
  refreshLibrarySection(sessionUser);

  bindButton("btnSignIn", async () => {
    const user = await Auth.signIn((step) => notify(step));
    const claimsEl = document.getElementById("authClaims") as HTMLElement;
    claimsEl.style.display = "block";
    claimsEl.textContent = JSON.stringify(user, null, 2);
    if (!user.email) {
      notify("Signed in, but no email claim was returned — check the Azure app registration's optional claims.", "error");
      return;
    }
    await establishSession(user.idToken);
    sessionUser = await getSessionUser();
    refreshLibrarySection(sessionUser);
    notify(`Signed in as ${user.email}`);
  });
  setSectionEnabled(
    "sectionAuth",
    Auth.isNestedAppAuthSupported(),
    "Nested App Authentication isn't supported on this PowerPoint build."
  );

  // Content Library
  document.getElementById("librarySelect")?.addEventListener("change", withErrorHandling(loadLibrary));
  bindButton("btnLibraryFinish", async () => {
    if (!currentFileInsertHandle) return;
    await Library.finishFileInsert(currentFileInsertHandle);
    currentFileInsertHandle = null;
    showLibraryFinishRow(false);
    notify("Done — temporary slide removed.");
  });

  // Default swatch colour comes from config/theme.json, not hardcoded in the HTML.
  for (const id of ["fillColorInput", "lineColorInput", "textColorInput"]) {
    const input = document.getElementById(id) as HTMLInputElement | null;
    if (input) input.value = theme.defaultColorSwatch;
  }

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

  notify(`Ready. Host: ${info.host} / ${info.platform} / NAA supported: ${Auth.isNestedAppAuthSupported()}`);
});
