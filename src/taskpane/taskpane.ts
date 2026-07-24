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
 * Wires a swatch + caret + hidden native <input type="color"> as one
 * control. The swatch is the default click target and applies the
 * currently-held color immediately — no picker in the way, so reusing the
 * same color across several shapes is one click each time. The caret opens
 * the native picker to actually change the color; picking a new one there
 * applies it too (listened on `input` rather than `change`, since macOS's
 * native colour panel has no explicit commit action — `change` is
 * unreliable in WKWebView-hosted Mac Office task panes, while `input`
 * fires live as the user moves around the picker) and updates the swatch.
 * Debounced so dragging around the picker doesn't fire a PowerPoint.run
 * call per pixel — it applies once movement settles for 150ms.
 */
function bindColorControl(baseId: string, handler: (hex: string) => Promise<void>): void {
  const input = document.getElementById(`${baseId}Input`) as HTMLInputElement | null;
  const swatch = document.getElementById(`${baseId}Swatch`) as HTMLButtonElement | null;
  const caret = document.getElementById(`${baseId}Caret`) as HTMLButtonElement | null;
  if (!input || !swatch || !caret) {
    console.warn(`Color control #${baseId} not found in taskpane.html`);
    return;
  }

  const syncSwatch = () => {
    swatch.style.backgroundColor = input.value;
  };
  syncSwatch();

  let debounceTimer: number | undefined;
  input.addEventListener("input", () => {
    syncSwatch();
    window.clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(withErrorHandling(() => handler(input.value)), 150);
  });

  swatch.addEventListener("click", withErrorHandling(() => handler(input.value)));
  caret.addEventListener("click", () => {
    if (typeof input.showPicker === "function") {
      input.showPicker();
    } else {
      input.click();
    }
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
  // Temporary: this inline grid is being replaced by the gallery dialog
  // (Phase 5) — minimal fix to keep it compiling/working against the new
  // { groups, items } response shape in the meantime, not a real feature.
  const { items } = await Library.fetchCatalog(select?.value ?? "text");
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

/**
 * fetch() with a hard timeout. Without this, a hung request (this task
 * pane's first-ever same-origin fetch to its own backend, from inside the
 * sideloaded WKWebView) can leave a Promise neither resolved nor rejected
 * forever — a plain .catch() doesn't help with that, since there's nothing
 * to catch. Confirmed as the root cause of the whole task pane appearing to
 * freeze on "Loading...": the pre-fix startup code awaited a session check
 * before wiring any buttons, so a hang there blocked everything, including
 * Sign In itself.
 */
async function fetchWithTimeout(input: string, init: RequestInit = {}, timeoutMs = 8000): Promise<Response> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, { ...init, signal: controller.signal });
  } finally {
    window.clearTimeout(timer);
  }
}

/** Exchanges a fresh Microsoft ID token for our own session cookie. */
async function establishSession(idToken: string): Promise<void> {
  const res = await fetchWithTimeout("/api/auth/session", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ idToken }),
  });
  if (!res.ok) throw new Error(`Failed to establish session (${res.status}).`);
}

/** Checks whether the session cookie from a previous sign-in is still valid, without forcing an interactive prompt. Never throws — a failed/timed-out check just means "not signed in". */
async function getSessionUser(): Promise<SessionUser | null> {
  try {
    const res = await fetchWithTimeout("/api/auth/me");
    if (!res.ok) return null;
    return await res.json();
  } catch {
    return null;
  }
}

function updateSignInStatus(user: SessionUser | null): void {
  const el = document.getElementById("signInStatus");
  if (!el) return;
  if (user) {
    el.textContent = `Signed in as ${user.email ?? user.displayName ?? "unknown user"}${user.isAdmin ? " (admin)" : ""}.`;
    el.classList.add("signed-in");
  } else {
    el.textContent = "Not signed in.";
    el.classList.remove("signed-in");
  }
}

/** Applies a change in sign-in state everywhere it matters — the status line and the Content Library gate. */
function applySessionState(user: SessionUser | null): void {
  updateSignInStatus(user);
  refreshLibrarySection(user);
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  const statusEl = document.getElementById("status");
  if (statusEl) bindStatusElement(statusEl);

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
    applySessionState(await getSessionUser());
    notify(`Signed in as ${user.email}`);
  });
  setSectionEnabled(
    "sectionAuth",
    Auth.isNestedAppAuthSupported(),
    "Nested App Authentication isn't supported on this PowerPoint build."
  );

  // Content Library
  // Temporary (Phase 5 skeleton) — just opens the gallery dialog to check
  // whether it shares this task pane's session cookie. The dialog reports
  // the answer directly on-screen; nothing is wired to messageParent yet.
  document.getElementById("btnBrowseLibraryTest")?.addEventListener("click", () => {
    Office.context.ui.displayDialogAsync(
      `${window.location.origin}/gallery.html?v=${Date.now()}`,
      { height: 80, width: 70 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          notify(`Failed to open dialog: ${result.error.message}`, "error");
        }
      }
    );
  });

  document.getElementById("librarySelect")?.addEventListener("change", withErrorHandling(loadLibrary));
  bindButton("btnLibraryFinish", async () => {
    if (!currentFileInsertHandle) return;
    await Library.finishFileInsert(currentFileInsertHandle);
    currentFileInsertHandle = null;
    showLibraryFinishRow(false);
    notify("Done — temporary slide removed.");
  });
  applySessionState(null); // starting state — the background check below updates this once it resolves, however long that takes

  // Default swatch colour comes from config/theme.json, not hardcoded in the HTML.
  for (const id of ["fillColorInput", "lineColorInput", "textColorInput"]) {
    const input = document.getElementById(id) as HTMLInputElement | null;
    if (input) input.value = theme.defaultColorSwatch;
  }

  // Fill, line & text color
  bindColorControl("fillColor", async (hex) => {
    await FillLineColors.fillColor(FillLineColors.hexToRgb(hex));
    notify(`Fill set to ${hex}.`);
  });
  bindColorControl("lineColor", async (hex) => {
    await FillLineColors.lineColor(FillLineColors.hexToRgb(hex));
    notify(`Line set to ${hex}.`);
  });
  bindColorControl("textColor", async (hex) => {
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

  // Checked in the background, only after everything above is already
  // wired and interactive — see fetchWithTimeout's comment for why this
  // must never be awaited directly in the startup path above. If there's
  // no existing backend session cookie, also try a passive MSAL restore
  // (ssoSilent, no popup) — this is what lets a returning user skip the
  // Sign In button entirely, and it's also what primes Auth's cached MSAL
  // instance/login hint so a later click on Sign In can go straight to
  // acquireTokenPopup without awaiting anything first (see msal.ts's
  // signIn() comment for why that ordering matters for the popup blocker).
  getSessionUser().then(async (user) => {
    applySessionState(user);
    if (user) return;
    const silentUser = await Auth.trySilentSignIn();
    if (silentUser?.email) {
      await establishSession(silentUser.idToken);
      applySessionState(await getSessionUser());
    }
  });
});
