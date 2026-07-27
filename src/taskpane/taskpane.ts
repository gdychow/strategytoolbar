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

// Grayscale + PowerPoint's own "Standard Colors" row values (same hex
// values PowerPoint's native fill/line/text color dropdown uses), so
// picks here look/feel like PowerPoint's own picker. This is deliberately
// NOT the only source of colors in the panel — see getThemeColors() below
// for the deck's actual theme colors, and the hex input for anything else.
const STANDARD_COLOR_PALETTE = [
  "#000000",
  "#FFFFFF",
  "#404040",
  "#808080",
  "#BFBFBF",
  "#D9D9D9",
  "#C00000",
  "#FF0000",
  "#FFC000",
  "#FFFF00",
  "#92D050",
  "#00B050",
  "#00B0F0",
  "#0070C0",
  "#002060",
  "#7030A0",
];

let activeColorHandler: ((hex: string) => Promise<void>) | null = null;
let activeColorInput: HTMLInputElement | null = null;
let activeColorSwatch: HTMLButtonElement | null = null;

function closeColorPickerPanel(): void {
  const panel = document.getElementById("colorPickerPanel");
  if (panel) panel.style.display = "none";
  activeColorHandler = null;
  activeColorInput = null;
  activeColorSwatch = null;
}

function applyPickedColor(hex: string): void {
  const input = activeColorInput;
  const swatch = activeColorSwatch;
  const handler = activeColorHandler;
  closeColorPickerPanel();
  if (input) input.value = hex;
  if (swatch) swatch.style.backgroundColor = hex;
  if (handler) withErrorHandling(() => handler(hex))();
}

function renderColorSwatches(container: Element, hexes: string[]): void {
  for (const hex of hexes) {
    const swatchBtn = document.createElement("button");
    swatchBtn.type = "button";
    swatchBtn.className = "color-picker-swatch";
    swatchBtn.style.backgroundColor = hex;
    swatchBtn.title = hex;
    swatchBtn.addEventListener("click", () => applyPickedColor(hex));
    container.appendChild(swatchBtn);
  }
}

/**
 * Theme colors get a labeled-row layout instead of the bare grid the
 * standard palette uses — a role name (e.g. "Accent 3") next to its own
 * swatch, visible at a glance rather than hidden behind a hover tooltip.
 * Deliberately more verbose than the standard grid: which named role maps
 * to which actual color is exactly the thing worth being able to check
 * directly against what PowerPoint's own picker shows for the same file.
 */
function renderThemeColorRows(container: Element, colors: { label: string; hex: string }[]): void {
  for (const { label, hex } of colors) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "color-picker-theme-row";
    row.title = hex;
    row.addEventListener("click", () => applyPickedColor(hex));

    const swatch = document.createElement("span");
    swatch.className = "color-picker-swatch";
    swatch.style.backgroundColor = hex;

    const text = document.createElement("span");
    text.className = "color-picker-theme-label";
    text.textContent = `${label} (${hex})`;

    row.appendChild(swatch);
    row.appendChild(text);
    container.appendChild(row);
  }
}

/**
 * Builds the shared color picker panel's contents once and wires its
 * dismiss behavior (click outside, Escape). See the HTML comment above
 * #colorPickerPanel in taskpane.html for why this exists instead of the
 * browser's native <input type="color"> picker: the native picker's
 * on-screen position isn't reliably controllable inside this WKWebView-
 * hosted task pane — two separate attempts at positioning it (a static
 * CSS pin, then a getBoundingClientRect()-computed one) both still opened
 * it at the screen's bottom-left. A plain in-page dropdown behaves like
 * any other positioned element instead, so it doesn't have that problem.
 *
 * Colors come from three places: the deck's own actual theme colors
 * (fetched once here, async — see FillLineColors.getThemeColors), a
 * fixed standard palette matching PowerPoint's own picker, and a hex
 * input for anything else. "More colors…" still opens the native
 * <input type="color"> as a last resort for full-spectrum picking — its
 * position isn't reliable (the same bug this panel exists to work around
 * everywhere else), but as a secondary, deliberately-chosen action that's
 * an acceptable trade for getting the full color spectrum back.
 */
function initColorPickerPanel(): void {
  const panel = document.getElementById("colorPickerPanel");
  const themeSection = document.getElementById("colorPickerThemeSection");
  const themeSwatches = document.getElementById("colorPickerThemeSwatches");
  const standardSwatches = document.getElementById("colorPickerStandardSwatches");
  const hexInput = document.getElementById("colorPickerHexInput") as HTMLInputElement | null;
  const applyBtn = document.getElementById("colorPickerApply");
  const moreBtn = document.getElementById("colorPickerMore");
  if (!panel || !themeSection || !themeSwatches || !standardSwatches || !hexInput || !applyBtn || !moreBtn) {
    console.warn("Color picker panel not found in taskpane.html");
    return;
  }

  renderColorSwatches(standardSwatches, STANDARD_COLOR_PALETTE);

  FillLineColors.getThemeColors()
    .then((colors) => {
      if (colors.length === 0) return; // unsupported PowerPoint build — leave the section hidden, not an error
      renderThemeColorRows(themeSwatches, colors);
      themeSection.style.display = "block";
    })
    .catch((err) => console.warn("Couldn't load theme colors:", err));

  applyBtn.addEventListener("click", () => {
    const hex = hexInput.value.trim();
    if (/^#[0-9a-fA-F]{6}$/.test(hex)) {
      applyPickedColor(hex);
    } else {
      notify("Enter a color as #RRGGBB.", "error");
    }
  });

  moreBtn.addEventListener("click", () => {
    if (!activeColorInput) return;
    if (typeof activeColorInput.showPicker === "function") {
      activeColorInput.showPicker();
    } else {
      activeColorInput.click();
    }
  });

  document.addEventListener("click", (e) => {
    if (panel.style.display === "none") return;
    const target = e.target as HTMLElement;
    if (panel.contains(target) || target.closest(".color-caret")) return;
    closeColorPickerPanel();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeColorPickerPanel();
  });
}

/**
 * Positions the panel at (left, top) but clamped so it never renders past
 * the viewport's right/bottom edge — the task pane's width varies (the
 * user can resize the PowerPoint window), and a fixed left/top computed
 * only from the caret's position can push the panel off-screen entirely
 * on a narrower pane. Flips above the caret instead of below it if there
 * isn't enough room underneath, same idea as a native dropdown.
 */
function positionPanelWithinViewport(panel: HTMLElement, caretRect: DOMRect): void {
  panel.style.left = "0px";
  panel.style.top = "0px";
  panel.style.display = "block";
  const panelWidth = panel.offsetWidth;
  const panelHeight = panel.offsetHeight;

  let left = caretRect.left;
  if (left + panelWidth > window.innerWidth) left = window.innerWidth - panelWidth - 8;
  left = Math.max(8, left);

  let top = caretRect.bottom + 4;
  if (top + panelHeight > window.innerHeight) top = caretRect.top - panelHeight - 4;
  top = Math.max(8, top);

  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
}

function openColorPickerPanel(
  caret: HTMLElement,
  input: HTMLInputElement,
  swatch: HTMLButtonElement,
  handler: (hex: string) => Promise<void>
): void {
  const panel = document.getElementById("colorPickerPanel");
  if (!panel) return;
  activeColorHandler = handler;
  activeColorInput = input;
  activeColorSwatch = swatch;
  const hexInput = document.getElementById("colorPickerHexInput") as HTMLInputElement | null;
  if (hexInput) hexInput.value = "";
  positionPanelWithinViewport(panel, caret.getBoundingClientRect());
}

/**
 * Wires a swatch + caret + hidden <input type="color"> as one control. The
 * swatch is the default click target and applies the currently-held color
 * immediately — no picker in the way, so reusing the same color across
 * several shapes is one click each time. The caret opens the shared custom
 * color picker panel (see initColorPickerPanel) to actually change the
 * color; that panel's "More colors…" button re-purposes this same hidden
 * input to reach the native OS picker for full-spectrum picking, so this
 * still needs to react when the input's value changes there (listened on
 * `input` rather than `change` — macOS's native colour panel has no
 * explicit commit action, and `change` is unreliable in WKWebView-hosted
 * Mac Office task panes, while `input` fires live as the user moves around
 * the picker). Debounced so dragging around that picker doesn't fire a
 * PowerPoint.run call per pixel — it applies once movement settles for
 * 150ms.
 */
function bindColorControl(baseId: string, handler: (hex: string) => Promise<void>): void {
  const input = document.getElementById(`${baseId}Input`) as HTMLInputElement | null;
  const swatch = document.getElementById(`${baseId}Swatch`) as HTMLButtonElement | null;
  const caret = document.getElementById(`${baseId}Caret`) as HTMLButtonElement | null;
  if (!input || !swatch || !caret) {
    console.warn(`Color control #${baseId} not found in taskpane.html`);
    return;
  }

  swatch.style.backgroundColor = input.value;

  let debounceTimer: number | undefined;
  input.addEventListener("input", () => {
    swatch.style.backgroundColor = input.value;
    window.clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(withErrorHandling(() => handler(input.value)), 150);
  });

  swatch.addEventListener("click", withErrorHandling(() => handler(input.value)));
  caret.addEventListener("click", () => openColorPickerPanel(caret, input, swatch, handler));
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
  const finishRow = document.getElementById("libraryFinishRow");
  const browseBtn = document.getElementById("btnBrowseLibrary") as HTMLButtonElement | null;
  if (finishRow) (finishRow as HTMLElement).style.display = show ? "block" : "none";
  if (browseBtn) browseBtn.disabled = show;
}

/** Called once the gallery dialog reports back which item the user picked (see the displayDialogAsync wiring below) — same insert engine as before, just triggered from the dialog instead of an inline grid click. */
async function insertPickedItem(item: Library.CatalogItem): Promise<void> {
  const handle = await Library.insertCatalogItem(item);
  if (handle) {
    currentFileInsertHandle = handle;
    showLibraryFinishRow(true);
    notify(`"${item.title}" added on a temporary slide — copy it across, then click Finish.`);
  } else {
    notify(`"${item.title}" inserted.`);
  }
}

/** Gates the Content Library section whenever sign-in state changes. The gallery dialog loads its own data lazily on open, so there's nothing to pre-fetch here. */
function refreshLibrarySection(user: SessionUser | null): void {
  const signedIn = !!user;
  const supported = Library.isLibraryInsertSupported();
  setSectionEnabled(
    "sectionLibrary",
    signedIn && supported,
    signedIn ? "Requires a newer PowerPoint build (PowerPointApi 1.2) than this one has." : "Sign in above to browse the content library."
  );
  if (!signedIn || !supported) {
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

/** Hides the Sign In button once signed in, and shows the admin-only Open Admin link — the only two other places sign-in/role state should be reflected in the UI. */
function updateAuthButtons(user: SessionUser | null): void {
  const signIn = document.getElementById("btnSignIn") as HTMLButtonElement | null;
  const openAdmin = document.getElementById("btnOpenAdmin") as HTMLButtonElement | null;
  if (signIn) signIn.style.display = user ? "none" : "";
  if (openAdmin) openAdmin.style.display = user?.isAdmin ? "" : "none";
}

/** Applies a change in sign-in state everywhere it matters — the status line, the auth buttons, and the Content Library gate. */
function applySessionState(user: SessionUser | null): void {
  updateSignInStatus(user);
  updateAuthButtons(user);
  refreshLibrarySection(user);
}

const SECTION_ORDER_STORAGE_KEY = "sectionOrder";

function getSections(): HTMLElement[] {
  return Array.from(document.querySelectorAll<HTMLElement>("body > section[id]"));
}

/**
 * Re-orders the existing <section> elements to match a previously-saved
 * order, if any — moves DOM nodes in place rather than re-rendering, so
 * none of the rest of this file's element lookups/wiring need to change.
 * Any section ID from a stale saved order no longer present is skipped;
 * any *current* section missing from a stale saved order (e.g. a future
 * new section) is appended at the end, keeping its relative position.
 */
function applyStoredSectionOrder(): void {
  const raw = localStorage.getItem(SECTION_ORDER_STORAGE_KEY);
  if (!raw) return;

  let storedOrder: string[];
  try {
    storedOrder = JSON.parse(raw);
  } catch {
    return;
  }

  const sections = getSections();
  const referenceNode = sections[sections.length - 1]?.nextSibling ?? null;
  const byId = new Map(sections.map((s) => [s.id, s]));

  const ordered = storedOrder.map((id) => byId.get(id)).filter((s): s is HTMLElement => !!s);
  const orderedSet = new Set(ordered);
  for (const s of sections) {
    if (!orderedSet.has(s)) ordered.push(s);
  }

  // insertBefore moves an already-attached node rather than duplicating it,
  // so repeating this in the desired final order builds that order in place.
  for (const s of ordered) {
    document.body.insertBefore(s, referenceNode);
  }
}

/**
 * User-customizable section order (per-user, per-machine — hence
 * localStorage, not shared catalog data). Native HTML5 drag-and-drop, no
 * library: only the small handle prepended to each <h2> is draggable
 * (see .drag-handle in taskpane.css), not the whole section, so dragging
 * never fights with clicking a button inside it.
 */
function initSectionReordering(): void {
  applyStoredSectionOrder();

  let draggedSection: HTMLElement | null = null;

  document.querySelectorAll<HTMLElement>(".drag-handle").forEach((handle) => {
    handle.addEventListener("dragstart", (e) => {
      const section = handle.closest("section");
      if (!section) return;
      draggedSection = section as HTMLElement;
      e.dataTransfer?.setData("text/plain", section.id);
      section.classList.add("dragging");
    });
    handle.addEventListener("dragend", () => {
      draggedSection?.classList.remove("dragging");
      draggedSection = null;
    });
  });

  getSections().forEach((section) => {
    section.addEventListener("dragover", (e) => {
      if (!draggedSection || draggedSection === section) return;
      e.preventDefault();
      const rect = section.getBoundingClientRect();
      const before = e.clientY < rect.top + rect.height / 2;
      section.parentElement?.insertBefore(draggedSection, before ? section : section.nextSibling);
    });
    section.addEventListener("drop", (e) => {
      e.preventDefault();
      localStorage.setItem(SECTION_ORDER_STORAGE_KEY, JSON.stringify(getSections().map((s) => s.id)));
    });
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  initSectionReordering();
  initColorPickerPanel();

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
  document.getElementById("btnOpenAdmin")?.addEventListener("click", () => {
    window.open("/admin", "_blank");
  });
  setSectionEnabled(
    "sectionAuth",
    Auth.isNestedAppAuthSupported(),
    "Nested App Authentication isn't supported on this PowerPoint build."
  );

  // Content Library — the dialog is a pure browse/search/select surface
  // (confirmed: a page opened via displayDialogAsync can only call
  // messageParent and requirements.isSetSupported, nothing PowerPoint-
  // specific), so all it ever sends back is which item was picked; this
  // task pane does the actual insert via the same Library.insertCatalogItem
  // used everywhere else in this app.
  document.getElementById("btnBrowseLibrary")?.addEventListener("click", () => {
    Office.context.ui.displayDialogAsync(
      `${window.location.origin}/gallery.html?v=${Date.now()}`,
      { height: 80, width: 70 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          notify(`Failed to open the library: ${result.error.message}`, "error");
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          if (!("message" in args)) return; // the other possible event is DialogEventReceived (closed/unloaded) — nothing to do here
          dialog.close();
          let item: Library.CatalogItem;
          try {
            item = JSON.parse(args.message);
          } catch (err) {
            notify(`Couldn't read the selected item: ${err instanceof Error ? err.message : String(err)}`, "error");
            return;
          }
          insertPickedItem(item).catch((err) =>
            notify(`Error: ${err instanceof Error ? err.message : String(err)}`, "error")
          );
        });
      }
    );
  });

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
