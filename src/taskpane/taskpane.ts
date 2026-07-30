import { bindStatusElement, notify, withErrorHandling, extractErrorMessage } from "../core/ui";
import { darken, THEME_SHADE_PERCENTS } from "../core/colorMath";
import * as Layout from "../features/layout";
import * as ObjectOrder from "../features/objectOrder";
import * as FillLineColors from "../features/fillLineColors";
import * as CustomColors from "../features/customColors";
import * as OtherTweaks from "../features/otherTweaks";
import * as TableFormat from "../features/tableFormat";
import * as Library from "../features/libraryInsert";
import * as Templates from "../features/templates";
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

type ColorControlKind = "fill" | "line" | "text";

const NO_COLOR_LABELS: Record<ColorControlKind, string | null> = {
  fill: "No Fill",
  line: "No Line",
  text: null, // native PowerPoint's Font Color dropdown has no "No Text Color" option
};

let activeColorHandler: ((hex: string) => Promise<void>) | null = null;
let activeColorInput: HTMLInputElement | null = null;
let activeColorSwatch: HTMLButtonElement | null = null;
let activeNoColorHandler: (() => Promise<void>) | null = null;

function closeColorPickerPanel(): void {
  const panel = document.getElementById("colorPickerPanel");
  if (panel) panel.style.display = "none";
  activeColorHandler = null;
  activeColorInput = null;
  activeColorSwatch = null;
  activeNoColorHandler = null;
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
  for (const hex of hexes) addSwatch(container, hex, hex);
}

function addSwatch(container: Element, hex: string, title: string): void {
  const swatchBtn = document.createElement("button");
  swatchBtn.type = "button";
  swatchBtn.className = "color-picker-swatch";
  swatchBtn.style.backgroundColor = hex;
  swatchBtn.title = title;
  swatchBtn.addEventListener("click", () => applyPickedColor(hex));
  container.appendChild(swatchBtn);
}

/**
 * Matches PowerPoint's own Theme Colors grid: the 10 base roles across the
 * top row, then 5 more rows of "Darker N%" shades beneath each — 60
 * swatches total. The role name (plus shade percent, for the shaded rows)
 * is on the hover tooltip rather than shown inline, same as the standard
 * palette below it.
 */
function renderThemeColorGrid(container: Element, colors: { label: string; hex: string }[]): void {
  for (const { label, hex } of colors) addSwatch(container, hex, label);
  for (const percent of THEME_SHADE_PERCENTS) {
    for (const { label, hex } of colors) addSwatch(container, darken(hex, percent), `${label}, Darker ${percent}%`);
  }
}

/** Custom Colors swatches are titled with their real name (e.g. "Wiley Blue"), not their hex — that's the whole point of a named palette. */
function renderNamedColorSwatches(container: Element, colors: { name: string; hex: string }[]): void {
  for (const { name, hex } of colors) addSwatch(container, hex, name);
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
 * Colors come from four places, top to bottom matching PowerPoint's own
 * dropdown: No Fill/No Line (kind-conditional, see openColorPickerPanel),
 * the deck's own actual theme colors plus their "Darker N%" shades
 * (fetched once here, async — see FillLineColors.getThemeColors and
 * renderThemeColorGrid), a fixed standard palette matching PowerPoint's
 * own picker, the deck's named Custom Colors if the template has any (see
 * CustomColors.getCustomColors), and a hex input for anything else.
 * "More [Fill/Line/Text] Colors…" still opens the native
 * <input type="color"> as a last resort for full-spectrum picking — its
 * position isn't reliable (the same bug this panel exists to work around
 * everywhere else), but as a secondary, deliberately-chosen action that's
 * an acceptable trade for getting the full color spectrum back. The
 * eyedropper button is shown only on browsers that actually support the
 * EyeDropper API (Chromium; not WebKit/Mac Office) — feature-detected
 * once here rather than assumed.
 */
function initColorPickerPanel(): void {
  const panel = document.getElementById("colorPickerPanel");
  const noColorBtn = document.getElementById("colorPickerNoColor") as HTMLButtonElement | null;
  const themeSection = document.getElementById("colorPickerThemeSection");
  const themeSwatches = document.getElementById("colorPickerThemeSwatches");
  const standardSwatches = document.getElementById("colorPickerStandardSwatches");
  const customSection = document.getElementById("colorPickerCustomSection");
  const customSwatches = document.getElementById("colorPickerCustomSwatches");
  const moreBtn = document.getElementById("colorPickerMore") as HTMLButtonElement | null;
  const eyedropperBtn = document.getElementById("colorPickerEyedropper") as HTMLButtonElement | null;
  if (
    !panel ||
    !noColorBtn ||
    !themeSection ||
    !themeSwatches ||
    !standardSwatches ||
    !customSection ||
    !customSwatches ||
    !moreBtn ||
    !eyedropperBtn
  ) {
    console.warn("Color picker panel not found in taskpane.html");
    return;
  }

  renderColorSwatches(standardSwatches, STANDARD_COLOR_PALETTE);

  FillLineColors.getThemeColors()
    .then((colors) => {
      if (colors.length === 0) return; // unsupported PowerPoint build — leave the section hidden, not an error
      renderThemeColorGrid(themeSwatches, colors);
      themeSection.style.display = "block";
    })
    .catch((err) => console.warn("Couldn't load theme colors:", err));

  CustomColors.getCustomColors()
    .then((colors) => {
      if (colors.length === 0) return; // unsupported build, or the template just has no custom colours defined
      renderNamedColorSwatches(customSwatches, colors);
      customSection.style.display = "block";
    })
    .catch((err) => console.warn("Couldn't load custom colors:", err));

  noColorBtn.addEventListener("click", () => {
    const handler = activeNoColorHandler;
    closeColorPickerPanel();
    if (handler) withErrorHandling(handler)();
  });

  moreBtn.addEventListener("click", () => {
    if (!activeColorInput) return;
    if (typeof activeColorInput.showPicker === "function") {
      activeColorInput.showPicker();
    } else {
      activeColorInput.click();
    }
  });

  if ("EyeDropper" in window) {
    eyedropperBtn.style.display = "block";
    eyedropperBtn.addEventListener("click", async () => {
      try {
        const result = await new window.EyeDropper!().open();
        applyPickedColor(result.sRGBHex);
      } catch {
        // user cancelled (Escape / clicked elsewhere) — nothing to do
      }
    });
  }

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

const MORE_COLORS_LABELS: Record<ColorControlKind, string> = {
  fill: "More Fill Colors…",
  line: "More Line Colors…",
  text: "More Text Colors…",
};

function openColorPickerPanel(
  caret: HTMLElement,
  input: HTMLInputElement,
  swatch: HTMLButtonElement,
  kind: ColorControlKind,
  handler: (hex: string) => Promise<void>,
  noColorHandler?: () => Promise<void>
): void {
  const panel = document.getElementById("colorPickerPanel");
  if (!panel) return;
  activeColorHandler = handler;
  activeColorInput = input;
  activeColorSwatch = swatch;
  activeNoColorHandler = noColorHandler ?? null;

  const noColorBtn = document.getElementById("colorPickerNoColor") as HTMLButtonElement | null;
  if (noColorBtn) {
    const label = NO_COLOR_LABELS[kind];
    noColorBtn.style.display = label && noColorHandler ? "block" : "none";
    noColorBtn.textContent = label ?? "";
  }
  const moreBtn = document.getElementById("colorPickerMore");
  if (moreBtn) moreBtn.textContent = MORE_COLORS_LABELS[kind];

  positionPanelWithinViewport(panel, caret.getBoundingClientRect());
}

/**
 * Wires a swatch + caret + hidden <input type="color"> as one control. The
 * swatch is the default click target and applies the currently-held color
 * immediately — no picker in the way, so reusing the same color across
 * several shapes is one click each time. The caret opens the shared custom
 * color picker panel (see initColorPickerPanel) to actually change the
 * color; that panel's "More [Fill/Line/Text] Colors…" button re-purposes
 * this same hidden input to reach the native OS picker for full-spectrum
 * picking, so this still needs to react when the input's value changes
 * there (listened on `input` rather than `change` — macOS's native colour
 * panel has no explicit commit action, and `change` is unreliable in
 * WKWebView-hosted Mac Office task panes, while `input` fires live as the
 * user moves around the picker). Debounced so dragging around that picker
 * doesn't fire a PowerPoint.run call per pixel — it applies once movement
 * settles for 150ms.
 *
 * `kind` drives the panel's per-control bits (see openColorPickerPanel):
 * which "No Fill"/"No Line" label (if any) to show, and the "More ...
 * Colors…" button's label. `noColorHandler`, when given, is what "No
 * Fill"/"No Line" actually calls — omitted for text, which has no
 * equivalent in PowerPoint's own Font Color dropdown.
 */
function bindColorControl(
  baseId: string,
  kind: ColorControlKind,
  handler: (hex: string) => Promise<void>,
  noColorHandler?: () => Promise<void>
): void {
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
  caret.addEventListener("click", () => openColorPickerPanel(caret, input, swatch, kind, handler, noColorHandler));
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
  // Task Pane Phase 14
  companyDomain: string | null;
  isCompanyAdmin: boolean;
  // Task Pane Phase 15 — the real access gate for every account-dependent
  // feature (Content Library, My Library); a session existing at all no
  // longer implies the account is actually usable.
  isRegistered: boolean;
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
    notify(`"${item.title}" added on a temporary slide — copy it across, then finish.`);
  } else {
    notify(`"${item.title}" inserted.`);
  }
}

/**
 * The gallery's "Insert as New Slide" route (#btnInsertAsSlide there) —
 * only ever sent for 'file'-mode items (the button is hidden for every
 * other mode), which is the only mode insertFileItemAsNewSlide handles.
 * Inserts directly as a permanent slide: no temp slide, no copy/paste, no
 * Finish row.
 */
async function insertPickedItemAsSlide(item: Library.CatalogItem): Promise<void> {
  if (item.insertMode !== "file") {
    throw new Error(`"Insert as New Slide" isn't available for "${item.title}".`);
  }
  await Library.insertFileItemAsNewSlide(item.id);
  notify(`"${item.title}" inserted as a new slide.`);
}

/**
 * Gates the Content Library section whenever sign-in/registration state
 * changes. The gallery dialog loads its own data lazily on open, so
 * there's nothing to pre-fetch here. Task Pane Phase 15: gated on
 * isRegistered, not just a valid session — an unregistered user is still
 * "signed in" (their session cookie is real) but has no access to any
 * account-dependent feature until they finish creating their account.
 */
function refreshLibrarySection(user: SessionUser | null): void {
  const unlocked = !!user?.isRegistered;
  const supported = Library.isLibraryInsertSupported();
  const reason = !user
    ? "Sign in above to browse the content library."
    : !user.isRegistered
      ? 'Finish creating your account above to browse the content library.'
      : "Requires a newer PowerPoint build (PowerPointApi 1.2) than this one has.";
  setSectionEnabled("sectionLibrary", unlocked && supported, reason);
  if (!unlocked || !supported) {
    showLibraryFinishRow(false);
    currentFileInsertHandle = null;
  }

  // Scope picker (label + select): admin/company-admin only — everyone
  // else's Add click always saves to their own personal library, so
  // there's nothing for them to pick (see TARGET_ENDPOINTS/
  // currentLibraryTarget, which already default to "personal" when the
  // select has no other option selected). The button itself always stays
  // in this same row (see taskpane.html) — only its label text changes.
  const showTargetPicker = unlocked && (!!user?.isAdmin || !!user?.isCompanyAdmin);
  const targetLabel = document.getElementById("libraryTargetLabel");
  if (targetLabel) (targetLabel as HTMLElement).style.display = showTargetPicker ? "" : "none";

  const targetSelect = document.getElementById("libraryTarget") as HTMLSelectElement | null;
  if (targetSelect) {
    (targetSelect as HTMLElement).style.display = showTargetPicker ? "" : "none";
    const previous = targetSelect.value;
    targetSelect.innerHTML = "";
    const addOption = (value: LibraryTarget, label: string) => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = label;
      targetSelect.appendChild(opt);
    };
    addOption("personal", "My Items");
    // Fixed short labels, not the real company domain — a long domain name
    // was pushing the Add button onto its own row (the row has nowhere
    // else to shrink once the option text is that long).
    if (user?.companyDomain && user?.isCompanyAdmin) addOption("company", "Company Library");
    if (user?.isAdmin) addOption("global", "Global Library");
    if (Array.from(targetSelect.options).some((o) => o.value === previous)) targetSelect.value = previous;
  }

  // Admins/company admins get a picker, so the button just says "Add"; a
  // plain user has no picker at all, so it spells out where content goes.
  setAddToLibraryLabel(showTargetPicker ? ADD_TO_LIBRARY_LABEL_ADMIN : ADD_TO_LIBRARY_LABEL_PLAIN);

  // Add-to-library needs a higher requirement set (PowerPointApi 1.8, for
  // exportAsBase64/getImageAsBase64) than the rest of this section (1.2) —
  // disabled with its own note instead of folding into the blanket
  // setSectionEnabled check above, which would incorrectly gate Browse
  // Library too.
  const editSupported = Library.isAdminLibraryEditSupported();
  const addBtn = document.getElementById("btnLibraryAdd") as HTMLButtonElement | null;
  if (addBtn) addBtn.disabled = !unlocked || !editSupported;
  const addNote = document.getElementById("libraryAddUnsupportedNote");
  if (addNote) {
    (addNote as HTMLElement).style.display = unlocked && !editSupported ? "block" : "none";
    if (unlocked && !editSupported) {
      addNote.textContent = "Adding to the library requires a newer PowerPoint build (PowerPointApi 1.8) than this one has.";
    }
  }
}

// ---------------------------------------------------------------------------
// My Library: add a new item from the current slide, or edit an existing
// item's graphic natively in PowerPoint (Task Pane Phase 12, generalized
// in Phase 13 from admin-only to every signed-in user, with a "Save to"
// target picking which scope a new item lands in). Mirrors the file-mode
// insert/Finish state machine above — currentFileInsertHandle/
// showLibraryFinishRow for end users, currentLibraryEdit/
// updateLibraryEditUI for this variant — since both are the same
// underlying "temp slide, do something, clean up" shape.
// ---------------------------------------------------------------------------

type LibraryTarget = "personal" | "global" | "company";

// Where a brand-new item gets POSTed, keyed by the #libraryTarget select's
// value. "global" and "company" both hit the same admin-gated route —
// which one the server allows is decided by isAdmin/isCompanyAdmin plus a
// "scope" field in the request body (see addSelectedSlideToLibrary) —
// "personal" requires only a signed-in user, at its own route entirely.
const TARGET_ENDPOINTS: Record<LibraryTarget, string> = {
  personal: "/api/personal/catalog",
  global: "/api/admin/catalog",
  company: "/api/admin/catalog",
};

function currentLibraryTarget(): LibraryTarget {
  const select = document.getElementById("libraryTarget") as HTMLSelectElement | null;
  const value = select?.value;
  return value === "global" || value === "company" ? value : "personal";
}

let currentLibraryEdit: { itemId: number; title: string; handle: Library.FileInsertHandle; isPersonal: boolean } | null =
  null;

// Cached purely so updateAdminSectionVisibility (called from the
// begin/save/cancel edit handlers, not just refreshMyLibrarySection) knows
// the current user's admin status without needing a fresh async fetch.
let lastKnownUser: SessionUser | null = null;

/**
 * The Administration section is shown to any admin — global or company —
 * by default, but a non-admin editing their own personal item's graphic
 * (via the gallery's Edit button) still needs to see #libraryEditRow's
 * Save/Cancel controls, so visibility is "isAdmin OR isCompanyAdmin OR an
 * edit is currently in progress", not a flat admin-only gate. A company
 * admin has real powers reachable via /admin/users (their own company's
 * user list) even though they can't reach the global catalog admin view,
 * so they need this entry point too, same as a global admin.
 */
function updateAdminSectionVisibility(): void {
  const section = document.getElementById("sectionMyLibrary");
  const show =
    !!lastKnownUser?.isRegistered && (!!lastKnownUser?.isAdmin || !!lastKnownUser?.isCompanyAdmin || !!currentLibraryEdit);
  if (section) (section as HTMLElement).style.display = show ? "" : "none";
}

function updateLibraryEditUI(): void {
  const editRow = document.getElementById("libraryEditRow");
  if (editRow) (editRow as HTMLElement).style.display = currentLibraryEdit ? "block" : "none";
  const titleInput = document.getElementById("libraryEditTitle") as HTMLInputElement | null;
  if (titleInput && !currentLibraryEdit) titleInput.style.display = "none";
  updateAdminSectionVisibility();
}

/**
 * "Open Admin…" — same admin-or-company-admin gate as the section itself
 * (see updateAdminSectionVisibility); a company admin's /admin/users view
 * is auto-scoped to their own company server-side, so there's nothing
 * unsafe about surfacing this link to them too.
 */
function refreshMyLibrarySection(user: SessionUser | null): void {
  lastKnownUser = user;
  if (!user?.isRegistered) {
    currentLibraryEdit = null;
  }

  const openAdminRow = document.getElementById("openAdminRow");
  if (openAdminRow) (openAdminRow as HTMLElement).style.display = user?.isAdmin || user?.isCompanyAdmin ? "" : "none";
  updateLibraryEditUI(); // also updates section visibility via updateAdminSectionVisibility
}

/**
 * Determines a clicked item's scope from ownerOid rather than the
 * #libraryTarget select — editing an *existing* item's scope is fixed by
 * that item, not user-chosen. The server re-checks real ownership
 * regardless (requireAdmin for shared items, owner_oid match for personal
 * ones), so this client-side inference only decides which endpoint to
 * call, never acts as the actual authorization.
 */
async function beginLibraryEdit(item: Library.CatalogItem): Promise<void> {
  if (currentLibraryEdit) {
    notify("Finish or cancel the current edit first.", "error");
    return;
  }
  if (item.insertMode === "unicode-char") {
    notify("Character items have no graphic to edit.", "error");
    return;
  }
  const handle =
    item.insertMode === "file"
      ? await Library.insertFileItem(item.id)
      : await Library.insertReconstructedItemOnTempSlide(item.reconstructSpec!);
  const isPersonal = !!item.ownerOid;
  currentLibraryEdit = { itemId: item.id, title: item.title, handle, isPersonal };
  const status = document.getElementById("libraryEditStatus");
  if (status) status.textContent = `Editing "${item.title}" — make your changes, then Save.`;
  // Personal items have no /admin-equivalent page, so title editing lives
  // here; shared items keep editing their title via /admin as before.
  const titleInput = document.getElementById("libraryEditTitle") as HTMLInputElement | null;
  if (titleInput) {
    titleInput.style.display = isPersonal ? "" : "none";
    titleInput.value = item.title;
  }
  updateLibraryEditUI();
  notify(`Editing "${item.title}" on a temporary slide.`);
}

async function saveLibraryEdit(): Promise<void> {
  if (!currentLibraryEdit) return;
  const { itemId, title, handle, isPersonal } = currentLibraryEdit;
  const endpointPrefix = isPersonal ? "/api/personal/catalog" : "/api/admin/catalog";
  const { pptxBase64, thumbnailBase64 } = await Library.exportSlideForAdmin(handle.tempSlideId);
  const res = await fetchWithTimeout(`${endpointPrefix}/${itemId}/content`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ pptxBase64, thumbnailBase64 }),
  });
  if (!res.ok) {
    notify(`Couldn't save "${title}" to the library (${res.status}).`, "error");
    return;
  }

  let finalTitle = title;
  if (isPersonal) {
    const titleInput = document.getElementById("libraryEditTitle") as HTMLInputElement | null;
    const newTitle = titleInput?.value.trim();
    if (newTitle && newTitle !== title) {
      const renameRes = await fetchWithTimeout(`/api/personal/catalog/${itemId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title: newTitle }),
      });
      if (renameRes.ok) finalTitle = newTitle;
    }
  }

  await Library.finishFileInsert(handle);
  currentLibraryEdit = null;
  updateLibraryEditUI();
  notify(`Saved "${finalTitle}" to the library.`);
}

async function cancelLibraryEdit(): Promise<void> {
  if (!currentLibraryEdit) return;
  await Library.finishFileInsert(currentLibraryEdit.handle);
  currentLibraryEdit = null;
  updateLibraryEditUI();
}

/**
 * Deletion has no in-task-pane UI (see deleteLibraryItem below, triggered
 * from the gallery dialog instead) since the gallery is an ordinary
 * browser context where window.confirm works fine — the arm/confirm
 * pattern here exists only because *this* task pane's embedded webview
 * doesn't support window.confirm/alert/prompt (confirmed directly: the
 * same call works in the gallery dialog and in /admin, both more ordinary
 * browser contexts, but throws "Function window.confirm is not supported"
 * here). Same fix already applied once this session for the color picker:
 * stop relying on the unsupported native thing, build a small in-page
 * equivalent instead — here, a two-step "click again to confirm" on the
 * button itself, no new markup needed.
 */
let addToLibraryArmed = false;
let addToLibraryArmedTimer: number | undefined;

// Admins/company admins have a scope picker next to the button, so it just
// says "Add"; a plain user has no picker at all, so the button spells out
// where the content is going. Whichever applies is set by
// refreshLibrarySection (via setAddToLibraryLabel) and remembered here so
// the arm/confirm two-click below can restore the right one afterward,
// rather than a single hardcoded constant.
const ADD_TO_LIBRARY_LABEL_ADMIN = "Add";
const ADD_TO_LIBRARY_LABEL_PLAIN = "Add to my library";
const ADD_TO_LIBRARY_CONFIRM_LABEL = "Click again to confirm — adds the current slide";
let currentAddToLibraryLabel = ADD_TO_LIBRARY_LABEL_PLAIN;

/** Sets #btnLibraryAdd's visible text without touching its icon <img> pair (a plain btn.textContent= would wipe those out). */
function setAddToLibraryLabel(label: string): void {
  currentAddToLibraryLabel = label;
  if (addToLibraryArmed) return; // don't clobber the "click again to confirm" prompt mid-arm
  const labelEl = document.querySelector("#btnLibraryAdd .btn-label");
  if (labelEl) labelEl.textContent = label;
}

async function addSelectedSlideToLibrary(): Promise<void> {
  const labelEl = document.querySelector("#btnLibraryAdd .btn-label");
  if (!addToLibraryArmed) {
    addToLibraryArmed = true;
    if (labelEl) labelEl.textContent = ADD_TO_LIBRARY_CONFIRM_LABEL;
    window.clearTimeout(addToLibraryArmedTimer);
    addToLibraryArmedTimer = window.setTimeout(() => {
      addToLibraryArmed = false;
      if (labelEl) labelEl.textContent = currentAddToLibraryLabel;
    }, 4000);
    return;
  }
  addToLibraryArmed = false;
  window.clearTimeout(addToLibraryArmedTimer);
  if (labelEl) labelEl.textContent = currentAddToLibraryLabel;

  const target = currentLibraryTarget();
  const { pptxBase64, thumbnailBase64 } = await Library.exportCurrentSlideForAdmin();
  // "company" and "global" share one endpoint (see TARGET_ENDPOINTS) — the
  // scope field is what the server uses to pick which admin check applies
  // and which columns the new row gets (see POST /api/admin/catalog).
  const body: { pptxBase64: string; thumbnailBase64: string; scope?: "company" } =
    target === "company" ? { pptxBase64, thumbnailBase64, scope: "company" } : { pptxBase64, thumbnailBase64 };
  const res = await fetchWithTimeout(TARGET_ENDPOINTS[target], {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    notify(`Couldn't add to the library (${res.status}).`, "error");
    return;
  }
  if (target === "global") {
    const { id, category } = await res.json();
    window.open(`/admin?category=${category}&highlight=${id}`, "_blank");
    notify("Added to the library — finish naming it in the Admin tab that just opened.");
  } else if (target === "company") {
    const { id, companyDomain } = await res.json();
    window.open(`/admin?scope=company:${encodeURIComponent(companyDomain)}&highlight=${id}`, "_blank");
    notify("Added to the company library — finish naming it in the Admin tab that just opened.");
  } else {
    notify('Added to your personal library — open "Browse Library…" to rename or insert it.');
  }
}

/** Triggered from the gallery dialog's Delete button (which does its own window.confirm — see the comment above addSelectedSlideToLibrary for why that's fine there but not here). */
async function deleteLibraryItem(item: Library.CatalogItem): Promise<void> {
  const url = item.ownerOid ? `/api/personal/catalog/${item.id}/delete` : `/admin/catalog/${item.id}/delete`;
  const res = await fetchWithTimeout(url, { method: "POST", headers: { Accept: "application/json" } });
  if (!res.ok) {
    notify(`Couldn't delete "${item.title}" (${res.status}).`, "error");
    return;
  }
  notify(`Deleted "${item.title}".`);
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
  // The server already knows exactly why a sign-in was refused (suspended,
  // deleted, an invalid token) and says so in the response body — surface
  // that instead of a bare status code, which told the user nothing.
  if (!res.ok) throw new Error(await extractErrorMessage(res, `Failed to sign in (${res.status}).`));
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
  const textEl = el?.querySelector(".sign-in-status-text");
  if (!el || !textEl) return;
  if (user && !user.isRegistered) {
    textEl.textContent = `Signed in as ${user.email ?? user.displayName ?? "unknown user"} — finish creating your account.`;
    el.classList.remove("signed-in");
  } else if (user) {
    textEl.textContent = `Signed in as ${user.email ?? user.displayName ?? "unknown user"}${user.isAdmin ? " (admin)" : ""}.`;
    el.classList.add("signed-in");
  } else {
    textEl.textContent = "Not signed in.";
    el.classList.remove("signed-in");
  }
}

/**
 * Task Pane Phase 15: three states, not two — no session at all shows
 * "Sign In"; a session that exists but hasn't finished registering shows
 * "Create Account" (clicking it opens the registration dialog directly,
 * see bindButton("btnSignIn", ...) below — no second Microsoft auth
 * needed); a fully registered session hides the whole section, not just
 * the button — there's nothing else in it worth keeping on-screen once
 * signed in (the actual "Signed in as…" status now lives in the page
 * footer, see updateSignInStatus, so it stays visible either way). The
 * admin-only Open Admin link lives in sectionMyLibrary (see
 * refreshMyLibrarySection), not here.
 */
function updateAuthButtons(user: SessionUser | null): void {
  const section = document.getElementById("sectionAuth");
  const signIn = document.getElementById("btnSignIn") as HTMLButtonElement | null;
  if (!signIn) return;
  if (!user) {
    if (section) section.style.display = "";
    signIn.textContent = "Sign In";
  } else if (!user.isRegistered) {
    if (section) section.style.display = "";
    signIn.textContent = "Create Account";
  } else if (section) {
    section.style.display = "none";
  }
}

/** Applies a change in sign-in state everywhere it matters — the status line, the auth buttons, and the Content Library gates. */
function applySessionState(user: SessionUser | null): void {
  updateSignInStatus(user);
  updateAuthButtons(user);
  refreshLibrarySection(user);
  refreshMyLibrarySection(user);
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

/**
 * Task Pane Phase 15: opens the account-creation dialog, same
 * displayDialogAsync mechanism as btnBrowseLibrary's gallery wiring below
 * — the dialog captures details and calls POST /api/auth/register itself
 * (it can make its own authenticated fetch calls directly, same as the
 * gallery already does for /api/catalog/...), then reports back a bare
 * { success: true } via messageParent so the task pane knows to refresh
 * its own session state and unlock the account-dependent sections.
 */
function openRegistrationDialog(): void {
  Office.context.ui.displayDialogAsync(
    `${window.location.origin}/register.html?v=${Date.now()}`,
    { height: 60, width: 40 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        notify(`Failed to open account creation: ${result.error.message}`, "error");
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        if (!("message" in args)) return;
        dialog.close();
        try {
          const parsed = JSON.parse(args.message);
          if (parsed?.success) {
            getSessionUser().then((user) => {
              applySessionState(user);
              notify("Account created.");
            });
          }
        } catch {
          // malformed message — nothing to do
        }
      });
    }
  );
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  initSectionReordering();
  initColorPickerPanel();

  const statusEl = document.getElementById("status");
  if (statusEl) bindStatusElement(statusEl);

  // Task Pane Phase 15: a returning-but-unregistered visitor already has a
  // valid session (checked first, cheaply) — that skips straight to the
  // registration dialog with no Microsoft popup at all. Only a visitor
  // with no session yet goes through the full Auth.signIn() flow; if the
  // resulting account also turns out to be unregistered (a brand-new
  // sign-in), the dialog opens immediately after, so one "Sign In" click
  // takes a new user straight through Microsoft auth into account
  // creation, no second click required.
  bindButton("btnSignIn", async () => {
    const errorEl = document.getElementById("signInError") as HTMLElement | null;
    if (errorEl) errorEl.style.display = "none";
    try {
      let user = await getSessionUser();
      if (!user) {
        const msalUser = await Auth.signIn((step) => notify(step));
        if (!msalUser.email) {
          // The one case where dumping raw claims is actually useful — an
          // admin misconfiguring the Azure app registration's optional
          // claims needs to see what Microsoft actually sent back. Never
          // shown on a normal sign-in (it would otherwise expose the live
          // idToken in the visible UI for no reason).
          const claimsEl = document.getElementById("authClaims") as HTMLElement;
          claimsEl.style.display = "block";
          claimsEl.textContent = JSON.stringify(msalUser, null, 2);
          notify("Signed in, but no email claim was returned — check the Azure app registration's optional claims.", "error");
          return;
        }
        await establishSession(msalUser.idToken);
        user = await getSessionUser();
        applySessionState(user);
      }
      if (user && !user.isRegistered) {
        openRegistrationDialog();
      } else if (user) {
        notify(`Signed in as ${user.email}`);
      }
    } catch (err) {
      // Surfaced here (right under the button, where the user's eyes
      // already are) in addition to the bottom-of-page status footer —
      // a suspended/deleted-account rejection was easy to miss down there.
      const message = err instanceof Error ? err.message : String(err);
      if (errorEl) {
        errorEl.textContent = message;
        errorEl.style.display = "block";
      }
      throw err;
    }
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
          let payload: { action?: string; item: Library.CatalogItem };
          try {
            const parsed = JSON.parse(args.message);
            // { action, item } as of Task Pane Phase 12 — tolerate a bare
            // item too (the old shape), defaulting to "insert", in case
            // anything still sends that.
            payload = "item" in parsed ? parsed : { action: "insert", item: parsed };
          } catch (err) {
            notify(`Couldn't read the selected item: ${err instanceof Error ? err.message : String(err)}`, "error");
            return;
          }
          const action =
            payload.action === "edit"
              ? beginLibraryEdit
              : payload.action === "delete"
                ? deleteLibraryItem
                : payload.action === "insert-as-slide"
                  ? insertPickedItemAsSlide
                  : insertPickedItem;
          action(payload.item).catch((err) =>
            notify(`Error: ${err instanceof Error ? err.message : String(err)}`, "error")
          );
        });
      }
    );
  });

  // Task Pane Phase 20 — Template Library. Same displayDialogAsync/
  // messageParent shape as btnBrowseLibrary above, kept as its own
  // handler (different dialog, different message shape) rather than
  // merged into that one. "Use This Template" is the dialog's only
  // outward action — everything else (browse/upload/rename/delete)
  // happens via plain fetch calls made directly from inside the dialog,
  // no PowerPoint API involved, so no round-trip needed for those.
  document.getElementById("btnBrowseTemplates")?.addEventListener("click", () => {
    Office.context.ui.displayDialogAsync(
      `${window.location.origin}/templates.html?v=${Date.now()}`,
      { height: 80, width: 70 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          notify(`Failed to open the template library: ${result.error.message}`, "error");
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          if (!("message" in args)) return;
          dialog.close();
          try {
            const payload = JSON.parse(args.message) as { action: string; templateId: number };
            if (payload.action !== "use-template") return;
            Templates.createFromTemplate(payload.templateId).catch((err) =>
              notify(`Error: ${err instanceof Error ? err.message : String(err)}`, "error")
            );
          } catch (err) {
            notify(`Couldn't read the selected template: ${err instanceof Error ? err.message : String(err)}`, "error");
          }
        });
      }
    );
  });

  bindButton("btnLibraryFinishDelete", async () => {
    if (!currentFileInsertHandle) return;
    await Library.finishFileInsert(currentFileInsertHandle);
    currentFileInsertHandle = null;
    showLibraryFinishRow(false);
    notify("Done — temporary slide removed.");
  });
  bindButton("btnLibraryFinishKeep", async () => {
    if (!currentFileInsertHandle) return;
    await Library.finishFileInsert(currentFileInsertHandle, { keep: true });
    currentFileInsertHandle = null;
    showLibraryFinishRow(false);
    notify("Done — temporary slide kept.");
  });

  bindButton("btnLibraryAdd", addSelectedSlideToLibrary);
  bindButton("btnLibrarySave", saveLibraryEdit);
  bindButton("btnLibraryCancel", cancelLibraryEdit);

  applySessionState(null); // starting state — the background check below updates this once it resolves, however long that takes

  // Default swatch colour comes from config/theme.json, not hardcoded in the HTML.
  for (const id of ["fillColorInput", "lineColorInput", "textColorInput"]) {
    const input = document.getElementById(id) as HTMLInputElement | null;
    if (input) input.value = theme.defaultColorSwatch;
  }

  // Fill, line & text color
  bindColorControl(
    "fillColor",
    "fill",
    async (hex) => {
      await FillLineColors.fillColor(FillLineColors.hexToRgb(hex));
      notify(`Fill set to ${hex}.`);
    },
    async () => {
      await FillLineColors.noFill();
      notify("Fill removed.");
    }
  );
  bindColorControl(
    "lineColor",
    "line",
    async (hex) => {
      await FillLineColors.lineColor(FillLineColors.hexToRgb(hex));
      notify(`Line set to ${hex}.`);
    },
    async () => {
      await FillLineColors.noLine();
      notify("Line removed.");
    }
  );
  bindColorControl("textColor", "text", async (hex) => {
    await FillLineColors.textColor(FillLineColors.hexToRgb(hex));
    notify(`Text color set to ${hex}.`);
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

  // Group & order (requires PowerPointApi 1.8 — addGroup/ungroup/setZOrder)
  bindButton("btnGroupShapes", ObjectOrder.groupShapes);
  bindButton("btnUngroupShapes", ObjectOrder.ungroupShapes);
  bindButton("btnBringToFront", ObjectOrder.bringToFront);
  bindButton("btnSendToBack", ObjectOrder.sendToBack);
  bindButton("btnBringForward", ObjectOrder.bringForward);
  bindButton("btnSendBackward", ObjectOrder.sendBackward);
  setSectionEnabled(
    "sectionObjectOrder",
    ObjectOrder.isGroupingSupported(),
    "Requires a newer PowerPoint build (PowerPointApi 1.8) than this one has."
  );

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
