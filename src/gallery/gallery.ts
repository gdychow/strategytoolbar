/**
 * Phase 5: the pop-out content library. A page opened via
 * Office.context.ui.displayDialogAsync can call exactly two Office.js
 * APIs — messageParent and requirements.isSetSupported (confirmed against
 * Microsoft's own docs source, office-js-docs-pr's
 * dialog-api-in-office-add-ins.md) — so this page is a pure browse/
 * search/select surface. It never touches PowerPoint APIs itself; it
 * reports the chosen item's id back to the task pane via messageParent,
 * and the task pane (which keeps full PowerPoint API access) performs
 * the actual insert via the existing, unmodified
 * src/features/libraryInsert.ts engine.
 *
 * Session cookie sharing between this dialog and the task pane was
 * verified directly against a real Office host before building this out
 * (both are same-origin, which was already required for messageParent to
 * work at all) — confirmed working, so this fetches exactly like any
 * other authenticated page, no special session hand-off needed.
 */
import { fetchCatalog, type CatalogItem, type CatalogResponse } from "../features/libraryInsert";

// Matches what's actually seeded (see db/seed/catalog-*.json). "My Items"
// (Task Pane Phase 13) is appended once sign-in state is known, not listed
// here statically — see Office.onReady, which awaits the auth check before
// the first render so the tab is present from the start rather than
// popping in after a delay.
let CATEGORIES: { value: string; label: string }[] = [
  { value: "text", label: "Text" },
  { value: "objects", label: "Objects" },
  { value: "shapes", label: "Shapes" },
  { value: "stamps", label: "Stamps" },
  { value: "tables", label: "Tables" },
  { value: "diagrams", label: "Diagrams" },
  { value: "symbols", label: "Symbols" },
  { value: "maps", label: "Maps" },
  { value: "clipart", label: "Clip Art" },
  { value: "frameworks", label: "Frameworks" },
  { value: "flags", label: "Flags" },
];

const cache = new Map<string, CatalogResponse>();
let activeCategory = CATEGORIES[0].value;
let selectedItem: CatalogItem | null = null;
// Task Pane Phase 16 follow-up: the last real (non-"__new__") value of
// #editGroup, so a cancelled/failed "+ Add new group…" attempt can revert
// the select instead of leaving it stuck on the placeholder option.
let editGroupPreviousValue = "";
// Task Pane Phase 16: the one card currently being dragged, mirroring
// /admin's own draggedCard/getDragAfterElement pattern (server.js) — same
// nearest-card-center heuristic, generalized from a dedicated drag handle
// to the whole card (these tiles are small, and only editable cards are
// draggable at all, so there's no accidental-drag risk from an ordinary
// browsing click).
let draggedCard: HTMLElement | null = null;
// Fetched once on load, before the first render (Task Pane Phase 13 moved
// this earlier — it used to run fire-and-forget after the initial paint,
// only affecting the Edit button, but the tab list now depends on it too).
let isAdmin = false;
let myOid: string | null = null;
let myTid: string | null = null;
// Task Pane Phase 14
let myCompanyDomain: string | null = null;
let isCompanyAdmin = false;

function statusEl(): HTMLElement | null {
  return document.getElementById("status");
}

function searchInput(): HTMLInputElement | null {
  return document.getElementById("searchInput") as HTMLInputElement | null;
}

function renderTabs(): void {
  const tabsEl = document.getElementById("tabs");
  if (!tabsEl) return;
  tabsEl.innerHTML = "";
  for (const cat of CATEGORIES) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "gallery-tab" + (cat.value === activeCategory ? " active" : "");
    btn.textContent = cat.label;
    btn.addEventListener("click", () => switchCategory(cat.value));
    tabsEl.appendChild(btn);
  }
}

/**
 * window.confirm() doesn't reliably produce a visible native dialog in
 * this displayDialogAsync-hosted webview — confirmed directly (the same
 * gap already found and fixed the same way in the Template Library
 * dialog, src/templates/templates.ts): clicking Delete with a plain
 * window.confirm() call silently did nothing, no dialog, no error, no
 * deletion. The comment this replaced claimed this dialog is "an ordinary
 * browser context where window.confirm works fine" — that was never
 * actually true for a displayDialogAsync-hosted page specifically (it was
 * confirmed only for /admin's separate, real-browser-tab context, a
 * genuinely different host). A two-click "click again to confirm" on the
 * button itself has no dependency on the host's native dialog support.
 */
let deleteArmed = false;
let deleteArmedTimer: number | undefined;
const DELETE_LABEL = "Delete";
const DELETE_CONFIRM_LABEL = "Click again to confirm";

function resetDeleteArm(): void {
  deleteArmed = false;
  window.clearTimeout(deleteArmedTimer);
  const btn = document.getElementById("btnDelete");
  if (btn) btn.textContent = DELETE_LABEL;
}

function hidePreview(): void {
  selectedItem = null;
  resetDeleteArm();
  const panel = document.getElementById("previewPanel");
  if (panel) panel.style.display = "none";
}

/**
 * Whether the signed-in viewer can Edit/Delete this item — a global admin
 * can touch any shared item; a company admin only their own company's
 * items (Task Pane Phase 14); anyone else only their own personal item
 * (Task Pane Phase 13). No "graphic" to edit/delete for a bare character
 * either way (see Task Pane Phase 12). The server re-checks real ownership
 * on every mutating request regardless of what this returns — this only
 * controls whether the buttons are offered.
 */
function canEdit(item: CatalogItem): boolean {
  if (item.insertMode === "unicode-char") return false;
  if (isAdmin) return true;
  if (item.ownerOid) return item.ownerOid === myOid && item.ownerTid === myTid;
  if (item.companyDomain) return isCompanyAdmin && item.companyDomain === myCompanyDomain;
  return false;
}

/**
 * Task Pane Phase 16 (redesigned): title/tags/group are always visible
 * once an item is selected — read-only display for anyone who can't edit
 * this item, editable inputs (with an explicit Save) for whoever can. No
 * more toggled "Edit Details" button/form — matches the Template Library
 * dialog's own always-visible manage row, generalized to also show
 * (rather than hide) the read-only case for non-owners, since tags/group
 * are useful context while browsing regardless of edit rights.
 */
function renderMetaSection(item: CatalogItem, editable: boolean): void {
  const metaSection = document.getElementById("metaSection");
  const scopedFields = document.getElementById("metaScopedFields");
  const tagsDisplay = document.getElementById("metaTagsDisplay");
  const groupDisplay = document.getElementById("metaGroupDisplay");
  const tagsInput = document.getElementById("editTags") as HTMLInputElement | null;
  const groupSelect = document.getElementById("editGroup") as HTMLSelectElement | null;
  const saveBtn = document.getElementById("btnSaveDetails");
  const errorEl = document.getElementById("editError");
  if (!metaSection || !scopedFields || !tagsDisplay || !groupDisplay || !tagsInput || !groupSelect || !saveBtn) return;

  if (errorEl) {
    errorEl.style.display = "none";
    errorEl.textContent = "";
  }

  const isPersonal = !!item.ownerOid;
  // Personal items have no tags/group at all — the section still shows
  // (with just a Save button) for the owner, since the title itself is
  // still editable; for anyone else it never applies (personal items
  // aren't visible to non-owners in the first place).
  metaSection.style.display = !isPersonal || editable ? "flex" : "none";
  scopedFields.style.display = isPersonal ? "none" : "flex";
  saveBtn.style.display = editable ? "" : "none";

  if (isPersonal) return;

  tagsDisplay.style.display = editable ? "none" : "";
  tagsInput.style.display = editable ? "" : "none";
  groupDisplay.style.display = editable ? "none" : "";
  groupSelect.style.display = editable ? "" : "none";
  tagsDisplay.textContent = item.tags.length ? item.tags.join(", ") : "—";
  groupDisplay.textContent = item.groupName ?? "—";

  if (!editable) return;
  tagsInput.value = item.tags.join(", ");
  groupSelect.innerHTML = "";
  const noneOption = document.createElement("option");
  noneOption.value = "";
  noneOption.textContent = "(none)";
  groupSelect.appendChild(noneOption);
  const groups = cache.get(activeCategory)?.groups ?? [];
  for (const group of [...groups].sort((a, b) => a.sortOrder - b.sortOrder)) {
    const opt = document.createElement("option");
    opt.value = String(group.id);
    opt.textContent = group.name;
    groupSelect.appendChild(opt);
  }
  const newGroupOption = document.createElement("option");
  newGroupOption.value = "__new__";
  newGroupOption.textContent = "+ Add new group…";
  groupSelect.appendChild(newGroupOption);
  groupSelect.value = item.groupId !== null ? String(item.groupId) : "";
  editGroupPreviousValue = groupSelect.value;
}

function showPreview(item: CatalogItem): void {
  const panel = document.getElementById("previewPanel");
  const img = document.getElementById("previewImg") as HTMLImageElement | null;
  const title = document.getElementById("previewTitle");
  const titleInput = document.getElementById("editTitle") as HTMLInputElement | null;
  const insertAsSlideBtn = document.getElementById("btnInsertAsSlide");
  const manageActionsRow = document.getElementById("manageActionsRow");
  const editBtn = document.getElementById("btnEdit");
  const deleteBtn = document.getElementById("btnDelete");
  if (!panel || !img || !title || !titleInput) return;
  resetDeleteArm(); // switching selection shouldn't leave a stale "click again to confirm" pointed at the wrong item
  const previewErrorEl = document.getElementById("previewError");
  if (previewErrorEl) {
    previewErrorEl.style.display = "none";
    previewErrorEl.textContent = "";
  }
  // Only 'file'-mode items go through the temp-slide/copy-paste/Finish
  // dance at all — 'reconstruct'/'unicode-char' items already insert
  // directly in one step, so a second "as new slide" route is meaningless
  // for them.
  if (insertAsSlideBtn) (insertAsSlideBtn as HTMLElement).style.display = item.insertMode === "file" ? "" : "none";
  const existingGlyph = panel.querySelector(".preview-glyph");
  if (existingGlyph) existingGlyph.remove();
  if (item.unicodeChar) {
    img.style.display = "none";
    const glyph = document.createElement("span");
    glyph.className = "preview-glyph";
    glyph.textContent = item.unicodeChar;
    panel.insertBefore(glyph, img);
  } else if (item.thumbnailUrl) {
    img.src = item.thumbnailUrl;
    img.style.display = "";
  } else {
    img.style.display = "none";
  }
  panel.style.display = "flex";
  const editable = canEdit(item);
  // Title: static heading for read-only viewers, an editable input (fed
  // straight into renderMetaSection's Save) for whoever can edit this item.
  title.style.display = editable ? "none" : "";
  titleInput.style.display = editable ? "" : "none";
  title.textContent = item.title;
  titleInput.value = item.title;
  if (manageActionsRow) manageActionsRow.style.display = editable ? "flex" : "none";
  if (editBtn) (editBtn as HTMLElement).style.display = editable ? "" : "none";
  if (deleteBtn) (deleteBtn as HTMLElement).style.display = editable ? "" : "none";
  renderMetaSection(item, editable);
}

async function saveEditDetails(item: CatalogItem): Promise<void> {
  const titleInput = document.getElementById("editTitle") as HTMLInputElement | null;
  const tagsInput = document.getElementById("editTags") as HTMLInputElement | null;
  const groupSelect = document.getElementById("editGroup") as HTMLSelectElement | null;
  const errorEl = document.getElementById("editError");
  if (!titleInput) return;

  const title = titleInput.value.trim();
  const isPersonal = !!item.ownerOid;

  try {
    if (isPersonal) {
      const res = await fetch(`/api/personal/catalog/${item.id}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
      item.title = title;
    } else {
      const tags = (tagsInput?.value ?? "")
        .split(",")
        .map((t) => t.trim())
        .filter(Boolean);
      const groupIdRaw = groupSelect?.value ?? "";
      const groupId = groupIdRaw ? Number(groupIdRaw) : null;
      const res = await fetch(`/admin/catalog/${item.id}`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ title, tags: tags.join(","), groupId }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
      item.title = title;
      item.tags = tags;
      item.groupId = groupId;
      const group = (cache.get(activeCategory)?.groups ?? []).find((g) => g.id === groupId);
      item.groupName = group?.name ?? null;
    }
    showPreview(item);
    rerenderGrid();
  } catch (err) {
    if (errorEl) {
      errorEl.textContent = err instanceof Error ? err.message : String(err);
      errorEl.style.display = "block";
    }
  }
}

/**
 * Maps the gallery's tab value to the shape POST /admin/catalog/reorder
 * expects. Every fixed category tab's value already matches a real
 * `category` string; the "company" tab's value is the literal string
 * "company", not the viewer's real domain, so it needs translating.
 */
function scopeForActiveCategory(): { category: string } | { companyDomain: string | null } {
  if (activeCategory === "company") return { companyDomain: myCompanyDomain };
  return { category: activeCategory };
}

/**
 * Handles the "+ Add new group…" option in #editGroup — mirrors /admin's
 * own inline group-create flow (server.js) but against the gallery's own
 * cached group list instead of a server-rendered <select>. Only new-group
 * *creation* lives here; renaming/reordering/deleting groups stays
 * /admin-only, same as category reassignment and thumbnail replacement.
 */
async function handleEditGroupChange(): Promise<void> {
  const groupSelect = document.getElementById("editGroup") as HTMLSelectElement | null;
  if (!groupSelect || groupSelect.value !== "__new__") {
    if (groupSelect) editGroupPreviousValue = groupSelect.value;
    return;
  }

  const name = window.prompt("New group name:")?.trim();
  if (!name) {
    groupSelect.value = editGroupPreviousValue;
    return;
  }

  const data = cache.get(activeCategory);
  const sortOrder = data?.groups.length ?? 0;
  try {
    const res = await fetch("/admin/groups", {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({ ...scopeForActiveCategory(), name, sortOrder }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);

    if (data) data.groups.push({ id: body.id, name, sortOrder });
    const opt = document.createElement("option");
    opt.value = String(body.id);
    opt.textContent = name;
    groupSelect.insertBefore(opt, groupSelect.querySelector('option[value="__new__"]'));
    groupSelect.value = String(body.id);
    editGroupPreviousValue = groupSelect.value;
  } catch (err) {
    window.alert(`Couldn't create group: ${err instanceof Error ? err.message : String(err)}`);
    groupSelect.value = editGroupPreviousValue;
  }
}

/**
 * The dialog's only way to communicate outward — see the module comment.
 * { action, item } instead of a bare item (Task Pane Phase 12) so the task
 * pane's DialogMessageReceived handler can tell an ordinary insert apart
 * from an edit/delete request. Sends the full item either way, not just
 * its id, so the task pane never needs a separate lookup/refetch.
 */
function insertItem(item: CatalogItem): void {
  Office.context.ui.messageParent(JSON.stringify({ action: "insert", item }));
}

/** The direct route for 'file'-mode items — see #btnInsertAsSlide's title text and insertFileItemAsNewSlide (libraryInsert.ts) for what this skips. */
function insertItemAsSlide(item: CatalogItem): void {
  Office.context.ui.messageParent(JSON.stringify({ action: "insert-as-slide", item }));
}

/**
 * Two-click arm/confirm (see resetDeleteArm's comment for why, not
 * window.confirm). Deletion happens via a direct fetch from right here —
 * not a messageParent round-trip to the task pane like insert/edit are —
 * the same "no Office.js involved, so no need to leave the dialog" pattern
 * saveEditDetails above already uses. Routing this through the task pane
 * instead (as an earlier version of this function did) meant the task
 * pane's shared DialogMessageReceived handler closed the dialog
 * unconditionally for every action, which made the whole gallery vanish
 * after a delete — not the expected "stay open, item's gone from the
 * grid" behavior a delete should have.
 */
async function deleteItem(item: CatalogItem): Promise<void> {
  const btn = document.getElementById("btnDelete");
  if (!deleteArmed) {
    deleteArmed = true;
    if (btn) btn.textContent = DELETE_CONFIRM_LABEL;
    window.clearTimeout(deleteArmedTimer);
    deleteArmedTimer = window.setTimeout(resetDeleteArm, 4000);
    return;
  }
  resetDeleteArm();

  const errorEl = document.getElementById("previewError");
  const isPersonal = !!item.ownerOid;
  const url = isPersonal ? `/api/personal/catalog/${item.id}/delete` : `/admin/catalog/${item.id}/delete`;
  try {
    const res = await fetch(url, { method: "POST", headers: { Accept: "application/json" } });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    const data = cache.get(activeCategory);
    if (data) data.items = data.items.filter((i) => i.id !== item.id);
    hidePreview();
    rerenderGrid();
  } catch (err) {
    if (errorEl) {
      errorEl.textContent = err instanceof Error ? err.message : String(err);
      errorEl.style.display = "block";
    }
  }
}

function editItem(item: CatalogItem): void {
  Office.context.ui.messageParent(JSON.stringify({ action: "edit", item }));
}

function selectItem(item: CatalogItem): void {
  selectedItem = item;
  showPreview(item);
  rerenderGrid();
}

function matchesSearch(item: CatalogItem, lowerFilter: string): boolean {
  if (!lowerFilter) return true;
  if (item.title.toLowerCase().includes(lowerFilter)) return true;
  return item.tags.some((t) => t.toLowerCase().includes(lowerFilter));
}

function renderItemCard(item: CatalogItem): HTMLElement {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "gallery-item" + (selectedItem?.id === item.id ? " selected" : "");
  card.dataset.itemId = String(item.id);

  // Drag-to-reorder only applies where groups exist at all — "My Items"
  // has no group concept (Task Pane Phase 13's own design) and no reorder
  // endpoint to call.
  if (activeCategory !== "personal" && canEdit(item)) {
    card.draggable = true;
    card.addEventListener("dragstart", (e) => {
      draggedCard = card;
      e.dataTransfer?.setData("text/plain", String(item.id));
      card.classList.add("dragging");
    });
    card.addEventListener("dragend", () => {
      draggedCard?.classList.remove("dragging");
      draggedCard = null;
    });
  }

  if (item.unicodeChar) {
    const glyph = document.createElement("span");
    glyph.className = "gallery-item-glyph";
    glyph.textContent = item.unicodeChar;
    card.appendChild(glyph);
  } else if (item.thumbnailUrl) {
    const img = document.createElement("img");
    img.src = item.thumbnailUrl;
    img.alt = item.title;
    card.appendChild(img);
  } else {
    const placeholder = document.createElement("span");
    placeholder.className = "gallery-item-placeholder";
    card.appendChild(placeholder);
  }

  const label = document.createElement("span");
  label.textContent = item.title;
  card.appendChild(label);

  // Single-click selects + shows the preview panel; double-click inserts
  // immediately (the click handler still runs first — harmless, it just
  // re-selects the same item before dblclick fires the actual insert).
  card.addEventListener("click", () => selectItem(item));
  card.addEventListener("dblclick", () => insertItem(item));
  return card;
}

/** Nearest-card-center heuristic, matching /admin's own getDragAfterElement (server.js) — good enough for a small drop target, not a strict row/column layout solve. */
function getDragAfterElement(container: HTMLElement, clientX: number, clientY: number): Element | null {
  const cards = Array.from(container.querySelectorAll(".gallery-item:not(.dragging)"));
  let closest: { card: Element; after: boolean } | null = null;
  let closestDistance = Infinity;
  for (const card of cards) {
    const box = card.getBoundingClientRect();
    const dx = clientX - (box.left + box.width / 2);
    const dy = clientY - (box.top + box.height / 2);
    const distance = dx * dx + dy * dy;
    if (distance < closestDistance) {
      closestDistance = distance;
      closest = { card, after: dx > 0 };
    }
  }
  if (!closest) return null;
  return closest.after ? closest.card.nextElementSibling : closest.card;
}

function persistReorder(row: HTMLElement): void {
  const groupIdRaw = row.dataset.groupId ?? "";
  const groupId = groupIdRaw ? Number(groupIdRaw) : null;
  const orderedIds = Array.from(row.querySelectorAll<HTMLElement>(".gallery-item")).map((c) => Number(c.dataset.itemId));
  const status = statusEl();

  fetch("/admin/catalog/reorder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...scopeForActiveCategory(), groupId, orderedIds }),
  })
    .then((res) => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = cache.get(activeCategory);
      if (!data) return;
      const idToItem = new Map(data.items.map((i) => [i.id, i]));
      const movedItems = orderedIds.map((id) => idToItem.get(id)).filter((i): i is CatalogItem => !!i);
      for (const it of movedItems) it.groupId = groupId;
      const group = data.groups.find((g) => g.id === groupId);
      for (const it of movedItems) it.groupName = group?.name ?? null;
      const firstIndex = data.items.findIndex((i) => orderedIds.includes(i.id));
      const remaining = data.items.filter((i) => !orderedIds.includes(i.id));
      remaining.splice(firstIndex, 0, ...movedItems);
      data.items = remaining;
      if (status) status.textContent = "";
    })
    .catch((err) => {
      if (status) status.textContent = `Couldn't save order: ${err instanceof Error ? err.message : String(err)}`;
    });
}

function renderGroupSection(container: HTMLElement, heading: string | null, groupId: number | null, items: CatalogItem[]): void {
  if (heading) {
    const h = document.createElement("h3");
    h.className = "gallery-group-heading";
    h.textContent = heading;
    container.appendChild(h);
  }
  const row = document.createElement("div");
  row.className = "gallery-item-row";
  row.dataset.groupId = groupId !== null ? String(groupId) : "";
  for (const item of items) {
    row.appendChild(renderItemCard(item));
  }
  if (activeCategory !== "personal") {
    row.addEventListener("dragover", (e) => {
      if (!draggedCard) return;
      e.preventDefault();
      const afterElement = getDragAfterElement(row, e.clientX, e.clientY);
      row.insertBefore(draggedCard, afterElement);
    });
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      if (!draggedCard) return;
      const finalRow = draggedCard.closest(".gallery-item-row") as HTMLElement | null;
      if (finalRow) persistReorder(finalRow);
    });
  }
  container.appendChild(row);
}

function renderGrid(data: CatalogResponse, filter: string): void {
  const grid = document.getElementById("galleryGrid");
  if (!grid) return;
  grid.innerHTML = "";

  const lowerFilter = filter.trim().toLowerCase();
  const sortedGroups = [...data.groups].sort((a, b) => a.sortOrder - b.sortOrder);

  // Buckets items by groupId (null = ungrouped) preserving each group's
  // own item order (already sort_order-sorted server-side).
  const grouped = new Map<number | null, CatalogItem[]>();
  for (const item of data.items) {
    if (!matchesSearch(item, lowerFilter)) continue;
    const key = item.groupId;
    const bucket = grouped.get(key);
    if (bucket) bucket.push(item);
    else grouped.set(key, [item]);
  }

  for (const group of sortedGroups) {
    const items = grouped.get(group.id);
    if (items && items.length > 0) renderGroupSection(grid, group.name, group.id, items);
  }
  const ungrouped = grouped.get(null);
  if (ungrouped && ungrouped.length > 0) {
    // Only label it "Other" if there's at least one real group to
    // distinguish it from — with no groups at all in this category, every
    // item is ungrouped and a heading would just be noise.
    renderGroupSection(grid, sortedGroups.length > 0 ? "Other" : null, null, ungrouped);
  }

  if (grid.children.length === 0) {
    const empty = document.createElement("p");
    empty.className = "gallery-empty";
    empty.textContent = "No items match your search.";
    grid.appendChild(empty);
  }
}

function rerenderGrid(): void {
  const data = cache.get(activeCategory);
  if (data) renderGrid(data, searchInput()?.value ?? "");
}

async function loadCategory(category: string): Promise<void> {
  const status = statusEl();
  if (!cache.has(category)) {
    if (status) status.textContent = "Loading…";
    try {
      cache.set(category, await fetchCatalog(category));
    } catch (err) {
      if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      return;
    }
  }
  if (status) status.textContent = "";
  rerenderGrid();
}

function switchCategory(category: string): void {
  if (category === activeCategory && cache.has(category)) return;
  activeCategory = category;
  hidePreview();
  renderTabs();
  loadCategory(category).catch((err) => {
    const status = statusEl();
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  });
}

Office.onReady(async () => {
  // Same-origin, session-cookie-authenticated fetch — confirmed working
  // from this dialog already (see the module comment). Awaited *before*
  // the first render (Task Pane Phase 13) since the tab list now depends
  // on it, not just the Edit button. A failed/timed-out check just means
  // "not signed in", same fail-safe default the task pane's own checks use.
  try {
    const res = await fetch("/api/auth/me");
    if (res.ok) {
      const user = await res.json();
      isAdmin = !!user?.isAdmin;
      myOid = user?.oid ?? null;
      myTid = user?.tid ?? null;
      myCompanyDomain = user?.companyDomain ?? null;
      isCompanyAdmin = !!user?.isCompanyAdmin;
      // Company tab before "My Items" — the company library is a shared,
      // browsable-by-anyone-at-that-company surface (Task Pane Phase 14),
      // closer in spirit to the fixed category tabs than to the strictly
      // personal one that follows it.
      if (myCompanyDomain) CATEGORIES = [...CATEGORIES, { value: "company", label: myCompanyDomain }];
      CATEGORIES = [...CATEGORIES, { value: "personal", label: "My Items" }];
    }
  } catch {
    // not signed in — leave CATEGORIES/isAdmin at their defaults
  }

  renderTabs();
  loadCategory(activeCategory).catch((err) => {
    const status = statusEl();
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  });

  searchInput()?.addEventListener("input", rerenderGrid);

  document.getElementById("btnInsert")?.addEventListener("click", () => {
    if (selectedItem) insertItem(selectedItem);
  });
  document.getElementById("btnInsertAsSlide")?.addEventListener("click", () => {
    if (selectedItem) insertItemAsSlide(selectedItem);
  });
  document.getElementById("btnEdit")?.addEventListener("click", () => {
    if (selectedItem) editItem(selectedItem);
  });
  document.getElementById("btnDelete")?.addEventListener("click", () => {
    if (selectedItem) deleteItem(selectedItem);
  });
  document.getElementById("btnSaveDetails")?.addEventListener("click", () => {
    if (selectedItem) saveEditDetails(selectedItem).catch(() => {});
  });
  document.getElementById("editGroup")?.addEventListener("change", () => {
    handleEditGroupChange().catch(() => {});
  });
});
