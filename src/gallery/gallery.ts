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

function hidePreview(): void {
  selectedItem = null;
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

function showPreview(item: CatalogItem): void {
  const panel = document.getElementById("previewPanel");
  const img = document.getElementById("previewImg") as HTMLImageElement | null;
  const title = document.getElementById("previewTitle");
  const editBtn = document.getElementById("btnEdit");
  const editDetailsBtn = document.getElementById("btnEditDetails");
  const deleteBtn = document.getElementById("btnDelete");
  if (!panel || !img || !title) return;
  closeEditDetails();
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
  title.textContent = item.title;
  panel.style.display = "flex";
  const editable = canEdit(item);
  if (editBtn) (editBtn as HTMLElement).style.display = editable ? "" : "none";
  if (editDetailsBtn) (editDetailsBtn as HTMLElement).style.display = editable ? "" : "none";
  if (deleteBtn) (deleteBtn as HTMLElement).style.display = editable ? "" : "none";
}

/**
 * Task Pane Phase 16: lightweight title/tags/group quick-edit, so a
 * company/global admin doesn't have to leave the gallery and open /admin
 * just to fix a typo or move an item between groups. Category reassignment,
 * thumbnail replacement, and new-group creation stay /admin-only — this is
 * deliberately the small subset worth editing without leaving the gallery.
 */
function openEditDetails(item: CatalogItem): void {
  const form = document.getElementById("editDetailsForm") as HTMLFormElement | null;
  const titleInput = document.getElementById("editTitle") as HTMLInputElement | null;
  const scopedFields = document.getElementById("editScopedFields");
  const tagsInput = document.getElementById("editTags") as HTMLInputElement | null;
  const groupSelect = document.getElementById("editGroup") as HTMLSelectElement | null;
  const errorEl = document.getElementById("editError");
  if (!form || !titleInput || !scopedFields || !tagsInput || !groupSelect) return;

  if (errorEl) {
    errorEl.style.display = "none";
    errorEl.textContent = "";
  }
  titleInput.value = item.title;

  const isPersonal = !!item.ownerOid;
  scopedFields.style.display = isPersonal ? "none" : "";
  if (!isPersonal) {
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

  form.style.display = "";
  titleInput.focus();
}

function closeEditDetails(): void {
  const form = document.getElementById("editDetailsForm") as HTMLFormElement | null;
  const errorEl = document.getElementById("editError");
  if (form) form.style.display = "none";
  if (errorEl) {
    errorEl.style.display = "none";
    errorEl.textContent = "";
  }
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
    closeEditDetails();
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

/**
 * Deletion asks for confirmation right here via window.confirm — unlike
 * the task pane's own embedded webview, this dialog is an ordinary browser
 * context where window.confirm works fine (confirmed directly this
 * session while building the color picker/admin-add flows), so there's no
 * need for the task pane's two-step arm/confirm workaround for this action.
 */
function deleteItem(item: CatalogItem): void {
  if (!window.confirm(`Delete "${item.title}" from the library? This can't be undone.`)) return;
  Office.context.ui.messageParent(JSON.stringify({ action: "delete", item }));
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
  document.getElementById("btnEdit")?.addEventListener("click", () => {
    if (selectedItem) editItem(selectedItem);
  });
  document.getElementById("btnDelete")?.addEventListener("click", () => {
    if (selectedItem) deleteItem(selectedItem);
  });
  document.getElementById("btnEditDetails")?.addEventListener("click", () => {
    if (selectedItem) openEditDetails(selectedItem);
  });
  document.getElementById("btnEditCancel")?.addEventListener("click", () => closeEditDetails());
  document.getElementById("editDetailsForm")?.addEventListener("submit", (e) => {
    e.preventDefault();
    if (selectedItem) saveEditDetails(selectedItem).catch(() => {});
  });
  document.getElementById("editGroup")?.addEventListener("change", () => {
    handleEditGroupChange().catch(() => {});
  });
});
