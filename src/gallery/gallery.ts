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

// Matches what's actually seeded (see db/seed/catalog-*.json) — Symbols
// isn't built yet (needs its own insert_mode, planned separately), so
// it's deliberately not listed here.
const CATEGORIES: { value: string; label: string }[] = [
  { value: "text", label: "Text" },
  { value: "objects", label: "Objects" },
  { value: "shapes", label: "Shapes" },
  { value: "stamps", label: "Stamps" },
  { value: "tables", label: "Tables" },
  { value: "diagrams", label: "Diagrams" },
];

const cache = new Map<string, CatalogResponse>();
let activeCategory = CATEGORIES[0].value;
let selectedItem: CatalogItem | null = null;

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

function showPreview(item: CatalogItem): void {
  const panel = document.getElementById("previewPanel");
  const img = document.getElementById("previewImg") as HTMLImageElement | null;
  const title = document.getElementById("previewTitle");
  if (!panel || !img || !title) return;
  if (item.thumbnailUrl) {
    img.src = item.thumbnailUrl;
    img.style.display = "";
  } else {
    img.style.display = "none";
  }
  title.textContent = item.title;
  panel.style.display = "flex";
}

/**
 * The dialog's only way to communicate outward — see the module comment.
 * Sends the full item, not just its id, so the task pane can call
 * Library.insertCatalogItem directly with no separate lookup/refetch —
 * it inserts exactly what the user saw and clicked here.
 */
function insertItem(item: CatalogItem): void {
  Office.context.ui.messageParent(JSON.stringify(item));
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

  if (item.thumbnailUrl) {
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

function renderGroupSection(container: HTMLElement, heading: string | null, items: CatalogItem[]): void {
  if (heading) {
    const h = document.createElement("h3");
    h.className = "gallery-group-heading";
    h.textContent = heading;
    container.appendChild(h);
  }
  const row = document.createElement("div");
  row.className = "gallery-item-row";
  for (const item of items) {
    row.appendChild(renderItemCard(item));
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
    if (items && items.length > 0) renderGroupSection(grid, group.name, items);
  }
  const ungrouped = grouped.get(null);
  if (ungrouped && ungrouped.length > 0) {
    // Only label it "Other" if there's at least one real group to
    // distinguish it from — with no groups at all in this category, every
    // item is ungrouped and a heading would just be noise.
    renderGroupSection(grid, sortedGroups.length > 0 ? "Other" : null, ungrouped);
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

Office.onReady(() => {
  renderTabs();
  loadCategory(activeCategory).catch((err) => {
    const status = statusEl();
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  });

  searchInput()?.addEventListener("input", rerenderGrid);

  document.getElementById("btnInsert")?.addEventListener("click", () => {
    if (selectedItem) insertItem(selectedItem);
  });
});
