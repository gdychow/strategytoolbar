/**
 * Task Pane Phase 20 — the Template Library dialog. Mirrors gallery.ts's
 * own displayDialogAsync/messageParent constraints (a dialog can call
 * exactly messageParent and requirements.isSetSupported — confirmed back
 * in Phase 5's research), but browsing/uploading/renaming/deleting all
 * happen via plain authenticated fetch calls directly from here (no
 * Office.js involved in any of that) — only "Use This Template" needs to
 * cross back into the task pane, since PowerPoint.createPresentation is a
 * PowerPoint-specific API a dialog can't call itself.
 */
import * as Templates from "../features/templates";
import type { TemplateItem, TemplateScope } from "../features/templates";

interface TabDef {
  scope: TemplateScope;
  label: string;
}

let tabs: TabDef[] = [{ scope: "global", label: "Global" }];
const cache = new Map<TemplateScope, TemplateItem[]>();
let activeScope: TemplateScope = "global";
let selectedItem: TemplateItem | null = null;

let isAdmin = false;
let isCompanyAdmin = false;
let myCompanyDomain: string | null = null;

/**
 * window.confirm() doesn't reliably produce a visible native dialog in
 * every Office-hosted webview this app runs in (this is the same class of
 * gap already worked around in the task pane itself, via its own arm/
 * confirm pattern on #btnLibraryAdd) — confirmed here directly: clicking
 * Delete with a plain window.confirm() call silently did nothing, no
 * dialog, no error, no deletion. A two-click "click again to confirm" on
 * the button itself has no such dependency on the host's dialog support.
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

function statusEl(): HTMLElement | null {
  return document.getElementById("status");
}

/** Whether the signed-in viewer can upload/rename/delete in this scope — personal is always manageable (it's the viewer's own), company/global mirror the same admin gating used everywhere else in this app. */
function canManageScope(scope: TemplateScope): boolean {
  if (scope === "personal") return true;
  if (scope === "company") return isCompanyAdmin || isAdmin;
  return isAdmin;
}

function renderTabs(): void {
  const tabsEl = document.getElementById("tabs");
  if (!tabsEl) return;
  tabsEl.innerHTML = "";
  for (const tab of tabs) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "template-tab" + (tab.scope === activeScope ? " active" : "");
    btn.textContent = tab.label;
    btn.addEventListener("click", () => switchScope(tab.scope));
    tabsEl.appendChild(btn);
  }
}

function hidePreview(): void {
  selectedItem = null;
  resetDeleteArm();
  const panel = document.getElementById("previewPanel");
  if (panel) panel.style.display = "none";
}

function showPreview(item: TemplateItem): void {
  const panel = document.getElementById("previewPanel");
  const title = document.getElementById("previewTitle");
  const manageRow = document.getElementById("manageRow");
  const editTitle = document.getElementById("editTitle") as HTMLInputElement | null;
  const editDescription = document.getElementById("editDescription") as HTMLTextAreaElement | null;
  const errorEl = document.getElementById("previewError");
  if (!panel || !title || !manageRow || !editTitle || !editDescription) return;

  resetDeleteArm(); // switching selection shouldn't leave a stale "click again to confirm" pointed at the wrong item
  title.textContent = item.title;
  if (errorEl) {
    errorEl.style.display = "none";
    errorEl.textContent = "";
  }

  const manageable = canManageScope(activeScope);
  manageRow.style.display = manageable ? "flex" : "none";
  if (manageable) {
    editTitle.value = item.title;
    editDescription.value = item.description ?? "";
  }
  panel.style.display = "flex";
}

function selectItem(item: TemplateItem): void {
  selectedItem = item;
  showPreview(item);
  rerenderGrid();
}

function renderCard(item: TemplateItem): HTMLElement {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "template-card" + (selectedItem?.id === item.id ? " selected" : "");
  const icon = document.createElement("span");
  icon.className = "template-card-icon";
  icon.textContent = "\u{1F4C4}"; // generic document glyph — no thumbnails this phase
  card.appendChild(icon);
  const label = document.createElement("span");
  label.textContent = item.title;
  card.appendChild(label);
  card.addEventListener("click", () => selectItem(item));
  return card;
}

function rerenderGrid(): void {
  const grid = document.getElementById("templateGrid");
  if (!grid) return;
  grid.innerHTML = "";
  const items = cache.get(activeScope) ?? [];
  if (items.length === 0) {
    const empty = document.createElement("p");
    empty.className = "template-empty";
    empty.textContent = "No templates here yet.";
    grid.appendChild(empty);
    return;
  }
  for (const item of items) grid.appendChild(renderCard(item));
}

function updateUploadRowVisibility(): void {
  const row = document.getElementById("uploadRow");
  if (row) row.style.display = canManageScope(activeScope) ? "flex" : "none";
}

async function loadScope(scope: TemplateScope): Promise<void> {
  const status = statusEl();
  if (!cache.has(scope)) {
    if (status) status.textContent = "Loading…";
    try {
      cache.set(scope, await Templates.fetchTemplates(scope));
    } catch (err) {
      if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      return;
    }
  }
  if (status) status.textContent = "";
  rerenderGrid();
}

function switchScope(scope: TemplateScope): void {
  if (scope === activeScope && cache.has(scope)) return;
  activeScope = scope;
  hidePreview();
  renderTabs();
  updateUploadRowVisibility();
  loadScope(scope).catch((err) => {
    const status = statusEl();
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  });
}

/** The dialog's only way out for this action — see the module comment. */
function useTemplate(item: TemplateItem): void {
  Office.context.ui.messageParent(JSON.stringify({ action: "use-template", templateId: item.id }));
}

async function saveEdit(): Promise<void> {
  if (!selectedItem) return;
  const editTitle = document.getElementById("editTitle") as HTMLInputElement | null;
  const editDescription = document.getElementById("editDescription") as HTMLTextAreaElement | null;
  const errorEl = document.getElementById("previewError");
  const title = editTitle?.value.trim() ?? "";
  if (!title) {
    if (errorEl) {
      errorEl.textContent = "Title can't be empty.";
      errorEl.style.display = "block";
    }
    return;
  }
  try {
    await Templates.renameTemplate(selectedItem.id, title, editDescription?.value ?? "");
    selectedItem.title = title;
    selectedItem.description = editDescription?.value || null;
    showPreview(selectedItem);
    rerenderGrid();
  } catch (err) {
    if (errorEl) {
      errorEl.textContent = err instanceof Error ? err.message : String(err);
      errorEl.style.display = "block";
    }
  }
}

async function deleteSelected(): Promise<void> {
  if (!selectedItem) return;
  const btn = document.getElementById("btnDelete");
  if (!deleteArmed) {
    deleteArmed = true;
    if (btn) btn.textContent = DELETE_CONFIRM_LABEL;
    window.clearTimeout(deleteArmedTimer);
    deleteArmedTimer = window.setTimeout(resetDeleteArm, 4000);
    return;
  }
  resetDeleteArm();

  const id = selectedItem.id;
  try {
    await Templates.deleteTemplate(id);
    const items = cache.get(activeScope);
    if (items) cache.set(activeScope, items.filter((i) => i.id !== id));
    hidePreview();
    rerenderGrid();
  } catch (err) {
    const errorEl = document.getElementById("previewError");
    if (errorEl) {
      errorEl.textContent = err instanceof Error ? err.message : String(err);
      errorEl.style.display = "block";
    }
  }
}

/** Title comes straight from the filename — there's no separate naming step before upload, only rename-after via the preview panel's existing Save action. */
async function handleUpload(file: File): Promise<void> {
  const status = statusEl();
  const title = file.name.replace(/\.potx$/i, "");
  if (status) status.textContent = "Uploading…";
  try {
    const created = await Templates.uploadTemplate(activeScope, file, title);
    const items = cache.get(activeScope) ?? [];
    items.push(created);
    cache.set(activeScope, items);
    rerenderGrid();
    if (status) status.textContent = "";
  } catch (err) {
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  }
}

Office.onReady(async () => {
  try {
    const res = await fetch("/api/auth/me");
    if (res.ok) {
      const user = await res.json();
      isAdmin = !!user?.isAdmin;
      isCompanyAdmin = !!user?.isCompanyAdmin;
      myCompanyDomain = user?.companyDomain ?? null;
      if (myCompanyDomain) tabs = [...tabs, { scope: "company", label: myCompanyDomain }];
      tabs = [...tabs, { scope: "personal", label: "My Templates" }];
    }
  } catch {
    // not signed in — leave tabs at just "Global"
  }

  renderTabs();
  updateUploadRowVisibility();
  loadScope(activeScope).catch((err) => {
    const status = statusEl();
    if (status) status.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
  });

  document.getElementById("btnUse")?.addEventListener("click", () => {
    if (selectedItem) useTemplate(selectedItem);
  });
  document.getElementById("btnSave")?.addEventListener("click", () => {
    saveEdit().catch(() => {});
  });
  document.getElementById("btnDelete")?.addEventListener("click", () => {
    deleteSelected().catch(() => {});
  });
  // Single button — clicking it just opens the native file picker; the
  // upload itself fires immediately once a file is chosen, no separate
  // "Upload" click needed.
  document.getElementById("btnUpload")?.addEventListener("click", () => {
    (document.getElementById("uploadFile") as HTMLInputElement | null)?.click();
  });
  document.getElementById("uploadFile")?.addEventListener("change", () => {
    const fileInput = document.getElementById("uploadFile") as HTMLInputElement | null;
    const file = fileInput?.files?.[0];
    if (!file) return;
    // The picker's accept=".potx" is only a hint — different OS/browser
    // combinations enforce it differently (some show an "All Files"
    // toggle regardless), so this is a real check, not decoration. The
    // server re-validates by mimetype + extension regardless; this just
    // gives instant feedback instead of a round trip for an obviously
    // wrong file.
    if (!/\.potx$/i.test(file.name)) {
      const status = statusEl();
      if (status) status.textContent = `"${file.name}" isn't a .potx file.`;
      if (fileInput) fileInput.value = "";
      return;
    }
    handleUpload(file)
      .catch(() => {})
      .finally(() => {
        if (fileInput) fileInput.value = ""; // reset so picking the same file again still fires "change"
      });
  });
});
