/**
 * Library Upload — bulk-import multiple .pptx files from /admin (or the
 * task pane's own "Bulk Upload…" button). Mirrors gallery.ts/templates.ts's
 * displayDialogAsync/messageParent constraints (a dialog can call exactly
 * messageParent and requirements.isSetSupported), but this dialog never
 * needs to cross back into the task pane mid-flow the way gallery's Insert
 * or templates' "Use This Template" do — the actual catalog writes happen
 * server-side on commit, no PowerPoint API involved at all. The only
 * message it ever sends is "done", once, so the task pane knows to close
 * the dialog.
 *
 * Three server round trips: POST to start a background conversion job
 * (render-sidecar/app.py does the real work), poll GET for progress, POST
 * commit with whatever the admin edited/excluded on the review screen.
 */

type Scope = "global" | "company";

const CATEGORIES: { value: string; label: string }[] = [
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

interface FileEntry {
  file: File;
  category: string;
}

interface JobFileStatus {
  filename: string;
  status: "pending" | "converting" | "done" | "error";
  slideCount?: number;
  error?: string;
}

interface JobItem {
  tempId: string;
  sourceFilename: string;
  slideIndex: number;
  category: string | null;
  title: string;
  tags: string[];
  included: boolean;
  insertMode: "reconstruct" | "file";
  thumbnailUrl: string;
}

interface ReviewItem extends JobItem {
  tagsRaw: string;
}

let isAdmin = false;
let isCompanyAdmin = false;
let companyDomain: string | null = null;
let scope: Scope = "global";

let chosenFiles: FileEntry[] = [];
let jobId: string | null = null;
let pollTimer: number | undefined;
let reviewItems: ReviewItem[] = [];

function showScreen(id: "setupScreen" | "progressScreen" | "reviewScreen" | "doneScreen"): void {
  for (const s of ["setupScreen", "progressScreen", "reviewScreen", "doneScreen"]) {
    const el = document.getElementById(s);
    if (el) (el as HTMLElement).style.display = s === id ? "" : "none";
  }
}

function setSetupError(message: string): void {
  const el = document.getElementById("setupError");
  if (!el) return;
  el.textContent = message;
  (el as HTMLElement).style.display = message ? "block" : "none";
}

function canUseGlobal(): boolean {
  return isAdmin;
}
function canUseCompany(): boolean {
  return isCompanyAdmin && !!companyDomain;
}

function renderFileList(): void {
  const list = document.getElementById("fileList");
  if (!list) return;
  list.innerHTML = "";
  chosenFiles.forEach((entry, index) => {
    const row = document.createElement("div");
    row.className = "upload-file-row";

    if (scope === "global") {
      const select = document.createElement("select");
      const blank = document.createElement("option");
      blank.value = "";
      blank.textContent = "Category…";
      if (entry.category === "") blank.selected = true;
      select.appendChild(blank);
      for (const cat of CATEGORIES) {
        const opt = document.createElement("option");
        opt.value = cat.value;
        opt.textContent = cat.label;
        if (cat.value === entry.category) opt.selected = true;
        select.appendChild(opt);
      }
      select.addEventListener("change", () => {
        entry.category = select.value;
        updateStartButton();
      });
      row.appendChild(select);
    }

    const name = document.createElement("span");
    name.className = "upload-file-name";
    name.textContent = entry.file.name;
    row.appendChild(name);

    // Absorbs the row's remaining width so filename+category sit together
    // on the left (immediately next to each other, since it's clear which
    // category belongs to which file) while the remove button still lands
    // on the right edge like every other row of controls in this dialog.
    const spacer = document.createElement("span");
    spacer.className = "upload-file-spacer";
    row.appendChild(spacer);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "upload-remove-btn";
    remove.textContent = "×";
    remove.title = "Remove";
    remove.addEventListener("click", () => {
      chosenFiles.splice(index, 1);
      renderFileList();
      updateStartButton();
    });
    row.appendChild(remove);

    list.appendChild(row);
  });
}

/** Global uploads need every file categorized before converting; company uploads have no per-file category picker at all, so only the file count matters there. */
function updateStartButton(): void {
  const btn = document.getElementById("btnStartUpload") as HTMLButtonElement | null;
  if (!btn) return;
  const allCategorized = scope !== "global" || chosenFiles.every((e) => e.category !== "");
  btn.disabled = chosenFiles.length === 0 || !allCategorized;
}

function addFiles(files: FileList): void {
  const rejected: string[] = [];
  for (const file of Array.from(files)) {
    if (!/\.pptx$/i.test(file.name)) {
      rejected.push(file.name);
      continue;
    }
    chosenFiles.push({ file, category: "" });
  }
  if (rejected.length) {
    setSetupError(`Skipped non-.pptx file(s): ${rejected.join(", ")}`);
  } else {
    setSetupError("");
  }
  renderFileList();
  updateStartButton();
}

async function startUpload(): Promise<void> {
  setSetupError("");
  const form = new FormData();
  for (const entry of chosenFiles) form.append("files", entry.file, entry.file.name);
  form.append("scope", scope);
  if (scope === "global") {
    form.append("categories", JSON.stringify(chosenFiles.map((e) => e.category)));
  }

  const btn = document.getElementById("btnStartUpload") as HTMLButtonElement | null;
  if (btn) btn.disabled = true;
  try {
    const res = await fetch("/api/admin/library-upload", { method: "POST", body: form });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    jobId = body.jobId;
    showScreen("progressScreen");
    pollJob();
  } catch (err) {
    setSetupError(err instanceof Error ? err.message : String(err));
    if (btn) btn.disabled = false;
  }
}

function renderFileProgress(files: JobFileStatus[]): void {
  const list = document.getElementById("fileProgressList");
  if (!list) return;
  list.innerHTML = "";
  for (const f of files) {
    const row = document.createElement("div");
    row.className = "upload-file-row";
    const name = document.createElement("span");
    name.className = "upload-file-name";
    name.textContent = f.filename;
    row.appendChild(name);
    const status = document.createElement("span");
    status.className = "upload-file-status" + (f.status === "error" ? " upload-file-status-error" : "");
    status.textContent =
      f.status === "pending"
        ? "Waiting…"
        : f.status === "converting"
          ? "Converting…"
          : f.status === "error"
            ? f.error || "Failed"
            : `${f.slideCount ?? 0} slide(s)`;
    row.appendChild(status);
    list.appendChild(row);
  }
}

function pollJob(): void {
  if (pollTimer !== undefined) window.clearInterval(pollTimer);
  pollTimer = window.setInterval(async () => {
    if (!jobId) return;
    try {
      const res = await fetch(`/api/admin/library-upload/${jobId}`);
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);

      renderFileProgress(body.files);
      const done = body.files.filter((f: JobFileStatus) => f.status === "done" || f.status === "error").length;
      const statusEl = document.getElementById("progressStatus");
      if (statusEl) statusEl.textContent = `Converting ${done} of ${body.files.length} file(s)…`;

      if (body.status === "error") {
        window.clearInterval(pollTimer);
        if (statusEl) statusEl.textContent = `Error: ${body.error || "Something went wrong."}`;
        const back = document.getElementById("btnProgressBack");
        if (back) (back as HTMLElement).style.display = "";
      } else if (body.status === "done") {
        window.clearInterval(pollTimer);
        enterReview(body.items);
      }
    } catch (err) {
      window.clearInterval(pollTimer);
      const statusEl = document.getElementById("progressStatus");
      if (statusEl) statusEl.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      const back = document.getElementById("btnProgressBack");
      if (back) (back as HTMLElement).style.display = "";
    }
  }, 1200);
}

function enterReview(items: JobItem[]): void {
  reviewItems = items.map((item) => ({ ...item, tagsRaw: "" }));
  if (reviewItems.length === 0) {
    showScreen("doneScreen");
    const status = document.getElementById("doneStatus");
    if (status) status.textContent = "No slides were extracted from the file(s) you uploaded.";
    return;
  }
  const summary = document.getElementById("reviewSummary");
  const fileCount = new Set(reviewItems.map((i) => i.sourceFilename)).size;
  if (summary) {
    summary.textContent = `${reviewItems.length} slide(s) from ${fileCount} file(s). Uncheck any you don't want to add.`;
  }
  renderReview();
  showScreen("reviewScreen");
}

function renderReview(): void {
  const list = document.getElementById("reviewList");
  if (!list) return;
  list.innerHTML = "";

  const groups = new Map<string, ReviewItem[]>();
  for (const item of reviewItems) {
    const group = groups.get(item.sourceFilename) ?? [];
    group.push(item);
    groups.set(item.sourceFilename, group);
  }

  for (const [filename, items] of groups) {
    const heading = document.createElement("h3");
    const categoryLabel = items[0].category
      ? CATEGORIES.find((c) => c.value === items[0].category)?.label ?? items[0].category
      : companyDomain;
    heading.textContent = `${filename} — ${categoryLabel}`;
    heading.className = "upload-review-group-heading";
    list.appendChild(heading);

    for (const item of items) {
      const row = document.createElement("div");
      row.className = "upload-review-row";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = item.included;
      checkbox.addEventListener("change", () => {
        item.included = checkbox.checked;
        row.classList.toggle("upload-review-row-excluded", !checkbox.checked);
      });
      row.appendChild(checkbox);

      const thumb = document.createElement("img");
      thumb.className = "upload-review-thumb";
      thumb.src = item.thumbnailUrl;
      thumb.alt = "";
      row.appendChild(thumb);

      // Auto-selected per slide by the render sidecar (same
      // classify_shape_tree logic scripts/slice-catalog-source.py uses):
      // "Native" shapes/text rebuild with one click; anything else (a
      // picture, table, or custom-drawn shape) needs the temp-slide/copy-
      // paste "Paste" path instead. Purely informational — not editable
      // here.
      const mode = document.createElement("span");
      mode.className = "upload-review-mode" + (item.insertMode === "reconstruct" ? " upload-review-mode-native" : "");
      mode.textContent = item.insertMode === "reconstruct" ? "Native" : "Paste";
      mode.title =
        item.insertMode === "reconstruct"
          ? "Inserts directly with one click, built from the shape's properties."
          : "Inserts via a temporary slide plus copy/paste — this slide has content (a picture, table, or custom-drawn shape) that can't be rebuilt natively.";
      row.appendChild(mode);

      const fields = document.createElement("div");
      fields.className = "upload-review-fields";

      const titleInput = document.createElement("input");
      titleInput.type = "text";
      titleInput.className = "upload-review-title";
      titleInput.value = item.title;
      titleInput.placeholder = "Title";
      titleInput.addEventListener("input", () => {
        item.title = titleInput.value;
      });
      fields.appendChild(titleInput);

      const tagsInput = document.createElement("input");
      tagsInput.type = "text";
      tagsInput.className = "upload-review-tags";
      tagsInput.placeholder = "Tags, comma separated";
      tagsInput.addEventListener("input", () => {
        item.tagsRaw = tagsInput.value;
      });
      fields.appendChild(tagsInput);

      row.appendChild(fields);
      list.appendChild(row);
    }
  }
}

function setCommitError(message: string): void {
  const el = document.getElementById("commitError");
  if (!el) return;
  el.textContent = message;
  (el as HTMLElement).style.display = message ? "block" : "none";
}

async function commit(): Promise<void> {
  if (!jobId) return;
  setCommitError("");
  const btn = document.getElementById("btnCommit") as HTMLButtonElement | null;
  if (btn) btn.disabled = true;
  try {
    const res = await fetch(`/api/admin/library-upload/${jobId}/commit`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        items: reviewItems.map((item) => ({
          tempId: item.tempId,
          title: item.title,
          category: item.category,
          included: item.included,
          tags: item.tagsRaw
            .split(",")
            .map((t) => t.trim())
            .filter(Boolean),
        })),
      }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    showScreen("doneScreen");
    const status = document.getElementById("doneStatus");
    if (status) status.textContent = `Added ${body.created} item(s) to the library.`;
  } catch (err) {
    setCommitError(err instanceof Error ? err.message : String(err));
    if (btn) btn.disabled = false;
  }
}

async function cancelUpload(): Promise<void> {
  if (jobId) {
    await fetch(`/api/admin/library-upload/${jobId}/cancel`, { method: "POST" }).catch(() => {});
  }
  finish();
}

/** The dialog's only way out — see the module comment. */
function finish(): void {
  Office.context.ui.messageParent(JSON.stringify({ action: "done" }));
}

function populateScope(): void {
  const scopeRow = document.getElementById("scopeRow");
  const select = document.getElementById("scopeSelect") as HTMLSelectElement | null;
  if (!scopeRow || !select) return;

  const options: { value: Scope; label: string }[] = [];
  if (canUseGlobal()) options.push({ value: "global", label: "Global Library" });
  if (canUseCompany()) options.push({ value: "company", label: companyDomain! });

  select.innerHTML = "";
  for (const opt of options) {
    const el = document.createElement("option");
    el.value = opt.value;
    el.textContent = opt.label;
    select.appendChild(el);
  }
  scope = options[0]?.value ?? "global";
  (scopeRow as HTMLElement).style.display = options.length > 1 ? "flex" : "none";

  select.addEventListener("change", () => {
    scope = select.value as Scope;
    renderFileList();
    updateStartButton();
  });
}

Office.onReady(async () => {
  try {
    const res = await fetch("/api/auth/me");
    if (res.ok) {
      const user = await res.json();
      isAdmin = !!user?.isAdmin;
      isCompanyAdmin = !!user?.isCompanyAdmin;
      companyDomain = user?.companyDomain ?? null;
    }
  } catch {
    // leave everything false — the check below reports it
  }

  if (!canUseGlobal() && !canUseCompany()) {
    setSetupError("You don't have permission to upload library content.");
    const btn = document.getElementById("btnStartUpload") as HTMLButtonElement | null;
    if (btn) btn.disabled = true;
    document.getElementById("btnChooseFiles")?.setAttribute("disabled", "true");
    return;
  }

  populateScope();

  document.getElementById("btnChooseFiles")?.addEventListener("click", () => {
    (document.getElementById("fileInput") as HTMLInputElement | null)?.click();
  });
  document.getElementById("fileInput")?.addEventListener("change", () => {
    const input = document.getElementById("fileInput") as HTMLInputElement | null;
    if (input?.files?.length) addFiles(input.files);
    if (input) input.value = ""; // reset so picking the same file(s) again still fires "change"
  });
  document.getElementById("btnStartUpload")?.addEventListener("click", () => {
    startUpload().catch((err) => setSetupError(err instanceof Error ? err.message : String(err)));
  });
  document.getElementById("btnProgressBack")?.addEventListener("click", () => {
    jobId = null;
    showScreen("setupScreen");
    const btn = document.getElementById("btnStartUpload") as HTMLButtonElement | null;
    if (btn) btn.disabled = chosenFiles.length === 0;
  });
  document.getElementById("btnCommit")?.addEventListener("click", () => {
    commit().catch((err) => setCommitError(err instanceof Error ? err.message : String(err)));
  });
  document.getElementById("btnCancelReview")?.addEventListener("click", () => {
    cancelUpload().catch(() => finish());
  });
  document.getElementById("btnClose")?.addEventListener("click", () => finish());
});
