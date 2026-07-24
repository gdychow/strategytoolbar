/**
 * Phase 5 skeleton. A page opened via Office.context.ui.displayDialogAsync
 * can call exactly two Office.js APIs — messageParent and
 * requirements.isSetSupported (confirmed against Microsoft's own docs
 * source, office-js-docs-pr's dialog-api-in-office-add-ins.md) — so this
 * page is a pure browse/select surface; the task pane performs the actual
 * insert once this dialog reports back which item was chosen.
 *
 * Right now this just answers one open question before the real UI (tabs,
 * search, grouped grid, preview) gets built on top of it: does the dialog
 * share the task pane's session cookie? Both are same-origin (required
 * for messageParent to work at all), which under normal browser rules
 * means yes — but this project has already been burned once this session
 * by an embedded-webview cookie-jar assumption turning out wrong (NAA's
 * task pane vs. a regular browser tab), so this checks it directly against
 * a real Office host instead of assuming.
 */
Office.onReady(() => {
  const statusEl = document.getElementById("status");
  if (!statusEl) return;

  fetch("/api/catalog/text")
    .then(async (res) => {
      if (res.status === 401) {
        statusEl.textContent = "NOT authenticated inside the dialog (401) — session cookie is not shared.";
        return;
      }
      if (!res.ok) {
        statusEl.textContent = `Unexpected response: ${res.status}`;
        return;
      }
      const data = await res.json();
      statusEl.textContent = `Authenticated inside the dialog — session cookie IS shared. Loaded ${data.items.length} item(s), ${data.groups.length} group(s).`;
    })
    .catch((err) => {
      statusEl.textContent = `Fetch failed: ${err instanceof Error ? err.message : String(err)}`;
    });
});
