/**
 * Task Pane Phase 20 — Template Library. Whole .potx presentation
 * templates, not individual slide content — a structurally different
 * action from the rest of libraryInsert.ts (creates a brand-new
 * presentation instead of touching the current one), so this stays a
 * separate module rather than folding into that file.
 *
 * "Create from template" needed no download/system-browser mechanism in
 * the end: PowerPoint.createPresentation(base64) (PowerPointApi 1.1)
 * creates and opens a new presentation in-app directly, once the server
 * has patched the stored .potx's internal content-type declaration to
 * look like an ordinary deck (see server/potxConvert.js — confirmed
 * necessary by testing directly, a real .potx's raw bytes are rejected
 * outright).
 */

export interface TemplateItem {
  id: number;
  title: string;
  description: string | null;
}

export type TemplateScope = "personal" | "company" | "global";

async function parseJsonOrThrow<T>(res: Response): Promise<T> {
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
  return body;
}

export async function fetchTemplates(scope: TemplateScope): Promise<TemplateItem[]> {
  const res = await fetch(`/api/templates/${scope}`);
  const body = await parseJsonOrThrow<{ templates: TemplateItem[] }>(res);
  return body.templates;
}

export async function uploadTemplate(scope: TemplateScope, file: File, title: string): Promise<TemplateItem> {
  const form = new FormData();
  form.append("file", file);
  form.append("title", title);
  const res = await fetch(`/api/templates/${scope}`, { method: "POST", body: form });
  const body = await parseJsonOrThrow<{ template: TemplateItem }>(res);
  return body.template;
}

export async function renameTemplate(id: number, title: string, description: string): Promise<void> {
  const res = await fetch(`/api/templates/${id}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, description }),
  });
  await parseJsonOrThrow<{ ok: true }>(res);
}

export async function deleteTemplate(id: number): Promise<void> {
  const res = await fetch(`/api/templates/${id}/delete`, { method: "POST" });
  await parseJsonOrThrow<{ ok: true }>(res);
}

/**
 * Fetches the converted bytes and hands them straight to
 * PowerPoint.createPresentation — no download, no second click. Must be
 * called from the task pane's own context, never from inside the
 * templates dialog itself (a dialog can only call
 * messageParent/requirements.isSetSupported, confirmed back in Phase 5's
 * own research — createPresentation is a PowerPoint-specific API, off
 * limits there).
 */
export async function createFromTemplate(id: number): Promise<void> {
  const res = await fetch(`/api/templates/${id}/create-payload`);
  const body = await parseJsonOrThrow<{ base64: string }>(res);
  await PowerPoint.createPresentation(body.base64);
}
