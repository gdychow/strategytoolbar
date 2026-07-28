/**
 * Task Pane Phase 15: the account-creation dialog, opened via
 * Office.context.ui.displayDialogAsync from the task pane's "Create
 * Account" button — same mechanism as the content gallery (see
 * src/gallery/gallery.ts's module comment for why a dialog can make its
 * own authenticated fetch calls directly, not just relay through
 * messageParent). A session already exists by the time this opens (either
 * from a fresh Microsoft sign-in or a returning-but-unregistered one), so
 * this page only ever needs to capture details and call
 * POST /api/auth/register — no auth flow of its own.
 */

interface Me {
  displayName: string | null;
  companyDomain: string | null;
}

function statusEl(): HTMLElement | null {
  return document.getElementById("status");
}

function errorEl(): HTMLElement | null {
  return document.getElementById("errorMsg");
}

function showError(message: string): void {
  const el = errorEl();
  if (!el) return;
  el.textContent = message;
  el.style.display = "block";
}

function clearError(): void {
  const el = errorEl();
  if (!el) return;
  el.style.display = "none";
  el.textContent = "";
}

async function loadDefaults(): Promise<void> {
  const status = statusEl();
  try {
    const res = await fetch("/api/auth/me");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const me: Me = await res.json();

    const fullNameInput = document.getElementById("fullName") as HTMLInputElement | null;
    if (fullNameInput && me.displayName) fullNameInput.value = me.displayName;

    const companyField = document.getElementById("companyField");
    const companyNameInput = document.getElementById("companyName") as HTMLInputElement | null;
    const noCompanyNote = document.getElementById("noCompanyNote");
    if (me.companyDomain) {
      if (companyNameInput) companyNameInput.value = me.companyDomain;
    } else {
      // Consumer-domain email (gmail.com, etc.) — no company step at all,
      // matching server/consumerDomains.js's denylist.
      if (companyField) (companyField as HTMLElement).style.display = "none";
      if (noCompanyNote) (noCompanyNote as HTMLElement).style.display = "block";
    }

    if (status) status.textContent = "";
  } catch (err) {
    if (status) status.textContent = `Couldn't load your details: ${err instanceof Error ? err.message : String(err)}`;
  }
}

async function handleSubmit(e: Event): Promise<void> {
  e.preventDefault();
  clearError();

  const submitBtn = document.getElementById("btnSubmit") as HTMLButtonElement | null;
  const fullName = (document.getElementById("fullName") as HTMLInputElement | null)?.value.trim() ?? "";
  const companyName = (document.getElementById("companyName") as HTMLInputElement | null)?.value.trim() ?? "";
  const jobTitle = (document.getElementById("jobTitle") as HTMLInputElement | null)?.value.trim() ?? "";
  const termsAccepted = (document.getElementById("termsAccepted") as HTMLInputElement | null)?.checked ?? false;
  const planInput = document.querySelector<HTMLInputElement>('input[name="plan"]:checked');
  const plan = planInput?.value ?? "monthly";

  if (submitBtn) submitBtn.disabled = true;
  try {
    const res = await fetch("/api/auth/register", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fullName, companyName, jobTitle, plan, termsAccepted }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
    Office.context.ui.messageParent(JSON.stringify({ success: true }));
  } catch (err) {
    showError(err instanceof Error ? err.message : String(err));
    if (submitBtn) submitBtn.disabled = false;
  }
}

Office.onReady(() => {
  loadDefaults();
  document.getElementById("registerForm")?.addEventListener("submit", (e) => {
    handleSubmit(e).catch((err) => showError(err instanceof Error ? err.message : String(err)));
  });
});
