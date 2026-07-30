/** Lightweight replacement for VBA's MsgBox — surfaces messages in the task pane. */

let statusEl: HTMLElement | null = null;

export function bindStatusElement(el: HTMLElement): void {
  statusEl = el;
}

export function notify(message: string, kind: "info" | "error" = "info"): void {
  console[kind === "error" ? "error" : "log"](message);
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.className = `status status-${kind}`;
}

/** Wraps a button handler: reports thrown errors instead of leaving the pane silent. */
export function withErrorHandling(fn: () => Promise<void>): () => void {
  return () => {
    fn().catch((err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      notify(`Error: ${message}`, "error");
    });
  };
}

/**
 * Every server route that can reject a request sends a real, specific
 * reason as JSON — `{ error: "This account has been suspended." }`, "Not
 * signed in.", "Finish creating your account first.", etc. — but a plain
 * `if (!res.ok) throw new Error(...(${res.status}))` throws that away and
 * shows the caller nothing but an HTTP status code. This reads the real
 * message back out (falling back to a generic one only if the body isn't
 * JSON or has no `error` field — e.g. a proxy-level 502) so error text
 * shown to the user always reflects why, not just that something failed.
 */
export async function extractErrorMessage(res: Response, fallback = `Request failed (${res.status}).`): Promise<string> {
  try {
    const body = await res.json();
    if (body && typeof body.error === "string" && body.error) return body.error;
  } catch {
    // Not JSON, or no body — fall through to the generic message.
  }
  return fallback;
}
