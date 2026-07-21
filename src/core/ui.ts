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
