// Chromium-only Web API (Chrome 95+); not yet in this project's installed
// TypeScript version's lib.dom.d.ts, and unsupported in WebKit (confirmed:
// open WebKit bug, no Safari/Mac-WKWebView support) — feature-detected via
// `"EyeDropper" in window` at every call site, never assumed present.
interface EyeDropperOpenResult {
  sRGBHex: string;
}

interface EyeDropper {
  open(options?: { signal?: AbortSignal }): Promise<EyeDropperOpenResult>;
}

interface EyeDropperConstructor {
  new (): EyeDropper;
}

interface Window {
  EyeDropper?: EyeDropperConstructor;
}
