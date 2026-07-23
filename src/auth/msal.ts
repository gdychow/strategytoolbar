/**
 * Phase 1 of the admin/auth plan: sign the user in via Nested App
 * Authentication (NAA) and surface their identity claims. This module is
 * intentionally client-only for now — no backend calls yet. See
 * ~/.claude/plans/fizzy-mixing-squid.md for the full design.
 *
 * NAA needs no manifest.xml changes (confirmed against Microsoft's current
 * docs — support is a pure runtime check, not a manifest declaration), but
 * it does need a real Azure App Registration; src/config/auth.json holds
 * the resulting client ID and starts out with a placeholder that must be
 * replaced before sign-in will work.
 *
 * cacheLocation is "sessionStorage", not MSAL's default "localStorage" —
 * NAA's real persistence lives in the OS/Office-brokered account, not in
 * MSAL's own cache.
 */

import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import authConfig from "../config/auth.json";

/** NAA is its own requirement set (not gated by a PowerPointApi version) — confirmed GA for PowerPoint on Windows/Mac/Web, but needs a WebView2-hosted Office runtime on Windows specifically. */
export function isNestedAppAuthSupported(): boolean {
  return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
}

let msalInstance: IPublicClientApplication | null = null;

async function getMsalInstance(): Promise<IPublicClientApplication> {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: authConfig.clientId,
        authority: authConfig.authority,
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    });
  }
  return msalInstance;
}

/** Needed for Office-on-web, where NAA identifies the account via a login hint rather than an ambient OS session. Harmless on desktop even though v1 targets desktop only. */
async function getLoginHint(): Promise<string | undefined> {
  try {
    const authContext = await Office.auth.getAuthContext();
    return authContext?.userPrincipalName;
  } catch {
    return undefined;
  }
}

export interface SignedInUser {
  email: string | null;
  oid: string | null;
  tid: string | null;
  idToken: string;
}

function toSignedInUser(result: AuthenticationResult): SignedInUser {
  const claims = (result.idTokenClaims ?? {}) as Record<string, unknown>;
  return {
    email: typeof claims.email === "string" ? claims.email : null,
    oid: typeof claims.oid === "string" ? claims.oid : null,
    tid: typeof claims.tid === "string" ? claims.tid : null,
    idToken: result.idToken,
  };
}

/**
 * Signs the user in, preferring a silent flow (ssoSilent) and falling back
 * to an interactive popup only when Microsoft requires it (first sign-in,
 * expired session, MFA, etc.) — matches "persistent with occasional
 * rechecking" rather than prompting every time.
 *
 * Takes an optional progress callback so the caller can surface each step
 * (rather than one opaque await) — useful while debugging exactly where a
 * sign-in attempt is stuck, since a hung popup and a hung MSAL init look
 * identical from the outside otherwise.
 */
export async function signIn(onProgress?: (step: string) => void): Promise<SignedInUser> {
  onProgress?.("Initializing MSAL...");
  const instance = await getMsalInstance();
  const loginHint = await getLoginHint();

  // MSAL requires at least one scope beyond openid/profile/email/offline_access.
  const request = {
    scopes: ["User.Read"],
    ...(loginHint ? { loginHint } : {}),
  };

  try {
    onProgress?.("Trying silent sign-in...");
    const result = await instance.ssoSilent(request);
    return toSignedInUser(result);
  } catch (error) {
    // Falls back to a popup for *any* ssoSilent failure, not just
    // InteractionRequiredAuthError — ssoSilent works via a hidden iframe,
    // which needs the same storage/cookie access a real top-level page
    // gets; inside Office-on-web the task pane is already an iframe, so a
    // second, nested hidden iframe can end up storage-partitioned and just
    // hang until MSAL's own timeout fires (confirmed: "timed_out", not
    // InteractionRequiredAuthError, so the old narrower check never
    // reached this fallback at all). A popup opens as its own top-level
    // window with normal storage access, so it isn't subject to the same
    // nested-iframe restriction.
    onProgress?.("Silent sign-in didn't complete — opening popup...");
    const result = await instance.acquireTokenPopup(request);
    return toSignedInUser(result);
  }
}
