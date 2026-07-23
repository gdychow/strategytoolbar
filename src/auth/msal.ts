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
  BrowserAuthError,
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

/**
 * Needed for Office-on-web, where NAA identifies the account via a login
 * hint rather than an ambient OS session. Harmless on desktop even though
 * v1 targets desktop only. Cached (module-level, like msalInstance) so a
 * click on Sign In never has to await this itself — see signIn()'s comment
 * for why that matters.
 */
let loginHintPromise: Promise<string | undefined> | null = null;
function getLoginHint(): Promise<string | undefined> {
  if (!loginHintPromise) {
    loginHintPromise = (async () => {
      try {
        const authContext = await Office.auth.getAuthContext();
        return authContext?.userPrincipalName;
      } catch {
        return undefined;
      }
    })();
  }
  return loginHintPromise;
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
 * Passively restores a prior sign-in via ssoSilent, with no popup and no
 * user interaction — call this once in the background on startup (never
 * from the Sign In button's click handler; see signIn()'s comment for why
 * that split matters). Never throws: any failure just means "not signed
 * in yet", which is a normal, expected outcome for a first-time user.
 */
export async function trySilentSignIn(): Promise<SignedInUser | null> {
  try {
    const instance = await getMsalInstance();
    const loginHint = await getLoginHint();
    if (!loginHint) return null; // ssoSilent needs a hint (or a cached account) to target
    const result = await instance.ssoSilent({ scopes: ["User.Read"], loginHint });
    return toSignedInUser(result);
  } catch {
    return null;
  }
}

/**
 * Signs the user in interactively via a popup. Deliberately does NOT try
 * ssoSilent first (use trySilentSignIn for that, in the background, on
 * startup): browsers only allow window.open() to bypass the popup blocker
 * within a short "trusted user activation" window right after a real
 * click, and ssoSilent's hidden-iframe round trip (several seconds before
 * it fails) blows straight through that window — confirmed as the actual
 * cause of the timed_out -> interaction_in_progress -> popup_window_error
 * chain seen when ssoSilent was awaited inside this same click handler.
 * getMsalInstance()/getLoginHint() are cheap here because trySilentSignIn
 * has normally already resolved and cached both by the time the user
 * clicks; if it hasn't (a very fast click), this still works, just with a
 * small risk of the same activation-window issue on that one click.
 *
 * Takes an optional progress callback so the caller can surface each step.
 */
export async function signIn(onProgress?: (step: string) => void): Promise<SignedInUser> {
  onProgress?.("Opening sign-in popup...");
  const instance = await getMsalInstance();
  const loginHint = await getLoginHint();

  // MSAL requires at least one scope beyond openid/profile/email/offline_access.
  const request = {
    scopes: ["User.Read"],
    ...(loginHint ? { loginHint } : {}),
  };

  try {
    const result = await instance.acquireTokenPopup(request);
    return toSignedInUser(result);
  } catch (error) {
    // A backgrounded trySilentSignIn() that's still in flight (or one that
    // timed out without its cleanup running — a known msal-browser gap,
    // confirmed against node_modules source) can leave MSAL's own
    // "interaction in progress" lock set, which would otherwise reject
    // this popup outright. Retry once with overrideInteractionInProgress —
    // a documented public PopupRequest option for exactly this case.
    if (error instanceof BrowserAuthError && error.errorCode === "interaction_in_progress") {
      onProgress?.("Clearing a stuck sign-in state and retrying...");
      const result = await instance.acquireTokenPopup({ ...request, overrideInteractionInProgress: true });
      return toSignedInUser(result);
    }
    throw error;
  }
}
