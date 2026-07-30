const { jwtVerify, createRemoteJWKSet, SignJWT } = require("jose");

const MICROSOFT_JWKS = createRemoteJWKSet(new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys"));

// Session cookie TTLs. Short-lived + refreshed on each successful silent
// re-auth is what actually delivers "persistent with occasional
// rechecking" — a flat 30-day cookie would never re-validate at all.
const SESSION_TTL_SECONDS = 24 * 60 * 60; // 24h per issuance
const SESSION_MAX_AGE_SECONDS = 30 * 24 * 60 * 60; // 30d absolute backstop
const REFRESH_THRESHOLD_SECONDS = 12 * 60 * 60; // reissue once more than half the TTL has elapsed

// Admin UI Phase 22 fix: suspend/delete originally only took effect at the
// next JWT refresh (~12h, REFRESH_THRESHOLD_SECONDS above) or the next
// sign-in — surprising in practice for an admin who just clicked Suspend
// and reasonably expects it to apply immediately. This in-memory cache of
// blocked (oid,tid) pairs is checked unconditionally on every request (see
// server.js's refresh middleware), not just at refresh time, so a
// suspension/deletion takes effect on the very next request regardless of
// how fresh the caller's session is. Rebuilt from the DB at startup
// (loadBlockedUsersCache) so a restart can't reopen a gap for someone
// already suspended; kept in sync afterward by the suspend/unsuspend/
// delete routes in server.js calling markUserBlocked/markUserUnblocked
// directly, since those actions and this cache live in the same single
// Node process — this app isn't horizontally scaled today; if it ever is,
// this cache would need to move to a shared store (e.g. Redis) instead.
const blockedUsers = new Set();

function userKey(oid, tid) {
  return `${oid}:${tid}`;
}

function isUserBlocked(oid, tid) {
  return blockedUsers.has(userKey(oid, tid));
}

function markUserBlocked(oid, tid) {
  blockedUsers.add(userKey(oid, tid));
}

function markUserUnblocked(oid, tid) {
  blockedUsers.delete(userKey(oid, tid));
}

/** Called once at server startup with the current suspended/deleted rows, so a restart doesn't temporarily un-block anyone. */
function loadBlockedUsersCache(rows) {
  for (const row of rows) blockedUsers.add(userKey(row.oid, row.tid));
}

function getSessionSecret() {
  const secret = process.env.SESSION_SECRET;
  if (!secret) {
    throw new Error("SESSION_SECRET is not set — refusing to sign sessions with no secret.");
  }
  return new TextEncoder().encode(secret);
}

/**
 * Verifies a Microsoft ID token: signature against Microsoft's live JWKS,
 * audience against our own Azure app registration, and issuer as a
 * *pattern* (https://login.microsoftonline.com/{tid}/v2.0) rather than a
 * single fixed value — personal Microsoft accounts don't necessarily use
 * the special "consumers" tenant GUID (confirmed against a real token
 * during the Phase 1 prototype; this account's tid was a distinct,
 * ordinary-looking tenant ID). `jwtVerify`'s built-in exp/nbf/iat checks
 * cover replay-by-expiry; nonce re-verification is not performed here —
 * that's a client-side protection MSAL already enforces during the
 * original auth exchange, and re-checking it server-side would need the
 * client to separately transmit its expected nonce, which nothing in this
 * design currently plumbs through.
 */
async function verifyMicrosoftIdToken(idToken) {
  const clientId = process.env.AZURE_CLIENT_ID;
  if (!clientId) {
    throw new Error("AZURE_CLIENT_ID is not set.");
  }

  const { payload } = await jwtVerify(idToken, MICROSOFT_JWKS, {
    audience: clientId,
  });

  const tid = payload.tid;
  const expectedIssuer = `https://login.microsoftonline.com/${tid}/v2.0`;
  if (typeof tid !== "string" || payload.iss !== expectedIssuer) {
    throw new Error(`Unexpected issuer: ${payload.iss}`);
  }

  return {
    oid: String(payload.oid),
    tid,
    email: typeof payload.email === "string" ? payload.email : null,
    displayName: typeof payload.name === "string" ? payload.name : null,
  };
}

/** Case-insensitive match against the comma-separated ADMIN_EMAILS env var. */
function isAdminEmail(email) {
  if (!email) return false;
  const admins = (process.env.ADMIN_EMAILS ?? "")
    .split(",")
    .map((e) => e.trim().toLowerCase())
    .filter(Boolean);
  return admins.includes(email.toLowerCase());
}

// companyDomain/isCompanyAdmin (Task Pane Phase 14) and isRegistered
// (Task Pane Phase 15) are all sourced from the caller (the DB row),
// unlike isAdmin — they're real mutable state (set by promote/demote
// actions or by completeRegistration), not something cheaply re-derivable
// from an env var on every call the way isAdminEmail is. isAdmin itself
// (Admin UI Phase 22) is now the OR of two independent sources: the
// env-file allowlist (isAdminEmail, permanent, never revocable via the UI)
// and user.isGlobalAdmin (a real DB column, set/cleared by the Users admin
// page's promote/demote-global-admin actions).
async function createSessionToken(user, sessionStart = Date.now()) {
  return new SignJWT({
    oid: user.oid,
    tid: user.tid,
    email: user.email,
    displayName: user.displayName,
    isAdmin: isAdminEmail(user.email) || !!user.isGlobalAdmin,
    companyDomain: user.companyDomain ?? null,
    isCompanyAdmin: !!user.isCompanyAdmin,
    isRegistered: !!user.isRegistered,
    sessionStart,
  })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime(Math.floor(Date.now() / 1000) + SESSION_TTL_SECONDS)
    .sign(getSessionSecret());
}

/**
 * Verifies our own session cookie (not a Microsoft token). Returns the
 * session claims plus whether it's due for a sliding-window refresh, or
 * null if invalid/expired/past the absolute max age.
 */
async function verifySessionToken(token) {
  try {
    const { payload } = await jwtVerify(token, getSessionSecret());
    const now = Date.now();
    if (now - payload.sessionStart > SESSION_MAX_AGE_SECONDS * 1000) {
      return null; // absolute backstop reached — force a fresh sign-in
    }
    const issuedAtMs = payload.iat * 1000;
    const shouldRefresh = now - issuedAtMs > REFRESH_THRESHOLD_SECONDS * 1000;
    return { claims: payload, shouldRefresh };
  } catch {
    return null;
  }
}

module.exports = {
  verifyMicrosoftIdToken,
  isAdminEmail,
  createSessionToken,
  verifySessionToken,
  SESSION_MAX_AGE_SECONDS,
  isUserBlocked,
  markUserBlocked,
  markUserUnblocked,
  loadBlockedUsersCache,
};
