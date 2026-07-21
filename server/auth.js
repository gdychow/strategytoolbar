const { jwtVerify, createRemoteJWKSet, SignJWT } = require("jose");

const MICROSOFT_JWKS = createRemoteJWKSet(new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys"));

// Session cookie TTLs. Short-lived + refreshed on each successful silent
// re-auth is what actually delivers "persistent with occasional
// rechecking" — a flat 30-day cookie would never re-validate at all.
const SESSION_TTL_SECONDS = 24 * 60 * 60; // 24h per issuance
const SESSION_MAX_AGE_SECONDS = 30 * 24 * 60 * 60; // 30d absolute backstop
const REFRESH_THRESHOLD_SECONDS = 12 * 60 * 60; // reissue once more than half the TTL has elapsed

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

async function createSessionToken(user, sessionStart = Date.now()) {
  return new SignJWT({
    oid: user.oid,
    tid: user.tid,
    email: user.email,
    displayName: user.displayName,
    isAdmin: isAdminEmail(user.email),
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
};
