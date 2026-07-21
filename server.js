const express = require("express");
const cookieParser = require("cookie-parser");
const path = require("path");
const fs = require("fs");
const https = require("https");
const os = require("os");

const { waitForDatabase, upsertUser } = require("./server/db");
const { verifyMicrosoftIdToken, createSessionToken, verifySessionToken } = require("./server/auth");

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, "dist");
const useTls = process.env.USE_TLS !== "false";

const app = express();
app.use(express.json());
app.use(cookieParser());
app.use(express.static(ROOT));

const SESSION_COOKIE = "session";
const cookieOptions = {
  httpOnly: true,
  secure: true,
  sameSite: "lax",
};

// Attaches req.user from the session cookie if present and valid, and
// transparently reissues the cookie on a sliding-expiration basis — this is
// what makes "persistent with occasional rechecking" actually happen,
// rather than a single long-lived token that never re-validates. Desktop
// task panes only (v1 scope) — a cookie set from an Office-on-web iframe
// context is a third-party cookie by browser rules and may not persist
// there; that's a documented limitation, not a bug, if it ever comes up.
app.use(async (req, res, next) => {
  const token = req.cookies?.[SESSION_COOKIE];
  if (!token) return next();

  const verified = await verifySessionToken(token);
  if (!verified) {
    res.clearCookie(SESSION_COOKIE, cookieOptions);
    return next();
  }

  req.user = verified.claims;
  if (verified.shouldRefresh) {
    const fresh = await createSessionToken(
      {
        oid: verified.claims.oid,
        tid: verified.claims.tid,
        email: verified.claims.email,
        displayName: verified.claims.displayName,
      },
      verified.claims.sessionStart
    );
    res.cookie(SESSION_COOKIE, fresh, cookieOptions);
  }
  next();
});

app.post("/api/auth/session", async (req, res) => {
  const { idToken } = req.body ?? {};
  if (typeof idToken !== "string") {
    return res.status(400).json({ error: "Missing idToken." });
  }

  let identity;
  try {
    identity = await verifyMicrosoftIdToken(idToken);
  } catch (err) {
    console.warn("ID token verification failed:", err.message);
    return res.status(401).json({ error: "Invalid token." });
  }

  const user = await upsertUser(identity);
  const sessionToken = await createSessionToken(identity);
  res.cookie(SESSION_COOKIE, sessionToken, cookieOptions);
  res.json({
    oid: user.oid,
    tid: user.tid,
    email: user.email,
    displayName: user.display_name,
  });
});

app.get("/api/auth/me", (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  const { oid, tid, email, displayName, isAdmin } = req.user;
  res.json({ oid, tid, email, displayName, isAdmin });
});

app.post("/api/auth/signout", (req, res) => {
  res.clearCookie(SESSION_COOKIE, cookieOptions);
  res.status(204).end();
});

// Phase 3: proves the authorization boundary works end-to-end. Deliberately
// a bare placeholder — the real catalog upload/management UI is Tier 3's
// own work, built on top of this once it exists.
app.get("/admin", (req, res) => {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isAdmin) return res.status(403).send("Not an admin.");
  res.send(`<!doctype html><html><body><h1>Welcome, admin</h1><p>Signed in as ${req.user.email}. Catalog management goes here.</p></body></html>`);
});

function startServer() {
  if (useTls) {
    const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
    const options = {
      key: fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
    };
    https.createServer(options, app).listen(PORT, () => {
      console.log(`Strategy Toolbar dev server running at https://localhost:${PORT}/taskpane.html`);
    });
  } else {
    app.listen(PORT, () => {
      console.log(`Strategy Toolbar server running on plain HTTP, port ${PORT} (TLS expected to terminate upstream)`);
    });
  }
}

waitForDatabase()
  .then(startServer)
  .catch((err) => {
    console.error("Failed to start:", err);
    process.exit(1);
  });
