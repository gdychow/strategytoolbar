const express = require("express");
const cookieParser = require("cookie-parser");
const path = require("path");
const fs = require("fs");
const https = require("https");
const os = require("os");

const {
  waitForDatabase,
  upsertUser,
  listSharedCatalogItems,
  listAllCatalogItems,
  getCatalogItem,
} = require("./server/db");
const { verifyMicrosoftIdToken, createSessionToken, verifySessionToken } = require("./server/auth");
const { resolveCatalogFilePath } = require("./server/catalog");

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, "dist");
const useTls = process.env.USE_TLS !== "false";

const app = express();
app.use(express.json());
app.use(cookieParser());
// Office Add-in task panes on Mac are known to cache their web content
// aggressively (WKWebView), which can silently leave an old taskpane.js
// running after a deploy with no visible sign anything is wrong.
// "no-cache" doesn't disable caching — it forces a revalidation request on
// every load, so a change always takes effect on the next reload instead of
// needing the user to manually clear the Office cache.
app.use(express.static(ROOT, { setHeaders: (res) => res.setHeader("Cache-Control", "no-cache") }));

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

// Tier 3: the shared content library. Any signed-in user can browse and
// insert from it — it's "shared", not "admin-only to read". source_file
// itself is never client-supplied: the client only ever sends a numeric
// item ID, and the server looks up which file (if any) that row points at.
app.get("/api/catalog/:category", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  const items = await listSharedCatalogItems(req.params.category);
  res.json(
    items.map((item) => ({
      id: item.id,
      title: item.title,
      insertMode: item.insert_mode,
      reconstructSpec: item.reconstruct_spec,
      thumbnailUrl: item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}` : null,
    }))
  );
});

app.get("/api/catalog/file/:itemId", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });

  const item = await getCatalogItem(req.params.itemId);
  if (!item || item.insert_mode !== "file") {
    return res.status(404).json({ error: "Not found." });
  }

  let filePath;
  try {
    filePath = resolveCatalogFilePath(item.source_file);
  } catch (err) {
    console.error("Catalog file path resolution failed:", err.message);
    return res.status(500).json({ error: "Server error." });
  }

  res.type("application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.sendFile(filePath, (err) => {
    if (err && !res.headersSent) res.status(404).json({ error: "File not found." });
  });
});

// Phase 3: proves the authorization boundary works end-to-end. Now also
// gives the owner a read-only view of what scripts/seed-catalog.js has
// loaded — real upload/edit UI is still out of scope (content is seeded
// via script for v1).
app.get("/admin", async (req, res) => {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isAdmin) return res.status(403).send("Not an admin.");

  const items = await listAllCatalogItems();
  const rows = items
    .map(
      (item) =>
        `<tr><td>${item.title}</td><td>${item.category}</td><td>${item.insert_mode}</td><td>${item.thumbnail_path ? "yes" : "no"}</td></tr>`
    )
    .join("");
  res.send(`<!doctype html><html><body>
    <h1>Welcome, admin</h1>
    <p>Signed in as ${req.user.email}.</p>
    <table border="1" cellpadding="4">
      <tr><th>Title</th><th>Category</th><th>Insert mode</th><th>Thumbnail</th></tr>
      ${rows}
    </table>
  </body></html>`);
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
