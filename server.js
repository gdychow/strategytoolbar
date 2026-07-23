const express = require("express");
const cookieParser = require("cookie-parser");
const multer = require("multer");
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
  updateCatalogItem,
  updateCatalogItemThumbnail,
  deleteCatalogItem,
} = require("./server/db");
const { verifyMicrosoftIdToken, createSessionToken, verifySessionToken } = require("./server/auth");
const { THUMBNAILS_DIR, CATALOG_CATEGORIES, resolveCatalogFilePath } = require("./server/catalog");

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, "dist");
const useTls = process.env.USE_TLS !== "false";

const app = express();
app.use(express.json());
app.use(cookieParser());
// Volume-backed thumbnails take precedence over the image-baked copy under
// dist/ — registered first so express.static's fallthrough-on-miss (it
// calls next() rather than erroring when a path doesn't resolve under its
// root) lets any thumbnail not yet on the volume (e.g. local dev, where
// docker-entrypoint.sh's one-time seed never runs) still resolve from the
// dist/ mount below, unchanged from today's behavior.
app.use("/assets/catalog/thumbnails", express.static(THUMBNAILS_DIR));

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

function requireAdmin(req, res, next) {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isAdmin) return res.status(403).send("Not an admin.");
  next();
}

function escapeHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Filename extension is always derived from this fixed table, never from
// the uploaded file's own name — the path-traversal defense for thumbnail
// uploads (see POST /admin/catalog/:id below).
const MIME_TO_EXT = { "image/png": "png", "image/jpeg": "jpg", "image/webp": "webp", "image/gif": "gif" };
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (!MIME_TO_EXT[file.mimetype]) return cb(new Error("Unsupported image type."));
    cb(null, true);
  },
});

// Phase 3 proved the authorization boundary; this now also lets the owner
// fix real curation mistakes (wrong title, wrong category, wrong
// thumbnail) directly, instead of hand-editing a seed JSON file and
// rerunning scripts/seed-catalog.js inside the container for every typo.
// insert_mode/source_file/reconstruct_spec stay out of this UI entirely —
// those are still owned by the slice+seed script pipeline. Creating new
// items is also out of scope here (deferred, not forgotten).
app.get("/admin", requireAdmin, async (req, res) => {
  const items = await listAllCatalogItems();
  const errorMsg = typeof req.query.error === "string" ? req.query.error : null;
  const rows = items
    .map((item) => {
      const thumbUrl = item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}` : null;
      const categoryOptions = CATALOG_CATEGORIES.map(
        (c) => `<option value="${c}"${c === item.category ? " selected" : ""}>${c}</option>`
      ).join("");
      return `
        <form id="edit-${item.id}" method="POST" action="/admin/catalog/${item.id}" enctype="multipart/form-data"></form>
        <tr>
          <td><input form="edit-${item.id}" name="title" value="${escapeHtml(item.title)}" size="30"></td>
          <td><select form="edit-${item.id}" name="category">${categoryOptions}</select></td>
          <td>${escapeHtml(item.insert_mode)}</td>
          <td><input form="edit-${item.id}" name="sortOrder" type="number" value="${item.sort_order}" style="width: 60px;"></td>
          <td>
            ${thumbUrl ? `<img src="${thumbUrl}" width="60" alt="">` : "(none)"}
            <input form="edit-${item.id}" name="thumbnail" type="file" accept="image/png,image/jpeg,image/webp,image/gif">
          </td>
          <td><button form="edit-${item.id}" type="submit">Save</button></td>
          <td>
            <form method="POST" action="/admin/catalog/${item.id}/delete" onsubmit="return confirm('Delete this catalog item permanently?')">
              <button type="submit">Delete</button>
            </form>
          </td>
        </tr>`;
    })
    .join("");
  res.send(`<!doctype html><html><body>
    <h1>Welcome, admin</h1>
    <p>Signed in as ${escapeHtml(req.user.email)}.</p>
    ${errorMsg ? `<p style="color: red;">${escapeHtml(errorMsg)}</p>` : ""}
    <table border="1" cellpadding="4">
      <tr><th>Title</th><th>Category</th><th>Insert mode</th><th>Sort order</th><th>Thumbnail</th><th></th><th></th></tr>
      ${rows}
    </table>
  </body></html>`);
});

app.post("/admin/catalog/:id", requireAdmin, upload.single("thumbnail"), async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid item id."));

  const title = typeof req.body.title === "string" ? req.body.title.trim() : "";
  const category = req.body.category;
  const sortOrder = Number(req.body.sortOrder);
  if (!title) return res.redirect(303, "/admin?error=" + encodeURIComponent("Title can't be empty."));
  if (!CATALOG_CATEGORIES.includes(category)) return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid category."));
  if (!Number.isInteger(sortOrder) || sortOrder < 0) {
    return res.redirect(303, "/admin?error=" + encodeURIComponent("Sort order must be a non-negative integer."));
  }

  const existing = await getCatalogItem(id);
  if (!existing) return res.redirect(303, "/admin?error=" + encodeURIComponent("Item not found."));

  const updated = await updateCatalogItem({ id, title, category, sortOrder });
  if (!updated) return res.redirect(303, "/admin?error=" + encodeURIComponent("Item not found."));

  if (req.file) {
    const ext = MIME_TO_EXT[req.file.mimetype];
    const newThumbnailPath = `item-${id}.${ext}`;
    await fs.promises.writeFile(path.join(THUMBNAILS_DIR, newThumbnailPath), req.file.buffer);
    await updateCatalogItemThumbnail({ id, thumbnailPath: newThumbnailPath });
    if (existing.thumbnail_path && existing.thumbnail_path !== newThumbnailPath) {
      await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
    }
  }

  res.redirect(303, "/admin");
});

app.post("/admin/catalog/:id/delete", requireAdmin, async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid item id."));

  const existing = await getCatalogItem(id);
  const deleted = await deleteCatalogItem(id);
  if (!deleted) return res.redirect(303, "/admin?error=" + encodeURIComponent("Item not found."));

  if (existing?.thumbnail_path) {
    await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
  }

  res.redirect(303, "/admin");
});

// Catches multer's file-size/type rejections (fileFilter's cb(new Error(...)))
// so a bad upload gets a clean redirect instead of Express's default HTML
// 500 page. Must have 4 params for Express to recognize it as error-handling
// middleware, even though `next` is unused.
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  if (req.path.startsWith("/admin/catalog/")) {
    return res.redirect(303, "/admin?error=" + encodeURIComponent(err.message || "Upload failed."));
  }
  console.error(err);
  res.status(500).send("Server error.");
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
