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
  listGroupsForCategory,
  getGroup,
  createGroup,
  updateGroup,
  deleteGroup,
  listAllTagNames,
  setItemTags,
} = require("./server/db");
const { verifyMicrosoftIdToken, createSessionToken, verifySessionToken } = require("./server/auth");
const { THUMBNAILS_DIR, CATALOG_CATEGORIES, resolveCatalogFilePath } = require("./server/catalog");
// Same clientId/authority the task pane's NAA client uses (src/auth/msal.ts)
// — reused as-is by /admin's separate, standard-MSAL browser sign-in flow
// below. Plain JSON, requirable directly from Node with no build step.
const authConfig = require("./src/config/auth.json");

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, "dist");
const useTls = process.env.USE_TLS !== "false";
// Cache-busting token for thumbnail URLs, appended as ?v=. Cache-Control:
// no-cache alone wasn't sufficient for WKWebView-hosted task panes on Mac
// earlier this session (taskpane.js/css needed an actual URL change, not
// just a header, to force a real reload) — thumbnails have the same
// problem once a category's images get regenerated, but no build step of
// their own to stamp a git commit into, so a per-process-start timestamp
// serves the same purpose: it changes on every deploy/restart.
const ASSET_VERSION = Date.now();

const app = express();
app.use(express.json());
// Needed for the plain (non-multipart) forms on /admin/groups — the
// catalog item edit form uses multipart/form-data (it has a file input,
// parsed by multer instead), but groups have no file upload, so their
// forms default to application/x-www-form-urlencoded, which nothing was
// parsing before this.
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
// Volume-backed thumbnails take precedence over the image-baked copy under
// dist/ — registered first so express.static's fallthrough-on-miss (it
// calls next() rather than erroring when a path doesn't resolve under its
// root) lets any thumbnail not yet on the volume (e.g. local dev, where
// docker-entrypoint.sh's one-time seed never runs) still resolve from the
// dist/ mount below, unchanged from today's behavior.
// Same "no-cache" reasoning as the dist/ mount below: WKWebView-hosted
// task panes on Mac cache aggressively and don't reliably revalidate on
// their own, confirmed earlier for taskpane.js/css — thumbnail images hit
// the exact same problem once a category's thumbnails get regenerated
// (e.g. the crop-to-content fix), since image URLs have no cache-busting
// query string the way taskpane.js/css do.
app.use("/assets/catalog/thumbnails", express.static(THUMBNAILS_DIR, { setHeaders: (res) => res.setHeader("Cache-Control", "no-cache") }));

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
//
// Response is { groups, items }, not a bare item array (Phase 5) — the
// gallery dialog needs the category's groups in their own admin-defined
// order to render group headers correctly, which can't be reliably
// inferred just from the first-occurrence order of group names within
// the already sort_order-sorted item list (a group's items don't have to
// be contiguous in that ordering).
app.get("/api/catalog/:category", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  const [items, groups] = await Promise.all([
    listSharedCatalogItems(req.params.category),
    listGroupsForCategory(req.params.category),
  ]);
  res.json({
    groups: groups.map((g) => ({ id: g.id, name: g.name, sortOrder: g.sort_order })),
    items: items.map((item) => ({
      id: item.id,
      title: item.title,
      insertMode: item.insert_mode,
      reconstructSpec: item.reconstruct_spec,
      thumbnailUrl: item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null,
      groupId: item.group_id,
      groupName: item.group_name,
      tags: item.tags,
    })),
  });
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

// Renders a standalone MSAL sign-in page for browser access to /admin — a
// second, independent auth flow from the task pane's NAA client
// (src/auth/msal.ts), since NAA specifically requires the Office host
// bridge (Office.auth) and won't function in a plain browser tab. Uses the
// UMD build vendored into dist/vendor by build.mjs rather than the esbuild
// bundle, since /admin has deliberately stayed framework-free. Posts the
// resulting idToken to the same POST /api/auth/session the task pane uses.
function renderSignInPage() {
  return `<!doctype html><html><body>
    <h1>Strategy Toolbar Admin</h1>
    <p id="status">Not signed in.</p>
    <button id="btnSignIn">Sign In</button>
    <script src="/vendor/msal-browser.min.js"></script>
    <script src="/vendor/msal-redirect-bridge.min.js"></script>
    <script>
      // loginPopup()'s opener window waits on a BroadcastChannel for the
      // auth response — this relays it there. Runs on every /admin load;
      // a no-op (rejects, ignored) on a normal visit with no pending auth
      // payload in the URL, and is what actually completes the flow when
      // this page is the popup that Microsoft just redirected back to.
      msalRedirectBridge.broadcastResponseToMainFrame().catch(() => {});

      const statusEl = document.getElementById("status");
      const msalInstance = new msal.PublicClientApplication({
        auth: {
          clientId: ${JSON.stringify(authConfig.clientId)},
          authority: ${JSON.stringify(authConfig.authority)},
          redirectUri: window.location.origin + "/admin",
        },
        cache: { cacheLocation: "sessionStorage" },
      });
      const ready = msalInstance.initialize();

      document.getElementById("btnSignIn").addEventListener("click", async () => {
        try {
          await ready;
          const result = await msalInstance.loginPopup({ scopes: ["User.Read"] });
          statusEl.textContent = "Signed in as " + (result.account?.username ?? "unknown") + " — finishing...";
          const res = await fetch("/api/auth/session", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ idToken: result.idToken }),
          });
          if (!res.ok) throw new Error("Failed to establish session (" + res.status + ").");
          window.location.reload();
        } catch (err) {
          statusEl.textContent = "Sign-in failed: " + (err && err.message ? err.message : String(err));
        }
      });
    </script>
  </body></html>`;
}

// Phase 3 proved the authorization boundary; this now also lets the owner
// fix real curation mistakes (wrong title, wrong category, wrong
// thumbnail) directly, instead of hand-editing a seed JSON file and
// rerunning scripts/seed-catalog.js inside the container for every typo.
// insert_mode/source_file/reconstruct_spec stay out of this UI entirely —
// those are still owned by the slice+seed script pipeline. Creating new
// items is also out of scope here (deferred, not forgotten).
app.get("/admin", async (req, res) => {
  if (!req.user) return res.send(renderSignInPage());
  if (!req.user.isAdmin) return res.status(403).send("Not an admin.");

  const [items, tagNames, groupsByCategory] = await Promise.all([
    listAllCatalogItems(),
    listAllTagNames(),
    Promise.all(CATALOG_CATEGORIES.map((c) => listGroupsForCategory(c))).then((lists) =>
      Object.fromEntries(CATALOG_CATEGORIES.map((c, i) => [c, lists[i]]))
    ),
  ]);
  const errorMsg = typeof req.query.error === "string" ? req.query.error : null;
  const rows = items
    .map((item) => {
      const thumbUrl = item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null;
      const categoryOptions = CATALOG_CATEGORIES.map(
        (c) => `<option value="${c}"${c === item.category ? " selected" : ""}>${c}</option>`
      ).join("");
      const groupOptions =
        `<option value="">(none)</option>` +
        groupsByCategory[item.category]
          .map((g) => `<option value="${g.id}"${g.id === item.group_id ? " selected" : ""}>${escapeHtml(g.name)}</option>`)
          .join("");
      return `
        <form id="edit-${item.id}" class="catalog-item-form" method="POST" action="/admin/catalog/${item.id}" enctype="multipart/form-data"></form>
        <tr>
          <td><input form="edit-${item.id}" name="title" value="${escapeHtml(item.title)}" size="30"></td>
          <td><select form="edit-${item.id}" name="category">${categoryOptions}</select></td>
          <td>${escapeHtml(item.insert_mode)}</td>
          <td><select form="edit-${item.id}" name="groupId">${groupOptions}</select></td>
          <td><input form="edit-${item.id}" name="tags" class="tags-input" value="${escapeHtml((item.tags || []).join(", "))}" size="20"></td>
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
  const groupNavLinks = CATALOG_CATEGORIES.map(
    (c) => `<a href="/admin/groups?category=${c}">${c}</a>`
  ).join(" · ");
  res.send(`<!doctype html><html><body>
    <h1>Welcome, admin</h1>
    <p>Signed in as ${escapeHtml(req.user.email)}.</p>
    ${errorMsg ? `<p style="color: red;">${escapeHtml(errorMsg)}</p>` : ""}
    <p>Manage groups: ${groupNavLinks}</p>
    <table border="1" cellpadding="4">
      <tr><th>Title</th><th>Category</th><th>Insert mode</th><th>Group</th><th>Tags</th><th>Sort order</th><th>Thumbnail</th><th></th><th></th></tr>
      ${rows}
    </table>
    <script>
      // Client-side typo-catching only, not a security boundary — the
      // server get-or-creates whatever tag names it's sent regardless
      // (see POST /admin/catalog/:id). This just makes an admin pause
      // before accidentally creating "arrows" next to an existing
      // "arrow". KNOWN_TAGS is embedded at render time rather than
      // fetched separately — the whole vocabulary is small (tens to a
      // couple hundred entries across ~230 items).
      const KNOWN_TAGS = ${JSON.stringify(tagNames)};

      function levenshtein(a, b) {
        const dp = [];
        for (let i = 0; i <= a.length; i++) dp.push([i]);
        for (let j = 1; j <= b.length; j++) dp[0][j] = j;
        for (let i = 1; i <= a.length; i++) {
          for (let j = 1; j <= b.length; j++) {
            dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
          }
        }
        return dp[a.length][b.length];
      }

      function nearMatches(tag) {
        const lower = tag.toLowerCase();
        const threshold = Math.max(2, Math.floor(lower.length * 0.3));
        return KNOWN_TAGS.filter((k) => k.toLowerCase() !== lower && levenshtein(lower, k.toLowerCase()) <= threshold);
      }

      document.querySelectorAll(".catalog-item-form").forEach((form) => {
        form.addEventListener("submit", (e) => {
          const tagsInput = document.querySelector('input.tags-input[form="' + form.id + '"]');
          if (!tagsInput) return;
          const known = new Set(KNOWN_TAGS.map((t) => t.toLowerCase()));
          const entered = tagsInput.value.split(",").map((t) => t.trim()).filter(Boolean);
          for (const tag of entered) {
            if (known.has(tag.toLowerCase())) continue;
            const suggestions = nearMatches(tag);
            const msg = suggestions.length
              ? '"' + tag + '" isn\\'t an existing tag. Did you mean: ' + suggestions.join(", ") + '? Click OK to create "' + tag + '" as a new tag anyway, or Cancel to fix it.'
              : 'Create new tag "' + tag + '"?';
            if (!confirm(msg)) {
              e.preventDefault();
              tagsInput.focus();
              return;
            }
          }
        });
      });
    </script>
  </body></html>`);
});

app.post("/admin/catalog/:id", requireAdmin, upload.single("thumbnail"), async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid item id."));

  const title = typeof req.body.title === "string" ? req.body.title.trim() : "";
  const category = req.body.category;
  const sortOrder = Number(req.body.sortOrder);
  const groupId = req.body.groupId ? Number(req.body.groupId) : null;
  const tags = typeof req.body.tags === "string" ? req.body.tags.split(",").map((t) => t.trim()).filter(Boolean) : [];
  if (!title) return res.redirect(303, "/admin?error=" + encodeURIComponent("Title can't be empty."));
  if (!CATALOG_CATEGORIES.includes(category)) return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid category."));
  if (!Number.isInteger(sortOrder) || sortOrder < 0) {
    return res.redirect(303, "/admin?error=" + encodeURIComponent("Sort order must be a non-negative integer."));
  }
  if (groupId !== null && (!Number.isInteger(groupId) || groupId <= 0)) {
    return res.redirect(303, "/admin?error=" + encodeURIComponent("Invalid group."));
  }

  const existing = await getCatalogItem(id);
  if (!existing) return res.redirect(303, "/admin?error=" + encodeURIComponent("Item not found."));

  const updated = await updateCatalogItem({ id, title, category, sortOrder, groupId });
  if (!updated) return res.redirect(303, "/admin?error=" + encodeURIComponent("Item not found."));

  await setItemTags(id, tags);

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

// Phase 5: admin-defined, admin-ordered sub-groupings within one category
// (e.g. "Pyramids" inside Diagrams), separate from the free-form tags
// above — a group is a single value with an explicit order, which tags
// deliberately don't try to be. One category's groups per page, matching
// the pattern of everything else in /admin being a plain per-row form.
app.get("/admin/groups", requireAdmin, async (req, res) => {
  const category = req.query.category;
  if (!CATALOG_CATEGORIES.includes(category)) {
    return res.status(400).send("Unknown category.");
  }
  const groups = await listGroupsForCategory(category);
  const errorMsg = typeof req.query.error === "string" ? req.query.error : null;
  const rows = groups
    .map(
      (g) => `
        <tr>
          <form id="edit-group-${g.id}" method="POST" action="/admin/groups/${g.id}"></form>
          <td><input form="edit-group-${g.id}" name="name" value="${escapeHtml(g.name)}"></td>
          <td><input form="edit-group-${g.id}" name="sortOrder" type="number" value="${g.sort_order}" style="width: 60px;"></td>
          <td><button form="edit-group-${g.id}" type="submit">Save</button></td>
          <td>
            <form method="POST" action="/admin/groups/${g.id}/delete" onsubmit="return confirm('Delete this group? Items in it become ungrouped.')">
              <button type="submit">Delete</button>
            </form>
          </td>
        </tr>`
    )
    .join("");
  res.send(`<!doctype html><html><body>
    <h1>Groups: ${escapeHtml(category)}</h1>
    <p><a href="/admin">&larr; Back to catalog</a></p>
    ${errorMsg ? `<p style="color: red;">${escapeHtml(errorMsg)}</p>` : ""}
    <table border="1" cellpadding="4">
      <tr><th>Name</th><th>Sort order</th><th></th><th></th></tr>
      ${rows}
    </table>
    <h2>Add group</h2>
    <form method="POST" action="/admin/groups">
      <input type="hidden" name="category" value="${escapeHtml(category)}">
      <input name="name" placeholder="Group name" required>
      <input name="sortOrder" type="number" value="0" style="width: 60px;">
      <button type="submit">Add</button>
    </form>
  </body></html>`);
});

app.post("/admin/groups", requireAdmin, async (req, res) => {
  const category = req.body.category;
  const name = typeof req.body.name === "string" ? req.body.name.trim() : "";
  const sortOrder = Number(req.body.sortOrder) || 0;
  if (!CATALOG_CATEGORIES.includes(category)) return res.status(400).send("Unknown category.");
  if (!name) {
    return res.redirect(303, `/admin/groups?category=${category}&error=` + encodeURIComponent("Group name can't be empty."));
  }
  try {
    await createGroup({ category, name, sortOrder });
  } catch (err) {
    return res.redirect(
      303,
      `/admin/groups?category=${category}&error=` + encodeURIComponent(`A group named "${name}" already exists in this category.`)
    );
  }
  res.redirect(303, `/admin/groups?category=${category}`);
});

app.post("/admin/groups/:id", requireAdmin, async (req, res) => {
  const id = Number(req.params.id);
  const name = typeof req.body.name === "string" ? req.body.name.trim() : "";
  const sortOrder = Number(req.body.sortOrder) || 0;
  // Category isn't submitted from this form (it's not editable — see
  // updateGroup's comment) — read it back from the row itself so we know
  // which category's page to redirect to.
  const existing = await getGroup(id);
  const category = existing?.category ?? CATALOG_CATEGORIES[0];
  if (!name) {
    return res.redirect(303, `/admin/groups?category=${category}&error=` + encodeURIComponent("Group name can't be empty."));
  }
  await updateGroup({ id, name, sortOrder });
  res.redirect(303, `/admin/groups?category=${category}`);
});

app.post("/admin/groups/:id/delete", requireAdmin, async (req, res) => {
  const id = Number(req.params.id);
  const existing = await getGroup(id);
  const category = existing?.category ?? CATALOG_CATEGORIES[0];
  await deleteGroup(id);
  res.redirect(303, `/admin/groups?category=${category}`);
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
