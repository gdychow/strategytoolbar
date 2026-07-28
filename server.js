const express = require("express");
const cookieParser = require("cookie-parser");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const https = require("https");
const os = require("os");
const crypto = require("crypto");

const {
  waitForDatabase,
  upsertUser,
  getUserByKey,
  completeRegistration,
  listUsersByCompanyDomain,
  setCompanyAdmin,
  listSharedCatalogItems,
  listCompanyCatalogItems,
  getCatalogItem,
  insertCatalogItem,
  updateCatalogItem,
  reorderCatalogItems,
  updateCatalogItemThumbnail,
  updateCatalogItemContent,
  deleteCatalogItem,
  listGroupsForCategory,
  listGroupsForCompany,
  getGroup,
  createGroup,
  updateGroup,
  reorderGroups,
  deleteGroup,
  listAllTagNames,
  setItemTags,
  listPersonalCatalogItems,
  getOwnedCatalogItem,
  updateOwnedCatalogItemContent,
  renameOwnedCatalogItem,
  deleteOwnedCatalogItem,
} = require("./server/db");
const { verifyMicrosoftIdToken, createSessionToken, verifySessionToken } = require("./server/auth");
const { THUMBNAILS_DIR, ADMIN_ADDED_DIR, PERSONAL_ADDED_DIR, CATALOG_CATEGORIES, resolveCatalogFilePath } = require("./server/catalog");
const { deriveCompanyDomain } = require("./server/consumerDomains");
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
// Default (100kb) is far too small for Task Pane Phase 12's admin
// content-save routes, whose JSON bodies carry a base64-encoded slide
// (Slide.exportAsBase64) plus a base64 thumbnail (Slide.getImageAsBase64)
// — raised globally rather than layered per-route, since body-parser
// middleware consumes the request stream once: a second express.json()
// later in the chain for just those routes would read nothing and
// silently produce an empty body, not add a second, larger limit.
app.use(express.json({ limit: "15mb" }));
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
    // Task Pane Phase 14/15: companyDomain/isCompanyAdmin/isRegistered are
    // re-fetched from the DB here (not just carried over from the old JWT
    // claims the way oid/tid/email/displayName are) so a promote/demote
    // or a just-completed registration takes effect within one refresh
    // cycle (~12h) instead of requiring a full sign-out — all three are
    // real mutable state, unlike isAdmin, which stays purely env-var-
    // derived and is safe to recompute cheaply inside createSessionToken
    // on every reissue.
    const dbUser = await getUserByKey(verified.claims.oid, verified.claims.tid);
    const fresh = await createSessionToken(
      {
        oid: verified.claims.oid,
        tid: verified.claims.tid,
        email: verified.claims.email,
        displayName: verified.claims.displayName,
        companyDomain: dbUser?.company_domain ?? null,
        isCompanyAdmin: dbUser?.is_company_admin ?? false,
        isRegistered: dbUser?.is_registered ?? false,
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

  const companyDomain = deriveCompanyDomain(identity.email);
  const user = await upsertUser({ ...identity, companyDomain });
  const sessionToken = await createSessionToken({
    ...identity,
    companyDomain: user.company_domain,
    isCompanyAdmin: user.is_company_admin,
    isRegistered: user.is_registered,
  });
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
  const { oid, tid, email, displayName, isAdmin, companyDomain, isCompanyAdmin, isRegistered } = req.user;
  res.json({ oid, tid, email, displayName, isAdmin, companyDomain, isCompanyAdmin, isRegistered });
});

// Task Pane Phase 15: completes registration for an already-signed-in but
// not-yet-registered user — deliberately doesn't require isRegistered
// (that's exactly what this sets). companyName is only required when the
// user actually has a companyDomain (consumer-domain users have no
// company step to fill in at all — see server/consumerDomains.js).
app.post("/api/auth/register", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Sign in first." });

  const fullName = typeof req.body?.fullName === "string" ? req.body.fullName.trim() : "";
  const companyName = typeof req.body?.companyName === "string" ? req.body.companyName.trim() : "";
  const jobTitle = typeof req.body?.jobTitle === "string" ? req.body.jobTitle.trim() : "";
  const plan = req.body?.plan;
  const termsAccepted = req.body?.termsAccepted === true;

  if (!fullName) return res.status(400).json({ error: "Full name is required." });
  if (req.user.companyDomain && !companyName) return res.status(400).json({ error: "Company name is required." });
  if (plan !== "monthly" && plan !== "annual") return res.status(400).json({ error: "Choose a plan." });
  if (!termsAccepted) return res.status(400).json({ error: "You must accept the Terms of Service." });

  const updated = await completeRegistration({
    oid: req.user.oid,
    tid: req.user.tid,
    fullName,
    companyName: req.user.companyDomain ? companyName : null,
    jobTitle: jobTitle || null,
    plan,
  });
  if (!updated) return res.status(404).json({ error: "Account not found." });

  // Reissues the cookie so the response of this very call already reflects
  // the new isRegistered/isCompanyAdmin state — the client doesn't have to
  // wait on the sliding-refresh cycle for its own registration to apply.
  const sessionToken = await createSessionToken(
    {
      oid: updated.oid,
      tid: updated.tid,
      email: updated.email,
      displayName: updated.display_name,
      companyDomain: updated.company_domain,
      isCompanyAdmin: updated.is_company_admin,
      isRegistered: updated.is_registered,
    },
    req.user.sessionStart
  );
  res.cookie(SESSION_COOKIE, sessionToken, cookieOptions);
  res.json({ ok: true, isCompanyAdmin: updated.is_company_admin });
});

app.post("/api/auth/signout", (req, res) => {
  res.clearCookie(SESSION_COOKIE, cookieOptions);
  res.status(204).end();
});

// Task Pane Phase 13: a signed-in user's own personal library — flat (no
// groups), owner-scoped instead of category-scoped. Registered *before*
// the /:category route below, same route-ordering fix already required
// once this session for /admin/catalog/reorder vs. /admin/catalog/:id —
// otherwise "personal" would be swallowed as a literal (and invalid)
// category value by the param route.
app.get("/api/catalog/personal", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const items = await listPersonalCatalogItems(req.user.oid, req.user.tid);
  res.json({
    groups: [],
    items: items.map((item) => ({
      id: item.id,
      title: item.title,
      insertMode: item.insert_mode,
      reconstructSpec: item.reconstruct_spec,
      unicodeChar: item.unicode_char,
      thumbnailUrl: item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null,
      groupId: item.group_id,
      groupName: item.group_name,
      tags: item.tags,
      ownerOid: req.user.oid,
      ownerTid: req.user.tid,
      companyDomain: null,
    })),
  });
});

// Task Pane Phase 14: a company member's shared library — any signed-in
// user with a companyDomain can browse (only company admins can add/edit,
// per the POST routes below). Registered before /:category for the same
// reason listPersonalCatalogItems's route is — "company" would otherwise
// be swallowed as a literal (invalid) category value.
app.get("/api/catalog/company", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  if (!req.user.companyDomain) return res.status(403).json({ error: "You don't have a company library." });

  const [items, groups] = await Promise.all([
    listCompanyCatalogItems(req.user.companyDomain),
    listGroupsForCompany(req.user.companyDomain),
  ]);
  res.json({
    groups: groups.map((g) => ({ id: g.id, name: g.name, sortOrder: g.sort_order })),
    items: items.map((item) => ({
      id: item.id,
      title: item.title,
      insertMode: item.insert_mode,
      reconstructSpec: item.reconstruct_spec,
      unicodeChar: item.unicode_char,
      thumbnailUrl: item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null,
      groupId: item.group_id,
      groupName: item.group_name,
      tags: item.tags,
      ownerOid: null,
      ownerTid: null,
      companyDomain: item.company_domain,
    })),
  });
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
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
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
      unicodeChar: item.unicode_char,
      thumbnailUrl: item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null,
      groupId: item.group_id,
      groupName: item.group_name,
      tags: item.tags,
      ownerOid: null,
      ownerTid: null,
      companyDomain: null,
    })),
  });
});

app.get("/api/catalog/file/:itemId", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });

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

// Task Pane Phase 12: JSON, called directly from the task pane (not a
// browser page) — see beginLibraryEdit/saveLibraryEdit in taskpane.ts.
// Replaces an existing item's underlying content with whatever the admin
// just exported from a live PowerPoint session (Slide.exportAsBase64 for
// the content, Slide.getImageAsBase64 for the thumbnail — both base64,
// no data: URL prefix). Always lands as insert_mode 'file', even if the
// item was 'reconstruct' before — updateCatalogItemContent forces that
// and clears reconstruct_spec. Task Pane Phase 14: fetches the existing
// row first and checks canManageRow against it, instead of the blanket
// requireAdmin middleware — a company admin can now reach this route too,
// for their own company's items.
app.post("/api/admin/catalog/:id/content", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Sign in first." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: "Invalid item id." });

  const { pptxBase64, thumbnailBase64 } = req.body ?? {};
  if (typeof pptxBase64 !== "string" || typeof thumbnailBase64 !== "string") {
    return res.status(400).json({ error: "Missing content." });
  }

  const existing = await getCatalogItem(id);
  if (!existing) return res.status(404).json({ error: "Item not found." });
  if (!canManageRow(req.user, existing)) return res.status(403).json({ error: "Not an admin for this item." });

  const sourceFile = `admin-added/item-${id}.pptx`;
  const thumbnailPath = `item-${id}.png`;
  await fs.promises.writeFile(path.join(ADMIN_ADDED_DIR, `item-${id}.pptx`), Buffer.from(pptxBase64, "base64"));
  await fs.promises.writeFile(path.join(THUMBNAILS_DIR, thumbnailPath), Buffer.from(thumbnailBase64, "base64"));

  const updated = await updateCatalogItemContent({ id, sourceFile, thumbnailPath });
  if (!updated) return res.status(404).json({ error: "Item not found." });

  // Cleans up whatever this item pointed at before — a Python-pipeline
  // category-prefixed file, a previous admin-added file, or (for a
  // 'reconstruct' item being migrated) nothing at all for source_file.
  if (existing.source_file && existing.source_file !== sourceFile) {
    await fs.promises.unlink(resolveCatalogFilePath(existing.source_file)).catch(() => {});
  }
  if (existing.thumbnail_path && existing.thumbnail_path !== thumbnailPath) {
    await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
  }

  res.json({ ok: true });
});

// JSON, called from the task pane — see addSelectedSlideToLibrary in
// taskpane.ts. Creates a brand-new item from whatever slide the admin
// currently has selected. Filenames are keyed by a fresh UUID rather than
// the not-yet-known row id, since the row can't be inserted until
// source_file already has a value (catalog_items' own CHECK constraint
// requires a 'file'-mode row to have one). Lands with a placeholder title
// (and, for the global scope, a placeholder category) — the admin
// finishes naming/categorizing it in /admin, which the task pane opens
// immediately after this succeeds. Task Pane Phase 14: `scope` in the
// body picks which of the two admin-gated destinations this goes to —
// "global" (today's only behavior, requires req.user.isAdmin) or
// "company" (requires req.user.isCompanyAdmin, lands with
// company_domain set and no category at all, per the schema's
// orthogonal-scoping-columns design).
app.post("/api/admin/catalog", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Sign in first." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const { pptxBase64, thumbnailBase64, scope } = req.body ?? {};
  if (typeof pptxBase64 !== "string" || typeof thumbnailBase64 !== "string") {
    return res.status(400).json({ error: "Missing content." });
  }
  const isCompanyScope = scope === "company";
  if (isCompanyScope) {
    if (!req.user.isCompanyAdmin || !req.user.companyDomain) return res.status(403).json({ error: "Not a company admin." });
  } else if (!req.user.isAdmin) {
    return res.status(403).json({ error: "Not an admin." });
  }

  const category = isCompanyScope ? null : CATALOG_CATEGORIES[0];
  const companyDomain = isCompanyScope ? req.user.companyDomain : null;
  const fileId = crypto.randomUUID();
  const sourceFile = `admin-added/${fileId}.pptx`;
  const thumbnailPath = `${fileId}.png`;
  await fs.promises.writeFile(path.join(ADMIN_ADDED_DIR, `${fileId}.pptx`), Buffer.from(pptxBase64, "base64"));
  await fs.promises.writeFile(path.join(THUMBNAILS_DIR, thumbnailPath), Buffer.from(thumbnailBase64, "base64"));

  const created = await insertCatalogItem({
    category,
    companyDomain,
    title: "Untitled item",
    insertMode: "file",
    sourceFile,
    thumbnailPath,
    sortOrder: 0,
  });

  res.json({ id: created.id, category, companyDomain });
});

// Task Pane Phase 13: any signed-in user, no admin check — see
// addSelectedSlideToLibrary in taskpane.ts with target "personal". Mirrors
// POST /api/admin/catalog exactly, but writes into PERSONAL_ADDED_DIR and
// stamps owner_oid/owner_tid instead of leaving them null.
app.post("/api/personal/catalog", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });

  const { pptxBase64, thumbnailBase64 } = req.body ?? {};
  if (typeof pptxBase64 !== "string" || typeof thumbnailBase64 !== "string") {
    return res.status(400).json({ error: "Missing content." });
  }

  const category = CATALOG_CATEGORIES[0];
  const fileId = crypto.randomUUID();
  const sourceFile = `personal-added/${fileId}.pptx`;
  const thumbnailPath = `${fileId}.png`;
  await fs.promises.writeFile(path.join(PERSONAL_ADDED_DIR, `${fileId}.pptx`), Buffer.from(pptxBase64, "base64"));
  await fs.promises.writeFile(path.join(THUMBNAILS_DIR, thumbnailPath), Buffer.from(thumbnailBase64, "base64"));

  const created = await insertCatalogItem({
    category,
    title: "Untitled item",
    insertMode: "file",
    sourceFile,
    thumbnailPath,
    sortOrder: 0,
    ownerOid: req.user.oid,
    ownerTid: req.user.tid,
  });

  res.json({ id: created.id });
});

// Owner-scoped equivalent of POST /api/admin/catalog/:id/content — the
// existing-row fetch and the update both filter by ownership instead of
// admin status (see getOwnedCatalogItem/updateOwnedCatalogItemContent).
app.post("/api/personal/catalog/:id/content", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: "Invalid item id." });

  const { pptxBase64, thumbnailBase64 } = req.body ?? {};
  if (typeof pptxBase64 !== "string" || typeof thumbnailBase64 !== "string") {
    return res.status(400).json({ error: "Missing content." });
  }

  const existing = await getOwnedCatalogItem(id, req.user.oid, req.user.tid);
  if (!existing) return res.status(404).json({ error: "Item not found." });

  const sourceFile = `personal-added/item-${id}.pptx`;
  const thumbnailPath = `item-${id}.png`;
  await fs.promises.writeFile(path.join(PERSONAL_ADDED_DIR, `item-${id}.pptx`), Buffer.from(pptxBase64, "base64"));
  await fs.promises.writeFile(path.join(THUMBNAILS_DIR, thumbnailPath), Buffer.from(thumbnailBase64, "base64"));

  const updated = await updateOwnedCatalogItemContent({ id, oid: req.user.oid, tid: req.user.tid, sourceFile, thumbnailPath });
  if (!updated) return res.status(404).json({ error: "Item not found." });

  if (existing.source_file && existing.source_file !== sourceFile) {
    await fs.promises.unlink(resolveCatalogFilePath(existing.source_file)).catch(() => {});
  }
  if (existing.thumbnail_path && existing.thumbnail_path !== thumbnailPath) {
    await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
  }

  res.json({ ok: true });
});

// Rename-only — personal items have no /admin-equivalent page, so this is
// their only metadata edit (no category/group/tags, unlike shared items).
app.post("/api/personal/catalog/:id", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: "Invalid item id." });

  const title = typeof req.body?.title === "string" ? req.body.title.trim() : "";
  if (!title) return res.status(400).json({ error: "Title is required." });

  const updated = await renameOwnedCatalogItem({ id, oid: req.user.oid, tid: req.user.tid, title });
  if (!updated) return res.status(404).json({ error: "Item not found." });
  res.json({ ok: true });
});

app.post("/api/personal/catalog/:id/delete", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Not signed in." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: "Invalid item id." });

  const existing = await getOwnedCatalogItem(id, req.user.oid, req.user.tid);
  if (!existing) return res.status(404).json({ error: "Item not found." });

  const deleted = await deleteOwnedCatalogItem(id, req.user.oid, req.user.tid);
  if (!deleted) return res.status(404).json({ error: "Item not found." });

  if (existing.source_file) await fs.promises.unlink(resolveCatalogFilePath(existing.source_file)).catch(() => {});
  if (existing.thumbnail_path) await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});

  res.json({ ok: true });
});

function requireAdmin(req, res, next) {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isRegistered) return res.status(403).send("Finish creating your account first.");
  if (!req.user.isAdmin) return res.status(403).send("Not an admin.");
  next();
}

// Task Pane Phase 14: a global admin can manage any company; a company
// admin only their own. Plain boolean helper rather than Express
// middleware — every call site needs to resolve `domain` per-request
// (from a query param, or from the target user's own row), so there's no
// single domain to bind a middleware factory to ahead of time.
function canAdminCompany(user, domain) {
  if (!user || !domain) return false;
  if (user.isAdmin) return true;
  return !!user.isCompanyAdmin && user.companyDomain === domain;
}

/**
 * Whether the signed-in user can manage a specific already-fetched
 * catalog item or group row — a global admin can touch anything; a
 * company admin only rows belonging to their own company; nobody else can
 * touch a global (non-company) row. Used by routes that fetch the
 * existing row first (to know its scope) before deciding, since
 * requireAdmin's blanket global-only check no longer applies now that
 * company-scoped rows exist alongside the shared catalog. Deliberately
 * NOT built on top of canAdminCompany — that helper treats a falsy domain
 * as "nothing to check, deny" (right for its own callers, which always
 * have a real domain in hand), but here a falsy company_domain means "this
 * is a global row," which a global admin must still be allowed to manage.
 */
function canManageRow(user, row) {
  if (!user) return false;
  if (user.isAdmin) return true;
  if (row?.company_domain) return !!user.isCompanyAdmin && user.companyDomain === row.company_domain;
  return false;
}

// Lists everyone at one company_domain with a promote/demote button per
// row — reachable from GET /admin's company-scope nav (see below).
// domain defaults to the viewer's own company when they aren't a global
// admin, so a company admin can't even construct a URL into a domain
// they don't belong to (canAdminCompany would 403 it anyway; this just
// avoids showing a dead link).
app.get("/admin/company-admins", async (req, res) => {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isRegistered) return res.status(403).send("Finish creating your account first.");
  const domain = typeof req.query.domain === "string" && req.query.domain ? req.query.domain : req.user.companyDomain;
  if (!domain) return res.status(400).send("No company to show.");
  if (!canAdminCompany(req.user, domain)) return res.status(403).send("Not an admin for this company.");

  const users = await listUsersByCompanyDomain(domain);
  const rows = users
    .map(
      (u) => `
      <tr>
        <td>${escapeHtml(u.email ?? "(no email)")}</td>
        <td>${u.is_company_admin ? "Yes" : "No"}</td>
        <td>
          <form method="POST" action="/admin/company-admins/${encodeURIComponent(u.oid)}/${encodeURIComponent(u.tid)}/${u.is_company_admin ? "demote" : "promote"}">
            <button type="submit">${u.is_company_admin ? "Demote" : "Promote"}</button>
          </form>
        </td>
      </tr>`
    )
    .join("");

  res.send(`<!doctype html><html><head><style>${ADMIN_STYLE}</style></head><body>
    <h1>Company Admins</h1>
    <h2>${escapeHtml(domain)}</h2>
    <p><a href="/admin?scope=company:${encodeURIComponent(domain)}">&larr; Back to library</a></p>
    <table>
      <thead><tr><th>Email</th><th>Company Admin</th><th></th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </body></html>`);
});

app.post("/admin/company-admins/:oid/:tid/promote", async (req, res) => {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isRegistered) return res.status(403).send("Finish creating your account first.");
  const { oid, tid } = req.params;
  const target = await getUserByKey(oid, tid);
  if (!target || !target.company_domain) return res.status(404).send("User not found.");
  if (!canAdminCompany(req.user, target.company_domain)) return res.status(403).send("Not an admin for this company.");

  await setCompanyAdmin(oid, tid, true);
  res.redirect(303, `/admin/company-admins?domain=${encodeURIComponent(target.company_domain)}`);
});

// The self-removal guard lives here, server-side — not just a disabled
// button client-side — so a company can't be talked down to zero admins
// by its own admins via a raw request. Global admins are exempt: they're
// the backstop if a company ever does end up with none.
app.post("/admin/company-admins/:oid/:tid/demote", async (req, res) => {
  if (!req.user) return res.status(401).send("Sign in first.");
  if (!req.user.isRegistered) return res.status(403).send("Finish creating your account first.");
  const { oid, tid } = req.params;
  const target = await getUserByKey(oid, tid);
  if (!target || !target.company_domain) return res.status(404).send("User not found.");
  if (!canAdminCompany(req.user, target.company_domain)) return res.status(403).send("Not an admin for this company.");
  if (oid === req.user.oid && tid === req.user.tid && !req.user.isAdmin) {
    return res.status(403).send("You can't remove your own company-admin status — ask another company admin or a global admin.");
  }

  await setCompanyAdmin(oid, tid, false);
  res.redirect(303, `/admin/company-admins?domain=${encodeURIComponent(target.company_domain)}`);
});

function escapeHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/** Appends ?error=/&error= to a redirect target, correctly whichever the target already has a query string (e.g. "/admin?category=text") or not (plain "/admin"). */
function redirectWithError(res, base, message) {
  res.redirect(303, base + (base.includes("?") ? "&" : "?") + "error=" + encodeURIComponent(message));
}

// Admin UI Phase 11: every admin edit form now submits via fetch (see the
// saveCards/group-management inline script in GET /admin) so it can save
// without a full page reload and report per-card success/failure — but
// each route keeps its original redirect-based behavior too, as a
// harmless fallback for anything that still posts here without this
// header (there's nothing left that does, but it costs nothing to keep).
function wantsJson(req) {
  return req.get("Accept") === "application/json";
}

// insert_mode is informational-only on a catalog item's card (see
// GET /admin) — never user-editable — but the raw DB value ("file" vs.
// "reconstruct" vs. "unicode-char") means nothing to an admin without
// context, so it's shown as a friendly label with an explanatory tooltip
// instead. "file" mode's tooltip in particular flags a real, deliberate
// limitation: the thumbnail file input below it only ever replaces the
// preview image, never the underlying inserted content — replacing that
// is tracked as its own future feature, not built here.
const MODE_INFO = {
  file: {
    label: "File",
    title:
      "Backed by a stored file. The image below only replaces the thumbnail preview — replacing the underlying content isn't built yet.",
  },
  reconstruct: { label: "Reconstructed", title: "Built from stored shape properties, not a file." },
  "unicode-char": { label: "Character", title: "A single character inserted at the cursor — no file or shape." },
};

// Shared by every /admin* page (Admin UI Phase 10) — /admin stays plain
// server-rendered HTML with no bundler, so this is one inline <style>
// block reused verbatim rather than a separate stylesheet file/build step.
const ADMIN_STYLE = `
  :root {
    --bg: #ffffff;
    --bg-alt: #fafafa;
    --bg-hover: #f0f0f0;
    --bg-subtle: #f5f5f5;
    --text: #222;
    --text-muted: #666;
    --text-faint: #999;
    --heading: #444;
    --border: #c8c8c8;
    --border-light: #ddd;
    --border-lighter: #eee;
    --accent: #1a5fb4;
    --accent-shadow: rgba(26, 95, 180, 0.6);
    --error-bg: #fdeaea;
    --error-text: #a12525;
    --success-text: #1a7a3c;
    --card-bg: #fff;
    --nav-text: #333;
    --disabled-bg: #f0f0f0;
    --disabled-border: #ccc;
    --disabled-text: #999;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #1e1e1e;
      --bg-alt: #2b2b2b;
      --bg-hover: #383838;
      --bg-subtle: #333;
      --text: #e8e8e8;
      --text-muted: #aaa;
      --text-faint: #888;
      --heading: #ccc;
      --border: #555;
      --border-light: #3a3a3a;
      --border-lighter: #333;
      --accent: #4c8fd6;
      --accent-shadow: rgba(76, 143, 214, 0.6);
      --error-bg: #4a2222;
      --error-text: #ff8f8f;
      --success-text: #4ade80;
      --card-bg: #262626;
      --nav-text: #ddd;
      --disabled-bg: #333;
      --disabled-border: #555;
      --disabled-text: #888;
    }
  }
  body { font-family: -apple-system, "Segoe UI", sans-serif; margin: 0; padding: 16px 24px; color: var(--text); background: var(--bg); color-scheme: light dark; }
  h1 { font-size: 20px; margin: 0 0 4px; }
  h2 { font-size: 15px; margin: 0 0 10px; }
  a { color: var(--accent); }
  .admin-error { color: var(--error-text); background: var(--error-bg); padding: 8px 10px; border-radius: 4px; }
  .admin-reorder-status { font-size: 12px; color: var(--success-text); margin-left: 8px; }
  .admin-category-nav { display: flex; flex-wrap: wrap; gap: 4px; margin: 14px 0; padding: 0; list-style: none; }
  .admin-category-nav a { padding: 6px 12px; border: 1px solid var(--border); border-radius: 4px; text-decoration: none; color: var(--nav-text); font-size: 13px; text-transform: capitalize; }
  .admin-category-nav a.active { background: var(--accent); color: #fff; border-color: var(--accent); }

  /* Sticky single Save bar (Admin UI Phase 11) — bleeds past body's own padding to span the full width while staying pinned during scroll. */
  .admin-save-bar {
    position: sticky; top: 0; z-index: 1000;
    display: flex; align-items: center; justify-content: flex-end; gap: 12px;
    margin: -16px -24px 16px; padding: 10px 24px;
    background: var(--bg-alt); border-bottom: 1px solid var(--border-light); box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
  }
  #adminSaveStatus { font-size: 12px; color: var(--text-muted); }
  #adminSaveAll { padding: 6px 16px; font-size: 13px; font-weight: 600; border: 1px solid var(--accent); border-radius: 4px; background: var(--accent); color: #fff; cursor: pointer; }
  #adminSaveAll:disabled { background: var(--disabled-bg); border-color: var(--disabled-border); color: var(--disabled-text); cursor: default; }

  .admin-cluster { margin-bottom: 8px; }
  .admin-cluster.dragging { opacity: 0.4; }
  .admin-group-heading-row { display: flex; align-items: center; gap: 6px; margin: 26px 0 10px; }
  .admin-cluster:first-of-type .admin-group-heading-row { margin-top: 18px; }
  .admin-group-drag { cursor: grab; color: var(--text-faint); font-size: 13px; user-select: none; line-height: 1; }
  .admin-group-drag:active { cursor: grabbing; }
  .admin-group-heading-input, .admin-group-heading-static {
    font-size: 13px; font-weight: 600; color: var(--heading); text-transform: uppercase; letter-spacing: 0.02em;
  }
  .admin-group-heading-input {
    border: 1px solid transparent; background: none; color: var(--heading); padding: 2px 4px; border-radius: 3px; flex: 0 1 auto; min-width: 40px;
  }
  .admin-group-heading-input:hover { background: var(--bg-subtle); }
  .admin-group-heading-input:focus { border-color: var(--border); background: var(--card-bg); outline: none; }
  .admin-group-heading-ungrouped { color: var(--text-faint); }
  .admin-group-delete {
    border: none; background: none; color: var(--error-text); font-size: 15px; line-height: 1; cursor: pointer; padding: 0 4px;
  }

  .admin-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }
  .admin-card-wrap { position: relative; border: 1px solid var(--border); border-radius: 6px; padding: 8px; background: var(--card-bg); display: flex; flex-direction: column; gap: 4px; }
  .admin-card-wrap.dragging { opacity: 0.4; }
  .admin-card-wrap.dirty { border-color: var(--accent); box-shadow: inset 3px 0 0 var(--accent); }
  @keyframes admin-card-highlight-pulse { from { box-shadow: 0 0 0 3px var(--accent-shadow); } to { box-shadow: 0 0 0 3px rgba(26, 95, 180, 0); } }
  .admin-card-wrap.highlight { animation: admin-card-highlight-pulse 2s ease-out 1; }
  .admin-card-drag { position: absolute; top: 4px; right: 6px; cursor: grab; color: var(--text-faint); font-size: 13px; user-select: none; line-height: 1; }
  .admin-card-drag:active { cursor: grabbing; }
  .admin-card { display: flex; flex-direction: column; gap: 4px; }
  .admin-card-thumb { width: 100%; height: 100px; object-fit: contain; background: var(--bg-alt); border: 1px solid var(--border-lighter); border-radius: 4px; cursor: zoom-in; }
  .admin-card-thumb-empty { display: flex; align-items: center; justify-content: center; color: var(--text-faint); font-size: 11px; cursor: default; }
  .admin-card-title, .admin-card-tags, .admin-card select, .admin-card input[type="file"] {
    font-size: 11px; padding: 3px 5px; border: 1px solid var(--border); border-radius: 3px; width: 100%; box-sizing: border-box;
    background: var(--card-bg); color: var(--text);
  }
  .admin-card-mode { font-size: 10px; color: var(--text-faint); text-transform: uppercase; align-self: flex-start; cursor: help; }
  .admin-card-error { font-size: 10px; color: var(--error-text); }
  .admin-card-delete button { font-size: 11px; padding: 4px 8px; cursor: pointer; border: 1px solid var(--border); border-radius: 3px; background: var(--bg-alt); color: var(--text); width: 100%; }
  .admin-card-delete { margin-top: -2px; }
  .admin-lightbox { display: none; position: fixed; inset: 0; background: rgba(0, 0, 0, 0.75); align-items: center; justify-content: center; z-index: 2000; cursor: zoom-out; }
  .admin-lightbox img { max-width: 90vw; max-height: 90vh; }
  table { border-collapse: collapse; color: var(--text); }
  table input, table select { font-size: 12px; padding: 3px 5px; background: var(--card-bg); color: var(--text); border: 1px solid var(--border); }
`;

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
  // Task Pane Phase 15: registration is required uniformly — even for a
  // global admin (ADMIN_EMAILS membership doesn't bypass it) — so this
  // stays a single point of enforcement rather than an exception. No
  // second registration form is built for the browser context; this just
  // points back to the one that exists, in the task pane.
  if (!req.user.isRegistered) {
    return res
      .status(403)
      .send("Your account isn't fully set up yet — open the task pane in PowerPoint, click \"Create Account\" to finish, then reload this page.");
  }

  // Task Pane Phase 14: ?scope=company:<domain> is the only new value this
  // takes — absent (or any other value) falls through to today's exact
  // ?category=-based behavior, unchanged. Kept as a separate query param
  // from ?category= rather than folding category into ?scope= too, since
  // every existing category link/reference stays valid with zero changes.
  const scopeParam = typeof req.query.scope === "string" ? req.query.scope : null;
  const companyScopeDomain = scopeParam && scopeParam.startsWith("company:") ? scopeParam.slice("company:".length) : null;
  const isCompanyScope = !!companyScopeDomain;

  if (isCompanyScope) {
    if (!canAdminCompany(req.user, companyScopeDomain)) return res.status(403).send("Not an admin for this company.");
  } else if (!req.user.isAdmin) {
    return res.status(403).send("Not an admin.");
  }

  const category = isCompanyScope
    ? null
    : CATALOG_CATEGORIES.includes(req.query.category)
      ? req.query.category
      : CATALOG_CATEGORIES[0];

  const [items, groups, tagNames] = await Promise.all([
    isCompanyScope ? listCompanyCatalogItems(companyScopeDomain) : listSharedCatalogItems(category),
    isCompanyScope ? listGroupsForCompany(companyScopeDomain) : listGroupsForCategory(category),
    listAllTagNames(),
  ]);

  // Cluster items by group, in each group's own sort_order, ungrouped
  // items trailing under "Ungrouped" — mirrors how the gallery
  // (src/gallery/gallery.ts) clusters the same data for end users, just
  // computed server-side here against the raw DB rows instead of the
  // client-side /api/catalog/:category JSON. Every group gets a section
  // even if currently empty (Admin UI Phase 11) — a newly-created group
  // needs somewhere to be renamed/deleted/dropped into before any item
  // has actually been saved into it.
  const clustersById = new Map(groups.map((g) => [g.id, { id: g.id, name: g.name, items: [] }]));
  const ungrouped = { id: null, name: "Ungrouped", items: [] };
  for (const item of items) {
    (clustersById.get(item.group_id) ?? ungrouped).items.push(item);
  }
  const clusters = [...groups.map((g) => clustersById.get(g.id)), ungrouped];

  const categoryOptions = CATALOG_CATEGORIES.map((c) => `<option value="${c}"${c === category ? " selected" : ""}>${c}</option>`).join("");
  const groupOptionsFor = (currentGroupId) =>
    `<option value="">(none)</option>` +
    groups.map((g) => `<option value="${g.id}"${g.id === currentGroupId ? " selected" : ""}>${escapeHtml(g.name)}</option>`).join("") +
    `<option value="__new__">+ Add new group…</option>`;

  function renderCard(item) {
    const thumbUrl = item.thumbnail_path ? `/assets/catalog/thumbnails/${item.thumbnail_path}?v=${ASSET_VERSION}` : null;
    const mode = MODE_INFO[item.insert_mode];
    return `
      <div class="admin-card-wrap" data-item-id="${item.id}">
        <div class="admin-card-drag" draggable="true" title="Drag to reorder">⠿</div>
        <form class="admin-card" method="POST" action="/admin/catalog/${item.id}" enctype="multipart/form-data">
          ${
            thumbUrl
              ? `<img class="admin-card-thumb" src="${thumbUrl}" data-full="${thumbUrl}" alt="">`
              : `<div class="admin-card-thumb admin-card-thumb-empty">(none)</div>`
          }
          <input class="admin-card-title" name="title" value="${escapeHtml(item.title)}">
          ${isCompanyScope ? "" : `<select name="category">${categoryOptions}</select>`}
          <select name="groupId">${groupOptionsFor(item.group_id)}</select>
          <input class="admin-card-tags tags-input" name="tags" placeholder="tags" value="${escapeHtml((item.tags || []).join(", "))}">
          <span class="admin-card-mode" title="${escapeHtml(mode?.title ?? "")}">${escapeHtml(mode?.label ?? item.insert_mode)}</span>
          <input name="thumbnail" type="file" accept="image/png,image/jpeg,image/webp,image/gif" title="Replaces the thumbnail preview image only — not the underlying inserted content.">
          <span class="admin-card-error"></span>
        </form>
        <form class="admin-card-delete" method="POST" action="/admin/catalog/${item.id}/delete">
          <button type="submit">Delete</button>
        </form>
      </div>`;
  }

  function renderCluster(cluster) {
    const isReal = cluster.id !== null;
    const headingRow = isReal
      ? `<div class="admin-group-heading-row">
           <span class="admin-group-drag" draggable="true" title="Drag to reorder group">⠿</span>
           <input class="admin-group-heading-input" data-group-id="${cluster.id}" value="${escapeHtml(cluster.name)}">
           <button type="button" class="admin-group-delete" data-group-id="${cluster.id}" title="Delete group">×</button>
         </div>`
      : `<div class="admin-group-heading-row admin-group-heading-ungrouped">
           <span class="admin-group-heading-static">${escapeHtml(cluster.name)}</span>
         </div>`;
    return `
      <section class="admin-cluster" data-group-id="${cluster.id ?? ""}">
        ${headingRow}
        <div class="admin-grid" data-group-id="${cluster.id ?? ""}">
          ${cluster.items.map(renderCard).join("")}
        </div>
      </section>`;
  }

  const clusterSections = clusters.map(renderCluster).join("");

  const categoryNav = CATALOG_CATEGORIES.map(
    (c) => `<a href="/admin?category=${c}"${!isCompanyScope && c === category ? ' class="active"' : ""}>${c}</a>`
  ).join("");
  // Task Pane Phase 14: one extra nav entry for the viewer's own company,
  // shown only when they're that company's admin (or a global admin) —
  // reaching a *different* company's library than your own is possible
  // for a global admin via a direct ?scope=company:<domain> URL, but not
  // surfaced in the nav (no "list every company" picker built for v1).
  const companyNavLink =
    req.user.companyDomain && canAdminCompany(req.user, req.user.companyDomain)
      ? `<a href="/admin?scope=company:${encodeURIComponent(req.user.companyDomain)}"${isCompanyScope ? ' class="active"' : ""}>${escapeHtml(req.user.companyDomain)}</a>`
      : "";
  const errorMsg = typeof req.query.error === "string" ? req.query.error : null;
  // Set by taskpane.ts's addSelectedSlideToLibrary after creating a new
  // item — lets this page scroll straight to it instead of leaving the
  // admin to hunt for "Untitled item" in a long, unsorted category.
  const highlightId = Number(req.query.highlight);
  const highlightItemId = Number.isInteger(highlightId) && highlightId > 0 ? highlightId : null;

  res.send(`<!doctype html><html><head><style>${ADMIN_STYLE}</style></head><body>
    <div id="adminSaveBar" class="admin-save-bar">
      <span id="adminSaveStatus">No changes</span>
      <button id="adminSaveAll" type="button" disabled>Save</button>
    </div>
    <h1>Welcome, admin</h1>
    <p>Signed in as ${escapeHtml(req.user.email)}.<span id="admin-reorder-status" class="admin-reorder-status"></span></p>
    ${isCompanyScope ? `<p>Viewing ${escapeHtml(companyScopeDomain)}'s library. <a href="/admin/company-admins?domain=${encodeURIComponent(companyScopeDomain)}">Manage company admins</a></p>` : ""}
    ${errorMsg ? `<p class="admin-error">${escapeHtml(errorMsg)}</p>` : ""}
    <nav class="admin-category-nav">${categoryNav}${companyNavLink}</nav>
    <div id="adminClusters">${clusterSections}</div>
    <div id="adminLightbox" class="admin-lightbox"><img id="adminLightboxImg" alt=""></div>
    <script>
      const CURRENT_SCOPE = ${JSON.stringify(isCompanyScope ? { companyDomain: companyScopeDomain } : { category })};
      const HIGHLIGHT_ITEM_ID = ${JSON.stringify(highlightItemId)};

      // Client-side typo-catching only, not a security boundary — the
      // server get-or-creates whatever tag names it's sent regardless
      // (see POST /admin/catalog/:id). This just makes an admin pause
      // before accidentally creating "arrows" next to an existing
      // "arrow". KNOWN_TAGS is embedded at render time rather than
      // fetched separately — the whole vocabulary is small (tens to a
      // couple hundred entries across ~230 items) — and grows in place as
      // tags get confirmed within a session, so a repeat mention of the
      // same new tag across several dirty cards only prompts once.
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

      // Checks every given form's tags field, prompting once per *unique*
      // unknown tag across the whole batch (not once per card) — so the
      // same typo appearing in five dirty cards doesn't fire five
      // confirm() dialogs. Returns false (abort the whole save, nothing
      // submitted) if any prompt is cancelled.
      function confirmTagsForForms(forms) {
        const known = new Set(KNOWN_TAGS.map((t) => t.toLowerCase()));
        const unique = new Map();
        for (const form of forms) {
          const tagsInput = form.querySelector("input.tags-input");
          if (!tagsInput) continue;
          for (const tag of tagsInput.value.split(",").map((t) => t.trim()).filter(Boolean)) {
            if (!known.has(tag.toLowerCase()) && !unique.has(tag.toLowerCase())) unique.set(tag.toLowerCase(), tag);
          }
        }
        for (const tag of unique.values()) {
          const suggestions = nearMatches(tag);
          const msg = suggestions.length
            ? '"' + tag + '" isn\\'t an existing tag. Did you mean: ' + suggestions.join(", ") + '? Click OK to create "' + tag + '" as a new tag anyway, or Cancel to fix it.'
            : 'Create new tag "' + tag + '"?';
          if (!confirm(msg)) return false;
          KNOWN_TAGS.push(tag);
        }
        return true;
      }

      // ---- Dirty tracking + single global Save (Admin UI Phase 11) ----
      // Every card's own Save button is gone — editing several items no
      // longer means saving them one at a time. groupId selects are
      // handled separately below (see "+ Add new group…"), since a
      // transient "__new__" selection shouldn't count as a real edit if
      // the admin cancels the create-group prompt.
      const dirtyIds = new Set();
      const saveStatus = document.getElementById("adminSaveStatus");
      const saveAllBtn = document.getElementById("adminSaveAll");

      function updateSaveBar() {
        if (dirtyIds.size === 0) {
          saveStatus.textContent = "No changes";
          saveAllBtn.disabled = true;
        } else {
          saveStatus.textContent = dirtyIds.size + " unsaved change" + (dirtyIds.size === 1 ? "" : "s");
          saveAllBtn.disabled = false;
        }
      }
      function markDirty(wrap) {
        dirtyIds.add(Number(wrap.dataset.itemId));
        wrap.classList.add("dirty");
        updateSaveBar();
      }
      function clearDirty(wrap) {
        dirtyIds.delete(Number(wrap.dataset.itemId));
        wrap.classList.remove("dirty");
        updateSaveBar();
      }

      document.querySelectorAll(".admin-card").forEach((form) => {
        const wrap = form.closest(".admin-card-wrap");
        form.querySelectorAll("input, select").forEach((field) => {
          if (field.name === "groupId") return;
          field.addEventListener("input", () => markDirty(wrap));
          field.addEventListener("change", () => markDirty(wrap));
        });
      });

      window.addEventListener("beforeunload", (e) => {
        if (dirtyIds.size === 0) return;
        e.preventDefault();
        e.returnValue = "";
      });

      async function saveCards(forms) {
        if (forms.length === 0) return;
        if (!confirmTagsForForms(forms)) return;

        const results = await Promise.allSettled(
          forms.map((form) =>
            fetch(form.action, { method: "POST", body: new FormData(form), headers: { Accept: "application/json" } }).then(
              async (res) => {
                const body = await res.json().catch(() => ({}));
                if (!res.ok) throw new Error(body.error || "HTTP " + res.status);
                return body;
              }
            )
          )
        );

        let succeeded = 0;
        let failed = 0;
        forms.forEach((form, i) => {
          const wrap = form.closest(".admin-card-wrap");
          const errorEl = wrap.querySelector(".admin-card-error");
          const result = results[i];
          if (result.status === "fulfilled") {
            succeeded++;
            clearDirty(wrap);
            if (errorEl) errorEl.textContent = "";
            const fileInput = form.querySelector('input[type="file"]');
            if (fileInput) fileInput.value = "";
            if (result.value.thumbnailUrl) {
              let thumb = wrap.querySelector(".admin-card-thumb");
              if (thumb.tagName !== "IMG") {
                const img = document.createElement("img");
                img.className = "admin-card-thumb";
                img.alt = "";
                thumb.replaceWith(img);
                thumb = img;
              }
              thumb.src = result.value.thumbnailUrl;
              thumb.dataset.full = result.value.thumbnailUrl;
            }
          } else {
            failed++;
            if (errorEl) errorEl.textContent = result.reason.message;
          }
        });

        saveStatus.textContent = failed === 0 ? "Saved " + succeeded + "." : "Saved " + succeeded + ", " + failed + " failed.";
        setTimeout(updateSaveBar, 2500);
      }

      document.querySelectorAll(".admin-card").forEach((form) => {
        form.addEventListener("submit", (e) => {
          e.preventDefault();
          saveCards([form]);
        });
      });
      saveAllBtn.addEventListener("click", () => {
        const dirtyForms = [...document.querySelectorAll(".admin-card-wrap.dirty")].map((wrap) => wrap.querySelector(".admin-card"));
        saveCards(dirtyForms);
      });

      // ---- Delete (fetch-based — a full navigation here would also
      // trip the beforeunload guard above if other cards are still
      // dirty, which would be a confusing prompt for an unrelated
      // action) ----
      document.querySelectorAll(".admin-card-delete").forEach((form) => {
        form.addEventListener("submit", (e) => {
          e.preventDefault();
          if (!confirm("Delete this catalog item permanently?")) return;
          fetch(form.action, { method: "POST", headers: { Accept: "application/json" } })
            .then((res) => {
              if (!res.ok) throw new Error("HTTP " + res.status);
              const wrap = form.closest(".admin-card-wrap");
              dirtyIds.delete(Number(wrap.dataset.itemId));
              updateSaveBar();
              wrap.remove();
            })
            .catch((err) => alert("Couldn't delete: " + err.message));
        });
      });

      // ---- Item drag-and-drop reorder (Admin UI Phase 10) — generalizes
      // the task pane's own single-list section-reorder pattern
      // (initSectionReordering in src/taskpane/taskpane.ts) to multiple
      // drop containers, one per group cluster: dragging a card into a
      // *different* cluster's .admin-grid both reorders it and reassigns
      // its group on drop. Persists immediately via fetch, no save step. ----
      let draggedCard = null;

      document.querySelectorAll(".admin-card-drag").forEach((handle) => {
        handle.addEventListener("dragstart", (e) => {
          const card = handle.closest(".admin-card-wrap");
          if (!card) return;
          draggedCard = card;
          e.dataTransfer.setData("text/plain", card.dataset.itemId);
          card.classList.add("dragging");
        });
        handle.addEventListener("dragend", () => {
          draggedCard?.classList.remove("dragging");
          draggedCard = null;
        });
      });

      // Nearest-card-center heuristic, not a strict row/column layout
      // solve — good enough for a low-traffic internal tool where a card
      // landing one position off just means dragging again.
      function getDragAfterElement(container, clientX, clientY) {
        const cards = [...container.querySelectorAll(".admin-card-wrap:not(.dragging)")];
        let closest = null;
        let closestDistance = Infinity;
        for (const card of cards) {
          const box = card.getBoundingClientRect();
          const dx = clientX - (box.left + box.width / 2);
          const dy = clientY - (box.top + box.height / 2);
          const distance = dx * dx + dy * dy;
          if (distance < closestDistance) {
            closestDistance = distance;
            closest = { card, after: dx > 0 };
          }
        }
        if (!closest) return null;
        return closest.after ? closest.card.nextElementSibling : closest.card;
      }

      document.querySelectorAll(".admin-grid").forEach((grid) => {
        grid.addEventListener("dragover", (e) => {
          if (!draggedCard) return;
          e.preventDefault();
          const afterElement = getDragAfterElement(grid, e.clientX, e.clientY);
          grid.insertBefore(draggedCard, afterElement);
        });
        grid.addEventListener("drop", (e) => {
          e.preventDefault();
          if (!draggedCard) return;
          const finalGrid = draggedCard.closest(".admin-grid");
          const groupIdRaw = finalGrid.dataset.groupId;
          const groupId = groupIdRaw === "" ? null : Number(groupIdRaw);
          const orderedIds = [...finalGrid.querySelectorAll(".admin-card-wrap")].map((c) => Number(c.dataset.itemId));
          const statusEl = document.getElementById("admin-reorder-status");
          fetch("/admin/catalog/reorder", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ ...CURRENT_SCOPE, groupId, orderedIds }),
          })
            .then((res) => {
              if (!res.ok) throw new Error("HTTP " + res.status);
              if (statusEl) {
                statusEl.textContent = " Saved.";
                setTimeout(() => (statusEl.textContent = ""), 1500);
              }
            })
            .catch((err) => {
              if (statusEl) statusEl.textContent = " Couldn't save order: " + err.message;
            });
        });
      });

      // ---- Lightbox — click any thumbnail to enlarge, dismiss on
      // outside click or Escape (same pattern as the task pane's own
      // closeColorPickerPanel in src/taskpane/taskpane.ts). ----
      const lightbox = document.getElementById("adminLightbox");
      const lightboxImg = document.getElementById("adminLightboxImg");
      document.querySelectorAll(".admin-grid").forEach((grid) => {
        grid.addEventListener("click", (e) => {
          const thumb = e.target.closest(".admin-card-thumb");
          if (!thumb || !thumb.dataset.full) return;
          lightboxImg.src = thumb.dataset.full;
          lightbox.style.display = "flex";
        });
      });
      lightbox.addEventListener("click", () => (lightbox.style.display = "none"));
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") lightbox.style.display = "none";
      });

      // ---- Inline group management (Admin UI Phase 11) — replaces the
      // old standalone /admin/groups page entirely. ----
      function addGroupOptionEverywhere(id, name) {
        document.querySelectorAll('select[name="groupId"]').forEach((select) => {
          const opt = document.createElement("option");
          opt.value = String(id);
          opt.textContent = name;
          select.insertBefore(opt, select.querySelector('option[value="__new__"]'));
        });
      }
      function renameGroupOptionEverywhere(id, name) {
        document.querySelectorAll('select[name="groupId"] option[value="' + id + '"]').forEach((opt) => {
          opt.textContent = name;
        });
      }
      function removeGroupOptionEverywhere(id) {
        document.querySelectorAll('select[name="groupId"] option[value="' + id + '"]').forEach((opt) => {
          const select = opt.parentElement;
          const wasSelected = opt.selected;
          opt.remove();
          if (wasSelected) select.value = "";
        });
      }

      // Create: the "+ Add new group…" option in any card's group select.
      document.querySelectorAll('select[name="groupId"]').forEach((select) => {
        select.dataset.prevValue = select.value;
        select.addEventListener("change", () => {
          if (select.value !== "__new__") {
            select.dataset.prevValue = select.value;
            markDirty(select.closest(".admin-card-wrap"));
            return;
          }
          const name = (prompt("New group name:") || "").trim();
          if (!name) {
            select.value = select.dataset.prevValue;
            return;
          }
          const sortOrder = document.querySelectorAll('.admin-cluster[data-group-id]:not([data-group-id=""])').length;
          fetch("/admin/groups", {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
            body: new URLSearchParams({ ...CURRENT_SCOPE, name, sortOrder: String(sortOrder) }),
          })
            .then(async (res) => {
              const body = await res.json().catch(() => ({}));
              if (!res.ok) throw new Error(body.error || "HTTP " + res.status);
              return body;
            })
            .then((body) => {
              addGroupOptionEverywhere(body.id, body.name);
              select.value = String(body.id);
              select.dataset.prevValue = select.value;
              markDirty(select.closest(".admin-card-wrap"));
            })
            .catch((err) => {
              alert("Couldn't create group: " + err.message);
              select.value = select.dataset.prevValue;
            });
        });
      });

      // Rename: click a group heading, edit, blur (or Enter) to save.
      document.querySelectorAll(".admin-group-heading-input").forEach((input) => {
        input.dataset.original = input.value;
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") input.blur();
        });
        input.addEventListener("blur", () => {
          const name = input.value.trim();
          if (!name || name === input.dataset.original) {
            input.value = input.dataset.original;
            return;
          }
          const id = input.dataset.groupId;
          fetch("/admin/groups/" + id, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
            body: new URLSearchParams({ name }),
          })
            .then(async (res) => {
              const body = await res.json().catch(() => ({}));
              if (!res.ok) throw new Error(body.error || "HTTP " + res.status);
            })
            .then(() => {
              input.dataset.original = name;
              renameGroupOptionEverywhere(id, name);
            })
            .catch((err) => {
              alert("Couldn't rename group: " + err.message);
              input.value = input.dataset.original;
            });
        });
      });

      // Delete: the × button next to a group heading.
      document.querySelectorAll(".admin-group-delete").forEach((button) => {
        button.addEventListener("click", () => {
          if (!confirm("Delete this group? Items in it become ungrouped.")) return;
          const id = button.dataset.groupId;
          fetch("/admin/groups/" + id + "/delete", { method: "POST", headers: { Accept: "application/json" } })
            .then((res) => {
              if (!res.ok) throw new Error("HTTP " + res.status);
              removeGroupOptionEverywhere(id);
              const cluster = document.querySelector('.admin-cluster[data-group-id="' + id + '"]');
              const ungroupedGrid = document.querySelector('.admin-grid[data-group-id=""]');
              if (cluster && ungroupedGrid) {
                const grid = cluster.querySelector(".admin-grid");
                while (grid.firstChild) ungroupedGrid.appendChild(grid.firstChild);
                cluster.remove();
              }
            })
            .catch((err) => alert("Couldn't delete group: " + err.message));
        });
      });

      // Reorder: drag a group heading. Plain 1-D vertical list among real
      // .admin-cluster sections — the synthetic Ungrouped cluster is
      // excluded from the draggable set and candidate list entirely, so
      // it can never be reordered or dropped past.
      let draggedCluster = null;
      const clustersContainer = document.getElementById("adminClusters");

      document.querySelectorAll(".admin-group-drag").forEach((handle) => {
        handle.addEventListener("dragstart", (e) => {
          const cluster = handle.closest(".admin-cluster");
          if (!cluster) return;
          draggedCluster = cluster;
          e.dataTransfer.setData("text/plain", cluster.dataset.groupId);
          cluster.classList.add("dragging");
        });
        handle.addEventListener("dragend", () => {
          draggedCluster?.classList.remove("dragging");
          draggedCluster = null;
        });
      });

      clustersContainer.addEventListener("dragover", (e) => {
        if (!draggedCluster) return;
        e.preventDefault();
        const candidates = [
          ...clustersContainer.querySelectorAll('.admin-cluster[data-group-id]:not([data-group-id=""]):not(.dragging)'),
        ];
        const after = candidates.find((c) => e.clientY < c.getBoundingClientRect().top + c.getBoundingClientRect().height / 2);
        clustersContainer.insertBefore(draggedCluster, after || document.querySelector('.admin-cluster[data-group-id=""]'));
      });
      clustersContainer.addEventListener("drop", (e) => {
        e.preventDefault();
        if (!draggedCluster) return;
        const orderedGroupIds = [
          ...clustersContainer.querySelectorAll('.admin-cluster[data-group-id]:not([data-group-id=""])'),
        ].map((c) => Number(c.dataset.groupId));
        fetch("/admin/groups/reorder", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ ...CURRENT_SCOPE, orderedGroupIds }),
        }).catch((err) => alert("Couldn't save group order: " + err.message));
      });

      // Scrolls to and briefly highlights a just-added item (Task Pane
      // Phase 12's "Add Selected Slide to Library" opens this page with
      // ?highlight=<id> so the admin doesn't have to hunt for "Untitled
      // item" in an unsorted category).
      if (HIGHLIGHT_ITEM_ID !== null) {
        const target = document.querySelector('.admin-card-wrap[data-item-id="' + HIGHLIGHT_ITEM_ID + '"]');
        if (target) {
          target.scrollIntoView({ behavior: "smooth", block: "center" });
          target.classList.add("highlight");
        }
      }
    </script>
  </body></html>`);
});

// The first JSON/fetch-based admin route — every other /admin* route is a
// redirect-based form post, but drag-and-drop reorder (Admin UI Phase 10)
// needs to persist on drop with no page reload. Re-sequences one
// category+group cluster's sort_order to match orderedIds, and reassigns
// group_id for every item in that list — covers both "reordered within
// its own cluster" and "dragged into a different cluster" in one call.
// Registered *before* POST /admin/catalog/:id below — Express matches
// routes in registration order, and :id would otherwise swallow this as
// a request to update an item literally named "reorder".
// Task Pane Phase 14: requireAdmin's blanket global-only check no longer
// applies to this route — it's shared by both global-admin and
// company-admin reorders now, so auth is checked per-request against
// whichever scope the request actually names (canAdminCompany for a
// companyDomain body, req.user.isAdmin for a category one).
app.post("/admin/catalog/reorder", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Sign in first." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const { category, companyDomain, groupId, orderedIds } = req.body ?? {};
  if (companyDomain) {
    if (!canAdminCompany(req.user, companyDomain)) return res.status(403).json({ error: "Not an admin for this company." });
  } else {
    if (!CATALOG_CATEGORIES.includes(category)) return res.status(400).json({ error: "Invalid category." });
    if (!req.user.isAdmin) return res.status(403).json({ error: "Not an admin." });
  }
  if (groupId !== null && !(Number.isInteger(groupId) && groupId > 0)) {
    return res.status(400).json({ error: "Invalid group." });
  }
  if (!Array.isArray(orderedIds) || orderedIds.length === 0 || !orderedIds.every((id) => Number.isInteger(id) && id > 0)) {
    return res.status(400).json({ error: "Invalid item list." });
  }
  await reorderCatalogItems({ category: companyDomain ? null : category, companyDomain: companyDomain || null, groupId, orderedIds });
  res.json({ ok: true });
});

// Task Pane Phase 14: fetches the existing row first and checks
// canManageRow against it, instead of the blanket requireAdmin middleware
// — a company admin can now reach this route too, for their own
// company's items. A company item's edit form has no category <select>
// at all (see renderCard), so `category` arrives undefined for those;
// updateCatalogItem treats that as "don't touch category" (see its own
// comment) rather than a validation failure.
app.post("/admin/catalog/:id", upload.single("thumbnail"), async (req, res) => {
  const id = Number(req.params.id);
  const json = wantsJson(req);
  // Redirect back to whichever scope's grid the edit was submitted from
  // once it's known (redirect-mode only — the JSON path never navigates,
  // so it has nothing to redirect to).
  const fail = (status, message, backTo) =>
    json ? res.status(status).json({ error: message }) : redirectWithError(res, backTo ?? "/admin", message);

  if (!req.user) return fail(401, "Sign in first.");
  if (!req.user.isRegistered) return fail(403, "Finish creating your account first.");
  if (!Number.isInteger(id) || id <= 0) return fail(400, "Invalid item id.");

  const existing = await getCatalogItem(id);
  if (!existing) return fail(404, "Item not found.");
  if (!canManageRow(req.user, existing)) return fail(403, "Not an admin for this item.");

  const isCompanyItem = !!existing.company_domain;
  const title = typeof req.body.title === "string" ? req.body.title.trim() : "";
  // Task Pane Phase 16: the gallery's lightweight quick-edit never sends
  // `category` at all (no category-reassignment UI there) — treat an
  // omitted category the same way a company item's always-omitted
  // category is already treated (updateCatalogItem's COALESCE leaves it
  // unchanged), rather than requiring it the way /admin's own full edit
  // form still does.
  const categoryProvided = !isCompanyItem && typeof req.body.category === "string";
  const category = categoryProvided ? req.body.category : undefined;
  const groupId = req.body.groupId ? Number(req.body.groupId) : null;
  const tags = typeof req.body.tags === "string" ? req.body.tags.split(",").map((t) => t.trim()).filter(Boolean) : [];
  const backTo = isCompanyItem
    ? `/admin?scope=company:${encodeURIComponent(existing.company_domain)}`
    : `/admin?category=${categoryProvided ? category : existing.category}`;
  if (!title) return fail(400, "Title can't be empty.", backTo);
  if (categoryProvided && !CATALOG_CATEGORIES.includes(category)) return fail(400, "Invalid category.");
  if (groupId !== null && (!Number.isInteger(groupId) || groupId <= 0)) {
    return fail(400, "Invalid group.", backTo);
  }

  const updated = await updateCatalogItem({ id, title, category, groupId });
  if (!updated) return fail(404, "Item not found.", backTo);

  await setItemTags(id, tags);

  let finalThumbnailPath = existing.thumbnail_path;
  if (req.file) {
    const ext = MIME_TO_EXT[req.file.mimetype];
    const newThumbnailPath = `item-${id}.${ext}`;
    await fs.promises.writeFile(path.join(THUMBNAILS_DIR, newThumbnailPath), req.file.buffer);
    await updateCatalogItemThumbnail({ id, thumbnailPath: newThumbnailPath });
    if (existing.thumbnail_path && existing.thumbnail_path !== newThumbnailPath) {
      await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
    }
    finalThumbnailPath = newThumbnailPath;
  }

  if (json) {
    const thumbnailUrl = finalThumbnailPath ? `/assets/catalog/thumbnails/${finalThumbnailPath}?v=${ASSET_VERSION}` : null;
    return res.json({ ok: true, thumbnailUrl });
  }
  res.redirect(303, backTo);
});

app.post("/admin/catalog/:id/delete", async (req, res) => {
  const id = Number(req.params.id);
  const json = wantsJson(req);
  const fail = (status, message, backTo) =>
    json ? res.status(status).json({ error: message }) : redirectWithError(res, backTo ?? "/admin", message);

  if (!req.user) return fail(401, "Sign in first.");
  if (!req.user.isRegistered) return fail(403, "Finish creating your account first.");
  if (!Number.isInteger(id) || id <= 0) return fail(400, "Invalid item id.");

  const existing = await getCatalogItem(id);
  if (!existing) return fail(404, "Item not found.");
  if (!canManageRow(req.user, existing)) return fail(403, "Not an admin for this item.");

  const backTo = existing.company_domain
    ? `/admin?scope=company:${encodeURIComponent(existing.company_domain)}`
    : existing.category && CATALOG_CATEGORIES.includes(existing.category)
      ? `/admin?category=${existing.category}`
      : "/admin";
  const deleted = await deleteCatalogItem(id);
  if (!deleted) return fail(404, "Item not found.", backTo);

  if (existing.thumbnail_path) {
    await fs.promises.unlink(path.join(THUMBNAILS_DIR, existing.thumbnail_path)).catch(() => {});
  }

  if (json) return res.json({ ok: true });
  res.redirect(303, backTo);
});

// Admin UI Phase 5 introduced groups as admin-defined, admin-ordered
// sub-groupings within one category (e.g. "Pyramids" inside Diagrams),
// separate from the free-form tags above. Phase 11 moved all group
// management (create/rename/delete/reorder) onto the main /admin grid
// page itself (see GET /admin's inline script) — this old standalone
// page is gone, kept only as a redirect for anything still linking here.
app.get("/admin/groups", requireAdmin, (req, res) => {
  const category = CATALOG_CATEGORIES.includes(req.query.category) ? req.query.category : CATALOG_CATEGORIES[0];
  res.redirect(301, `/admin?category=${category}`);
});

// Reused by GET /admin's inline "+ Add new group…" dropdown option (via
// fetch, Accept: application/json) — the redirect-based response stays as
// a fallback for anything posting here without that header. Task Pane
// Phase 14: accepts either category (global scope) or companyDomain
// (company scope) in the body — CURRENT_SCOPE's spread in GET /admin's
// inline script sends whichever applies.
app.post("/admin/groups", async (req, res) => {
  const category = req.body.category;
  const companyDomain = req.body.companyDomain;
  const name = typeof req.body.name === "string" ? req.body.name.trim() : "";
  const sortOrder = Number(req.body.sortOrder) || 0;
  const json = wantsJson(req);
  const backTo = companyDomain ? `/admin?scope=company:${encodeURIComponent(companyDomain)}` : `/admin?category=${category}`;
  const fail = (status, message) => (json ? res.status(status).json({ error: message }) : redirectWithError(res, backTo, message));

  if (!req.user) return fail(401, "Sign in first.");
  if (!req.user.isRegistered) return fail(403, "Finish creating your account first.");
  if (companyDomain) {
    if (!canAdminCompany(req.user, companyDomain)) return fail(403, "Not an admin for this company.");
  } else {
    if (!CATALOG_CATEGORIES.includes(category)) return fail(400, "Unknown category.");
    if (!req.user.isAdmin) return fail(403, "Not an admin.");
  }
  if (!name) return fail(400, "Group name can't be empty.");

  let created;
  try {
    created = companyDomain
      ? await createGroup({ companyDomain, name, sortOrder })
      : await createGroup({ category, name, sortOrder });
  } catch (err) {
    return fail(409, `A group named "${name}" already exists here.`);
  }
  if (json) return res.json({ ok: true, id: created.id, name });
  res.redirect(303, backTo);
});

// Persists dragging a group heading to a new position (GET /admin's inline
// script). Registered *before* POST /admin/groups/:id below — same
// Express route-ordering pitfall already caught once this session for
// /admin/catalog/reorder vs. /admin/catalog/:id: :id would otherwise
// swallow this as a request to rename a group literally named "reorder".
app.post("/admin/groups/reorder", async (req, res) => {
  if (!req.user) return res.status(401).json({ error: "Sign in first." });
  if (!req.user.isRegistered) return res.status(403).json({ error: "Finish creating your account first." });
  const { category, companyDomain, orderedGroupIds } = req.body ?? {};
  if (companyDomain) {
    if (!canAdminCompany(req.user, companyDomain)) return res.status(403).json({ error: "Not an admin for this company." });
  } else {
    if (!CATALOG_CATEGORIES.includes(category)) return res.status(400).json({ error: "Invalid category." });
    if (!req.user.isAdmin) return res.status(403).json({ error: "Not an admin." });
  }
  if (!Array.isArray(orderedGroupIds) || orderedGroupIds.length === 0 || !orderedGroupIds.every((id) => Number.isInteger(id) && id > 0)) {
    return res.status(400).json({ error: "Invalid group list." });
  }
  await reorderGroups({ category: companyDomain ? null : category, companyDomain: companyDomain || null, orderedGroupIds });
  res.json({ ok: true });
});

// Reused by GET /admin's inline group-heading rename (click to edit,
// blur/Enter to save). sortOrder is optional now that group order is set
// by drag-and-drop (see POST /admin/groups/reorder above) — omitting it
// leaves the group's current position untouched.
app.post("/admin/groups/:id", async (req, res) => {
  const id = Number(req.params.id);
  const name = typeof req.body.name === "string" ? req.body.name.trim() : "";
  const sortOrder = req.body.sortOrder !== undefined ? Number(req.body.sortOrder) : null;
  const json = wantsJson(req);
  // Neither category nor companyDomain is submitted from this form (scope
  // isn't editable — see updateGroup's comment) — read it back from the
  // row itself so the redirect-mode fallback knows which page to send
  // back to.
  const existing = await getGroup(id);
  const backTo = existing?.company_domain
    ? `/admin?scope=company:${encodeURIComponent(existing.company_domain)}`
    : `/admin?category=${existing?.category ?? CATALOG_CATEGORIES[0]}`;
  const fail = (status, message) => (json ? res.status(status).json({ error: message }) : redirectWithError(res, backTo, message));

  if (!req.user) return fail(401, "Sign in first.");
  if (!req.user.isRegistered) return fail(403, "Finish creating your account first.");
  if (!existing) return fail(404, "Group not found.");
  if (!canManageRow(req.user, existing)) return fail(403, "Not an admin for this group.");
  if (!name) return fail(400, "Group name can't be empty.");

  await updateGroup({ id, name, sortOrder });
  if (json) return res.json({ ok: true });
  res.redirect(303, backTo);
});

// Reused by GET /admin's inline group-heading delete button.
app.post("/admin/groups/:id/delete", async (req, res) => {
  const id = Number(req.params.id);
  const json = wantsJson(req);
  const existing = await getGroup(id);
  const backTo = existing?.company_domain
    ? `/admin?scope=company:${encodeURIComponent(existing.company_domain)}`
    : `/admin?category=${existing?.category ?? CATALOG_CATEGORIES[0]}`;
  const fail = (status, message) => (json ? res.status(status).json({ error: message }) : redirectWithError(res, backTo, message));

  if (!req.user) return fail(401, "Sign in first.");
  if (!req.user.isRegistered) return fail(403, "Finish creating your account first.");
  if (!existing) return fail(404, "Group not found.");
  if (!canManageRow(req.user, existing)) return fail(403, "Not an admin for this group.");

  await deleteGroup(id);
  if (json) return res.json({ ok: true });
  res.redirect(303, backTo);
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
