const fs = require("fs");
const path = require("path");

const CATALOG_DIR = path.resolve(process.env.CATALOG_DIR || path.join(__dirname, "..", "data", "catalog"));

// Persistent-volume home for thumbnail images (see server.js's second
// express.static mount) — separate from CATALOG_DIR's category folders
// since admin-uploaded thumbnails aren't scoped by category the way
// source .pptx files are (see the flat item-{id}.{ext} naming in
// server.js's POST /admin/catalog/:id).
const THUMBNAILS_DIR = path.join(CATALOG_DIR, "thumbnails");
fs.mkdirSync(THUMBNAILS_DIR, { recursive: true });

// Task Pane Phase 12: source files for items edited or added directly from
// the task pane (via Slide.exportAsBase64) — flat and separate from the
// Python pipeline's category-prefixed paths (e.g. "text/text-001.pptx"),
// mirroring how THUMBNAILS_DIR above already sidesteps that same coupling.
const ADMIN_ADDED_DIR = path.join(CATALOG_DIR, "admin-added");
fs.mkdirSync(ADMIN_ADDED_DIR, { recursive: true });

// Task Pane Phase 13: source files for personal (owner-scoped) library
// items added or edited from the task pane — same flat, non-category-
// prefixed layout as ADMIN_ADDED_DIR above, just a separate directory so
// personal and admin-added content never collide by coincidence of naming.
const PERSONAL_ADDED_DIR = path.join(CATALOG_DIR, "personal-added");
fs.mkdirSync(PERSONAL_ADDED_DIR, { recursive: true });

// Hand-kept in sync with db/init.sql's catalog_items.category CHECK
// constraint — validated here before any DB write, so a bad value gets a
// clean 400 instead of a raw constraint-violation error page.
const CATALOG_CATEGORIES = ["text", "objects", "shapes", "stamps", "tables", "symbols", "diagrams", "maps", "clipart", "frameworks", "flags"];

// Task Pane Phase 20: whole .potx template files — a structurally
// different entity from catalog_items (no category-prefixed or
// admin/personal-added split needed, since every template is a real
// uploaded file regardless of scope; scope itself lives on the DB row,
// not the directory layout).
const TEMPLATES_DIR = path.join(CATALOG_DIR, "templates");
fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

/**
 * Resolves a catalog item's source_file (as stored in the DB, e.g.
 * "text/text-010.pptx") to an absolute path under CATALOG_DIR. The client
 * never supplies this value directly — it only ever sends a numeric item
 * ID, and the server looks up source_file from the DB row — but this still
 * refuses to resolve outside CATALOG_DIR as defense in depth.
 */
function resolveCatalogFilePath(sourceFile) {
  const resolved = path.resolve(CATALOG_DIR, sourceFile);
  if (resolved !== CATALOG_DIR && !resolved.startsWith(CATALOG_DIR + path.sep)) {
    throw new Error("Refusing to resolve a catalog file outside CATALOG_DIR.");
  }
  return resolved;
}

/** Same never-escape-the-directory defense as resolveCatalogFilePath, scoped to TEMPLATES_DIR instead. */
function resolveTemplateFilePath(sourceFile) {
  const resolved = path.resolve(TEMPLATES_DIR, sourceFile);
  if (resolved !== TEMPLATES_DIR && !resolved.startsWith(TEMPLATES_DIR + path.sep)) {
    throw new Error("Refusing to resolve a template file outside TEMPLATES_DIR.");
  }
  return resolved;
}

module.exports = {
  CATALOG_DIR,
  THUMBNAILS_DIR,
  ADMIN_ADDED_DIR,
  PERSONAL_ADDED_DIR,
  TEMPLATES_DIR,
  CATALOG_CATEGORIES,
  resolveCatalogFilePath,
  resolveTemplateFilePath,
};
