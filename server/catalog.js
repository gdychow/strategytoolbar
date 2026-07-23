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

// Hand-kept in sync with db/init.sql's catalog_items.category CHECK
// constraint — validated here before any DB write, so a bad value gets a
// clean 400 instead of a raw constraint-violation error page.
const CATALOG_CATEGORIES = ["text", "objects", "shapes", "stamps", "tables", "symbols", "diagrams"];

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

module.exports = { CATALOG_DIR, THUMBNAILS_DIR, CATALOG_CATEGORIES, resolveCatalogFilePath };
