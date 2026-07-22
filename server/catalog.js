const path = require("path");

const CATALOG_DIR = path.resolve(process.env.CATALOG_DIR || path.join(__dirname, "..", "data", "catalog"));

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

module.exports = { CATALOG_DIR, resolveCatalogFilePath };
