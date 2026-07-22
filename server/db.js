const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
});

/**
 * Waits for Postgres to accept connections, retrying instead of crash-
 * looping — docker-compose's `depends_on: condition: service_healthy`
 * already waits for the DB's own healthcheck, but this is a cheap second
 * line of defense against startup races.
 */
async function waitForDatabase(maxAttempts = 10, delayMs = 1000) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await pool.query("SELECT 1");
      console.log("Connected to Postgres.");
      return;
    } catch (err) {
      console.log(`Postgres not ready yet (attempt ${attempt}/${maxAttempts}): ${err.message}`);
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }
  throw new Error("Could not connect to Postgres after repeated attempts.");
}

/** Upserts a user by (oid, tid) and bumps last_seen_at — called on every successful token verification. */
async function upsertUser({ oid, tid, email, displayName }) {
  const result = await pool.query(
    `INSERT INTO users (oid, tid, email, display_name, last_seen_at)
     VALUES ($1, $2, $3, $4, now())
     ON CONFLICT (oid, tid)
     DO UPDATE SET email = EXCLUDED.email, display_name = EXCLUDED.display_name, last_seen_at = now()
     RETURNING oid, tid, email, display_name, created_at, last_seen_at`,
    [oid, tid, email, displayName]
  );
  return result.rows[0];
}

/** Shared/admin catalog items in one category, in display order. */
async function listSharedCatalogItems(category) {
  const result = await pool.query(
    `SELECT id, category, title, insert_mode, reconstruct_spec, thumbnail_path, sort_order
     FROM catalog_items
     WHERE category = $1 AND owner_oid IS NULL
     ORDER BY sort_order, id`,
    [category]
  );
  return result.rows;
}

/** All shared/admin catalog items across every category — used by the /admin read-only table. */
async function listAllCatalogItems() {
  const result = await pool.query(
    `SELECT id, category, title, insert_mode, thumbnail_path, sort_order
     FROM catalog_items
     WHERE owner_oid IS NULL
     ORDER BY category, sort_order, id`
  );
  return result.rows;
}

/** A single catalog item by ID, including source_file — used to resolve the file for a 'file'-mode insert. */
async function getCatalogItem(id) {
  const result = await pool.query(
    `SELECT id, category, title, insert_mode, source_file, reconstruct_spec, thumbnail_path, sort_order
     FROM catalog_items WHERE id = $1`,
    [id]
  );
  return result.rows[0] ?? null;
}

/** Removes every shared item in a category, so scripts/seed-catalog.js can reseed it from scratch — simplest way to stay idempotent across both insert_mode shapes (source_file is UNIQUE but NULL for 'reconstruct' items, so ON CONFLICT can't dedupe those). */
async function deleteCatalogItemsByCategory(category) {
  await pool.query(`DELETE FROM catalog_items WHERE category = $1 AND owner_oid IS NULL`, [category]);
}

/** Inserts one shared catalog item. Used only by scripts/seed-catalog.js, always after a same-category deleteCatalogItemsByCategory(). */
async function insertCatalogItem({ category, title, insertMode, sourceFile, reconstructSpec, thumbnailPath, sortOrder }) {
  const result = await pool.query(
    `INSERT INTO catalog_items (category, title, insert_mode, source_file, reconstruct_spec, thumbnail_path, sort_order)
     VALUES ($1, $2, $3, $4, $5, $6, $7)
     RETURNING id`,
    [
      category,
      title,
      insertMode,
      sourceFile ?? null,
      reconstructSpec ? JSON.stringify(reconstructSpec) : null,
      thumbnailPath ?? null,
      sortOrder ?? 0,
    ]
  );
  return result.rows[0];
}

module.exports = {
  pool,
  waitForDatabase,
  upsertUser,
  listSharedCatalogItems,
  listAllCatalogItems,
  getCatalogItem,
  deleteCatalogItemsByCategory,
  insertCatalogItem,
};
