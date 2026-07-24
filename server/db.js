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

// Shared by every catalog-item read query below: joins in the group name
// and an aggregated array of tag names. LEFT JOINs throughout so an item
// with no group and/or no tags still returns one row (tags as '{}', not
// dropped) rather than being excluded by an inner join.
const CATALOG_ITEM_GROUP_TAGS_JOIN = `
  LEFT JOIN catalog_groups cg ON cg.id = ci.group_id
  LEFT JOIN catalog_item_tags cit ON cit.item_id = ci.id
  LEFT JOIN tags t ON t.id = cit.tag_id
`;
const CATALOG_ITEM_GROUP_TAGS_SELECT = `
  ci.group_id, cg.name AS group_name,
  COALESCE(array_agg(t.name ORDER BY t.name) FILTER (WHERE t.name IS NOT NULL), '{}') AS tags
`;

/** Shared/admin catalog items in one category, in display order, with group + tags joined in. */
async function listSharedCatalogItems(category) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.reconstruct_spec, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.category = $1 AND ci.owner_oid IS NULL
     GROUP BY ci.id, cg.name
     ORDER BY ci.sort_order, ci.id`,
    [category]
  );
  return result.rows;
}

/** All shared/admin catalog items across every category — used by the /admin table. */
async function listAllCatalogItems() {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.owner_oid IS NULL
     GROUP BY ci.id, cg.name
     ORDER BY ci.category, ci.sort_order, ci.id`
  );
  return result.rows;
}

/** A single catalog item by ID, including source_file — used to resolve the file for a 'file'-mode insert, and by /admin's edit form. */
async function getCatalogItem(id) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.source_file, ci.reconstruct_spec, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.id = $1
     GROUP BY ci.id, cg.name`,
    [id]
  );
  return result.rows[0] ?? null;
}

/** A category's groups, in display order — used by /admin/groups and the group <select> on the item edit form. */
async function listGroupsForCategory(category) {
  const result = await pool.query(
    `SELECT id, category, name, sort_order FROM catalog_groups WHERE category = $1 ORDER BY sort_order, name`,
    [category]
  );
  return result.rows;
}

/** A single group by ID, mainly so /admin/groups' edit/delete routes know which category page to redirect back to (category isn't resubmitted from those forms — see updateGroup's comment on why it's not editable). */
async function getGroup(id) {
  const result = await pool.query(`SELECT id, category, name, sort_order FROM catalog_groups WHERE id = $1`, [id]);
  return result.rows[0] ?? null;
}

/** Creates a new group within a category. UNIQUE (category, name) surfaces a clear conflict on a duplicate name rather than silently duplicating. */
async function createGroup({ category, name, sortOrder }) {
  const result = await pool.query(
    `INSERT INTO catalog_groups (category, name, sort_order) VALUES ($1, $2, $3) RETURNING id`,
    [category, name, sortOrder ?? 0]
  );
  return result.rows[0];
}

/** Renames/reorders an existing group. Category is not editable here — a group moving categories would strand every item currently assigned to it, so that's a delete-and-recreate, not an edit. */
async function updateGroup({ id, name, sortOrder }) {
  const result = await pool.query(
    `UPDATE catalog_groups SET name = $2, sort_order = $3 WHERE id = $1 RETURNING id`,
    [id, name, sortOrder]
  );
  return result.rows[0] ?? null;
}

/** Deletes a group. Referencing items fall back to ungrouped (group_id NULL) via the FK's ON DELETE SET NULL, not deleted themselves. */
async function deleteGroup(id) {
  const result = await pool.query(`DELETE FROM catalog_groups WHERE id = $1 RETURNING id`, [id]);
  return result.rows[0] ?? null;
}

/** Looks up a tag by exact (case-insensitive) name, creating it if it doesn't exist yet. Used when saving an item's tags from /admin — the client already ran a near-match confirmation pass, so this just get-or-creates the name it was given, no further fuzzy matching server-side. */
async function getOrCreateTag(name) {
  const result = await pool.query(
    `INSERT INTO tags (name) VALUES ($1)
     ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
     RETURNING id, name`,
    [name]
  );
  return result.rows[0];
}

/** Fetches every known tag name, for embedding into /admin's near-match-confirmation script. */
async function listAllTagNames() {
  const result = await pool.query(`SELECT name FROM tags ORDER BY name`);
  return result.rows.map((r) => r.name);
}

/** Replaces an item's full set of tags in one go (delete-then-reinsert, matching the same idempotent style scripts/seed-catalog.js already uses for whole-category reseeds) — simpler than diffing old vs. new for a handful of tags per item. */
async function setItemTags(itemId, tagNames) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    await client.query(`DELETE FROM catalog_item_tags WHERE item_id = $1`, [itemId]);
    for (const name of tagNames) {
      const tag = await client.query(
        `INSERT INTO tags (name) VALUES ($1) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id`,
        [name]
      );
      await client.query(`INSERT INTO catalog_item_tags (item_id, tag_id) VALUES ($1, $2) ON CONFLICT DO NOTHING`, [
        itemId,
        tag.rows[0].id,
      ]);
    }
    await client.query("COMMIT");
  } catch (err) {
    await client.query("ROLLBACK");
    throw err;
  } finally {
    client.release();
  }
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

/**
 * Updates a shared item's title/category/sort_order/group — never
 * insert_mode/source_file/reconstruct_spec, which stay owned by the
 * slice+seed script pipeline. owner_oid IS NULL mirrors
 * listAllCatalogItems' scope, as defense in depth against this admin-only
 * surface ever touching a future private item.
 *
 * If category is changing, group_id is forced to NULL in the same
 * statement rather than trusting the caller's groupId - a group from the
 * item's old category almost certainly doesn't belong to the new one, and
 * this is cheaper/safer than validating cross-category group ownership.
 * The CASE's bare `category` reference is the pre-update row value (all
 * SET expressions in one UPDATE see the same pre-update row), so this is
 * a correct same-statement comparison, not a race.
 */
async function updateCatalogItem({ id, title, category, sortOrder, groupId }) {
  const result = await pool.query(
    `UPDATE catalog_items
     SET title = $2,
         group_id = CASE WHEN category = $3 THEN $5::integer ELSE NULL END,
         category = $3,
         sort_order = $4
     WHERE id = $1 AND owner_oid IS NULL
     RETURNING id`,
    [id, title, category, sortOrder, groupId ?? null]
  );
  return result.rows[0] ?? null;
}

/** Points thumbnail_path at a newly-uploaded file. Kept separate from updateCatalogItem so an edit that doesn't touch the thumbnail never overwrites it with NULL. */
async function updateCatalogItemThumbnail({ id, thumbnailPath }) {
  const result = await pool.query(
    `UPDATE catalog_items SET thumbnail_path = $2 WHERE id = $1 AND owner_oid IS NULL RETURNING id`,
    [id, thumbnailPath]
  );
  return result.rows[0] ?? null;
}

/** Deletes one shared catalog item. RETURNING id lets the caller distinguish "deleted" from "already gone" (404 vs. success) instead of silently no-op'ing. */
async function deleteCatalogItem(id) {
  const result = await pool.query(
    `DELETE FROM catalog_items WHERE id = $1 AND owner_oid IS NULL RETURNING id`,
    [id]
  );
  return result.rows[0] ?? null;
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
  updateCatalogItem,
  updateCatalogItemThumbnail,
  deleteCatalogItem,
  listGroupsForCategory,
  getGroup,
  createGroup,
  updateGroup,
  deleteGroup,
  getOrCreateTag,
  listAllTagNames,
  setItemTags,
};
