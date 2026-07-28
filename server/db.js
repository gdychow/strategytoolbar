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
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
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

/** A single catalog item by ID, including source_file — used to resolve the file for a 'file'-mode insert, and by /admin's edit form. */
async function getCatalogItem(id) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.source_file, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
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

/**
 * Renames and/or reorders an existing group. Category is not editable
 * here — a group moving categories would strand every item currently
 * assigned to it, so that's a delete-and-recreate, not an edit.
 *
 * sortOrder is optional (pass null/undefined to leave it unchanged) —
 * since Admin UI Phase 11, a plain rename (click the heading, edit,
 * blur) doesn't also resubmit position; reordering is its own drag-driven
 * action (see reorderGroups below).
 */
async function updateGroup({ id, name, sortOrder }) {
  const result = await pool.query(
    `UPDATE catalog_groups SET name = $2, sort_order = COALESCE($3, sort_order) WHERE id = $1 RETURNING id`,
    [id, name, sortOrder ?? null]
  );
  return result.rows[0] ?? null;
}

/**
 * Re-sequences a category's groups to match orderedGroupIds' order
 * (0..n-1) — mirrors reorderCatalogItems' transaction shape exactly. The
 * synthetic "Ungrouped" cluster has no backing row and is never part of
 * orderedGroupIds; it always renders last, client-side, unconditionally.
 */
async function reorderGroups({ category, orderedGroupIds }) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    for (let i = 0; i < orderedGroupIds.length; i++) {
      await client.query(`UPDATE catalog_groups SET sort_order = $1 WHERE id = $2 AND category = $3`, [
        i,
        orderedGroupIds[i],
        category,
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

/** Deletes a group. Referencing items fall back to ungrouped (group_id NULL) via the FK's ON DELETE SET NULL, not deleted themselves. */
async function deleteGroup(id) {
  const result = await pool.query(`DELETE FROM catalog_groups WHERE id = $1 RETURNING id`, [id]);
  return result.rows[0] ?? null;
}

/**
 * Looks up a group by (category, name), creating it if it doesn't exist -
 * mirrors getOrCreateTag's upsert pattern, keyed off catalog_groups' own
 * UNIQUE (category, name) constraint. Used by scripts/seed-catalog.js so a
 * seed file can assign items straight into a named group (e.g. Phase 7's
 * Maps category, auto-grouped by source file) without a separate /admin
 * step to create those groups first.
 */
async function getOrCreateGroupIdForCategory(category, name, sortOrder) {
  const result = await pool.query(
    `INSERT INTO catalog_groups (category, name, sort_order) VALUES ($1, $2, $3)
     ON CONFLICT (category, name) DO UPDATE SET name = EXCLUDED.name
     RETURNING id`,
    [category, name, sortOrder ?? 0]
  );
  return result.rows[0].id;
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

/**
 * Inserts one catalog item. Used by scripts/seed-catalog.js (always shared,
 * always after a same-category deleteCatalogItemsByCategory()) and, since
 * Task Pane Phase 13, by POST /api/personal/catalog for a brand-new
 * personal item — ownerOid/ownerTid default to null (shared/admin item,
 * today's only behavior) unless explicitly passed.
 */
async function insertCatalogItem({
  category,
  title,
  insertMode,
  sourceFile,
  reconstructSpec,
  unicodeChar,
  thumbnailPath,
  sortOrder,
  groupId,
  ownerOid,
  ownerTid,
}) {
  const result = await pool.query(
    `INSERT INTO catalog_items (category, title, insert_mode, source_file, reconstruct_spec, unicode_char, thumbnail_path, sort_order, group_id, owner_oid, owner_tid)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
     RETURNING id`,
    [
      category,
      title,
      insertMode,
      sourceFile ?? null,
      reconstructSpec ? JSON.stringify(reconstructSpec) : null,
      unicodeChar ?? null,
      thumbnailPath ?? null,
      sortOrder ?? 0,
      groupId ?? null,
      ownerOid ?? null,
      ownerTid ?? null,
    ]
  );
  return result.rows[0];
}

/**
 * Updates a shared item's title/category/group — never
 * insert_mode/source_file/reconstruct_spec, which stay owned by the
 * slice+seed script pipeline. owner_oid IS NULL mirrors
 * listSharedCatalogItems' scope, as defense in depth against this
 * admin-only surface ever touching a future private item.
 *
 * sort_order is deliberately not settable here — drag-and-drop (see
 * reorderCatalogItems below) is the only way to reorder items as of Admin
 * UI Phase 10, replacing the old manual number input.
 *
 * If category is changing, group_id is forced to NULL in the same
 * statement rather than trusting the caller's groupId - a group from the
 * item's old category almost certainly doesn't belong to the new one, and
 * this is cheaper/safer than validating cross-category group ownership.
 * The CASE's bare `category` reference is the pre-update row value (all
 * SET expressions in one UPDATE see the same pre-update row), so this is
 * a correct same-statement comparison, not a race.
 */
async function updateCatalogItem({ id, title, category, groupId }) {
  const result = await pool.query(
    `UPDATE catalog_items
     SET title = $2,
         group_id = CASE WHEN category = $3 THEN $4::integer ELSE NULL END,
         category = $3
     WHERE id = $1 AND owner_oid IS NULL
     RETURNING id`,
    [id, title, category, groupId ?? null]
  );
  return result.rows[0] ?? null;
}

/**
 * Re-sequences one category+group cluster's sort_order to match
 * orderedIds' order (0..n-1), and reassigns group_id for every item in
 * that list — covers both "reordered within its own cluster" and
 * "dragged into a different cluster" in one call, since the destination
 * cluster's full membership is always what the client sends. The source
 * cluster (if an item moved out of it) is left with a gap in its
 * sequence, which is harmless — ordering only depends on relative order,
 * not contiguity. `category = $4` is defense in depth against a crafted
 * request smuggling an id from a different category into this call.
 */
async function reorderCatalogItems({ category, groupId, orderedIds }) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    for (let i = 0; i < orderedIds.length; i++) {
      await client.query(
        `UPDATE catalog_items SET sort_order = $1, group_id = $2
         WHERE id = $3 AND category = $4 AND owner_oid IS NULL`,
        [i, groupId, orderedIds[i], category]
      );
    }
    await client.query("COMMIT");
  } catch (err) {
    await client.query("ROLLBACK");
    throw err;
  } finally {
    client.release();
  }
}

/** Points thumbnail_path at a newly-uploaded file. Kept separate from updateCatalogItem so an edit that doesn't touch the thumbnail never overwrites it with NULL. */
async function updateCatalogItemThumbnail({ id, thumbnailPath }) {
  const result = await pool.query(
    `UPDATE catalog_items SET thumbnail_path = $2 WHERE id = $1 AND owner_oid IS NULL RETURNING id`,
    [id, thumbnailPath]
  );
  return result.rows[0] ?? null;
}

/**
 * Points source_file/thumbnail_path at freshly admin-exported content
 * (Task Pane Phase 12 — see POST /api/admin/catalog/:id/content) and
 * forces insert_mode = 'file' with reconstruct_spec cleared, which is
 * what actually migrates a 'reconstruct'-mode item to 'file'-mode the
 * first time it's edited this way. Kept separate from updateCatalogItem
 * (title/category/group) and updateCatalogItemThumbnail (thumbnail only)
 * — neither of those covers replacing the underlying content itself.
 */
async function updateCatalogItemContent({ id, sourceFile, thumbnailPath }) {
  const result = await pool.query(
    `UPDATE catalog_items
     SET source_file = $2, thumbnail_path = $3, insert_mode = 'file', reconstruct_spec = NULL
     WHERE id = $1 AND owner_oid IS NULL
     RETURNING id, category`,
    [id, sourceFile, thumbnailPath]
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

// ---------------------------------------------------------------------------
// Task Pane Phase 13 — personal (owner-scoped) library items. Kept as their
// own small explicit functions rather than parameterizing the
// shared/admin functions above with an optional owner filter, matching
// this file's existing style (e.g. updateCatalogItem vs.
// updateCatalogItemThumbnail are already kept separate for the same
// reason: one clear WHERE clause per function beats one function with a
// conditional WHERE).
// ---------------------------------------------------------------------------

/** One user's own personal items, in display order — no category/group scoping, since personal libraries are a flat list (see gallery.ts's "My Items" tab). */
async function listPersonalCatalogItems(oid, tid) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.owner_oid = $1 AND ci.owner_tid = $2
     GROUP BY ci.id, cg.name
     ORDER BY ci.sort_order, ci.id`,
    [oid, tid]
  );
  return result.rows;
}

/** A single personal item by ID, scoped to its owner — mirrors getCatalogItem, used to read the old source_file/thumbnail_path before an owner-driven edit/delete. Returns null both when the id doesn't exist and when it exists but isn't this user's, so callers can't distinguish "not found" from "not yours" (an intentional 404-not-403 shape, matching how the rest of this API never leaks another user's data via response codes). */
async function getOwnedCatalogItem(id, oid, tid) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.title, ci.insert_mode, ci.source_file, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.id = $1 AND ci.owner_oid = $2 AND ci.owner_tid = $3
     GROUP BY ci.id, cg.name`,
    [id, oid, tid]
  );
  return result.rows[0] ?? null;
}

/** Owner-scoped equivalent of updateCatalogItemContent (Task Pane Phase 12) — replaces a personal item's underlying graphic, always landing as insert_mode 'file' with reconstruct_spec cleared, same migration behavior as the admin version. */
async function updateOwnedCatalogItemContent({ id, oid, tid, sourceFile, thumbnailPath }) {
  const result = await pool.query(
    `UPDATE catalog_items
     SET source_file = $4, thumbnail_path = $5, insert_mode = 'file', reconstruct_spec = NULL
     WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3
     RETURNING id`,
    [id, oid, tid, sourceFile, thumbnailPath]
  );
  return result.rows[0] ?? null;
}

/** Renames a personal item — personal items have no /admin-equivalent page, so this is their only metadata-edit surface (unlike shared items, which also get category/group/tags via updateCatalogItem). */
async function renameOwnedCatalogItem({ id, oid, tid, title }) {
  const result = await pool.query(
    `UPDATE catalog_items SET title = $4 WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3 RETURNING id`,
    [id, oid, tid, title]
  );
  return result.rows[0] ?? null;
}

/** Deletes one personal item. RETURNING id lets the caller distinguish "deleted" from "not yours/doesn't exist". */
async function deleteOwnedCatalogItem(id, oid, tid) {
  const result = await pool.query(
    `DELETE FROM catalog_items WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3 RETURNING id`,
    [id, oid, tid]
  );
  return result.rows[0] ?? null;
}

module.exports = {
  pool,
  waitForDatabase,
  upsertUser,
  listSharedCatalogItems,
  getCatalogItem,
  deleteCatalogItemsByCategory,
  insertCatalogItem,
  updateCatalogItem,
  reorderCatalogItems,
  updateCatalogItemThumbnail,
  updateCatalogItemContent,
  deleteCatalogItem,
  listGroupsForCategory,
  getGroup,
  createGroup,
  updateGroup,
  reorderGroups,
  deleteGroup,
  getOrCreateGroupIdForCategory,
  getOrCreateTag,
  listAllTagNames,
  setItemTags,
  listPersonalCatalogItems,
  getOwnedCatalogItem,
  updateOwnedCatalogItemContent,
  renameOwnedCatalogItem,
  deleteOwnedCatalogItem,
};
