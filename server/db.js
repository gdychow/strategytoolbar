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

/**
 * Upserts a user by (oid, tid) and bumps last_seen_at — called on every
 * successful token verification. companyDomain (Task Pane Phase 14,
 * already resolved by the caller via server/consumerDomains.js's
 * deriveCompanyDomain) is refreshed on every call — an email's domain
 * could theoretically change across logins.
 *
 * Task Pane Phase 15: a brand-new row always starts is_company_admin =
 * false and is_registered = false — the "first user of a domain becomes
 * that domain's admin" auto-promotion no longer happens here. It moved to
 * completeRegistration (below), scoped to is_registered = true rows only,
 * so an abandoned/incomplete sign-in can never permanently squat a
 * company's admin slot. The ON CONFLICT branch only ever refreshes
 * email/display_name/company_domain/last_seen_at — it must never touch
 * is_registered/full_name/company_name/job_title/terms_accepted_at/
 * registered_at/is_company_admin, all of which are owned exclusively by
 * completeRegistration and the promote/demote routes in server.js. Same
 * for is_global_admin/is_suspended/deleted_at (Admin UI Phase 22) — a
 * deleted user's row still gets its email/display_name briefly
 * re-populated here on a sign-in attempt before server.js's post-upsert
 * deleted_at check refuses to issue a session; harmless since no cookie
 * ever results and every list query filters WHERE deleted_at IS NULL.
 */
async function upsertUser({ oid, tid, email, displayName, companyDomain }) {
  const result = await pool.query(
    `INSERT INTO users (oid, tid, email, display_name, company_domain, last_seen_at)
     VALUES ($1, $2, $3, $4, $5::text, now())
     ON CONFLICT (oid, tid)
     DO UPDATE SET email = EXCLUDED.email, display_name = EXCLUDED.display_name, company_domain = EXCLUDED.company_domain, last_seen_at = now()
     RETURNING oid, tid, email, display_name, company_domain, is_company_admin, is_registered, full_name, company_name, job_title, created_at, last_seen_at,
               is_global_admin, is_suspended, deleted_at`,
    [oid, tid, email, displayName, companyDomain ?? null]
  );
  return result.rows[0];
}

/** Plain lookup by primary key — used by the session-refresh middleware to pick up a promote/demote, a suspension/deletion, or a just-completed registration without forcing a full sign-out. */
async function getUserByKey(oid, tid) {
  const result = await pool.query(
    `SELECT oid, tid, email, display_name, company_domain, is_company_admin, is_registered, full_name, company_name, job_title,
            is_global_admin, is_suspended, deleted_at
     FROM users WHERE oid = $1 AND tid = $2`,
    [oid, tid]
  );
  return result.rows[0] ?? null;
}

/**
 * Task Pane Phase 15: completes registration for an already-signed-in but
 * not-yet-registered user — the only place is_registered ever flips to
 * true. Also where the "first user of a domain becomes that domain's
 * admin" auto-promotion now happens (relocated from upsertUser's old
 * first-INSERT check), scoped to is_registered = true rows so it only
 * ever considers people who've actually finished setting up an account.
 * One transaction: updates the user row, then upserts a subscriptions row
 * for the chosen plan (ON CONFLICT covers a retried/duplicate submission,
 * not a real re-subscribe flow — there is no real billing yet).
 */
async function completeRegistration({ oid, tid, fullName, companyName, jobTitle, plan }) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const userResult = await client.query(
      `UPDATE users
       SET full_name = $3,
           company_name = $4,
           job_title = $5,
           terms_accepted_at = now(),
           registered_at = now(),
           is_registered = true,
           is_company_admin = (
             company_domain IS NOT NULL
             AND NOT EXISTS (
               SELECT 1 FROM users u2
               WHERE u2.company_domain = users.company_domain
                 AND u2.is_registered = true
                 AND (u2.oid, u2.tid) <> (users.oid, users.tid)
             )
           )
       WHERE oid = $1 AND tid = $2
       RETURNING oid, tid, email, display_name, company_domain, is_company_admin, is_registered, full_name, company_name, job_title,
                 is_global_admin, is_suspended, deleted_at`,
      [oid, tid, fullName, companyName ?? null, jobTitle ?? null]
    );
    const user = userResult.rows[0];
    if (!user) {
      await client.query("ROLLBACK");
      return null;
    }
    await client.query(
      `INSERT INTO subscriptions (owner_oid, owner_tid, plan, status)
       VALUES ($1, $2, $3, 'pending')
       ON CONFLICT (owner_oid, owner_tid) DO UPDATE SET plan = EXCLUDED.plan, updated_at = now()`,
      [oid, tid, plan]
    );
    await client.query("COMMIT");
    return user;
  } catch (err) {
    await client.query("ROLLBACK");
    throw err;
  } finally {
    client.release();
  }
}

// Shared by listAllUsers/listUsersByCompanyDomain (Admin UI Phase 22) — the
// subscriptions LEFT JOIN is 1:1 (owner_oid/owner_tid is UNIQUE there), so
// this never duplicates a user row; a user with no subscriptions row at all
// (pre-Phase-15 grandfathered accounts) just gets plan/status = NULL.
const USER_ADMIN_LIST_SELECT = `
  SELECT u.oid, u.tid, u.email, u.display_name, u.company_domain,
         u.is_company_admin, u.is_global_admin, u.is_registered, u.is_suspended,
         u.created_at, u.registered_at, u.last_seen_at,
         s.plan, s.status
  FROM users u
  LEFT JOIN subscriptions s ON s.owner_oid = u.oid AND s.owner_tid = u.tid
`;

/** Every non-deleted user sharing a company_domain, with the dates/flags the Users admin page needs. */
async function listUsersByCompanyDomain(companyDomain) {
  const result = await pool.query(
    `${USER_ADMIN_LIST_SELECT} WHERE u.company_domain = $1 AND u.deleted_at IS NULL ORDER BY u.email`,
    [companyDomain]
  );
  return result.rows;
}

/** Every non-deleted user, unscoped — the global admin's full user list. Nothing else in this file lists users without a company_domain filter. */
async function listAllUsers() {
  const result = await pool.query(
    `${USER_ADMIN_LIST_SELECT} WHERE u.deleted_at IS NULL ORDER BY u.company_domain NULLS LAST, u.email`
  );
  return result.rows;
}

/** Distinct company domains with a member count — backs the global admin's company picker on the Users page (nothing like this existed before Phase 22; company-scoped views previously only worked if you already knew the domain string). */
async function listCompanyDomains() {
  const result = await pool.query(
    `SELECT company_domain, count(*)::int AS user_count
     FROM users WHERE company_domain IS NOT NULL AND deleted_at IS NULL
     GROUP BY company_domain ORDER BY company_domain`
  );
  return result.rows;
}

/** Backs the "always at least one company admin" guard in the demote route. */
async function countCompanyAdmins(companyDomain) {
  const result = await pool.query(
    `SELECT count(*)::int AS count FROM users WHERE company_domain = $1 AND is_company_admin = true AND deleted_at IS NULL`,
    [companyDomain]
  );
  return result.rows[0].count;
}

/** Sets (or clears) one user's company-admin status — the only place is_company_admin is ever written after the initial auto-promotion in upsertUser. */
async function setCompanyAdmin(oid, tid, isCompanyAdmin) {
  const result = await pool.query(
    `UPDATE users SET is_company_admin = $3 WHERE oid = $1 AND tid = $2 RETURNING oid, tid, company_domain, is_company_admin`,
    [oid, tid, isCompanyAdmin]
  );
  return result.rows[0] ?? null;
}

/** Sets (or clears) one user's global-admin status. Separate from (and OR'd with) the ADMIN_EMAILS env-var check — see isAdminEmail in server/auth.js — so this only ever affects UI-granted admins, never the env-file ones. */
async function setGlobalAdmin(oid, tid, isGlobalAdmin) {
  const result = await pool.query(
    `UPDATE users SET is_global_admin = $3 WHERE oid = $1 AND tid = $2 RETURNING oid, tid, email, is_global_admin`,
    [oid, tid, isGlobalAdmin]
  );
  return result.rows[0] ?? null;
}

/** Toggles suspension — enforced at session-issuance/refresh time in server.js, not per-request. */
async function setSuspended(oid, tid, isSuspended) {
  const result = await pool.query(
    `UPDATE users SET is_suspended = $3, suspended_at = CASE WHEN $3 THEN now() ELSE NULL END
     WHERE oid = $1 AND tid = $2 RETURNING oid, tid, email, is_suspended`,
    [oid, tid, isSuspended]
  );
  return result.rows[0] ?? null;
}

/**
 * The real "delete" — the row itself is never removed (catalog_items/
 * templates/subscriptions all FK-reference users with no cascade, so a hard
 * DELETE would fail the moment this user owns anything). deleted_at +
 * scrubbed PII is what "deleted" means here; company_domain is deliberately
 * left alone since it's needed for existing owned company-scoped content
 * and for the audit log's domain-scoping to keep working. The
 * `deleted_at IS NULL` guard makes a repeat call a safe no-op.
 */
async function softDeleteUser(oid, tid) {
  const result = await pool.query(
    `UPDATE users
     SET deleted_at = now(), email = NULL, display_name = NULL, full_name = NULL,
         company_name = NULL, job_title = NULL, is_company_admin = false, is_global_admin = false
     WHERE oid = $1 AND tid = $2 AND deleted_at IS NULL
     RETURNING oid, tid`,
    [oid, tid]
  );
  return result.rows[0] ?? null;
}

/**
 * "Free access tier" lives on the existing subscriptions table (plan/status
 * columns already exist, alongside the unused Stripe placeholder columns
 * for real billing later) rather than a new users column — one source of
 * truth for "what access level does this user have." Enabling sets
 * plan='free'/status='active'; disabling reverts status to 'pending' (the
 * original registration-time default) and leaves plan untouched, since
 * plan is NOT NULL and can't be cleared outright.
 */
async function setFreeTier(oid, tid, enabled) {
  const result = await pool.query(
    enabled
      ? `INSERT INTO subscriptions (owner_oid, owner_tid, plan, status)
         VALUES ($1, $2, 'free', 'active')
         ON CONFLICT (owner_oid, owner_tid) DO UPDATE SET plan = 'free', status = 'active', updated_at = now()
         RETURNING owner_oid, owner_tid, plan, status`
      : `UPDATE subscriptions SET status = 'pending', updated_at = now()
         WHERE owner_oid = $1 AND owner_tid = $2
         RETURNING owner_oid, owner_tid, plan, status`,
    [oid, tid]
  );
  return result.rows[0] ?? null;
}

/** Admin UI Phase 22 audit log — one row per promote/demote/suspend/delete/free-tier action. actor_email/target_email are captured at write-time (not joined live) so the log stays legible even after a target's PII is scrubbed by a later delete. */
async function logAdminAction({ actorOid, actorTid, actorEmail, action, targetOid, targetTid, targetEmail }) {
  await pool.query(
    `INSERT INTO admin_actions (actor_oid, actor_tid, actor_email, action, target_oid, target_tid, target_email)
     VALUES ($1, $2, $3, $4, $5, $6, $7)`,
    [actorOid, actorTid, actorEmail ?? null, action, targetOid, targetTid, targetEmail ?? null]
  );
}

/** domain=null → unscoped (global admin's view of every action); domain set → only actions whose target currently belongs to that company. */
async function listRecentAdminActions({ domain = null, limit = 100 } = {}) {
  const result = await pool.query(
    domain
      ? `SELECT a.* FROM admin_actions a
         JOIN users target ON target.oid = a.target_oid AND target.tid = a.target_tid
         WHERE target.company_domain = $1
         ORDER BY a.created_at DESC LIMIT $2`
      : `SELECT * FROM admin_actions ORDER BY created_at DESC LIMIT $1`,
    domain ? [domain, limit] : [limit]
  );
  return result.rows;
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

/** Task Pane Phase 14 counterpart of listSharedCatalogItems, scoped by company_domain instead of category — a genuinely different WHERE and result set (no category shown), so kept as its own function. */
async function listCompanyCatalogItems(companyDomain) {
  const result = await pool.query(
    `SELECT ci.id, ci.company_domain, ci.title, ci.insert_mode, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.company_domain = $1
     GROUP BY ci.id, cg.name
     ORDER BY ci.sort_order, ci.id`,
    [companyDomain]
  );
  return result.rows;
}

/**
 * A single catalog item by ID, including source_file — used to resolve the
 * file for a 'file'-mode insert, and by /admin's edit form. company_domain
 * (Task Pane Phase 14) is included so the route-level scope-aware auth
 * check (req.user.isAdmin || (isCompanyAdmin && companyDomain matches))
 * has something to compare against.
 */
async function getCatalogItem(id) {
  const result = await pool.query(
    `SELECT ci.id, ci.category, ci.company_domain, ci.title, ci.insert_mode, ci.source_file, ci.reconstruct_spec, ci.unicode_char, ci.thumbnail_path, ci.sort_order,
            ${CATALOG_ITEM_GROUP_TAGS_SELECT}
     FROM catalog_items ci
     ${CATALOG_ITEM_GROUP_TAGS_JOIN}
     WHERE ci.id = $1
     GROUP BY ci.id, cg.name`,
    [id]
  );
  return result.rows[0] ?? null;
}

/** A category's groups, in display order — used by GET /admin and the group <select> on the item edit form. */
async function listGroupsForCategory(category) {
  const result = await pool.query(
    `SELECT id, category, name, sort_order FROM catalog_groups WHERE category = $1 ORDER BY sort_order, name`,
    [category]
  );
  return result.rows;
}

/** Task Pane Phase 14 counterpart of listGroupsForCategory, scoped by company_domain instead — a genuinely different WHERE (and result set), so kept as its own function rather than parameterized. */
async function listGroupsForCompany(companyDomain) {
  const result = await pool.query(
    `SELECT id, company_domain, name, sort_order FROM catalog_groups WHERE company_domain = $1 ORDER BY sort_order, name`,
    [companyDomain]
  );
  return result.rows;
}

/** A single group by ID, mainly so /admin/groups' edit/delete routes know which scope's page to redirect back to (neither category nor company_domain is resubmitted from those forms — see updateGroup's comment on why category isn't editable; the same now holds for company_domain). */
async function getGroup(id) {
  const result = await pool.query(
    `SELECT id, category, company_domain, name, sort_order FROM catalog_groups WHERE id = $1`,
    [id]
  );
  return result.rows[0] ?? null;
}

/**
 * Creates a new group, either within a category (global scope) or within a
 * company (Task Pane Phase 14) — exactly one of category/companyDomain is
 * expected to be set, matching catalog_groups' own CHECK constraint.
 * Whichever UNIQUE partial index applies (global vs. company) surfaces a
 * clear conflict on a duplicate name rather than silently duplicating.
 */
async function createGroup({ category, companyDomain, name, sortOrder }) {
  const result = await pool.query(
    `INSERT INTO catalog_groups (category, company_domain, name, sort_order) VALUES ($1, $2, $3, $4) RETURNING id`,
    [category ?? null, companyDomain ?? null, name, sortOrder ?? 0]
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
 * Re-sequences a category's (or, since Task Pane Phase 14, a company's)
 * groups to match orderedGroupIds' order (0..n-1) — mirrors
 * reorderCatalogItems' transaction shape exactly. Exactly one of
 * category/companyDomain is expected to be set, branching which WHERE
 * predicate applies rather than duplicating the whole function, since the
 * two scopes only ever differ in that one clause. The synthetic
 * "Ungrouped" cluster has no backing row and is never part of
 * orderedGroupIds; it always renders last, client-side, unconditionally.
 */
async function reorderGroups({ category, companyDomain, orderedGroupIds }) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    for (let i = 0; i < orderedGroupIds.length; i++) {
      if (companyDomain) {
        await client.query(`UPDATE catalog_groups SET sort_order = $1 WHERE id = $2 AND company_domain = $3`, [
          i,
          orderedGroupIds[i],
          companyDomain,
        ]);
      } else {
        await client.query(`UPDATE catalog_groups SET sort_order = $1 WHERE id = $2 AND category = $3`, [
          i,
          orderedGroupIds[i],
          category,
        ]);
      }
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
  companyDomain,
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
    `INSERT INTO catalog_items (category, company_domain, title, insert_mode, source_file, reconstruct_spec, unicode_char, thumbnail_path, sort_order, group_id, owner_oid, owner_tid)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
     RETURNING id`,
    [
      category ?? null,
      companyDomain ?? null,
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
 *
 * Task Pane Phase 14: category is optional — a company-scoped item's edit
 * form has no category <select> at all (its gallery tab is the company,
 * not a category), so `category` arrives as null/undefined for those
 * edits. When it's omitted, the category/group_id columns are left
 * exactly as they are (COALESCE keeps the existing category — always
 * null for a company item — and group_id is set to whatever groupId was
 * given, with no category-changed reset since category never changes on
 * this path).
 */
async function updateCatalogItem({ id, title, category, groupId }) {
  const result = await pool.query(
    `UPDATE catalog_items
     SET title = $2,
         group_id = CASE WHEN $3::text IS NULL OR category = $3 THEN $4::integer ELSE NULL END,
         category = COALESCE($3, category)
     WHERE id = $1 AND owner_oid IS NULL
     RETURNING id`,
    [id, title, category ?? null, groupId ?? null]
  );
  return result.rows[0] ?? null;
}

/**
 * Re-sequences one category+group (or, since Task Pane Phase 14, one
 * company+group) cluster's sort_order to match orderedIds' order
 * (0..n-1), and reassigns group_id for every item in that list — covers
 * both "reordered within its own cluster" and "dragged into a different
 * cluster" in one call, since the destination cluster's full membership
 * is always what the client sends. The source cluster (if an item moved
 * out of it) is left with a gap in its sequence, which is harmless —
 * ordering only depends on relative order, not contiguity. The
 * category/company_domain equality check in each branch is defense in
 * depth against a crafted request smuggling an id from a different
 * scope into this call.
 */
async function reorderCatalogItems({ category, companyDomain, groupId, orderedIds }) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    for (let i = 0; i < orderedIds.length; i++) {
      if (companyDomain) {
        await client.query(
          `UPDATE catalog_items SET sort_order = $1, group_id = $2
           WHERE id = $3 AND company_domain = $4`,
          [i, groupId, orderedIds[i], companyDomain]
        );
      } else {
        await client.query(
          `UPDATE catalog_items SET sort_order = $1, group_id = $2
           WHERE id = $3 AND category = $4 AND owner_oid IS NULL`,
          [i, groupId, orderedIds[i], category]
        );
      }
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

// ---------------------------------------------------------------------------
// Task Pane Phase 20 — whole .potx templates. Own small table, not folded
// into catalog_items (no category/insert_mode/groups/tags to carry). Same
// personal-owned-vs-shared-unscoped split as the catalog_items functions
// above: personal rows are SQL-scoped by owner_oid/owner_tid; company/
// global rows are fetched/mutated unscoped, with authorization (canManageRow
// for managing, canAccessTemplate for browsing/using) checked at the route
// layer in server.js against the row this file already returns.
// ---------------------------------------------------------------------------

async function listPersonalTemplates(oid, tid) {
  const result = await pool.query(
    `SELECT id, title, description, sort_order FROM templates WHERE owner_oid = $1 AND owner_tid = $2 ORDER BY sort_order, id`,
    [oid, tid]
  );
  return result.rows;
}

async function listCompanyTemplates(companyDomain) {
  const result = await pool.query(
    `SELECT id, title, description, sort_order FROM templates WHERE company_domain = $1 ORDER BY sort_order, id`,
    [companyDomain]
  );
  return result.rows;
}

async function listGlobalTemplates() {
  const result = await pool.query(
    `SELECT id, title, description, sort_order FROM templates WHERE owner_oid IS NULL AND company_domain IS NULL ORDER BY sort_order, id`
  );
  return result.rows;
}

/** A single template by ID, any scope — used both to check canAccessTemplate/canManageRow against the row and to resolve source_file for the create-payload route. */
async function getTemplate(id) {
  const result = await pool.query(
    `SELECT id, owner_oid, owner_tid, company_domain, title, description, source_file, sort_order FROM templates WHERE id = $1`,
    [id]
  );
  return result.rows[0] ?? null;
}

async function insertTemplate({ ownerOid, ownerTid, companyDomain, title, description, sourceFile, sortOrder }) {
  const result = await pool.query(
    `INSERT INTO templates (owner_oid, owner_tid, company_domain, title, description, source_file, sort_order)
     VALUES ($1, $2, $3, $4, $5, $6, $7)
     RETURNING id, title, description`,
    [ownerOid ?? null, ownerTid ?? null, companyDomain ?? null, title, description ?? null, sourceFile, sortOrder ?? 0]
  );
  return result.rows[0];
}

async function getOwnedTemplate(id, oid, tid) {
  const result = await pool.query(
    `SELECT id, title, description, source_file FROM templates WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3`,
    [id, oid, tid]
  );
  return result.rows[0] ?? null;
}

async function renameOwnedTemplate({ id, oid, tid, title, description }) {
  const result = await pool.query(
    `UPDATE templates SET title = $4, description = $5, updated_at = now() WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3 RETURNING id`,
    [id, oid, tid, title, description ?? null]
  );
  return result.rows[0] ?? null;
}

async function deleteOwnedTemplate(id, oid, tid) {
  const result = await pool.query(`DELETE FROM templates WHERE id = $1 AND owner_oid = $2 AND owner_tid = $3 RETURNING id`, [
    id,
    oid,
    tid,
  ]);
  return result.rows[0] ?? null;
}

/** Unscoped rename, used for company/global templates once the route layer has already checked canManageRow against the fetched row. */
async function renameTemplate({ id, title, description }) {
  const result = await pool.query(`UPDATE templates SET title = $2, description = $3, updated_at = now() WHERE id = $1 RETURNING id`, [
    id,
    title,
    description ?? null,
  ]);
  return result.rows[0] ?? null;
}

/** Unscoped delete, used for company/global templates once the route layer has already checked canManageRow. */
async function deleteTemplate(id) {
  const result = await pool.query(`DELETE FROM templates WHERE id = $1 RETURNING id`, [id]);
  return result.rows[0] ?? null;
}

module.exports = {
  pool,
  waitForDatabase,
  upsertUser,
  getUserByKey,
  completeRegistration,
  listUsersByCompanyDomain,
  setCompanyAdmin,
  listAllUsers,
  listCompanyDomains,
  countCompanyAdmins,
  setGlobalAdmin,
  setSuspended,
  softDeleteUser,
  setFreeTier,
  logAdminAction,
  listRecentAdminActions,
  listSharedCatalogItems,
  listCompanyCatalogItems,
  getCatalogItem,
  deleteCatalogItemsByCategory,
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
  getOrCreateGroupIdForCategory,
  getOrCreateTag,
  listAllTagNames,
  setItemTags,
  listPersonalCatalogItems,
  getOwnedCatalogItem,
  updateOwnedCatalogItemContent,
  renameOwnedCatalogItem,
  deleteOwnedCatalogItem,
  listPersonalTemplates,
  listCompanyTemplates,
  listGlobalTemplates,
  getTemplate,
  insertTemplate,
  getOwnedTemplate,
  renameOwnedTemplate,
  deleteOwnedTemplate,
  renameTemplate,
  deleteTemplate,
};
