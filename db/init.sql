-- Runs automatically on first startup of the postgres container (official
-- image convention: anything in /docker-entrypoint-initdb.d/ executes once,
-- against an empty data directory — see docker-compose.yml's volume mount).
--
-- users is keyed by (oid, tid), not oid alone — Microsoft's own guidance,
-- since oid isn't guaranteed globally unique across tenants.
--
-- is_admin is deliberately NOT a column here — it's computed live from the
-- ADMIN_EMAILS env var at session-issuance time (see server/auth.js), so
-- there's no admin-list/database drift to manage.
CREATE TABLE IF NOT EXISTS users (
  oid TEXT NOT NULL,
  tid TEXT NOT NULL,
  email TEXT,
  display_name TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  last_seen_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (oid, tid)
);

-- Tier 3's shared/admin content catalog. owner_oid/owner_tid NULL means a
-- shared/admin-managed item; non-null means a private item scoped to that
-- user (not populated yet — private items are a later fast-follow).
--
-- Each item is inserted one of two ways, chosen at content-prep time by a
-- mechanical check of the source shape's XML (prstGeom vs custGeom), not a
-- subjective fidelity judgment call:
--   'reconstruct' — built directly on the user's current slide via
--     addGeometricShape/addTextBox/addGroup from reconstruct_spec. True
--     one-click insert, used whenever the source is plain preset geometry.
--   'file' — inserted via insertSlidesFromBase64 from source_file (a
--     pre-sliced single-slide pptx), since PowerPoint JS has no API for
--     custom/freeform vector geometry. Requires a manual copy/paste-finish
--     step in the task pane (see src/features/libraryInsert.ts).
CREATE TABLE IF NOT EXISTS catalog_items (
  id SERIAL PRIMARY KEY,
  owner_oid TEXT,
  owner_tid TEXT,
  category TEXT NOT NULL CHECK (
    category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams')
  ),
  title TEXT NOT NULL,
  insert_mode TEXT NOT NULL CHECK (insert_mode IN ('reconstruct', 'file')),
  source_file TEXT UNIQUE,  -- 'file' mode only, e.g. 'text/text-010.pptx', resolved under CATALOG_DIR
  reconstruct_spec JSONB,   -- 'reconstruct' mode only: preset type(s), position/size, fill, line, rotation, adjustments, text runs
  thumbnail_path TEXT,      -- relative path under CATALOG_DIR/thumbnails/ (persistent volume, not the image), nullable
  sort_order INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  FOREIGN KEY (owner_oid, owner_tid) REFERENCES users (oid, tid),
  CHECK (
    (insert_mode = 'file' AND source_file IS NOT NULL AND reconstruct_spec IS NULL) OR
    (insert_mode = 'reconstruct' AND reconstruct_spec IS NOT NULL AND source_file IS NULL)
  )
);
CREATE INDEX IF NOT EXISTS idx_catalog_items_category_shared
  ON catalog_items (category, sort_order) WHERE owner_oid IS NULL;

-- Tier 3 Phase 5: admin-defined, admin-ordered sub-groupings within one
-- category (e.g. "Pyramids" inside Diagrams) — deliberately NOT derived
-- from tags below, since a group needs exactly one value and an explicit
-- order, which a free multi-value tag doesn't naturally give you.
CREATE TABLE IF NOT EXISTS catalog_groups (
  id SERIAL PRIMARY KEY,
  category TEXT NOT NULL CHECK (
    category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams')
  ),
  name TEXT NOT NULL,
  sort_order INTEGER NOT NULL DEFAULT 0,
  UNIQUE (category, name)
);

-- An item's group must belong to the same category as the item itself -
-- enforced at the application layer (server.js), not here, since a plain
-- CHECK can't reference another table. If an admin edit changes an item's
-- category, the server resets group_id to NULL in that same update rather
-- than leaving a group reference from the old category dangling.
-- ON DELETE SET NULL: deleting a group is a simple admin action that
-- shouldn't require first reassigning every item in it - they just fall
-- back to ungrouped (the gallery's "Other" bucket) automatically.
ALTER TABLE catalog_items ADD COLUMN IF NOT EXISTS group_id INTEGER REFERENCES catalog_groups (id) ON DELETE SET NULL;

-- Fixed-ish vocabulary, multiple per item, but admins can add new ones
-- from the /admin item-edit form (with a client-side near-match
-- confirmation step - see server.js - to discourage near-duplicates like
-- "arrow" vs "arrows"; not enforced here beyond plain uniqueness).
CREATE TABLE IF NOT EXISTS tags (
  id SERIAL PRIMARY KEY,
  name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS catalog_item_tags (
  item_id INTEGER NOT NULL REFERENCES catalog_items (id) ON DELETE CASCADE,
  tag_id INTEGER NOT NULL REFERENCES tags (id) ON DELETE CASCADE,
  PRIMARY KEY (item_id, tag_id)
);
