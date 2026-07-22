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
  thumbnail_path TEXT,      -- relative path under assets/catalog/thumbnails/, nullable
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
