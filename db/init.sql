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
    category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams', 'maps', 'clipart', 'frameworks', 'flags')
  ),
  title TEXT NOT NULL,
  insert_mode TEXT NOT NULL CHECK (insert_mode IN ('reconstruct', 'file', 'unicode-char')),
  source_file TEXT UNIQUE,  -- 'file' mode only, e.g. 'text/text-010.pptx', resolved under CATALOG_DIR
  reconstruct_spec JSONB,   -- 'reconstruct' mode only: preset type(s), position/size, fill, line, rotation, adjustments, text runs
  unicode_char TEXT,        -- 'unicode-char' mode only: the literal character (e.g. '≈') to insert into the current text selection
  thumbnail_path TEXT,      -- relative path under CATALOG_DIR/thumbnails/ (persistent volume, not the image), nullable
  sort_order INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  FOREIGN KEY (owner_oid, owner_tid) REFERENCES users (oid, tid),
  CHECK (
    (insert_mode = 'file' AND source_file IS NOT NULL AND reconstruct_spec IS NULL AND unicode_char IS NULL) OR
    (insert_mode = 'reconstruct' AND reconstruct_spec IS NOT NULL AND source_file IS NULL AND unicode_char IS NULL) OR
    (insert_mode = 'unicode-char' AND unicode_char IS NOT NULL AND source_file IS NULL AND reconstruct_spec IS NULL)
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
    category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams', 'maps', 'clipart', 'frameworks', 'flags')
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

-- Task Pane Phase 14: company libraries. company_domain is a plain TEXT
-- scoping key (the domain itself, e.g. 'acmecorp.com'), not a synthetic id
-- into a separate companies table — consistent with how is_admin already
-- avoids a users-table column in favor of something derived rather than
-- stored, just here it's the domain, not a boolean, and it IS stored since
-- (unlike ADMIN_EMAILS) it isn't cheaply re-derivable from an env var.
-- Consumer-domain emails (gmail.com, etc. — see server/consumerDomains.js)
-- get company_domain = NULL, same as everyone did before this phase.
ALTER TABLE users ADD COLUMN IF NOT EXISTS company_domain TEXT;
-- Set once, either by auto-promotion (the first user ever seen for a given
-- company_domain) or by an explicit promote/demote action (see
-- server.js's /admin/company-admins routes) — never silently recomputed
-- on login the way is_admin is, since it's real mutable state, not a live
-- env-var lookup.
ALTER TABLE users ADD COLUMN IF NOT EXISTS is_company_admin BOOLEAN NOT NULL DEFAULT false;

-- A company-scoped item (no owner, no global category) alongside the two
-- existing scopes (global: both NULL; personal: owner_oid/tid set).
ALTER TABLE catalog_items ADD COLUMN IF NOT EXISTS company_domain TEXT;
ALTER TABLE catalog_items ADD CONSTRAINT catalog_items_owner_or_company_not_both
  CHECK (owner_oid IS NULL OR company_domain IS NULL);

-- category becomes conditionally required — global items keep a real
-- category (their gallery tab IS a category), company items don't (their
-- gallery tab is the company itself) — same "orthogonal scoping columns"
-- pattern as owner_oid vs. company_domain above. Personal items may or
-- may not have one (Phase 13 already defaults new personal items to a
-- category; harmless, not special-cased here).
ALTER TABLE catalog_items ALTER COLUMN category DROP NOT NULL;
ALTER TABLE catalog_items DROP CONSTRAINT IF EXISTS catalog_items_category_check;
ALTER TABLE catalog_items ADD CONSTRAINT catalog_items_category_check CHECK (
  category IS NULL OR category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams', 'maps', 'clipart', 'frameworks', 'flags')
);
ALTER TABLE catalog_items ADD CONSTRAINT catalog_items_category_scope
  CHECK ((company_domain IS NULL AND owner_oid IS NULL) = (category IS NOT NULL) OR owner_oid IS NOT NULL);

ALTER TABLE catalog_groups ADD COLUMN IF NOT EXISTS company_domain TEXT;
ALTER TABLE catalog_groups ALTER COLUMN category DROP NOT NULL;
ALTER TABLE catalog_groups DROP CONSTRAINT IF EXISTS catalog_groups_category_check;
ALTER TABLE catalog_groups ADD CONSTRAINT catalog_groups_category_check CHECK (
  category IS NULL OR category IN ('text', 'objects', 'shapes', 'stamps', 'tables', 'symbols', 'diagrams', 'maps', 'clipart', 'frameworks', 'flags')
);
ALTER TABLE catalog_groups ADD CONSTRAINT catalog_groups_scope_check
  CHECK ((category IS NOT NULL) <> (company_domain IS NOT NULL)); -- exactly one set
-- Postgres treats NULL <> NULL as unknown (not true) in a plain
-- UNIQUE(category, name), which would silently fail to catch two
-- different companies' groups both named e.g. "Templates" — hence two
-- partial unique indexes instead of one constraint, mirroring how
-- idx_catalog_items_category_shared above already uses a partial index
-- rather than a plain one.
ALTER TABLE catalog_groups DROP CONSTRAINT IF EXISTS catalog_groups_category_name_key;
CREATE UNIQUE INDEX IF NOT EXISTS catalog_groups_global_unique ON catalog_groups (category, name) WHERE company_domain IS NULL;
CREATE UNIQUE INDEX IF NOT EXISTS catalog_groups_company_unique ON catalog_groups (company_domain, name) WHERE company_domain IS NOT NULL;

-- Task Pane Phase 15: explicit account registration. A row existing in
-- `users` no longer implies the person has actually finished setting up
-- an account (see server/db.js's completeRegistration) — is_registered is
-- the real access gate for every account-dependent route from this phase
-- on. full_name/company_name/job_title are captured at registration and
-- are deliberately separate from display_name/company_domain (the raw,
-- always-refreshed Microsoft/email-derived values) so an edit here never
-- gets silently overwritten on a later sign-in.
ALTER TABLE users ADD COLUMN IF NOT EXISTS is_registered BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE users ADD COLUMN IF NOT EXISTS full_name TEXT;
ALTER TABLE users ADD COLUMN IF NOT EXISTS company_name TEXT;
ALTER TABLE users ADD COLUMN IF NOT EXISTS job_title TEXT;
ALTER TABLE users ADD COLUMN IF NOT EXISTS terms_accepted_at TIMESTAMPTZ;
ALTER TABLE users ADD COLUMN IF NOT EXISTS registered_at TIMESTAMPTZ;

-- One row per user for now (UNIQUE(owner_oid, owner_tid)) — seat/company-
-- wide billing would be a bigger redesign, deliberately not built ahead
-- of need. status starts 'pending' at registration since no real payment
-- happens yet; the stripe_* columns are nullable placeholders, a seam for
-- real billing later, not wired to anything in this phase. Subscription
-- status does not gate access anywhere yet — is_registered on users is
-- the only access gate this phase wires up.
CREATE TABLE IF NOT EXISTS subscriptions (
  id SERIAL PRIMARY KEY,
  owner_oid TEXT NOT NULL,
  owner_tid TEXT NOT NULL,
  plan TEXT NOT NULL CHECK (plan IN ('monthly', 'annual')),
  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'trial', 'active', 'past_due', 'canceled')),
  current_period_start TIMESTAMPTZ,
  current_period_end TIMESTAMPTZ,
  stripe_customer_id TEXT,
  stripe_subscription_id TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (owner_oid, owner_tid),
  FOREIGN KEY (owner_oid, owner_tid) REFERENCES users (oid, tid)
);

-- Task Pane Phase 20: whole .potx presentation templates (not individual
-- slide content — no category/insert_mode/groups/tags, a materially
-- simpler shape than catalog_items). Same three-way scope as catalog_items
-- (personal/company/global) but only two scoping columns needed — there's
-- no separate "shared, no owner, no company, but still has a category"
-- case here, so both-null just means global directly, no third column.
-- No thumbnail_path: upload is a plain file picker with no PowerPoint API
-- involvement, so there's no live-document capture to generate one from;
-- the column can be added additively later if real thumbnails are wanted.
CREATE TABLE IF NOT EXISTS templates (
  id SERIAL PRIMARY KEY,
  owner_oid TEXT,
  owner_tid TEXT,
  company_domain TEXT,
  title TEXT NOT NULL,
  description TEXT,
  source_file TEXT NOT NULL UNIQUE,
  sort_order INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  FOREIGN KEY (owner_oid, owner_tid) REFERENCES users (oid, tid),
  CHECK (owner_oid IS NULL OR company_domain IS NULL)
);
CREATE INDEX IF NOT EXISTS idx_templates_owner ON templates (owner_oid, owner_tid, sort_order) WHERE owner_oid IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_templates_company ON templates (company_domain, sort_order) WHERE company_domain IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_templates_global ON templates (sort_order) WHERE owner_oid IS NULL AND company_domain IS NULL;

-- Admin UI Phase 22: user administration. is_global_admin is separate from
-- (and OR'd with) the ADMIN_EMAILS env-var check in isAdminEmail() —
-- env-file admins stay admin regardless of this column and can never be
-- demoted via the UI; this column is only for admins granted through the
-- UI, which *are* revocable. is_suspended/deleted_at are enforced at
-- session-issuance/refresh time (see server.js), not per-request.
ALTER TABLE users ADD COLUMN IF NOT EXISTS is_global_admin BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE users ADD COLUMN IF NOT EXISTS is_suspended BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE users ADD COLUMN IF NOT EXISTS suspended_at TIMESTAMPTZ;
-- A deleted user's row is never actually removed (catalog_items/templates/
-- subscriptions all FK-reference users with no cascade) — deleted_at set +
-- PII columns scrubbed by softDeleteUser() is the real "delete."
ALTER TABLE users ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMPTZ;

-- "Free access tier" reuses the existing subscriptions table (already has
-- the Stripe placeholder columns for real billing later) rather than a new
-- users column — one source of truth for "what access level does this
-- user have." A free-tier user gets plan='free', status='active'.
ALTER TABLE subscriptions DROP CONSTRAINT subscriptions_plan_check;
ALTER TABLE subscriptions ADD CONSTRAINT subscriptions_plan_check CHECK (plan IN ('monthly', 'annual', 'free'));

-- actor_email/target_email are captured at write-time, not joined live, so
-- the log stays legible even after a target's PII is scrubbed by a later
-- delete. Real FKs are safe here since rows are only ever soft-deleted.
CREATE TABLE IF NOT EXISTS admin_actions (
  id SERIAL PRIMARY KEY,
  actor_oid TEXT NOT NULL,
  actor_tid TEXT NOT NULL,
  actor_email TEXT,
  action TEXT NOT NULL CHECK (action IN (
    'promote_company_admin', 'demote_company_admin',
    'promote_global_admin', 'demote_global_admin',
    'suspend', 'unsuspend', 'delete', 'set_free_tier', 'unset_free_tier'
  )),
  target_oid TEXT NOT NULL,
  target_tid TEXT NOT NULL,
  target_email TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  FOREIGN KEY (actor_oid, actor_tid) REFERENCES users (oid, tid),
  FOREIGN KEY (target_oid, target_tid) REFERENCES users (oid, tid)
);
CREATE INDEX IF NOT EXISTS idx_admin_actions_target ON admin_actions (target_oid, target_tid);
CREATE INDEX IF NOT EXISTS idx_admin_actions_created ON admin_actions (created_at DESC);
