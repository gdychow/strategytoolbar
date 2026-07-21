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

-- Placeholder shape for Tier 3's catalog — not built out further in this
-- phase. owner_oid/owner_tid NULL means a shared/admin-managed item;
-- non-null means a private item scoped to that user.
CREATE TABLE IF NOT EXISTS catalog_items (
  id SERIAL PRIMARY KEY,
  owner_oid TEXT,
  owner_tid TEXT,
  title TEXT NOT NULL,
  category TEXT,
  source_ref TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  FOREIGN KEY (owner_oid, owner_tid) REFERENCES users (oid, tid)
);
