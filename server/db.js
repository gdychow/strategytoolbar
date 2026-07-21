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

module.exports = { pool, waitForDatabase, upsertUser };
