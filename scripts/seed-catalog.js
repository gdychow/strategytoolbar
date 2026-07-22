/**
 * One-time/rerunnable seed step: loads a db/seed/catalog-<category>.json
 * file into the catalog_items table. Deletes every existing shared item in
 * that category first, then inserts everything fresh — simpler than trying
 * to make ON CONFLICT idempotent across two insert_mode shapes, since
 * source_file is UNIQUE but NULL for 'reconstruct' items (and NULL never
 * conflicts with NULL in Postgres). A failure partway through just leaves
 * the category needing a rerun — low-frequency manual step, not worth a
 * real transaction for.
 *
 * Usage: node scripts/seed-catalog.js db/seed/catalog-text.json
 */
const fs = require("fs");
const path = require("path");

const { pool, deleteCatalogItemsByCategory, insertCatalogItem } = require("../server/db");

async function main() {
  const seedPath = process.argv[2];
  if (!seedPath) {
    console.error("Usage: node scripts/seed-catalog.js db/seed/catalog-<category>.json");
    process.exit(1);
  }

  const { category, items } = JSON.parse(fs.readFileSync(path.resolve(seedPath), "utf8"));

  await deleteCatalogItemsByCategory(category);
  for (const item of items) {
    await insertCatalogItem({
      category,
      title: item.title,
      insertMode: item.insertMode,
      sourceFile: item.sourceFile,
      reconstructSpec: item.reconstructSpec,
      thumbnailPath: item.thumbnail,
      sortOrder: item.sortOrder,
    });
  }
  console.log(`Seeded ${items.length} item(s) into category "${category}".`);
  await pool.end();
}

main().catch((err) => {
  console.error("Seed failed:", err);
  process.exit(1);
});
