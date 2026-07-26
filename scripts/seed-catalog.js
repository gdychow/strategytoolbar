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

const { pool, deleteCatalogItemsByCategory, insertCatalogItem, getOrCreateGroupIdForCategory } = require("../server/db");

async function main() {
  const seedPath = process.argv[2];
  if (!seedPath) {
    console.error("Usage: node scripts/seed-catalog.js db/seed/catalog-<category>.json");
    process.exit(1);
  }

  const { category, items } = JSON.parse(fs.readFileSync(path.resolve(seedPath), "utf8"));

  // Optional per-item groupName (e.g. Phase 7's Maps category, auto-grouped
  // by source file) - resolved to a group id once per distinct name here,
  // in first-appearance order (so e.g. "World Maps" from the first source
  // file seeded lands before "3D Globes" from the last), rather than
  // requiring the groups to already exist via /admin first.
  const groupNameToId = new Map();
  let nextGroupSortOrder = 0;
  for (const item of items) {
    if (item.groupName && !groupNameToId.has(item.groupName)) {
      nextGroupSortOrder += 1;
      groupNameToId.set(item.groupName, await getOrCreateGroupIdForCategory(category, item.groupName, nextGroupSortOrder));
    }
  }

  await deleteCatalogItemsByCategory(category);
  for (const item of items) {
    await insertCatalogItem({
      category,
      title: item.title,
      insertMode: item.insertMode,
      sourceFile: item.sourceFile,
      reconstructSpec: item.reconstructSpec,
      unicodeChar: item.unicodeChar,
      thumbnailPath: item.thumbnail,
      sortOrder: item.sortOrder,
      groupId: item.groupName ? groupNameToId.get(item.groupName) : null,
    });
  }
  console.log(`Seeded ${items.length} item(s) into category "${category}".`);
  await pool.end();
}

main().catch((err) => {
  console.error("Seed failed:", err);
  process.exit(1);
});
