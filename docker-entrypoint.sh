#!/bin/sh
# Substitutes the real domain into the pre-built manifest at container
# startup (not image build time), so the same image can be pointed at a
# different domain just by changing DOMAIN in the compose file — no rebuild.
set -e

DOMAIN="${DOMAIN:-toolbar.gavinchow.me}"
sed -i "s#__DOMAIN__#${DOMAIN}#g" /app/dist/manifest.xml

# One-time seed of the persistent catalog volume's thumbnails/ dir from the
# image's baked-in copy (dist/assets/catalog/thumbnails, put there by
# build.mjs). Guarded on the dir not already existing, so this only fires
# on a genuinely fresh volume — never clobbers a thumbnail an admin has
# since replaced via /admin, which now lives on the volume, not the image.
CATALOG_DIR="${CATALOG_DIR:-/app/data/catalog}"
if [ ! -d "$CATALOG_DIR/thumbnails" ]; then
  mkdir -p "$CATALOG_DIR/thumbnails"
  cp -r /app/dist/assets/catalog/thumbnails/. "$CATALOG_DIR/thumbnails/"
fi

exec node server.js
