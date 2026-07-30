#!/bin/sh
# Substitutes the real domain into the pre-built manifest at container
# startup (not image build time), so the same image can be pointed at a
# different domain just by changing DOMAIN in the compose file — no rebuild.
set -e

DOMAIN="${DOMAIN:-toolbar.gavinchow.me}"
sed -i "s#__DOMAIN__#${DOMAIN}#g" /app/dist/manifest.xml

# Same "same image, no rebuild" substitution pattern as DOMAIN above —
# lets multiple instances built from this one image (e.g. a beta and a
# production instance) be sideloaded/installed side by side as genuinely
# distinct add-ins instead of colliding as "the same add-in, different
# source." Defaults to the original single-instance GUID, so any existing
# deployment that never sets MANIFEST_ID keeps its current identity
# unchanged.
MANIFEST_ID="${MANIFEST_ID:-15e2608e-07d9-4f12-8a7c-18158275f61b}"
sed -i "s#__MANIFEST_ID__#${MANIFEST_ID}#g" /app/dist/manifest.xml

# Same pattern again for the Azure App Registration client ID — NOT a
# secret, but IS environment-specific (different domains can be registered
# under different App Registrations). Two targets: src/config/auth.json
# (read fresh at Node startup by server.js's require()) and dist/
# taskpane.js (esbuild bakes auth.json's clientId in as a literal string
# at build time, via src/auth/msal.ts's import — so this one's a source-
# code substitution, not just a data file). Defaults to the original
# gavinchow.me App Registration's client ID, so any existing deployment
# that never sets AZURE_CLIENT_ID keeps authenticating against the same
# app it always has.
AZURE_CLIENT_ID="${AZURE_CLIENT_ID:-2d499447-2b3d-4da8-ba78-ff5d1b1699b1}"
sed -i "s#__AZURE_CLIENT_ID__#${AZURE_CLIENT_ID}#g" /app/src/config/auth.json /app/dist/taskpane.js

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
