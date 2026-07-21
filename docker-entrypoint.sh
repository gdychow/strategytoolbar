#!/bin/sh
# Substitutes the real domain into the pre-built manifest at container
# startup (not image build time), so the same image can be pointed at a
# different domain just by changing DOMAIN in the compose file — no rebuild.
set -e

DOMAIN="${DOMAIN:-toolbar.gavinchow.me}"
sed -i "s#__DOMAIN__#${DOMAIN}#g" /app/dist/manifest.xml

exec node server.js
