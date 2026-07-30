#!/usr/bin/env sh
# Canonical way to rebuild and restart the app container. A bare
# `docker compose up --build` leaves GIT_COMMIT unset, which
# docker-compose.yml's build arg silently defaults to "unknown" for — the
# task pane's build-info line (top-right, visible even if Office.js never
# initializes) then shows "unknown" instead of a real commit hash. This
# script exists so that's never forgotten again.
set -e
cd "$(dirname "$0")"

# data/catalog is a host bind mount (not a named volume, since 2026-07-29),
# so unlike a named volume it never inherits the image's baked-in node:node
# ownership — on a fresh host where this directory doesn't exist yet,
# Docker auto-creates it as root:root on first `up`, and the container's
# non-root node user then can't write into it (mkdir/thumbnail-seed fails
# with "Permission denied"). Pre-creating it here, world-writable, sidesteps
# needing to know or coordinate the container user's UID across every host
# this runs on (Mac dev, and now more than one Linux box) — harmless no-op
# if it already exists with the right permissions.
mkdir -p data/catalog/thumbnails
chmod -R 777 data/catalog

export GIT_COMMIT="$(git rev-parse --short HEAD)"
docker compose up --build -d
