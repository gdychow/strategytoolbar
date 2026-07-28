#!/usr/bin/env sh
# Canonical way to rebuild and restart the app container. A bare
# `docker compose up --build` leaves GIT_COMMIT unset, which
# docker-compose.yml's build arg silently defaults to "unknown" for — the
# task pane's build-info line (top-right, visible even if Office.js never
# initializes) then shows "unknown" instead of a real commit hash. This
# script exists so that's never forgotten again.
set -e
cd "$(dirname "$0")"
export GIT_COMMIT="$(git rev-parse --short HEAD)"
docker compose up --build -d
