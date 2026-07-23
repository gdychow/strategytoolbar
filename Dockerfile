# Multi-stage build: compile with the full Node toolchain, then ship a
# runtime image with only production dependencies (Express/pg/jose/
# cookie-parser — no esbuild/typescript). Two containers total with
# docker-compose (this app + Postgres) — no other dependencies beyond that.

FROM node:20-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json build.mjs manifest.xml manifest.prod.xml ./
COPY src ./src
COPY assets ./assets
# .git is excluded from the build context (see .dockerignore), so the git
# commit for the visible on-page build stamp is passed in from the host
# (which does have .git) rather than computed inside the image.
ARG GIT_COMMIT=unknown
ENV GIT_COMMIT=$GIT_COMMIT
RUN npx tsc --noEmit && node build.mjs --prod

FROM node:20-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV USE_TLS=false
ENV PORT=8080
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY server.js docker-entrypoint.sh ./
COPY server ./server
COPY scripts ./scripts
COPY db/seed ./db/seed
RUN chmod +x docker-entrypoint.sh
COPY --from=builder --chown=node:node /app/dist ./dist
# Pre-creates the catalog_files mount point owned by node:node so Docker
# propagates that ownership onto a fresh named volume on first mount —
# named volumes otherwise default to root ownership, which the node user
# (see USER below) can't write into (the entrypoint's thumbnails/ seed and
# server/catalog.js's own mkdirSync both write here at startup).
RUN mkdir -p /app/data/catalog && chown -R node:node /app/data/catalog
EXPOSE 8080
USER node
ENTRYPOINT ["./docker-entrypoint.sh"]
