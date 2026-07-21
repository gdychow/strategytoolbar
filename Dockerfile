# Multi-stage build: compile with the full Node toolchain, then ship only
# the static output (dist/) in a minimal runtime image. No other containers
# required — TLS is expected to terminate upstream (e.g. Cloudflare Tunnel).

FROM node:20-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json build.mjs manifest.xml manifest.prod.xml ./
COPY src ./src
COPY assets ./assets
RUN npx tsc --noEmit && node build.mjs --prod

FROM node:20-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV USE_TLS=false
ENV PORT=8080
COPY server.js docker-entrypoint.sh ./
RUN chmod +x docker-entrypoint.sh
COPY --from=builder --chown=node:node /app/dist ./dist
EXPOSE 8080
USER node
ENTRYPOINT ["./docker-entrypoint.sh"]
