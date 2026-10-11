# Builder stage
# Build natively on the runner; all deps are pure JS, so node_modules
# can be copied into the target-platform image without QEMU emulation.
FROM --platform=$BUILDPLATFORM node:20-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
RUN npm run build

# Production stage
FROM node:20-alpine AS production
# Docker-Secrets: NAME_FILE=/run/secrets/... wird beim Start zu NAME geladen
COPY --chmod=0755 env-secrets.sh /usr/local/bin/env-secrets

WORKDIR /app
COPY --from=builder /app/build ./build
COPY --from=builder /app/node_modules ./node_modules
COPY --from=builder /app/package.json ./package.json
COPY --from=builder /app/package-lock.json ./package-lock.json

EXPOSE 3000
ENTRYPOINT ["/usr/local/bin/env-secrets", "node", "build/index.js", "--http", "--port", "3000"]
