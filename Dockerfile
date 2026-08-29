# syntax=docker/dockerfile:1

# ============================================================================
# Build stage: install all deps and run the webpack production build
# ============================================================================
FROM node:22-alpine AS builder

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci

COPY . .
RUN npm run build

# ============================================================================
# Production dependencies stage: runtime deps only (no webpack/jest/eslint)
# ============================================================================
FROM node:22-alpine AS deps

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# ============================================================================
# Runtime stage: minimal image, non-root user
# ============================================================================
FROM node:22-alpine

WORKDIR /app
ENV NODE_ENV=production

COPY --from=deps /app/node_modules ./node_modules
COPY --from=builder /app/dist ./dist

# Pull current alpine security patches (e.g. libcrypto3/openssl) into the
# runtime image; the node:22-alpine digest lags the alpine repos.
RUN apk upgrade --no-cache
COPY --from=builder /app/scripts ./scripts
COPY --from=builder /app/manifest.template.xml ./manifest.template.xml
COPY --from=builder /app/package.json ./package.json

# manifest.xml and .manifest-guid are generated at startup, so /app must be
# writable; the node user owns it after the chown below.
RUN chown -R node:node /app
USER node

EXPOSE 3000

# Docker's HEALTHCHECK runs without curl in alpine images, so probe with node.
# rejectUnauthorized:false keeps the check working with self-signed certs.
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "const p=process.env.PROTOCOL||'https',port=process.env.PORT||'3000',m=p==='https'?require('https'):require('http');m.get(p+'://127.0.0.1:'+port+'/healthz',{rejectUnauthorized:false},r=>process.exit(r.statusCode===200?0:1)).on('error',()=>process.exit(1))"

CMD ["node", "scripts/docker-server.cjs"]
