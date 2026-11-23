# syntax=docker/dockerfile:1

# --- build: compile TypeScript to dist/ ---
FROM node:26-alpine AS build
WORKDIR /app
COPY package.json yarn.lock ./
RUN yarn install --frozen-lockfile
COPY tsconfig.json ./
COPY src ./src
RUN yarn build

# --- deps: production-only node_modules ---
FROM node:26-alpine AS deps
WORKDIR /app
COPY package.json yarn.lock ./
RUN yarn install --frozen-lockfile --production && yarn cache clean

# --- runtime ---
FROM node:26-alpine AS runtime
ENV NODE_ENV=production
ENV ONENOTE_MCP_HTTP_HOST=0.0.0.0
ENV ONENOTE_MCP_HTTP_PORT=3000
ENV XDG_CONFIG_HOME=/data
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
COPY package.json ./
RUN mkdir -p /data && chown -R node:node /data
USER node
EXPOSE 3000
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:3000/healthz').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"
CMD ["node", "dist/cli.js", "--transport", "http"]
