FROM node:22.12-alpine AS builder

WORKDIR /app

# Copy necessary files for the build
COPY package*.json tsconfig.json ./
COPY ./src ./src

RUN --mount=type=cache,target=/root/.npm npm ci

RUN --mount=type=cache,target=/root/.npm npm run build

FROM node:22-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist ./dist
COPY --from=builder /app/package.json ./package.json
COPY --from=builder /app/package-lock.json ./package-lock.json

ENV NODE_ENV=production
LABEL org.opencontainers.image.title="Microsoft Teams MCP Server"
LABEL org.opencontainers.image.description="MCP server for interacting with Microsoft Teams"

RUN npm ci --ignore-scripts --omit-dev

ENTRYPOINT ["node", "dist/index.js", "--transport", "http"]
