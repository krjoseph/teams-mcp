FROM node:22.12-alpine AS builder

# Copy necessary files for the build
COPY package*.json tsconfig.json ./
COPY ./src ./src

WORKDIR /

RUN --mount=type=cache,target=/root/.npm npm install

RUN --mount=type=cache,target=/root/.npm-production npm ci --ignore-scripts --omit-dev

RUN --mount=type=cache,target=/root/.npm npm run build

FROM node:22-alpine AS release

COPY --from=builder ./dist /app/dist
COPY --from=builder ./package.json /app/package.json
COPY --from=builder ./package-lock.json /app/package-lock.json

ENV NODE_ENV=production
LABEL org.opencontainers.image.title="Microsoft Teams MCP Server"
LABEL org.opencontainers.image.description="MCP server for interacting with Microsoft Teams"

WORKDIR /app

RUN npm ci --ignore-scripts --omit-dev

ENTRYPOINT ["node", "dist/index.js", "--transport", "http", "--port", "$PORT"]
