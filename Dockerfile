# ============================================================
# Global build arguments (declared before any FROM so they can
# be used in FROM lines; must be re-declared inside each stage
# to be visible there).
# ============================================================

# OnlyOffice DocumentServer version — controls both the source image
# tag AND the versioned asset directory prefix (/v<DS_VERSION>-<HASH>).
ARG DS_VERSION=9.3.1

# Revision counter. Bump this (--build-arg HASH=2) whenever you want
# to bust the browser cache for the OnlyOffice assets without changing
# the DocumentServer version itself.
ARG HASH=2

# ============================================================
# Stage 1: OnlyOffice DocumentServer assets source
# ============================================================
FROM onlyoffice/documentserver:${DS_VERSION} AS documentserver

# AllFonts.js and themes.js are NOT present in the image — they are
# generated at container startup by documentserver-generate-allfonts.sh.
# We run that script here (passing `false` so it skips the data-container
# wait branch) so the files exist before the COPY in the final stage.
RUN documentserver-generate-allfonts.sh false

# ============================================================
# Stage 2: Next.js website builder
# ============================================================
FROM node:24-alpine AS builder

# Re-declare args inside this stage to make them visible here.
ARG DS_VERSION
ARG HASH

# Expose the versioned asset path to Next.js at build time.
ENV NEXT_PUBLIC_APP_ROOT=/v${DS_VERSION}-${HASH}

WORKDIR /app

# Install pnpm
RUN corepack enable && corepack prepare pnpm@latest --activate

# Copy dependency manifests first for better layer caching.
# pnpm-workspace.yaml carries allowBuilds (pnpm 10+); must be present before install.
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./

# Install dependencies (frozen lockfile for reproducibility).
RUN pnpm install --frozen-lockfile

# Copy the website source code (explicit paths isolate Next.js from plugins tree).
COPY app ./app
COPY assets ./assets
COPY components ./components
COPY hooks ./hooks
COPY i18n ./i18n
COPY lib ./lib
COPY messages ./messages
COPY public ./public
COPY scripts ./scripts
COPY store ./store
COPY types ./types
COPY utils ./utils
COPY components.json eslint.config.mjs next.config.ts postcss.config.mjs svgo.config.js tsconfig.json ./

# Bake the live OnlyOffice asset prefix so the Agent plugin can find plugins.js.
RUN printf '{"appRoot":"/v%s-%s"}\n' "$DS_VERSION" "$HASH" > public/app-root.json

# Run the Next.js static export build.
RUN pnpm build

# ============================================================
# Stage 3: Caddy production server
# ============================================================
FROM caddy:2-alpine AS final

# Re-declare args inside this stage.
ARG DS_VERSION
ARG HASH

WORKDIR /srv

# Copy the Next.js static export output.
COPY --from=builder /app/out ./

# Copy OnlyOffice DocumentServer assets directly from the source stage
# into the versioned directory — assets never pass through the builder,
# so there is no redundant copy of the large asset tree.
COPY --from=documentserver /var/www/onlyoffice/documentserver/fonts         ./v${DS_VERSION}-${HASH}/fonts
COPY --from=documentserver /var/www/onlyoffice/documentserver/sdkjs         ./v${DS_VERSION}-${HASH}/sdkjs
COPY --from=documentserver /var/www/onlyoffice/documentserver/web-apps      ./v${DS_VERSION}-${HASH}/web-apps
COPY --from=documentserver /var/www/onlyoffice/documentserver/sdkjs-plugins ./v${DS_VERSION}-${HASH}/sdkjs-plugins

# api.js is generated from a template at runtime in a full DocumentServer
# deployment, but here we serve it statically — copy the template as-is.
RUN cp "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js.tpl" \
       "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js"

# public/v9.3* is dockerignored (huge OO tree). Bake Agent autostart list here so
# DocsAPI's sync plugins.json XHR works even when the embed XHR proxy loses the race.
# Use an absolute URL — path-only entries are resolved under /v9.3.1-2/ and 404.
RUN printf '%s\n' \
  '{' \
  '  "url": "",' \
  '  "pluginsData": ["https://editor.app.br/office-plugins/agent/config.json"],' \
  '  "autostart": ["asc.{7E4A1C90-2B6D-4F11-9A33-8C0E5D71B2A4}"]' \
  '}' \
  > "./v${DS_VERSION}-${HASH}/plugins.json"

# Copy editor plugins tree to /srv/plugins for plugins.editor.app.br (:8081)
COPY plugins /srv/plugins

# Copy Caddyfile.
COPY Caddyfile /etc/caddy/Caddyfile

EXPOSE 80 8081 443

