# ============================================================
# Global build arguments (declared before any FROM so they can
# be used in FROM lines; must be re-declared inside each stage
# to be visible there).
# ============================================================

# OnlyOffice DocumentServer version — controls both the source image
# tag AND the versioned asset directory prefix (/v<DS_VERSION>-<HASH>).
ARG DS_VERSION=9.4.0.1

# Revision counter. Bump this (--build-arg HASH=2) whenever you want
# to bust the browser cache for the OnlyOffice assets without changing
# the DocumentServer version itself.
ARG HASH=1

# Pin pnpm so local, Cloudflare and GitHub builds interpret the lockfile and
# dependency build-script policy identically.
ARG PNPM_VERSION=11.16.0

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
FROM node:22-alpine AS builder

# Re-declare args inside this stage to make them visible here.
ARG DS_VERSION
ARG HASH
ARG PNPM_VERSION

# Expose the versioned asset path to Next.js at build time.
ENV NEXT_PUBLIC_APP_ROOT=/v${DS_VERSION}-${HASH}

WORKDIR /app

# Install pnpm
RUN corepack enable && corepack prepare pnpm@${PNPM_VERSION} --activate

# Copy dependency manifests and the build-script allowlist first for better
# layer caching and deterministic installs in clean Docker builders.
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./

# Install dependencies (frozen lockfile for reproducibility).
RUN pnpm install --frozen-lockfile

# Copy the rest of the source code.
COPY . .

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
# api.js is generated from a template at runtime in a full DocumentServer
# deployment, but here we serve it statically — copy the template as-is.
RUN cp "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js.tpl" \
       "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js"

# Community DocumentServer ships dormant Google Analytics loaders inside some
# editor bundles. The site CSP already blocks them; rewrite the hostnames to
# the reserved .invalid TLD as an additional fail-closed privacy safeguard.
RUN analytics_files="$(grep -RIlE 'googletagmanager\.com|google-analytics\.com' \
      "./v${DS_VERSION}-${HASH}" || true)" && \
    if [ -n "$analytics_files" ]; then \
      printf '%s\n' "$analytics_files" | \
        xargs sed -i -E 's/(googletagmanager|google-analytics)\.com/analytics.invalid/g'; \
    fi

# Fail closed if the generated site or editor resources still reference the
# previous operator or other explicitly forbidden third-party services.
RUN matches="$(grep -RIlE 'office-editor\.ziziyi\.com|office-plugins\.ziziyi\.com|googletagmanager\.com|google-analytics\.com|api\.producthunt\.com|chromewebstore\.google\.com' /srv || true)" && \
    if [ -n "$matches" ]; then echo "$matches"; exit 1; fi

# Copy Caddyfile.
COPY Caddyfile /etc/caddy/Caddyfile

EXPOSE 80 443
