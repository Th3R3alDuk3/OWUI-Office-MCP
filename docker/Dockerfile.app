FROM ghcr.io/astral-sh/uv:python3.13-trixie-slim

# OfficeCLI screenshots need a `chromium`; Landlock denies /dev/shm. The fonts
# are metric-compatible with Office's, so line breaks match. tini reaps the
# helper processes Chromium leaves behind.
RUN apt-get update \
 && apt-get install --yes --no-install-recommends \
    tini chromium-headless-shell fonts-liberation fonts-crosextra-carlito \
    fonts-crosextra-caladea \
 && rm -rf /var/lib/apt/lists/* \
 && printf '#!/bin/sh\nexec /usr/lib/chromium/chromium-headless-shell --disable-dev-shm-usage "$@"\n' \
    > /usr/bin/chromium \
 && chmod 755 /usr/bin/chromium

# Pinned by version and checksum in bin/download.sh.
COPY --chmod=755 bin/officecli bin/landrun /usr/local/bin/

WORKDIR /app/

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy

COPY pyproject.toml uv.lock ./
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --no-install-project

# The binaries already sit in /usr/local/bin.
COPY --exclude=bin . .

# Plain log lines, no banner: the container's stdout is the log.
ENV PATH="/app/.venv/bin:$PATH" \
    FASTMCP_ENABLE_RICH_LOGGING=false \
    FASTMCP_SHOW_SERVER_BANNER=false

RUN useradd --system --uid 1000 app
USER app

ENTRYPOINT ["tini", "--"]
CMD ["python", "main.py"]
