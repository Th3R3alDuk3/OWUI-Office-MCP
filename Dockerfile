FROM ghcr.io/astral-sh/uv:python3.14-trixie-slim

# OfficeCLI screenshots need a `chromium`; Landlock denies /dev/shm. The fonts
# are metric-compatible with Office's, so line breaks match.
RUN apt-get update \
 && apt-get install --yes --no-install-recommends \
    chromium-headless-shell fonts-liberation fonts-crosextra-carlito \
    fonts-crosextra-caladea \
 && rm -rf /var/lib/apt/lists/* \
 && printf '#!/bin/sh\nexec /usr/lib/chromium/chromium-headless-shell --disable-dev-shm-usage "$@"\n' \
    > /usr/bin/chromium \
 && chmod 755 /usr/bin/chromium

# Pinned by version and checksum in bin/download.sh.
COPY --chmod=755 bin/officecli bin/landrun /usr/local/bin/

WORKDIR /app/

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=0

COPY pyproject.toml uv.lock ./
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --no-install-project

# The binaries already sit in /usr/local/bin.
COPY --exclude=bin . .

ENV PATH="/app/.venv/bin:$PATH"

RUN useradd --system --uid 1000 app
USER app

CMD ["python", "main.py"]
