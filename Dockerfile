FROM ghcr.io/astral-sh/uv:python3.14-trixie-slim AS build

WORKDIR /app/

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=0

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project


FROM python:3.14-slim-trixie

# OfficeCLI renders screenshots through a headless browser it finds as
# `chromium`. Landlock denies /dev/shm, so shared memory goes to TMPDIR. The
# fonts are metric-compatible with Arial, Times New Roman, Courier New,
# Calibri and Cambria, so line breaks match Office.
RUN apt-get update \
 && apt-get install --yes --no-install-recommends \
    chromium-headless-shell fonts-liberation fonts-crosextra-carlito \
    fonts-crosextra-caladea \
 && rm -rf /var/lib/apt/lists/* \
 && printf '#!/bin/sh\nexec /usr/lib/chromium/chromium-headless-shell --disable-dev-shm-usage "$@"\n' \
    > /usr/bin/chromium \
 && chmod 755 /usr/bin/chromium

ARG OFFICECLI_VERSION=1.0.149
ARG OFFICECLI_SHA256=ba0f397351ca3c31109ddc8e9690b848304da77594fd7573a76f2b1eb7e430e7

ADD --checksum=sha256:${OFFICECLI_SHA256} --chmod=755 \
    https://github.com/iOfficeAI/OfficeCLI/releases/download/v${OFFICECLI_VERSION}/officecli-linux-x64 \
    /usr/local/bin/officecli

WORKDIR /app/

COPY --from=build /app/.venv .venv
COPY . .

ENV PATH="/app/.venv/bin:$PATH"

# Projects and prepared templates live here; mount it to keep them.
RUN useradd --system --uid 1000 app \
 && mkdir data \
 && chown app data
USER app

CMD ["python", "main.py"]
