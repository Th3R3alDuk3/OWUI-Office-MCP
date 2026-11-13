#!/usr/bin/env bash
# Fetches the pinned binaries here, so the image builds without GitHub.
set -euo pipefail

cd "$(dirname "$0")"

OFFICECLI_VERSION=1.0.161
OFFICECLI_SHA256=ed551216f7a62c562d8d231c0416374566684fcc6163176f2ba4ad721c62900a
LANDRUN_VERSION=0.1.17
LANDRUN_SHA256=6ada66a06669e8994e174a7271af2db636308e55a0d6ec896cc7d326b46727f6

download() {  # name url sha256

    if [ -f "$1" ] && sha256sum --check --status <<< "$3  $1"; then
        echo "$1 is up to date"
        return
    fi

    curl --fail --location --progress-bar --output "$1" "$2"
    sha256sum --check <<< "$3  $1"
    chmod 755 "$1"
}

download officecli \
    "https://github.com/iOfficeAI/OfficeCLI/releases/download/v${OFFICECLI_VERSION}/officecli-linux-x64" \
    "$OFFICECLI_SHA256"
download landrun \
    "https://github.com/Zouuup/landrun/releases/download/v${LANDRUN_VERSION}/landrun-linux-amd64" \
    "$LANDRUN_SHA256"
