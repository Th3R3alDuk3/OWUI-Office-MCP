#!/usr/bin/env bash
# Fetches the pinned OfficeCLI and landrun binaries into this directory, so
# the image builds without reaching GitHub. Run it once before building.
set -euo pipefail

cd "$(dirname "$0")"

OFFICECLI_VERSION=1.0.155
OFFICECLI_SHA256=1a5da5cb7e9d995db7a6560beb7466760c79617648f72e65f4ce55897c407f67
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
