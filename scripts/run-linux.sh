#!/usr/bin/env bash
# Runs the whole matrix on Linux x64 inside Docker, without touching the host's node_modules.
# Usage: bash scripts/run-linux.sh [extra run-all.mjs args]
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
IMAGE="node:24-bookworm"
MSYS_NO_PATHCONV=1; export MSYS_NO_PATHCONV
docker run --rm -v "$ROOT:/src" "$IMAGE" bash -lc '
  set -e
  mkdir -p /w && cd /src && tar --exclude=node_modules --exclude=out -cf - . | (cd /w && tar -xf -)
  cd /w
  npm ci --no-audit --no-fund >/dev/null
  for v in sdk/*/; do npm ci --prefix "$v" --no-audit --no-fund >/dev/null; done
  node run-all.mjs --platform linux-x64 --quiet '"$*"'
  cp results/linux-x64.json /src/results/linux-x64.json
'
# Regenerate RESULTS.md on the host so it merges every platform file.
cd "$ROOT" && node run-all.mjs --merge-only --quiet >/dev/null
echo "results/linux-x64.json written"
