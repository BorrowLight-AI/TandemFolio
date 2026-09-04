#!/usr/bin/env bash
set -euo pipefail

bundle_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
if ! command -v node >/dev/null 2>&1; then
  echo 'Install Node.js 22.12+ from https://nodejs.org/, then reopen Terminal.' >&2
  exit 1
fi
node "$bundle_root/verify.mjs"
if [[ "${1:-}" == '--check' ]]; then exit 0; fi
if [[ $# -gt 1 || ( $# -eq 1 && "$1" != '--update' ) ]]; then
  echo 'Usage: bash install.sh [--check|--update]' >&2; exit 1
fi
if ! command -v codex >/dev/null 2>&1; then
  echo 'Codex CLI is required. Install it, then run this script again.' >&2
  exit 1
fi
codex --version
if [[ "${1:-}" != '--update' ]]; then codex plugin marketplace add "$bundle_root"; fi
codex plugin add tandemfolio@tandemfolio-releases
echo 'TandemFolio installed. Keep this folder in place and start a new Codex task.'
