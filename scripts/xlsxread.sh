#!/usr/bin/env bash
# xlsxread.sh — Inspect an XLSX file using xlsxlite.
#
# Usage:
#   ./scripts/xlsxread.sh <file.xlsx>
#
# Prints sheet count, row/column counts, sample rows, and timing.

set -euo pipefail

if [ $# -lt 1 ]; then
  echo "Usage: $0 <file.xlsx>" >&2
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

exec go run "${PROJECT_DIR}/internal/xlsxread" "$1"
