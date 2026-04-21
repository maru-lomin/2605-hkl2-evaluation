#!/usr/bin/env bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

OUT_DIR="${OUT_DIR:-.}"
uv run result_to_excel.py ./dataset --out-dir result