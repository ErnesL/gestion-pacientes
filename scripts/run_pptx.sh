#!/usr/bin/env bash
set -euo pipefail

EXCEL_PATH="${1:-src-material/test.xlsx}"
TEMPLATE_PATH="${2:-templates/plan-de-alimentacion-base.pptx}"
OUTPUT_PATH="${3:-output/Plan Alimentacion - output.pptx}"

PYTHON_BIN="${PYTHON_BIN:-}"
if [[ -z "$PYTHON_BIN" ]]; then
  if [[ -x ".venv/bin/python" ]]; then
    PYTHON_BIN=".venv/bin/python"
  else
    PYTHON_BIN="python3"
  fi
fi

"$PYTHON_BIN" src/generate_pptx.py "$EXCEL_PATH" \
  --template "$TEMPLATE_PATH" \
  --output "$OUTPUT_PATH"
