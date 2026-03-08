#!/usr/bin/env bash
set -euo pipefail

EXCEL_PATH="${1:-src-material/test.xlsx}"
TEMPLATE_PATH="${2:-templates/informe-antropometrico-base.pptx}"
OUTPUT_PATH="${3:-output/Informe Antropometrico - output.pptx}"
TODAY="${4:-}"

PYTHON_BIN="${PYTHON_BIN:-}"
if [[ -z "$PYTHON_BIN" ]]; then
  if [[ -x ".venv/bin/python" ]]; then
    PYTHON_BIN=".venv/bin/python"
  else
    PYTHON_BIN="python3"
  fi
fi

if [[ -n "$TODAY" ]]; then
  "$PYTHON_BIN" src/generate_anthro_pptx.py "$EXCEL_PATH" \
    --template "$TEMPLATE_PATH" \
    --output "$OUTPUT_PATH" \
    --today "$TODAY"
else
  "$PYTHON_BIN" src/generate_anthro_pptx.py "$EXCEL_PATH" \
    --template "$TEMPLATE_PATH" \
    --output "$OUTPUT_PATH"
fi
