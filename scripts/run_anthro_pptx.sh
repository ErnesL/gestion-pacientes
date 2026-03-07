#!/usr/bin/env bash
set -euo pipefail

EXCEL_PATH="${1:-src-material/test.xlsx}"
TEMPLATE_PATH="${2:-src-material/Informe Antropométrico base.pptx}"
OUTPUT_PATH="${3:-output/Informe Antropometrico - output.pptx}"
TODAY="${4:-}"

TODAY_ARGS=()
if [[ -n "$TODAY" ]]; then
  TODAY_ARGS=(--today "$TODAY")
fi

python src/generate_anthro_pptx.py "$EXCEL_PATH" \
  --template "$TEMPLATE_PATH" \
  --output "$OUTPUT_PATH" \
  "${TODAY_ARGS[@]}"
