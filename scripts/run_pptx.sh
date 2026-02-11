#!/usr/bin/env bash
set -euo pipefail

EXCEL_PATH="${1:-src-material/test.xlsx}"
TEMPLATE_PATH="${2:-src-material/Plan de Alimentación Base.pptx}"
OUTPUT_PATH="${3:-output/Plan Alimentacion - output.pptx}"

python src/generate_pptx.py "$EXCEL_PATH" \
  --template "$TEMPLATE_PATH" \
  --output "$OUTPUT_PATH"
