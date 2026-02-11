from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

from openpyxl import load_workbook
from pptx import Presentation

from excel_helpers import (
    MEAL_DEFS,
    ValidationError,
    build_meal_replacements,
    build_replacements,
    build_totals_replacements,
    load_patient_info,
)
from pptx_helpers import (
    align_marked_shapes,
    apply_vertical_stack,
    replace_in_shape,
    slide_contains_tokens,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genera PPTX desde Excel")
    parser.add_argument("excel", help="Ruta al archivo Excel")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "src-material" /
                    "Plan de Alimentación Base.pptx"),
        help="Ruta al PPTX base con placeholders",
    )
    parser.add_argument(
        "--output",
        default=str(PROJECT_ROOT / "output" / "Plan Alimentacion.pptx"),
        help="Ruta de salida PPTX",
    )
    return parser.parse_args()


def remove_slides_by_index(presentation: Presentation, indices: List[int]) -> None:
    if not indices:
        return
    slide_id_list = presentation.slides._sldIdLst
    for index in sorted(indices, reverse=True):
        slide_id_list.remove(slide_id_list[index])


def main() -> int:
    args = parse_args()
    excel_path = Path(args.excel)
    template_path = Path(args.template)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        print(f"No existe el archivo: {excel_path}")
        return 1
    if not template_path.exists():
        print(f"No existe el PPTX base: {template_path}")
        return 1

    wb = load_workbook(excel_path, data_only=True)

    try:
        patient = load_patient_info(wb)
    except ValidationError as exc:
        print(f"Error: {exc}")
        return 1

    replacements = build_replacements(patient)
    ws_req = wb["REQUERIMIENTOS"]
    replacements.update(build_totals_replacements(ws_req))

    presentation = Presentation(str(template_path))
    slides_to_remove: List[int] = []
    meal_tokens_to_remove: List[List[str]] = []
    placeholder_values: Dict[str, int] = {}

    for meal_def in MEAL_DEFS:
        meal_repl, _, include_meal, tokens, meal_placeholder_values = build_meal_replacements(
            ws_req, meal_def
        )
        replacements.update(meal_repl)
        placeholder_values.update(meal_placeholder_values)
        if not include_meal:
            meal_tokens_to_remove.append(tokens)

    for idx, slide in enumerate(presentation.slides):
        for tokens in meal_tokens_to_remove:
            if slide_contains_tokens(slide, tokens):
                slides_to_remove.append(idx)
                break
        if idx in slides_to_remove:
            continue
        table_maps = []
        for shape in slide.shapes:
            replace_in_shape(
                shape,
                replacements,
                placeholder_values,
                presentation.slide_width,
                slide.shapes,
                table_maps,
            )
        align_marked_shapes(slide, placeholder_values, table_maps)
        apply_vertical_stack(slide)

    remove_slides_by_index(presentation, slides_to_remove)

    presentation.save(str(output_path))
    print(f"PPTX generado: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
