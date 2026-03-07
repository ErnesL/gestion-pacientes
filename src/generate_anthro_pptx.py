from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation

from excel_helpers import (
    ValidationError,
    build_anthropometric_replacements,
    build_measurements_table_replacements,
    build_summary_table_replacements,
    load_anthropometric_data,
)
from pptx_helpers import replace_in_shape


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def parse_iso_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "Formato invalido para --today. Usa YYYY-MM-DD."
        ) from exc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera Informe Antropometrico PPTX desde Excel"
    )
    parser.add_argument("excel", help="Ruta al archivo Excel")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "src-material" / "Informe Antropométrico base.pptx"),
        help="Ruta al PPTX antropometrico base",
    )
    parser.add_argument(
        "--output",
        default=str(PROJECT_ROOT / "output" / "Informe Antropometrico.pptx"),
        help="Ruta de salida PPTX",
    )
    parser.add_argument(
        "--today",
        type=parse_iso_date,
        default=None,
        help="Fecha base para calculos (YYYY-MM-DD). Si no se indica, usa hoy.",
    )
    return parser.parse_args()


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
        anthro_data = load_anthropometric_data(wb)
    except ValidationError as exc:
        print(f"Error: {exc}")
        return 1

    today: date = args.today if args.today is not None else date.today()
    base_replacements = build_anthropometric_replacements(anthro_data, today)
    summary_table_replacements = build_summary_table_replacements(anthro_data)
    measurement_table_replacements = build_measurements_table_replacements(anthro_data)

    presentation = Presentation(str(template_path))
    if len(presentation.slides) < 4:
        print("Error: El template antropometrico debe tener al menos 4 diapositivas.")
        return 1

    for slide_idx, slide in enumerate(presentation.slides):
        slide_replacements = dict(base_replacements)
        if slide_idx == 2:
            slide_replacements.update(summary_table_replacements)
        elif slide_idx == 3:
            slide_replacements.update(measurement_table_replacements)

        table_maps = []
        for shape in list(slide.shapes):
            replace_in_shape(
                shape,
                slide_replacements,
                {},
                presentation.slide_width,
                slide.shapes,
                table_maps,
            )

    presentation.save(str(output_path))
    print(f"PPTX antropometrico generado: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
