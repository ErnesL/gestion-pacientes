from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from add_template_sheets import iter_excel_files
from app_support import get_template_paths, inspect_excel_file, sanitize_filename_component
from excel_helpers import build_validation_warning_message
from generate_anthro_pptx import generate_anthro_pptx
from generate_pptx import generate_plan_pptx


@dataclass
class BatchPptxResult:
    excel_path: Path
    patient_name: str
    plan_pptx: Path | None = None
    anthro_pptx: Path | None = None
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def parse_iso_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "Formato invalido para --today. Usa YYYY-MM-DD."
        ) from exc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Genera recursivamente los PPTX de plan e informe antropometrico "
            "al lado de cada Excel dentro de una carpeta."
        )
    )
    parser.add_argument("folder", help="Carpeta con archivos .xlsx o .xlsm")
    parser.add_argument(
        "--today",
        type=parse_iso_date,
        default=None,
        help="Fecha base para el informe antropometrico (YYYY-MM-DD).",
    )
    return parser.parse_args()


def collect_excel_files(folder: Path) -> list[Path]:
    return list(iter_excel_files(folder, recursive=True))


def build_output_paths(excel_path: Path, patient_name: str) -> tuple[Path, Path]:
    safe_name = sanitize_filename_component(patient_name or excel_path.stem)
    output_dir = excel_path.parent
    plan_output = output_dir / f"Plan Alimentacion - {safe_name}.pptx"
    anthro_output = output_dir / f"Informe Antropometrico - {safe_name}.pptx"
    return plan_output, anthro_output


def print_progress_start(
    root_folder: Path,
    excel_path: Path,
    *,
    index: int,
    total: int,
) -> None:
    remaining = total - index
    print(
        f"[{index}/{total}] Voy por: {excel_path.relative_to(root_folder)} | quedan {remaining}",
        flush=True,
    )


def print_progress_result(result: BatchPptxResult) -> None:
    if result.errors:
        status = "ERROR PARCIAL" if (result.plan_pptx or result.anthro_pptx) else "ERROR"
        print(f"  -> {status}", flush=True)
        for error in result.errors:
            print(f"     {error}", flush=True)
        return

    warnings_text = f" | advertencias={len(result.warnings)}" if result.warnings else ""
    print(
        "  -> OK"
        f"{warnings_text}"
        f" | plan={result.plan_pptx.name if result.plan_pptx else 'no'}"
        f" | informe={result.anthro_pptx.name if result.anthro_pptx else 'no'}",
        flush=True,
    )


def generate_pptx_for_excel(
    excel_path: Path,
    *,
    today: date | None = None,
) -> BatchPptxResult:
    parsed_data = inspect_excel_file(excel_path)
    patient_name = parsed_data.patient.name or excel_path.stem
    result = BatchPptxResult(
        excel_path=excel_path,
        patient_name=patient_name,
    )
    if parsed_data.issues:
        result.warnings.append(build_validation_warning_message(parsed_data.issues))

    template_paths = get_template_paths()
    plan_template = template_paths["plan"]
    anthro_template = template_paths["anthro"]
    plan_output, anthro_output = build_output_paths(excel_path, patient_name)

    try:
        generate_plan_pptx(
            excel_path=excel_path,
            template_path=plan_template,
            output_path=plan_output,
            parsed_data=parsed_data,
        )
        result.plan_pptx = plan_output
    except Exception as exc:
        result.errors.append(f"Plan: {exc}")

    try:
        generate_anthro_pptx(
            excel_path=excel_path,
            template_path=anthro_template,
            output_path=anthro_output,
            today=today,
            parsed_data=parsed_data,
        )
        result.anthro_pptx = anthro_output
    except Exception as exc:
        result.errors.append(f"Informe: {exc}")

    return result


def write_reports(root_folder: Path, results: list[BatchPptxResult]) -> None:
    csv_path = root_folder / "reporte-generacion-pptx.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "excel_path",
                "patient_name",
                "plan_pptx",
                "anthro_pptx",
                "warnings",
                "errors",
            ]
        )
        for result in results:
            writer.writerow(
                [
                    str(result.excel_path),
                    result.patient_name,
                    str(result.plan_pptx) if result.plan_pptx else "",
                    str(result.anthro_pptx) if result.anthro_pptx else "",
                    " | ".join(result.warnings),
                    " | ".join(result.errors),
                ]
            )

    txt_path = root_folder / "resumen-generacion-pptx.txt"
    with txt_path.open("w", encoding="utf-8") as handle:
        handle.write("Generacion batch de PPTX\n")
        handle.write(f"Total excels: {len(results)}\n")
        handle.write(
            f"Planes generados: {sum(1 for result in results if result.plan_pptx is not None)}\n"
        )
        handle.write(
            f"Informes generados: {sum(1 for result in results if result.anthro_pptx is not None)}\n"
        )
        handle.write(
            f"Excels con errores: {sum(1 for result in results if result.errors)}\n"
        )


def generate_pptx_batch(
    folder: Path | str,
    *,
    today: date | None = None,
    show_progress: bool = False,
) -> list[BatchPptxResult]:
    root_folder = Path(folder)
    if not root_folder.exists():
        raise FileNotFoundError(f"No existe la carpeta: {root_folder}")
    if not root_folder.is_dir():
        raise NotADirectoryError(f"La ruta no es una carpeta: {root_folder}")

    workbook_paths = collect_excel_files(root_folder)
    results: list[BatchPptxResult] = []
    total = len(workbook_paths)
    for index, excel_path in enumerate(workbook_paths, start=1):
        if show_progress:
            print_progress_start(root_folder, excel_path, index=index, total=total)
        result = generate_pptx_for_excel(excel_path, today=today)
        results.append(result)
        if show_progress:
            print_progress_result(result)

    write_reports(root_folder, results)
    return results


def main() -> int:
    args = parse_args()
    try:
        results = generate_pptx_batch(
            args.folder,
            today=args.today,
            show_progress=True,
        )
    except (FileNotFoundError, NotADirectoryError, ValueError) as exc:
        print(f"Error: {exc}")
        return 1

    print(f"Procesados: {len(results)}")
    print(f"Planes generados: {sum(1 for result in results if result.plan_pptx is not None)}")
    print(f"Informes generados: {sum(1 for result in results if result.anthro_pptx is not None)}")
    print(f"Con errores: {sum(1 for result in results if result.errors)}")
    print(f"Reporte CSV: {Path(args.folder) / 'reporte-generacion-pptx.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
