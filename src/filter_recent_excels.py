from __future__ import annotations

import argparse
import csv
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from add_template_sheets import (
    DEFAULT_TEMPLATE_WORKBOOK,
    WorkbookFormulaResolver,
    iter_excel_files,
    load_excel_workbook,
    resolve_anthro_backfill_rows,
    update_workbook_with_template_sheets,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_ROOT = Path("/Users/ernestolugo/Work/excels")
DEFAULT_OUTPUT_ROOT = PROJECT_ROOT / "output" / "excels-2025-en-adelante"
SUPPORTED_ORGANIZATIONS = ("CARACAS", "ELEMENT BOX", "HIDROCAVEN")
MIN_YEAR_DEFAULT = 2025
CLINICAL_DATE_LABELS = {"Fecha", "Fecha de evaluación"}
DATE_FORMATS = (
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d/%m/%y",
    "%d-%m-%Y",
    "%d-%m-%y",
)


@dataclass
class WorkbookSelectionResult:
    source_path: Path
    organization: str
    matched: bool
    selected_date: date | None
    date_source: str
    destination_path: Path | None = None
    backfill_applied: bool = False
    added_sheets: tuple[str, ...] = ()
    replaced_sheets: tuple[str, ...] = ()
    error: str = ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Filtra historias clinico-nutricionales del 2025 en adelante y las "
            "copia a una carpeta nueva preservando la estructura por organizacion."
        )
    )
    parser.add_argument(
        "input_root",
        nargs="?",
        default=str(DEFAULT_INPUT_ROOT),
        help=f"Carpeta raiz a escanear. Default: {DEFAULT_INPUT_ROOT}",
    )
    parser.add_argument(
        "output_root",
        nargs="?",
        default=str(DEFAULT_OUTPUT_ROOT),
        help=f"Carpeta de salida. Default: {DEFAULT_OUTPUT_ROOT}",
    )
    parser.add_argument(
        "--min-year",
        type=int,
        default=MIN_YEAR_DEFAULT,
        help=f"Ano minimo incluido. Default: {MIN_YEAR_DEFAULT}",
    )
    return parser.parse_args()


def parse_date_like(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text).date()
        except ValueError:
            return None
    return None


def safe_cell_value(
    value_wb,
    formula_resolver: WorkbookFormulaResolver,
    *,
    sheet_name: str,
    coordinate: str,
):
    if sheet_name not in value_wb.sheetnames:
        return None
    value = value_wb[sheet_name][coordinate].value
    if value is not None:
        return value
    if sheet_name not in formula_resolver.formula_wb.sheetnames:
        return None
    return formula_resolver.resolve_coordinate(sheet_name, coordinate)


def extract_history_date(
    value_wb,
    formula_resolver: WorkbookFormulaResolver,
) -> date | None:
    return parse_date_like(
        safe_cell_value(
            value_wb,
            formula_resolver,
            sheet_name="HISTORIA",
            coordinate="B2",
        )
    )


def extract_anthro_dates(
    value_wb,
    formula_wb,
) -> list[date]:
    rows = resolve_anthro_backfill_rows(
        value_wb,
        source_wb_formulas=formula_wb,
    )
    if rows is None:
        return []

    summary_rows, measurement_rows = rows
    dates: list[date] = []
    for label, values in [*summary_rows, *measurement_rows]:
        if label not in CLINICAL_DATE_LABELS:
            continue
        for value in values:
            parsed = parse_date_like(value)
            if parsed is not None:
                dates.append(parsed)
    return dates


def detect_reference_date(
    workbook_path: Path,
    value_wb,
    formula_wb,
) -> tuple[date | None, str]:
    resolver = WorkbookFormulaResolver(value_wb, formula_wb)

    clinical_candidates: list[tuple[date, str]] = []
    history_date = extract_history_date(value_wb, resolver)
    if history_date is not None:
        clinical_candidates.append((history_date, "HISTORIA!B2"))

    for anthro_date in extract_anthro_dates(value_wb, formula_wb):
        clinical_candidates.append((anthro_date, "antropometria"))

    if clinical_candidates:
        selected_date, source = max(clinical_candidates, key=lambda item: item[0])
        return selected_date, source

    modified_date = parse_date_like(value_wb.properties.modified)
    if modified_date is not None:
        return modified_date, "workbook.modified"

    created_date = parse_date_like(value_wb.properties.created)
    if created_date is not None:
        return created_date, "workbook.created"

    return datetime.fromtimestamp(workbook_path.stat().st_mtime).date(), "file.mtime"


def determine_organization(input_root: Path, workbook_path: Path) -> str:
    relative_parts = workbook_path.relative_to(input_root).parts
    if not relative_parts:
        return ""
    return relative_parts[0]


def collect_supported_workbooks(input_root: Path) -> list[Path]:
    return [
        workbook_path
        for workbook_path in iter_excel_files(input_root, recursive=True)
        if determine_organization(input_root, workbook_path) in SUPPORTED_ORGANIZATIONS
    ]


def print_progress_start(
    input_root: Path,
    workbook_path: Path,
    *,
    index: int,
    total: int,
) -> None:
    remaining = total - index
    relative_path = workbook_path.relative_to(input_root)
    print(
        f"[{index}/{total}] Voy por: {relative_path} | quedan {remaining}",
        flush=True,
    )


def print_progress_result(result: WorkbookSelectionResult) -> None:
    if result.error:
        print(
            f"  -> ERROR | {result.error}",
            flush=True,
        )
        return

    if result.matched:
        sheet_changes: list[str] = []
        if result.added_sheets:
            sheet_changes.append("agregadas=" + ", ".join(result.added_sheets))
        if result.replaced_sheets:
            sheet_changes.append("reemplazadas=" + ", ".join(result.replaced_sheets))
        changes_text = (
            " | " + " | ".join(sheet_changes)
            if sheet_changes
            else ""
        )
        print(
            f"  -> COPIADO+ACTUALIZADO | fecha {result.selected_date.isoformat()} "
            f"| fuente {result.date_source}{changes_text}",
            flush=True,
        )
        return

    rendered_date = result.selected_date.isoformat() if result.selected_date else "sin fecha"
    print(
        f"  -> OMITIDO | fecha {rendered_date} | fuente {result.date_source}",
        flush=True,
    )


def scan_workbook(
    input_root: Path,
    workbook_path: Path,
    *,
    min_year: int,
) -> WorkbookSelectionResult:
    organization = determine_organization(input_root, workbook_path)
    value_wb = load_excel_workbook(workbook_path, data_only=True)
    formula_wb = load_excel_workbook(workbook_path, data_only=False)
    try:
        selected_date, source = detect_reference_date(
            workbook_path,
            value_wb,
            formula_wb,
        )
        matched = selected_date is not None and selected_date.year >= min_year
        return WorkbookSelectionResult(
            source_path=workbook_path,
            organization=organization,
            matched=matched,
            selected_date=selected_date,
            date_source=source,
        )
    except Exception as exc:
        return WorkbookSelectionResult(
            source_path=workbook_path,
            organization=organization,
            matched=False,
            selected_date=None,
            date_source="error",
            error=str(exc),
        )
    finally:
        formula_wb.close()
        value_wb.close()


def copy_selected_workbooks(
    input_root: Path,
    output_root: Path,
    *,
    min_year: int,
    show_progress: bool = False,
    template_workbook_path: Path | str = DEFAULT_TEMPLATE_WORKBOOK,
) -> list[WorkbookSelectionResult]:
    results: list[WorkbookSelectionResult] = []
    workbook_paths = collect_supported_workbooks(input_root)
    total = len(workbook_paths)
    for index, workbook_path in enumerate(workbook_paths, start=1):
        if show_progress:
            print_progress_start(
                input_root,
                workbook_path,
                index=index,
                total=total,
            )
        result = scan_workbook(input_root, workbook_path, min_year=min_year)
        if result.matched:
            destination_path = output_root / workbook_path.relative_to(input_root)
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(workbook_path, destination_path)
            result.destination_path = destination_path
            try:
                update_result = update_workbook_with_template_sheets(
                    destination_path,
                    template_workbook_path=template_workbook_path,
                )
                result.backfill_applied = True
                result.added_sheets = update_result.added_sheets
                result.replaced_sheets = update_result.replaced_sheets
            except Exception as exc:
                result.error = f"backfill: {exc}"
        results.append(result)
        if show_progress:
            print_progress_result(result)
    return results


def write_report(
    output_root: Path,
    results: Iterable[WorkbookSelectionResult],
) -> None:
    rows = list(results)
    output_root.mkdir(parents=True, exist_ok=True)

    csv_path = output_root / "reporte.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "organization",
                "matched",
                "selected_date",
                "date_source",
                "source_path",
                "destination_path",
                "backfill_applied",
                "added_sheets",
                "replaced_sheets",
                "error",
            ]
        )
        for result in rows:
            writer.writerow(
                [
                    result.organization,
                    "yes" if result.matched else "no",
                    result.selected_date.isoformat() if result.selected_date else "",
                    result.date_source,
                    str(result.source_path),
                    str(result.destination_path) if result.destination_path else "",
                    "yes" if result.backfill_applied else "no",
                    ", ".join(result.added_sheets),
                    ", ".join(result.replaced_sheets),
                    result.error,
                ]
            )

    summary_path = output_root / "resumen.txt"
    totals_by_org = {
        organization: {
            "scanned": 0,
            "matched": 0,
            "errors": 0,
        }
        for organization in SUPPORTED_ORGANIZATIONS
    }
    for result in rows:
        if result.organization not in totals_by_org:
            continue
        totals_by_org[result.organization]["scanned"] += 1
        if result.matched:
            totals_by_org[result.organization]["matched"] += 1
        if result.error:
            totals_by_org[result.organization]["errors"] += 1

    with summary_path.open("w", encoding="utf-8") as handle:
        handle.write("Filtro de excels 2025 en adelante\n")
        handle.write(f"Total escaneados: {len(rows)}\n")
        handle.write(f"Total seleccionados: {sum(1 for result in rows if result.matched)}\n")
        handle.write(
            f"Total actualizados: {sum(1 for result in rows if result.backfill_applied and not result.error)}\n"
        )
        handle.write(
            f"Total con error: {sum(1 for result in rows if result.error)}\n"
        )
        handle.write("\n")
        for organization in SUPPORTED_ORGANIZATIONS:
            totals = totals_by_org[organization]
            handle.write(f"{organization}\n")
            handle.write(f"  escaneados: {totals['scanned']}\n")
            handle.write(f"  copiados: {totals['matched']}\n")
            handle.write(f"  errores: {totals['errors']}\n")
            handle.write("\n")


def main() -> int:
    args = parse_args()
    input_root = Path(args.input_root)
    output_root = Path(args.output_root)

    if not input_root.exists():
        print(f"Error: no existe la carpeta {input_root}")
        return 1
    if not input_root.is_dir():
        print(f"Error: {input_root} no es una carpeta")
        return 1

    results = copy_selected_workbooks(
        input_root,
        output_root,
        min_year=args.min_year,
        show_progress=True,
    )
    write_report(output_root, results)

    print(f"Escaneados: {len(results)}")
    print(f"Seleccionados: {sum(1 for result in results if result.matched)}")
    print(f"Actualizados: {sum(1 for result in results if result.backfill_applied and not result.error)}")
    print(f"Errores: {sum(1 for result in results if result.error)}")
    print(f"Salida: {output_root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
