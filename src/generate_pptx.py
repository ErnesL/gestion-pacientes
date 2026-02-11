from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass
class PatientInfo:
    name: str
    ci: str
    sex: str
    age: str
    discipline: str


class ValidationError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genera PPTX desde Excel")
    parser.add_argument("excel", help="Ruta al archivo Excel")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "Source material" /
                    "Plan de Alimentación Base.pptx"),
        help="Ruta al PPTX base con placeholders",
    )
    parser.add_argument(
        "--output",
        default=str(PROJECT_ROOT / "output" / "Plan Alimentacion.pptx"),
        help="Ruta de salida PPTX",
    )
    return parser.parse_args()


def load_patient_info(wb) -> PatientInfo:
    ws = wb["HISTORIA"]
    name = str(ws["C4"].value or "").strip()
    ci = str(ws["C5"].value or "").strip()
    age_val = ws["C7"].value
    sex = str(ws["C10"].value or "").strip()
    discipline = str(ws["I8"].value or "").strip()

    age = ""
    if isinstance(age_val, (int, float)):
        age = str(int(round(age_val)))
    elif age_val:
        age = str(age_val)

    missing = []
    if not name:
        missing.append("Nombre y Apellido (HISTORIA!C4)")
    if not ci:
        missing.append("Cédula (HISTORIA!C5)")
    if not sex:
        missing.append("Sexo (HISTORIA!C10)")
    if not age:
        missing.append("Edad (HISTORIA!C7)")

    if missing:
        raise ValidationError("Faltan campos: " + "; ".join(missing))

    return PatientInfo(
        name=name,
        ci=ci,
        sex=sex,
        age=age,
        discipline=discipline,
    )


def build_replacements(patient: PatientInfo) -> Dict[str, str]:
    placeholder = "____________________"
    display_name = format_short_name(patient.name)
    return {
        "{{PACIENTE}}": display_name,
        "{{DISCIPLINA}}": patient.discipline or placeholder,
        "{{OBJETIVO}}": placeholder,
        "{{SEXO}}": patient.sex,
        "{{EDAD}}": patient.age,
    }


GROUP_ROWS = {
    "L": 48,  # Lacteos (Leche)
    "V": 49,  # Vegetales
    "F": 50,  # Frutas
    "A": 51,  # Almidones
    "P": 53,  # Proteinas (Carnes semi)
    "G": 54,  # Grasas
}

GROUP_SUFFIX = {
    "L": "LACTEOS",
    "V": "VEGETALES",
    "F": "FRUTAS",
    "A": "ALMIDONES",
    "P": "PROTEINAS",
    "G": "GRASAS",
}

MEAL_DEFS = [
    {"name": "PRE", "col": "K", "groups": ["L", "V", "F", "A", "P", "G"]},
    {"name": "DES", "col": "L", "groups": ["L", "F", "A", "P", "G"]},
    {"name": "MAM", "col": "M", "groups": ["L", "F", "A", "P", "G"]},
    {"name": "ALM", "col": "N", "groups": ["V", "F", "A", "P", "G"]},
    {"name": "MTP", "col": "P", "groups": ["L", "F", "A", "P", "G"]},
    {"name": "CEN", "col": "R", "groups": ["V", "F", "A", "P", "G"]},
]


def to_int(value) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(round(value))
    if isinstance(value, str):
        try:
            return int(round(float(value.replace(",", "."))))
        except ValueError:
            return 0
    return 0


def build_meal_replacements(
    ws, meal_def
) -> Tuple[Dict[str, str], Dict[str, int], bool, List[str], Dict[str, int]]:
    values = {}
    for code, row in GROUP_ROWS.items():
        values[code] = to_int(ws[f"{meal_def['col']}{row}"].value)

    replacements = {}
    for code, suffix in GROUP_SUFFIX.items():
        placeholder = f"{{{{{meal_def['name']}_{suffix}}}}}"
        replacements[placeholder] = "" if values[code] == 0 else str(
            values[code])

    include = any(values[code] > 0 for code in meal_def["groups"])
    tokens = [
        f"{{{{{meal_def['name']}_{GROUP_SUFFIX[code]}}}}}" for code in meal_def["groups"]]
    placeholder_values = {
        f"{{{{{meal_def['name']}_{suffix}}}}}": values[code]
        for code, suffix in GROUP_SUFFIX.items()
    }
    return replacements, values, include, tokens, placeholder_values


def format_short_name(full_name: str) -> str:
    parts = [p for p in full_name.split() if p.strip()]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) >= 3:
        return f"{parts[0]} {parts[-2]}"
    return f"{parts[0]} {parts[1]}"


def replace_in_text(text: str, replacements: Dict[str, str]) -> str:
    updated = text
    for key, value in replacements.items():
        if key in updated:
            updated = updated.replace(key, value)
    return updated


def replace_in_text_frame(text_frame, replacements: Dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        if not paragraph.text:
            continue
        replaced_run = False
        for run in paragraph.runs:
            new_text = replace_in_text(run.text, replacements)
            if new_text != run.text:
                run.text = new_text
                replaced_run = True
        if not replaced_run:
            new_para_text = replace_in_text(paragraph.text, replacements)
            if new_para_text != paragraph.text:
                paragraph.text = new_para_text


COL_MARKER_RE = re.compile(r"{{COL[:_](?P<key>[A-Z_]+)}}")
SHAPE_COL_MARKER_RE = re.compile(r"COL_(?P<key>[A-Z_]+)")
SHAPE_KEY_SUFFIXES = ("_ARROW", "_FLECHA", "_ICON")


def normalize_shape_key(key: str) -> str:
    for suffix in SHAPE_KEY_SUFFIXES:
        if key.endswith(suffix):
            return key[: -len(suffix)]
    return key


def should_hide_shape(shape, placeholder_values: Dict[str, int]) -> bool:
    name = getattr(shape, "name", "") or ""
    for key in SHAPE_COL_MARKER_RE.findall(name):
        normalized = normalize_shape_key(key)
        placeholder = f"{{{{{normalized}}}}}"
        if placeholder_values.get(placeholder, 0) == 0:
            return True
    return False


def remove_shape(shape) -> None:
    element = shape._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def replace_in_shape(
    shape,
    replacements: Dict[str, str],
    placeholder_values: Dict[str, int],
    slide_width: int,
    slide_shapes,
) -> None:
    if should_hide_shape(shape, placeholder_values):
        remove_shape(shape)
        return

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            replace_in_shape(
                subshape, replacements, placeholder_values, slide_width, slide_shapes
            )
        return

    if shape.has_table:
        table = shape.table
        cols_to_hide = set()
        for col_idx in range(len(table.columns)):
            for row in table.rows:
                cell_text = row.cells[col_idx].text
                markers = COL_MARKER_RE.findall(cell_text)
                if markers:
                    for key in markers:
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder_values.get(placeholder, 0) == 0:
                            cols_to_hide.add(col_idx)
                    cleaned = COL_MARKER_RE.sub("", cell_text).strip()
                    if cleaned != cell_text:
                        row.cells[col_idx].text = cleaned
                for placeholder, value in placeholder_values.items():
                    if value == 0 and placeholder in cell_text:
                        cols_to_hide.add(col_idx)

        if cols_to_hide:
            remove_table_columns(table, sorted(cols_to_hide, reverse=True))
            if slide_width and shape._parent is slide_shapes:
                table_width = sum(col.width for col in table.columns)
                shape.width = table_width
                shape.left = int((slide_width - shape.width) / 2)

        for row in table.rows:
            for cell in row.cells:
                replace_in_text_frame(cell.text_frame, replacements)
        return

    if shape.has_text_frame:
        replace_in_text_frame(shape.text_frame, replacements)


def text_frame_contains(text_frame, tokens: List[str]) -> bool:
    for paragraph in text_frame.paragraphs:
        if any(token in paragraph.text for token in tokens):
            return True
    return False


def shape_contains_tokens(shape, tokens: List[str]) -> bool:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(shape_contains_tokens(subshape, tokens) for subshape in shape.shapes)
    if shape.has_text_frame and text_frame_contains(shape.text_frame, tokens):
        return True
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if text_frame_contains(cell.text_frame, tokens):
                    return True
    return False


def remove_table_columns(table, col_indices: List[int]) -> None:
    tbl = table._tbl
    grid = tbl.tblGrid
    for col_idx in col_indices:
        grid_cols = getattr(grid, "gridCol_lst", None)
        if grid_cols is None:
            grid_cols = list(grid.iterchildren())
        if col_idx < 0 or col_idx >= len(grid_cols):
            continue
        grid.remove(grid_cols[col_idx])
        for row in table.rows:
            cells = getattr(row._tr, "tc_lst", None)
            if cells is None:
                cells = list(row._tr.iterchildren())
            if col_idx < len(cells):
                row._tr.remove(cells[col_idx])


def slide_contains_tokens(slide, tokens: List[str]) -> bool:
    return any(shape_contains_tokens(shape, tokens) for shape in slide.shapes)


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
        for shape in slide.shapes:
            replace_in_shape(
                shape,
                replacements,
                placeholder_values,
                presentation.slide_width,
                slide.shapes,
            )

    remove_slides_by_index(presentation, slides_to_remove)

    presentation.save(str(output_path))
    print(f"PPTX generado: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
