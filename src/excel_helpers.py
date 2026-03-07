from __future__ import annotations

from datetime import date, datetime, timedelta
from dataclasses import dataclass
from typing import Dict, List, Tuple

from openpyxl.utils.cell import column_index_from_string


@dataclass
class PatientInfo:
    name: str
    ci: str
    sex: str
    age: str
    discipline: str


class ValidationError(Exception):
    pass


@dataclass
class AnthropometricReportData:
    patient: PatientInfo
    peso_corporal_kg: str
    estatura_m: str
    masa_grasa_kg: str
    pct_grasa_carter: str
    table_resumen: List[List[str]]
    table_medidas: List[List[str]]


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


MONTH_NAMES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


def require_sheet(wb, sheet_name: str):
    if sheet_name not in wb.sheetnames:
        raise ValidationError(f"No existe la hoja requerida: {sheet_name}")
    return wb[sheet_name]


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


def load_patient_info(wb) -> PatientInfo:
    ws = require_sheet(wb, "HISTORIA")
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
        missing.append("Cedula (HISTORIA!C5)")
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


def value_is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def require_cell_value(ws, cell_ref: str, field_label: str):
    value = ws[cell_ref].value
    if value_is_missing(value):
        raise ValidationError(
            f"Falta campo: {field_label} ({ws.title}!{cell_ref})")
    return value


def to_age_text(value) -> str:
    if isinstance(value, (int, float)):
        return str(int(round(value)))
    if value is None:
        return ""
    text = str(value).strip()
    return text


def format_decimal(value, decimals: int = 2, decimal_comma: bool = True) -> str:
    if isinstance(value, (int, float)):
        formatted = f"{float(value):.{decimals}f}"
    else:
        text = str(value or "").strip()
        if not text:
            return ""
        try:
            parsed = float(text.replace(",", "."))
        except ValueError:
            return text
        formatted = f"{parsed:.{decimals}f}"
    if decimal_comma:
        return formatted.replace(".", ",")
    return formatted


def format_table_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{value.day}/{value.month}/{value.year % 100:02d}"
    if isinstance(value, date):
        return f"{value.day}/{value.month}/{value.year % 100:02d}"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value).strip()


def read_range_values(
    ws,
    start_col: str,
    end_col: str,
    start_row: int,
    end_row: int,
) -> List[List[str]]:
    start_col_idx = column_index_from_string(start_col)
    end_col_idx = column_index_from_string(end_col)
    values: List[List[str]] = []
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=end_row,
        min_col=start_col_idx,
        max_col=end_col_idx,
    ):
        values.append([format_table_value(cell.value) for cell in row])
    return values


def read_selected_columns(
    ws,
    cols: List[str],
    start_row: int,
    end_row: int,
) -> List[List[str]]:
    values: List[List[str]] = []
    for row_idx in range(start_row, end_row + 1):
        values.append([format_table_value(ws[f"{col}{row_idx}"].value) for col in cols])
    return values


def load_anthropometric_data(wb) -> AnthropometricReportData:
    ws_history = require_sheet(wb, "HISTORIA")
    ws_summary = require_sheet(wb, "RESUMEN ANTROPOMETRICO")

    full_name = str(require_cell_value(
        ws_history, "C4", "Nombre y Apellido")).strip()
    ci = str(require_cell_value(ws_history, "C5", "Cedula")).strip()
    age = to_age_text(require_cell_value(ws_history, "C7", "Edad"))
    discipline = str(ws_history["I8"].value or "").strip()
    sex = str(ws_history["C10"].value or "").strip()

    patient = PatientInfo(
        name=full_name,
        ci=ci,
        sex=sex,
        age=age,
        discipline=discipline,
    )

    peso_corporal = format_decimal(
        require_cell_value(
            ws_summary,
            "F6",
            "Peso corporal (RESUMEN ANTROPOMETRICO!F6)",
        )
    )
    estatura = format_decimal(
        require_cell_value(
            ws_summary,
            "F38",
            "Estatura (RESUMEN ANTROPOMETRICO!F38)",
        )
    )
    masa_grasa = format_decimal(
        require_cell_value(
            ws_summary,
            "F12",
            "Masa grasa (RESUMEN ANTROPOMETRICO!F12)",
        )
    )
    pct_grasa = format_decimal(
        require_cell_value(
            ws_summary,
            "F8",
            "% Grasa Carter (RESUMEN ANTROPOMETRICO!F8)",
        )
    )

    table_resumen = read_selected_columns(
        ws_summary,
        cols=["D", "F"],
        start_row=4,
        end_row=16,
    )
    table_medidas = read_range_values(
        ws_summary,
        start_col="E",
        end_col="F",
        start_row=36,
        end_row=61,
    )

    return AnthropometricReportData(
        patient=patient,
        peso_corporal_kg=peso_corporal,
        estatura_m=estatura,
        masa_grasa_kg=masa_grasa,
        pct_grasa_carter=pct_grasa,
        table_resumen=table_resumen,
        table_medidas=table_medidas,
    )


def month_name_es(reference_date: date) -> str:
    return MONTH_NAMES_ES[reference_date.month]


def build_anthropometric_replacements(
    data: AnthropometricReportData, today: date
) -> Dict[str, str]:
    next_control = today + timedelta(days=42)
    discipline = data.patient.discipline or "____________________"
    return {
        "{{PACIENTE}}": format_short_name(data.patient.name),
        "{{EDAD}}": data.patient.age,
        "{{CI}}": data.patient.ci,
        "{{DISCIPLINA}}": discipline,
        "{{OBJETIVO}}": "PERDER GRASA",
        "{{PESO_CORPORAL_KG}}": data.peso_corporal_kg,
        "{{ESTATURA_M}}": data.estatura_m,
        "{{MASA_GRASA_KG}}": data.masa_grasa_kg,
        "{{PCT_GRASA_CARTER}}": data.pct_grasa_carter,
        "{{MES_ACTUAL}}": month_name_es(today),
        "{{PROXIMO_CONTROL}}": next_control.strftime("%d/%m/%Y"),
    }


def build_summary_table_replacements(data: AnthropometricReportData) -> Dict[str, str]:
    replacements: Dict[str, str] = {}
    for row_idx, row_values in enumerate(data.table_resumen, start=1):
        for col_idx, cell_value in enumerate(row_values, start=1):
            replacements[f"{{{{R{row_idx}C{col_idx}}}}}"] = cell_value
    return replacements


def build_measurements_table_replacements(
    data: AnthropometricReportData,
) -> Dict[str, str]:
    replacements: Dict[str, str] = {}
    for row_idx, row_values in enumerate(data.table_medidas, start=1):
        for col_idx, cell_value in enumerate(row_values, start=1):
            replacements[f"{{{{M{row_idx}C{col_idx}}}}}"] = cell_value
            replacements[f"{{{{R{row_idx}C{col_idx}}}}}"] = cell_value
    return replacements


def format_short_name(full_name: str) -> str:
    parts = [p for p in full_name.split() if p.strip()]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) >= 3:
        return f"{parts[0]} {parts[-2]}"
    return f"{parts[0]} {parts[1]}"


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
        f"{{{{{meal_def['name']}_{GROUP_SUFFIX[code]}}}}}"
        for code in meal_def["groups"]
    ]
    placeholder_values = {
        f"{{{{{meal_def['name']}_{suffix}}}}}": values[code]
        for code, suffix in GROUP_SUFFIX.items()
    }
    return replacements, values, include, tokens, placeholder_values


def build_totals_replacements(ws) -> Dict[str, str]:
    totals_col = "T"
    totals = {
        "L": to_int(ws[f"{totals_col}{GROUP_ROWS['L']}"].value),
        "V": to_int(ws[f"{totals_col}{GROUP_ROWS['V']}"].value),
        "F": to_int(ws[f"{totals_col}{GROUP_ROWS['F']}"].value),
        "A": to_int(ws[f"{totals_col}{GROUP_ROWS['A']}"].value),
        "P": to_int(ws[f"{totals_col}{GROUP_ROWS['P']}"].value),
        "G": to_int(ws[f"{totals_col}{GROUP_ROWS['G']}"].value),
    }
    return {
        "{{TOTAL_LACTEOS}}": "" if totals["L"] == 0 else str(totals["L"]),
        "{{TOTAL_VEGETALES}}": "" if totals["V"] == 0 else str(totals["V"]),
        "{{TOTAL_FRUTAS}}": "" if totals["F"] == 0 else str(totals["F"]),
        "{{TOTAL_ALMIDONES}}": "" if totals["A"] == 0 else str(totals["A"]),
        "{{TOTAL_PROTEINAS}}": "" if totals["P"] == 0 else str(totals["P"]),
        "{{TOTAL_GRASAS}}": "" if totals["G"] == 0 else str(totals["G"]),
    }
