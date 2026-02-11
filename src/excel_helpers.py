from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple


@dataclass
class PatientInfo:
    name: str
    ci: str
    sex: str
    age: str
    discipline: str


class ValidationError(Exception):
    pass


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
