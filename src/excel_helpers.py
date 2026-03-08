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


@dataclass(frozen=True)
class ExampleFood:
    code: str
    description: str
    group_code: str
    amount_per_serving: float
    use_decimal: bool
    singular_text: str
    plural_text: str


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

GROUP_NAMES = {
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

EXAMPLE_GROUP_HEADERS = {
    "LACTEOS": "L",
    "VEGETALES": "V",
    "FRUTAS": "F",
    "ALMIDONES": "A",
    "PROTEINAS": "P",
    "GRASAS": "G",
}

MEAL_EXAMPLE_ORDER = {
    "PRE": ["P", "A", "F", "L", "G", "V"],
    "DES": ["P", "A", "G", "L", "F", "V"],
    "MAM": ["L", "A", "P", "F", "G", "V"],
    "ALM": ["P", "A", "V", "G", "F", "L"],
    "MTP": ["P", "A", "L", "F", "G", "V"],
    "CEN": ["P", "A", "V", "G", "F", "L"],
}

DEFAULT_EXAMPLE_FOOD_DEFS = {
    "PAN": ExampleFood(
        code="PAN",
        description="pan",
        group_code="A",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 reb. de pan",
        plural_text="{amount} reb. de pan",
    ),
    "AREPA": ExampleFood(
        code="AREPA",
        description="arepa",
        group_code="A",
        amount_per_serving=30,
        use_decimal=False,
        singular_text="30 g de arepa",
        plural_text="{amount} g de arepa",
    ),
    "GRANOLA": ExampleFood(
        code="GRANOLA",
        description="granola",
        group_code="A",
        amount_per_serving=15,
        use_decimal=False,
        singular_text="15 g de granola",
        plural_text="{amount} g de granola",
    ),
    "ARROZ": ExampleFood(
        code="ARROZ",
        description="arroz",
        group_code="A",
        amount_per_serving=50,
        use_decimal=False,
        singular_text="50 g de arroz",
        plural_text="{amount} g de arroz",
    ),
    "PURE DE PAPA": ExampleFood(
        code="PURE DE PAPA",
        description="pure de papa",
        group_code="A",
        amount_per_serving=60,
        use_decimal=False,
        singular_text="60 g de pure de papa",
        plural_text="{amount} g de pure de papa",
    ),
    "HUEVO": ExampleFood(
        code="HUEVO",
        description="huevo",
        group_code="P",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 huevo",
        plural_text="{amount} huevos",
    ),
    "JAMON": ExampleFood(
        code="JAMON",
        description="jamon",
        group_code="P",
        amount_per_serving=30,
        use_decimal=False,
        singular_text="30 g de jamon",
        plural_text="{amount} g de jamon",
    ),
    "POLLO": ExampleFood(
        code="POLLO",
        description="pollo",
        group_code="P",
        amount_per_serving=30,
        use_decimal=False,
        singular_text="30 g de pollo",
        plural_text="{amount} g de pollo",
    ),
    "ATUN": ExampleFood(
        code="ATUN",
        description="atun",
        group_code="P",
        amount_per_serving=30,
        use_decimal=False,
        singular_text="30 g de atun",
        plural_text="{amount} g de atun",
    ),
    "QUESO BLANCO": ExampleFood(
        code="QUESO BLANCO",
        description="queso blanco",
        group_code="P",
        amount_per_serving=30,
        use_decimal=False,
        singular_text="30 g de queso blanco",
        plural_text="{amount} g de queso blanco",
    ),
    "PROTEINA LIQUIDA": ExampleFood(
        code="PROTEINA LIQUIDA",
        description="proteina liquida",
        group_code="P",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 servicio de proteina liquida",
        plural_text="{amount} servicios de proteina liquida",
    ),
    "YOGURT GRIEGO": ExampleFood(
        code="YOGURT GRIEGO",
        description="yogurt griego",
        group_code="L",
        amount_per_serving=170,
        use_decimal=False,
        singular_text="170 g de yogurt griego",
        plural_text="{amount} g de yogurt griego",
    ),
    "LECHE": ExampleFood(
        code="LECHE",
        description="leche",
        group_code="L",
        amount_per_serving=240,
        use_decimal=False,
        singular_text="240 ml de leche",
        plural_text="{amount} ml de leche",
    ),
    "AGUACATE": ExampleFood(
        code="AGUACATE",
        description="aguacate",
        group_code="G",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 lonja de aguacate",
        plural_text="{amount} lonjas de aguacate",
    ),
    "ACEITE DE OLIVA": ExampleFood(
        code="ACEITE DE OLIVA",
        description="aceite de oliva",
        group_code="G",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 cdta de aceite de oliva",
        plural_text="{amount} cdtas de aceite de oliva",
    ),
    "CAMBUR": ExampleFood(
        code="CAMBUR",
        description="cambur",
        group_code="F",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 cambur",
        plural_text="{amount} cambures",
    ),
    "MANZANA": ExampleFood(
        code="MANZANA",
        description="manzana",
        group_code="F",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 manzana",
        plural_text="{amount} manzanas",
    ),
    "PERA": ExampleFood(
        code="PERA",
        description="pera",
        group_code="F",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 pera",
        plural_text="{amount} peras",
    ),
    "FRESAS": ExampleFood(
        code="FRESAS",
        description="fresas",
        group_code="F",
        amount_per_serving=80,
        use_decimal=False,
        singular_text="80 g de fresas",
        plural_text="{amount} g de fresas",
    ),
    "ENSALADA CRUDA": ExampleFood(
        code="ENSALADA CRUDA",
        description="ensalada cruda",
        group_code="V",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 taza de ensalada cruda",
        plural_text="{amount} tazas de ensalada cruda",
    ),
    "VEGETALES SALTEADOS": ExampleFood(
        code="VEGETALES SALTEADOS",
        description="vegetales salteados",
        group_code="V",
        amount_per_serving=1,
        use_decimal=False,
        singular_text="1 taza de vegetales salteados",
        plural_text="{amount} tazas de vegetales salteados",
    ),
}

EXAMPLE_FOOD_ALIASES = {
    "PAN BLANCO": "PAN",
    "PAN INTEGRAL": "PAN",
    "QUESO": "QUESO BLANCO",
    "YOGUR GRIEGO": "YOGURT GRIEGO",
    "ACEITE": "ACEITE DE OLIVA",
    "ENSALADA": "ENSALADA CRUDA",
}


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


def to_number(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", "."))
        except ValueError:
            return 0.0
    return 0.0


def format_quantity(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


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


def normalize_lookup_label(value: str) -> str:
    normalized = value.replace("_", " ").strip().upper()
    for source, target in (
        ("Á", "A"),
        ("É", "E"),
        ("Í", "I"),
        ("Ó", "O"),
        ("Ú", "U"),
        ("Ü", "U"),
    ):
        normalized = normalized.replace(source, target)
    return " ".join(normalized.split())


def find_labeled_value(
    ws,
    target_label: str,
    label_col: str,
    value_col: str,
    start_row: int,
    end_row: int,
):
    expected = normalize_lookup_label(target_label)
    for row_idx in range(start_row, end_row + 1):
        current_label = ws[f"{label_col}{row_idx}"].value
        if current_label is None:
            continue
        if normalize_lookup_label(str(current_label)) == expected:
            return ws[f"{value_col}{row_idx}"].value
    return None


def normalize_food_name(value: str) -> str:
    normalized = normalize_lookup_label(value)
    return EXAMPLE_FOOD_ALIASES.get(normalized, normalized)


def build_food_lookup_keys(value: str) -> List[str]:
    exact_key = normalize_lookup_label(value)
    alias_key = EXAMPLE_FOOD_ALIASES.get(exact_key)
    if alias_key and alias_key != exact_key:
        return [exact_key, alias_key]
    return [exact_key]


def parse_bool_like(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    normalized = normalize_lookup_label(str(value))
    return normalized in {"SI", "S", "YES", "Y", "TRUE", "1"}


def parse_float_like(value, field_label: str) -> float:
    if value_is_missing(value):
        raise ValidationError(f"Falta campo: {field_label}")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    try:
        return float(text)
    except ValueError as exc:
        raise ValidationError(
            f"Valor inválido en {field_label}: {value}"
        ) from exc


def format_example_amount(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def build_example_fragment(food: ExampleFood, servings: float) -> str:
    amount = servings * food.amount_per_serving
    if amount == food.amount_per_serving:
        return food.singular_text
    rendered_amount = format_example_amount(amount)
    return food.plural_text.replace("{amount}", rendered_amount).replace("{n}", rendered_amount)


def build_example_food_lookup(wb) -> Dict[str, ExampleFood]:
    if "EQUIVALENCIAS_EJEMPLOS" not in wb.sheetnames:
        return DEFAULT_EXAMPLE_FOOD_DEFS.copy()

    ws = wb["EQUIVALENCIAS_EJEMPLOS"]
    headers: Dict[str, int] = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        headers[normalize_lookup_label(str(cell.value))] = cell.column

    required_headers = [
        "CODIGO ALIMENTO",
        "GRUPO",
        "CANTIDAD POR RACION",
        "TEXTO SINGULAR",
        "TEXTO PLURAL",
    ]
    missing_headers = [
        header for header in required_headers if header not in headers]
    if missing_headers:
        raise ValidationError(
            "Faltan columnas en EQUIVALENCIAS_EJEMPLOS: " +
            ", ".join(missing_headers)
        )

    food_lookup: Dict[str, ExampleFood] = {}
    duplicated_keys: set[str] = set()

    for row_idx in range(2, ws.max_row + 1):
        code_value = ws.cell(
            row=row_idx, column=headers["CODIGO ALIMENTO"]).value
        if value_is_missing(code_value):
            continue

        code = normalize_lookup_label(str(code_value))
        group_value = ws.cell(row=row_idx, column=headers["GRUPO"]).value
        description_col = headers.get("DESCRIPCION BASE")
        description_value = ws.cell(
            row=row_idx, column=description_col).value if description_col else code_value
        quantity_value = ws.cell(
            row=row_idx, column=headers["CANTIDAD POR RACION"]).value
        singular_value = ws.cell(
            row=row_idx, column=headers["TEXTO SINGULAR"]).value
        plural_value = ws.cell(
            row=row_idx, column=headers["TEXTO PLURAL"]).value
        decimal_col = headers.get("USA DECIMAL")
        decimal_value = ws.cell(
            row=row_idx, column=decimal_col).value if decimal_col else None

        if value_is_missing(group_value):
            raise ValidationError(
                f"Falta campo: GRUPO (EQUIVALENCIAS_EJEMPLOS!B{row_idx})"
            )
        if value_is_missing(singular_value):
            raise ValidationError(
                f"Falta campo: TEXTO_SINGULAR (EQUIVALENCIAS_EJEMPLOS!G{row_idx})"
            )
        if value_is_missing(plural_value):
            raise ValidationError(
                f"Falta campo: TEXTO_PLURAL (EQUIVALENCIAS_EJEMPLOS!H{row_idx})"
            )

        group_label = normalize_lookup_label(str(group_value))
        if group_label not in EXAMPLE_GROUP_HEADERS:
            raise ValidationError(
                f"Grupo no soportado en EQUIVALENCIAS_EJEMPLOS fila {row_idx}: {group_value}"
            )

        food = ExampleFood(
            code=code,
            description=str(description_value or code_value).strip(),
            group_code=EXAMPLE_GROUP_HEADERS[group_label],
            amount_per_serving=parse_float_like(
                quantity_value,
                f"CANTIDAD_POR_RACION (EQUIVALENCIAS_EJEMPLOS!D{row_idx})",
            ),
            use_decimal=parse_bool_like(decimal_value),
            singular_text=str(singular_value).strip(),
            plural_text=str(plural_value).strip(),
        )

        lookup_keys = {
            code,
            normalize_lookup_label(str(description_value or code_value)),
        }
        for lookup_key in lookup_keys:
            if not lookup_key:
                continue
            if lookup_key in duplicated_keys:
                continue
            if lookup_key in food_lookup:
                duplicated_keys.add(lookup_key)
                food_lookup.pop(lookup_key, None)
                continue
            food_lookup[lookup_key] = food

    if not food_lookup:
        return DEFAULT_EXAMPLE_FOOD_DEFS.copy()

    return food_lookup


def load_examples_sheet(wb) -> Dict[str, Dict[str, str]]:
    if "EJEMPLOS_COMIDAS" not in wb.sheetnames:
        return {}

    ws = wb["EJEMPLOS_COMIDAS"]
    headers: Dict[str, int] = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        headers[normalize_lookup_label(str(cell.value))] = cell.column

    if "COMIDA" not in headers:
        raise ValidationError(
            "La hoja EJEMPLOS_COMIDAS debe incluir una columna COMIDA en la fila 1"
        )

    meal_rows: Dict[str, Dict[str, str]] = {}
    for row_idx in range(2, ws.max_row + 1):
        meal_value = ws.cell(row=row_idx, column=headers["COMIDA"]).value
        if value_is_missing(meal_value):
            continue
        meal_name = str(meal_value).strip().upper()
        if meal_name not in {meal["name"] for meal in MEAL_DEFS}:
            raise ValidationError(
                f"Comida no reconocida en EJEMPLOS_COMIDAS!A{row_idx}: {meal_value}"
            )
        if meal_name in meal_rows:
            raise ValidationError(
                f"La comida {meal_name} está repetida en EJEMPLOS_COMIDAS"
            )

        row_data: Dict[str, str] = {}
        for header_name, group_code in EXAMPLE_GROUP_HEADERS.items():
            col_idx = headers.get(header_name)
            value = ws.cell(
                row=row_idx, column=col_idx).value if col_idx else None
            row_data[group_code] = str(value).strip() if value else ""

        obs_col = headers.get("OBSERVACION")
        obs_value = ws.cell(
            row=row_idx, column=obs_col).value if obs_col else None
        row_data["OBSERVACION"] = str(obs_value).strip() if obs_value else ""
        meal_rows[meal_name] = row_data

    return meal_rows


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
        values.append(
            [format_table_value(ws[f"{col}{row_idx}"].value) for col in cols])
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
    estatura_value = find_labeled_value(
        ws_summary,
        target_label="Talla (m)",
        label_col="E",
        value_col="F",
        start_row=36,
        end_row=61,
    )
    if value_is_missing(estatura_value):
        estatura_value = ws_summary["F38"].value
    if value_is_missing(estatura_value):
        raise ValidationError(
            "Falta campo: Estatura (etiqueta 'Talla (m)' en RESUMEN ANTROPOMETRICO!E36:F61)"
        )
    estatura = format_decimal(estatura_value)
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
) -> Tuple[Dict[str, str], Dict[str, float], bool, List[str], Dict[str, float]]:
    values = {}
    for code, row in GROUP_ROWS.items():
        values[code] = to_number(ws[f"{meal_def['col']}{row}"].value)

    replacements = {}
    for code, suffix in GROUP_SUFFIX.items():
        placeholder = f"{{{{{meal_def['name']}_{suffix}}}}}"
        replacements[placeholder] = "" if values[code] == 0 else format_quantity(
            values[code]
        )

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


def build_meal_example_texts(wb, ws_req) -> Dict[str, str]:
    meal_rows = load_examples_sheet(wb)
    if not meal_rows:
        return {}
    food_lookup = build_example_food_lookup(wb)

    example_texts: Dict[str, str] = {}
    for meal_def in MEAL_DEFS:
        meal_name = meal_def["name"]
        servings = {
            code: to_number(
                ws_req[f"{meal_def['col']}{GROUP_ROWS[code]}"].value)
            for code in GROUP_ROWS
        }
        needs_example = any(
            servings[group_code] > 0 for group_code in meal_def["groups"]
        )
        if meal_name not in meal_rows:
            continue

        row_data = meal_rows[meal_name]

        fragments: List[str] = []
        for group_code in MEAL_EXAMPLE_ORDER.get(meal_name, meal_def["groups"]):
            if group_code not in meal_def["groups"]:
                continue
            serving_count = servings[group_code]
            if serving_count <= 0:
                continue

            food_name = row_data.get(group_code, "")
            if not food_name:
                continue

            food = None
            for lookup_key in build_food_lookup_keys(food_name):
                food = food_lookup.get(lookup_key)
                if food is not None:
                    break
            if food is None:
                fragments.append(food_name)
            else:
                fragments.append(build_example_fragment(food, serving_count))

        if not fragments and not row_data.get("OBSERVACION", "") and not needs_example:
            continue

        example_text = "EJEMPLO:"
        if fragments:
            example_text += " " + " + ".join(fragments)
        observation = row_data.get("OBSERVACION", "")
        if observation:
            example_text += f" | {observation}"
        example_texts[meal_name] = example_text

    return example_texts


def build_totals_replacements(ws) -> Dict[str, str]:
    totals_col = "T"
    totals = {
        "L": to_number(ws[f"{totals_col}{GROUP_ROWS['L']}"].value),
        "V": to_number(ws[f"{totals_col}{GROUP_ROWS['V']}"].value),
        "F": to_number(ws[f"{totals_col}{GROUP_ROWS['F']}"].value),
        "A": to_number(ws[f"{totals_col}{GROUP_ROWS['A']}"].value),
        "P": to_number(ws[f"{totals_col}{GROUP_ROWS['P']}"].value),
        "G": to_number(ws[f"{totals_col}{GROUP_ROWS['G']}"].value),
    }
    return {
        "{{TOTAL_LACTEOS}}": "" if totals["L"] == 0 else format_quantity(totals["L"]),
        "{{TOTAL_VEGETALES}}": "" if totals["V"] == 0 else format_quantity(totals["V"]),
        "{{TOTAL_FRUTAS}}": "" if totals["F"] == 0 else format_quantity(totals["F"]),
        "{{TOTAL_ALMIDONES}}": "" if totals["A"] == 0 else format_quantity(totals["A"]),
        "{{TOTAL_PROTEINAS}}": "" if totals["P"] == 0 else format_quantity(totals["P"]),
        "{{TOTAL_GRASAS}}": "" if totals["G"] == 0 else format_quantity(totals["G"]),
    }
