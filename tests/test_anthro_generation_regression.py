from __future__ import annotations
from generate_anthro_pptx import generate_anthro_pptx
from excel_helpers import inspect_workbook

import unittest
from datetime import datetime
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))


PLAN_HEADERS = [
    "COMIDA",
    "LACTEOS",
    "VEGETALES",
    "FRUTAS",
    "ALMIDONES",
    "PROTEINAS",
    "GRASAS",
]

SUMMARY_ROWS = [
    ("Evaluación", "1era evaluación"),
    ("Fecha", datetime(2026, 3, 25)),
    ("Peso (Kg)", 73.1),
    ("Talla parada (cm)", 173.5),
    ("%grasa carter", 26.0263),
    ("% Grasa (Durnin y W. 1974)", 35.56225646887326),
    ("Interpretación", "Sobrepeso II"),
    ("Kg de Masa Magra", 54.0747747),
    ("Kg de Grasa", 19.0252253),
    ("Masa Muscular (Kg)", 29.19015240045747),
    ("Masa Adiposa (Kg)", 31.7276355194793),
    ("Sumatoria de 6 pliegues", 145),
    ("Somatotipo", "Endomorfo"),
]

MEASUREMENT_ROWS = [
    ("Fecha de evaluación", datetime(2026, 3, 25)),
    ("Peso actual (kg)", 73.1),
    ("Talla (m)", 1.735),
    ("Talla (cm)", 173.5),
    ("Circunferencias(cm)", None),
    ("Brazo relajado (cm)", 31.5),
    ("Brazo Flexionado en Tensión (cm)", 32.5),
    ("Antebrazo máximo (cm)", 25),
    ("Tórax (Mesoesternal) (cm)", 93.8),
    ("Cintura mínimo (cm)", 80.7),
    ("Cadera máximo (cm)", 103.9),
    ("Muslo máximo (cm)", 63.6),
    ("Muslo medial (cm)", 55.7),
    ("Pantorrilla máximo (cm)", 35.1),
    ("Pliegues (mm)", None),
    ("Bíceps (mm)", 8.5),
    ("Tríceps (mm)", 21.5),
    ("Subescapular (mm)", 27.5),
    ("Ileo-crestal (mm)", 33.5),
    ("Supra-espinal (mm)", 22.5),
    ("Abdominal (mm)", 27.5),
    ("Muslo frontal (mm)", 32.5),
    ("Pantorrilla (mm)", 13.5),
    ("Diametros óseos (cm)", None),
    ("Humeral (Biepicondilar)", 6.1),
    ("Femoral (Biepicondilar)", 8.8),
]


def build_client_style_workbook(path: Path) -> None:
    wb = Workbook()

    history = wb.active
    history.title = "HISTORIA"
    history["C4"] = "Victoria Juliac"
    history["C5"] = "30322716"
    history["C7"] = 23
    history["C10"] = "Femenino"
    history["I8"] = ""

    plan = wb.create_sheet("PLAN_ALIMENTACION_TEMPLATE")
    plan.append(PLAN_HEADERS)
    plan.append(["DES", 1, 0, 1, 2, 1, 1])

    anthro = wb.create_sheet("ANTROPOMETRIA_TEMPLATE")
    anthro.append(["SECCION", "ETIQUETA", "VALOR"])
    for label, value in SUMMARY_ROWS:
        anthro.append(["RESUMEN", label, value])
    for label, value in MEASUREMENT_ROWS:
        anthro.append(["MEDIDAS", label, value])

    wb.save(path)


class AnthroGenerationRegressionTest(unittest.TestCase):
    def test_generation_accepts_pct_grasa_carter_alias(self) -> None:
        with TemporaryDirectory() as tmpdir:
            temp_dir = Path(tmpdir)
            excel_path = temp_dir / "Historia Clínica - Victoria Juliac.xlsx"
            output_path = temp_dir / "Informe Antropometrico.pptx"
            build_client_style_workbook(excel_path)

            workbook = load_workbook(excel_path, data_only=True)
            parsed_data = inspect_workbook(workbook)

            self.assertFalse(parsed_data.has_blocking_issues)
            self.assertEqual(parsed_data.anthro_data.pct_grasa_carter, "26,03")

            generated_path = generate_anthro_pptx(
                excel_path=excel_path,
                output_path=output_path,
                parsed_data=parsed_data,
            )

            self.assertEqual(generated_path, output_path)
            self.assertTrue(output_path.exists())


if __name__ == "__main__":
    unittest.main()
