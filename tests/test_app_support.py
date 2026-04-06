from __future__ import annotations

from pathlib import Path
import sys
import unittest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from app_support import default_output_dir_for_excel


class AppSupportTest(unittest.TestCase):
    def test_default_output_dir_for_excel_uses_excel_parent(self) -> None:
        excel_path = Path("C:/clientes/acme/Paciente 01.xlsx")

        output_dir = default_output_dir_for_excel(excel_path)

        self.assertEqual(output_dir, Path("C:/clientes/acme"))


if __name__ == "__main__":
    unittest.main()
