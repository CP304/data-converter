"""Schneller Selbsttest ohne GUI: python data_converter_gui.py --self-test"""

import tempfile
from pathlib import Path

from . import filetools, tabular, xlsx_io


def run_self_test():
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)

        # XLSX-Roundtrip
        xlsx = root / "probe.xlsx"
        xlsx_io.write_xlsx(xlsx, ["SKU", "Preis"], [["A-100", "12.4"], ["007", "8"]])
        rows = xlsx_io.read_xlsx(xlsx)
        assert rows == [["SKU", "Preis"], ["A-100", "12.4"], ["007", "8"]], rows

        # CSV -> Tabelle -> JSON
        csv_file = root / "probe.csv"
        csv_file.write_text("sku;preis\nA-100;12,40\n", encoding="utf-8")
        table = tabular.read_table(csv_file)
        assert table.headers == ["sku", "preis"], table.headers
        transformed, _ = tabular.apply_transform(table, tabular.TransformOptions(decimal="comma_to_dot"))
        assert transformed.rows[0][1] == "12.40", transformed.rows

        # Umbenennen-Plan
        sample = root / "ALT_liste.csv"
        sample.write_text("x", encoding="utf-8")
        plan = filetools.plan_rename([sample], filetools.RenameOptions(search="ALT", replace="NEU"))
        assert plan and plan[0][1] == "NEU_liste.csv", plan

    print("Selbsttest OK")
