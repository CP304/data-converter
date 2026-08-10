import tempfile
import zipfile
from pathlib import Path

from data_converter_gui import ConverterBackend, FileLogger, JobConfig, SpecialOptions, suggested_outputs


def run_mock_conversion_test():
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        source_dir = root / "input"
        output_dir = root / "output"
        source_dir.mkdir()
        output_dir.mkdir()

        (source_dir / "ALT_supplier_prices.csv").write_text("sku,price\nA-100,12.40\nB-200,8.10\n", encoding="utf-8")
        (source_dir / "ignore_notes.txt").write_text("not part of this job", encoding="utf-8")

        log_file = output_dir / "conversion_log_test.txt"
        logger = FileLogger(log_file)
        backend = ConverterBackend(logger, lambda: False)
        job = JobConfig(
            source=source_dir,
            source_kind="folder",
            input_exts=[".csv"],
            output_exts=[".json", ".tsv", ".zip"],
            contains="supplier",
            max_mb="1",
            special=SpecialOptions(
                search_text="ALT",
                replace_text="NEW",
                prefix="PO_",
                suffix="_2026",
                zip_outputs=True,
            ),
        )

        backend.run_job(job, output_dir, same_folder=False, recursive=True)

        expected_json = output_dir / "PO_NEW_supplier_prices_2026.json"
        expected_tsv = output_dir / "PO_NEW_supplier_prices_2026.tsv"
        expected_zip = output_dir / "PO_NEW_supplier_prices_2026.zip"
        expected_bundle = output_dir / "PO_NEW_supplier_prices_2026_converted.zip"

        assert expected_json.exists(), expected_json
        assert expected_tsv.exists(), expected_tsv
        assert expected_zip.exists(), expected_zip
        assert expected_bundle.exists(), expected_bundle
        assert "A-100" in expected_json.read_text(encoding="utf-8")
        assert "OK:" in log_file.read_text(encoding="utf-8")

        with zipfile.ZipFile(expected_bundle) as archive:
            names = set(archive.namelist())
            assert expected_json.name in names
            assert expected_tsv.name in names


def run_format_catalog_test():
    checks = {
        ".sldasm": ".step",
        ".catproduct": ".stp",
        ".dwg": ".pdf",
        ".gbr": ".zip",
        ".heic": ".webp",
        ".xlsb": ".xlsx",
        ".msg": ".pdf",
        ".svg": ".png",
        ".rar": ".zip",
        ".mp4": ".webm",
    }
    for input_ext, output_ext in checks.items():
        assert output_ext in suggested_outputs([input_ext]), (input_ext, output_ext)


if __name__ == "__main__":
    run_format_catalog_test()
    run_mock_conversion_test()
    print("Mock conversion tests OK")
