from pathlib import Path

from src.core.file_classifier import classify_file
from src.core.manifest import MANIFEST_COLUMNS, build_manifest


def test_manifest_generates_hashes_and_required_columns(tmp_path: Path):
    data_file = tmp_path / "table.csv"
    data_file.write_text("a,b\n1,2\n", encoding="utf-8")

    manifest = build_manifest(tmp_path)

    assert list(manifest.columns) == MANIFEST_COLUMNS
    row = manifest.iloc[0]
    assert row["file_name"] == "table.csv"
    assert row["file_type"] == "csv"
    assert len(row["md5"]) == 32
    assert len(row["sha256"]) == 64


def test_file_classifier_handles_required_extensions():
    assert classify_file(Path("a.pdf")) == "pdf"
    assert classify_file(Path("a.xlsx")) == "excel"
    assert classify_file(Path("a.fastq.gz")) == "omics_raw"
    assert classify_file(Path("script.R")) == "script"
    assert classify_file(Path("unknown.bin")) == "unknown"
