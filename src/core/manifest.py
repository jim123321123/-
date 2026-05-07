from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd

from .file_classifier import classify_file, normalized_suffix


MANIFEST_COLUMNS = [
    "file_id",
    "file_name",
    "relative_path",
    "absolute_path",
    "extension",
    "file_type",
    "file_size_bytes",
    "modified_time",
    "md5",
    "sha256",
]


def file_hashes(path: Path) -> tuple[str, str]:
    md5 = hashlib.md5()
    sha256 = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            md5.update(chunk)
            sha256.update(chunk)
    return md5.hexdigest(), sha256.hexdigest()


def build_manifest(root: Path) -> pd.DataFrame:
    rows = []
    files = sorted(path for path in root.rglob("*") if path.is_file())
    for index, path in enumerate(files, start=1):
        md5, sha256 = file_hashes(path)
        stat = path.stat()
        rows.append(
            {
                "file_id": f"FILE{index:04d}",
                "file_name": path.name,
                "relative_path": str(path.relative_to(root)),
                "absolute_path": str(path.resolve()),
                "extension": normalized_suffix(path).lstrip("."),
                "file_type": classify_file(path),
                "file_size_bytes": stat.st_size,
                "modified_time": pd.Timestamp.fromtimestamp(stat.st_mtime).isoformat(),
                "md5": md5,
                "sha256": sha256,
            }
        )
    return pd.DataFrame(rows, columns=MANIFEST_COLUMNS)


def write_manifests(root: Path, output_dir: Path) -> tuple[pd.DataFrame, Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = build_manifest(root)
    raw_path = output_dir / "raw_file_manifest.csv"
    checksum_path = output_dir / "checksum_manifest.csv"
    manifest.to_csv(raw_path, index=False, encoding="utf-8-sig")
    manifest[["file_id", "relative_path", "md5", "sha256"]].to_csv(
        checksum_path, index=False, encoding="utf-8-sig"
    )
    return manifest, raw_path, checksum_path
