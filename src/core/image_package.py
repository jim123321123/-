from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pandas as pd

from .file_classifier import classify_file


IMAGE_INVENTORY_COLUMNS = [
    "image_id",
    "file_name",
    "relative_path",
    "extension",
    "source_type",
    "suggested_check",
    "included_in_package",
]


def _source_type(path: Path) -> str:
    suffix = path.suffix.lower()
    name = path.name.lower()
    if suffix == ".pdf":
        return "manuscript_pdf" if "manuscript" in name or "paper" in name else "supplement_pdf"
    if suffix in {".czi", ".nd2"}:
        return "raw_image"
    if "fig" in name or "panel" in name:
        return "figure_image"
    return "raw_image"


def _suggested_check(path: Path) -> str:
    if path.suffix.lower() in {".czi", ".nd2"}:
        return "unsupported raw format, export preview recommended"
    if path.suffix.lower() == ".pdf":
        return "manual review recommended"
    return "Proofig AI / Imagetwin recommended"


def create_image_check_package(extracted_dir: Path, output_dir: Path) -> tuple[pd.DataFrame, Path]:
    package_dir = output_dir / "package"
    package_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    for index, path in enumerate(sorted(p for p in extracted_dir.rglob("*") if p.is_file()), start=1):
        if classify_file(path) not in {"image", "pdf"}:
            continue
        relative = path.relative_to(extracted_dir)
        target = package_dir / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
        rows.append(
            {
                "image_id": f"IMG{index:04d}",
                "file_name": path.name,
                "relative_path": str(relative),
                "extension": path.suffix.lower().lstrip("."),
                "source_type": _source_type(path),
                "suggested_check": _suggested_check(path),
                "included_in_package": "Yes",
            }
        )
    inventory = pd.DataFrame(rows, columns=IMAGE_INVENTORY_COLUMNS)
    inventory_path = output_dir / "image_inventory.csv"
    inventory.to_csv(inventory_path, index=False, encoding="utf-8-sig")
    zip_path = output_dir / "image_check_package.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in package_dir.rglob("*"):
            if file.is_file():
                archive.write(file, file.relative_to(package_dir))
        archive.write(inventory_path, inventory_path.name)
    return inventory, zip_path
