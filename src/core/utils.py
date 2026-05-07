from __future__ import annotations

import json
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml


RUN_SUBDIRS = [
    "input",
    "extracted",
    "outputs/tables",
    "outputs/reports",
    "outputs/figures",
    "outputs/image_check",
    "outputs/external_ai",
    "logs",
]


def safe_project_name(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_\-\u4e00-\u9fff]+", "_", name.strip())
    return cleaned.strip("_") or "project"


def create_run_dir(base_dir: Path, project_name: str) -> Path:
    run_name = f"{datetime.now():%Y%m%d_%H%M%S}_{safe_project_name(project_name)}"
    run_dir = base_dir / "runs" / run_name
    for subdir in RUN_SUBDIRS:
        (run_dir / subdir).mkdir(parents=True, exist_ok=True)
    return run_dir


def copy_input_file(source: Path | None, destination_dir: Path, name: str | None = None) -> Path | None:
    if not source:
        return None
    destination_dir.mkdir(parents=True, exist_ok=True)
    target = destination_dir / (name or source.name)
    shutil.copy2(source, target)
    return target


def extract_zip(zip_path: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            target = (target_dir / member.filename).resolve()
            if not str(target).startswith(str(target_dir.resolve())):
                raise ValueError(f"Unsafe zip path: {member.filename}")
        archive.extractall(target_dir)


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)


def open_path(path: Path) -> None:
    import os

    os.startfile(str(path))  # type: ignore[attr-defined]
