from __future__ import annotations

import sys
import shutil
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from src.gui.main_window import MainWindow


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bundled_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")).resolve()
    return Path(__file__).resolve().parent


def ensure_runtime_files(base_dir: Path) -> None:
    source = bundled_base_dir()
    (base_dir / "runs").mkdir(parents=True, exist_ok=True)
    (base_dir / "config").mkdir(parents=True, exist_ok=True)
    for name in ("qc_thresholds.yaml", "table_profiles.yaml"):
        src = source / "config" / name
        dst = base_dir / "config" / name
        if src.exists() and not dst.exists():
            shutil.copy2(src, dst)
    asset_dir = base_dir / "src" / "assets"
    asset_dir.mkdir(parents=True, exist_ok=True)
    disclaimer_src = source / "src" / "assets" / "disclaimer.txt"
    disclaimer_dst = asset_dir / "disclaimer.txt"
    if disclaimer_src.exists() and not disclaimer_dst.exists():
        shutil.copy2(disclaimer_src, disclaimer_dst)


def main() -> int:
    app = QApplication(sys.argv)
    base_dir = runtime_base_dir()
    try:
        ensure_runtime_files(base_dir)
        window = MainWindow(base_dir)
        window.show()
        return app.exec()
    except Exception as exc:
        QMessageBox.critical(None, "启动失败", str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
