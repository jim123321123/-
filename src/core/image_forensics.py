from __future__ import annotations

import hashlib
from itertools import combinations
from pathlib import Path
from typing import Any

import pandas as pd
from PIL import Image, ImageOps

from .file_classifier import classify_file


SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff"}

FUTURE_IMAGE_RULES = {
    "I003": "splicing_boundary_detector：预留给拼接边界/局部纹理异常检查，需要接入更稳定的图像取证算法后启用。",
    "I004": "copy_move_detector：预留给同图内部复制移动检查，需要避免把标尺、文字和重复结构误判为异常。",
    "I005": "metadata_consistency_detector：预留给显微镜/相机元数据一致性检查，需要先定义不同仪器的白名单字段。",
    "I006": "panel_layout_consistency_detector：预留给图版排版一致性检查，需要结合人工标注或外部图像平台结果。",
}


def _issue(
    issues: list[dict[str, Any]],
    issue_type: str,
    severity: str,
    file_name: str,
    evidence: str,
    related_files: str,
    details: dict[str, Any],
) -> None:
    risk_level = {"CRITICAL": "Red", "HIGH": "Red", "MEDIUM": "Orange", "LOW": "Yellow"}.get(severity, "Yellow")
    issues.append(
        {
            "issue_id": f"IMG{len(issues) + 1:03d}",
            "module": "Image Forensics",
            "rule_id": "I001" if issue_type == "exact_duplicate_image_detector" else "I002",
            "severity": severity,
            "risk_level": risk_level,
            "issue_type": issue_type,
            "file_name": file_name,
            "sheet_name": "",
            "row_index": "",
            "column_name": "",
            "related_columns": related_files,
            "evidence": evidence,
            "recommended_action": "建议核对原始图像、采集参数、图像命名和实验条件，确认是否为同一图片被重复使用。",
            "details": details,
            "need_human_review": "Yes" if risk_level in {"Red", "Orange"} else "Recommended",
            "affects_submission": "Yes" if risk_level in {"Red", "Orange"} else "Review",
        }
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _condition_hint(paths: list[str]) -> bool:
    text = " ".join(paths).lower()
    tokens = ("control", "treat", "wt", "ko", "day1", "day2", "vehicle", "drug", "apoe", "gl2")
    return sum(1 for token in tokens if token in text) >= 2


def _average_hash(path: Path, size: int = 8) -> int | None:
    try:
        with Image.open(path) as image:
            gray = ImageOps.grayscale(image).resize((size, size))
            pixels = list(gray.getdata())
    except Exception:
        return None
    avg = sum(pixels) / len(pixels)
    bits = 0
    for pixel in pixels:
        bits = (bits << 1) | int(pixel >= avg)
    return bits


def _hamming(left: int, right: int) -> int:
    return (left ^ right).bit_count()


def _image_files(root: Path) -> list[Path]:
    files = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or classify_file(path) != "image":
            continue
        if path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS:
            files.append(path)
    return files


def run_image_forensics(root: Path, thresholds: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    cfg = (thresholds or {}).get("image", {})
    exact_enabled = cfg.get("exact_duplicate_enabled", True)
    perceptual_enabled = cfg.get("perceptual_duplicate_enabled", True)
    high_distance = int(cfg.get("phash_high_distance", 5))
    medium_distance = int(cfg.get("phash_medium_distance", 10))
    issues: list[dict[str, Any]] = []
    files = _image_files(root)

    if exact_enabled:
        by_hash: dict[str, list[Path]] = {}
        for path in files:
            by_hash.setdefault(_sha256(path), []).append(path)
        for digest, group in by_hash.items():
            if len(group) < 2:
                continue
            rels = [str(path.relative_to(root)) for path in group]
            severity = "CRITICAL" if _condition_hint(rels) else "HIGH"
            _issue(
                issues,
                "exact_duplicate_image_detector",
                severity,
                group[0].name,
                f"发现 {len(group)} 个 SHA256 完全相同的图片文件，但文件名或路径不同。",
                "; ".join(rels),
                {"sha256": digest, "files": rels, "image_1": rels[0], "image_2": rels[1] if len(rels) > 1 else ""},
            )

    if perceptual_enabled:
        hashes = [(path, _average_hash(path)) for path in files]
        hashes = [(path, value) for path, value in hashes if value is not None]
        reported: set[tuple[str, str]] = set()
        for (left_path, left_hash), (right_path, right_hash) in combinations(hashes, 2):
            if _sha256(left_path) == _sha256(right_path):
                continue
            distance = _hamming(left_hash, right_hash)
            if distance > medium_distance:
                continue
            rels = [str(left_path.relative_to(root)), str(right_path.relative_to(root))]
            key = tuple(sorted(rels))
            if key in reported:
                continue
            reported.add(key)
            severity = "HIGH" if distance <= high_distance else "MEDIUM"
            if severity == "HIGH" and _condition_hint(rels):
                severity = "CRITICAL"
            _issue(
                issues,
                "perceptual_duplicate_detector",
                severity,
                left_path.name,
                f"发现两张图片感知哈希距离为 {distance}，可能是缩放、压缩或轻微亮度变化后的近重复图片。",
                "; ".join(rels),
                {"hamming_distance": distance, "distance": distance, "files": rels, "image_1": rels[0], "image_2": rels[1]},
            )
    return issues


def write_image_forensics_results(issues: list[dict[str, Any]], output_path: Path) -> pd.DataFrame:
    df = pd.DataFrame(issues)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return df
