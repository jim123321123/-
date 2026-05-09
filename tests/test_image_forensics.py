from pathlib import Path

from PIL import Image, ImageEnhance

from src.core.image_forensics import run_image_forensics


def test_exact_duplicate_image_detector(tmp_path: Path):
    image = Image.new("RGB", (24, 24), "red")
    left = tmp_path / "WT_control.png"
    right = tmp_path / "KO_treat.png"
    image.save(left)
    image.save(right)

    issues = run_image_forensics(tmp_path)

    assert any(issue["rule_id"] == "I001" for issue in issues)
    assert any(issue["severity"] == "CRITICAL" for issue in issues)


def test_perceptual_duplicate_image_detector(tmp_path: Path):
    image = Image.new("RGB", (32, 32), "blue")
    for x in range(8, 24):
        for y in range(8, 24):
            image.putpixel((x, y), (220, 220, 255))
    left = tmp_path / "day1.png"
    right = tmp_path / "day2.png"
    image.save(left)
    ImageEnhance.Brightness(image).enhance(1.03).save(right)

    issues = run_image_forensics(tmp_path)

    assert any(issue["rule_id"] == "I002" for issue in issues)
