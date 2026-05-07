import pandas as pd

from src.core.report_summary import generate_raw_data_overview


def test_raw_data_overview_describes_package_without_api():
    manifest = pd.DataFrame(
        [
            {"file_type": "excel", "file_name": "data.xlsx"},
            {"file_type": "pdf", "file_name": "paper.pdf"},
            {"file_type": "image", "file_name": "figure.png"},
        ]
    )
    sheets = pd.DataFrame(
        [
            {"qc_profile": "generic_numeric", "parse_status": "ok"},
            {"qc_profile": "rna_seq_de", "parse_status": "ok"},
        ]
    )
    issues = pd.DataFrame(
        [
            {"risk_level": "Red", "module": "Numeric Forensics"},
            {"risk_level": "Orange", "module": "Numeric Forensics"},
            {"risk_level": "Yellow", "module": "External AI"},
        ]
    )
    external = pd.DataFrame(
        [
            {"tool": "Proofig AI", "status": "skipped"},
            {"tool": "LLM", "status": "skipped"},
        ]
    )

    text = generate_raw_data_overview(manifest, sheets, issues, external)

    assert "本次上传的压缩包共识别 3 个文件" in text
    assert "excel 1 个" in text
    assert "共解析 2 个 sheet" in text
    assert "Red 1 项" in text
    assert "未配置可用的外部 AI API" in text
