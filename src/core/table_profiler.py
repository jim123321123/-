from __future__ import annotations

import re
from collections.abc import Iterable


PROFILE_KEYWORDS = {
    "rna_seq_de": {
        "gene",
        "symbol",
        "locus",
        "log2fc",
        "log2foldchange",
        "pvalue",
        "p",
        "qvalue",
        "padj",
        "fdr",
        "fpkm",
        "counts",
        "expression",
    },
    "metabolomics": {
        "metabolite",
        "biochemical",
        "biochemicalname",
        "kegg",
        "hmdb",
        "foldchange",
        "pvalue",
        "superpathway",
        "subpathway",
    },
    "enrichment": {"term", "pathway", "ontology", "count", "pvalue", "fdr", "qvalue", "genes", "geneid", "enrichment"},
    "gene_list": {"gene", "symbol", "tag", "category", "list", "group"},
    "figure_source": {"figure", "panel", "mean", "sd", "sem", "n", "pvalue", "source"},
    "generic_numeric": set(),
}

TYPE_LABELS = {
    "rna_seq_de": "DEG / RNA-seq table",
    "metabolomics": "Metabolomics table",
    "enrichment": "GO / KEGG enrichment table",
    "gene_list": "Gene tag / gene list table",
    "figure_source": "Figure source table",
    "generic_numeric": "Generic numeric table",
}


def normalize_column_name(name: object) -> str:
    text = "" if name is None else str(name).strip().lower()
    text = re.sub(r"[\s_\-\(\)\.]+", "", text)
    return text


def make_unique_columns(columns: Iterable[object]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for index, column in enumerate(columns, start=1):
        value = "" if column is None else str(column).strip()
        if not value or value.lower().startswith("unnamed"):
            value = f"column_{index}"
        count = seen.get(value, 0)
        seen[value] = count + 1
        result.append(value if count == 0 else f"{value}_{count + 1}")
    return result


def detect_profile(columns: Iterable[object]) -> tuple[str, str]:
    normalized = {normalize_column_name(column) for column in columns}
    scores = {
        profile: len(normalized.intersection(keywords))
        for profile, keywords in PROFILE_KEYWORDS.items()
        if profile != "generic_numeric"
    }
    best_profile, best_score = max(scores.items(), key=lambda item: item[1])
    if best_score >= 2:
        return TYPE_LABELS[best_profile], best_profile
    return TYPE_LABELS["generic_numeric"], "generic_numeric"
