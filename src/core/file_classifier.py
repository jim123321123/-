from pathlib import Path


FILE_TYPE_EXTENSIONS = {
    "pdf": {".pdf"},
    "docx": {".docx"},
    "excel": {".xlsx", ".xls"},
    "csv": {".csv", ".tsv", ".txt"},
    "image": {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".czi", ".nd2"},
    "script": {".r", ".py", ".ipynb"},
    "omics_raw": {".fastq", ".fq", ".fastq.gz", ".fq.gz", ".bam", ".sam", ".raw", ".mzml", ".mzxml"},
    "archive": {".zip", ".gz"},
}


def normalized_suffix(path: Path) -> str:
    name = path.name.lower()
    for compound in (".fastq.gz", ".fq.gz"):
        if name.endswith(compound):
            return compound
    return path.suffix.lower()


def classify_file(path: Path) -> str:
    suffix = normalized_suffix(path)
    for file_type, extensions in FILE_TYPE_EXTENSIONS.items():
        if suffix in extensions:
            return file_type
    return "unknown"
