"""
Safe file operations utilities
"""

from pathlib import Path
import zipfile
from typing import BinaryIO

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

def file_exists(path: str | Path) -> bool:
    """Check that a path exists and is a file."""
    return Path(path).is_file()


def get_file_size_mb(path: str | Path) -> float:
    """Return file size in megabytes."""
    p = Path(path)
    return p.stat().st_size / 1_048_576 if p.exists() else 0.0


def is_excel_file(path: str | Path) -> bool:
    """Basic extension-based check for Excel file."""
    return Path(path).suffix.lower() in ALLOWED_EXTENSIONS


def open_readonly(path: str | Path) -> BinaryIO:
    """Open a file in binary read-only mode with safety checks."""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(path)
    return p.open("rb")


def contains_macros(xlsx_path: str | Path) -> bool:
    """Detect presence of VBA project inside an xlsm/xlsx file using zipfile scan."""
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        return any(name.upper().startswith("XL/VBA") for name in zf.namelist())
