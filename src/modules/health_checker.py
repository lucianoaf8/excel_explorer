"""
File integrity validation module
"""

from ..core.base_analyzer import BaseAnalyzer
from pathlib import Path
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from ..utils import file_handler as fh
from ..utils.error_handler import ExplorerError


class HealthChecker(BaseAnalyzer):
    def __init__(self, config=None):
        super().__init__(config)
        self._results = {}

    def _check_corruption(self, file_path: Path):
        try:
            openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            return False
        except InvalidFileException:
            return True
        except Exception:  # broader corruption / unreadable
            return True

    def analyze(self, workbook_path: str | Path):
        """Validate an Excel file and populate self._results."""
        file_path = Path(workbook_path)
        if not fh.file_exists(file_path):
            raise ExplorerError(f"File not found: {file_path}")
        if not fh.is_excel_file(file_path):
            raise ExplorerError("Unsupported file type – expected Excel")

        max_mb = self.config.get("max_file_size_mb", 500)
        size_mb = fh.get_file_size_mb(file_path)

        issues = []
        if size_mb > max_mb:
            issues.append(f"File size {size_mb:.1f}MB exceeds limit {max_mb}MB")

        corrupted = self._check_corruption(file_path)
        if corrupted:
            issues.append("File appears to be corrupted or unreadable")

        if fh.contains_macros(file_path):
            issues.append("File contains VBA macros")

        # openpyxl raises InvalidFileException for encrypted files
        try:
            openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        except InvalidFileException:
            issues.append("File may be password protected or encrypted")

        status = "OK" if not issues else "FAIL"
        score = max(0, 100 - len(issues) * 25)  # simple scoring algo

        self._results = {
            "status": status,
            "issues": issues,
            "size_mb": size_mb,
            "score": score,
        }

    def get_results(self):
        """Return structured results"""
        return self._results
