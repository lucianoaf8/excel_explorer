"""
Workbook architecture analysis module
"""

from ..core.base_analyzer import BaseAnalyzer
from pathlib import Path
import openpyxl
from openpyxl.worksheet.table import Table
from ..utils.error_handler import ExplorerError


class StructureMapper(BaseAnalyzer):
    """Analyze workbook structure: sheets, named ranges, tables, props."""

    def __init__(self, config=None):
        super().__init__(config or {})
        self._results = {}

    def analyze(self, workbook_path: str | Path):
        file_path = Path(workbook_path)
        if not file_path.is_file():
            raise ExplorerError(f"StructureMapper: file not found: {file_path}")

        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=False)

        include_hidden = self.config.get("include_hidden_sheets", True)

        sheets = []
        tables = []
        for ws in wb.worksheets:
            visible = ws.sheet_state == "visible"
            if not visible and not include_hidden:
                continue

            sheet_info = {
                "name": ws.title,
                "visible": visible,
                "protected": bool(ws.protection.sheet),
            }
            sheets.append(sheet_info)

            # Tables on this sheet
            for tbl_name, tbl_obj in ws.tables.items():
                if isinstance(tbl_obj, Table):
                    tables.append({
                        "sheet": ws.title,
                        "name": tbl_name,
                        "ref": tbl_obj.ref,
                    })

        # Named ranges (global definitions only for now)
        named_ranges = []
        # Iterate through defined names (openpyxl 3.1+ uses iterator)
        for defn in getattr(wb.defined_names, "definedName", wb.defined_names):
            # Skip autoFilter or print titles etc.
            if defn.name.startswith("_xlnm"):
                continue
            named_ranges.append({
                "name": defn.name,
                "attr_text": defn.attr_text,
            })

        props = wb.properties
        workbook_properties = {
            "sheet_count": len(wb.sheetnames),
            "creator": props.creator,
            "created": str(props.created) if props.created else None,
            "modified": str(props.modified) if props.modified else None,
        }

        self._results = {
            "sheets": sheets,
            "tables": tables,
            "named_ranges": named_ranges,
            "workbook_properties": workbook_properties,
        }

    def get_results(self):
        return self._results
