"""
Markdown utilities for Excel Explorer using mdutils
Provides GitHub-flavored markdown generation capabilities
"""

from typing import List, Dict, Optional
from mdutils.mdutils import MdUtils


class MarkdownReportBuilder:
    """Production-ready markdown report builder using mdutils"""
    
    def __init__(self, filename: str, title: str):
        self.md = MdUtils(file_name=filename, title=title)

    def add_section(self, title: str, level: int = 2):
        """Add a section header"""
        self.md.new_header(level=level, title=title)

    def add_paragraph(self, text: str):
        """Add a paragraph of text"""
        self.md.new_paragraph(text)

    def add_bullet_list(self, items: List[str]):
        """Add a bulleted list"""
        self.md.new_list(items)

    def add_numbered_list(self, items: List[str]):
        """Add a numbered list"""
        self.md.new_list(items, marked_with_number=True)

    def add_table(self, headers: List[str], rows: List[List[str]]):
        """Add a table with headers and rows"""
        # Flatten rows into a single list as required by mdutils
        flat = headers + [cell for row in rows for cell in row]
        row_count = len(rows) + 1
        col_count = len(headers)
        self.md.new_table(columns=col_count, rows=row_count, text=flat)

    def add_key_value_table(self, data: Dict[str, str]):
        """Add a key-value table"""
        headers = ["Property", "Value"]
        rows = [[k, str(v)] for k, v in data.items()]
        self.add_table(headers, rows)

    def add_horizontal_line(self):
        """Add a horizontal rule"""
        self.md.new_line("---")

    def add_code_block(self, code: str, language: str = ""):
        """Add a code block with optional language"""
        self.md.new_line("```" + language)
        self.md.new_line(code)
        self.md.new_line("```")

    def add_newline(self):
        """Add a blank line"""
        self.md.new_line()

    def add_bold_text(self, text: str):
        """Add bold text"""
        self.md.new_line(self.md.new_bold(text))

    def add_italic_text(self, text: str):
        """Add italic text"""
        self.md.new_line(self.md.new_italic(text))

    def save(self) -> str:
        """Save the markdown file and return the content"""
        return self.md.create_md_file()

    def get_content(self) -> str:
        """Get the markdown content without saving"""
        return self.md.get_md_text()
