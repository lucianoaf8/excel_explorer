#!/usr/bin/env python3
"""
Excel Explorer v2.0 - GUI Entry Point
"""

import tkinter as tk
from excel_explorer_gui import ExcelExplorerApp


def main():
    """Launch GUI application"""
    try:
        root = tk.Tk()
        app = ExcelExplorerApp(root)
        root.mainloop()
    except ImportError as e:
        print(f"Import failed: {e}")
        print("Install: pip install openpyxl")
        return 1
    except Exception as e:
        print(f"Startup failed: {e}")
        return 1
    return 0


if __name__ == "__main__":
    exit(main())
