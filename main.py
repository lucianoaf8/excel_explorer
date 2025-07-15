#!/usr/bin/env python3
"""
Excel Explorer - Modern GUI Entry Point
A comprehensive Excel file analysis tool with real-time progress tracking.
"""

import sys
import os
import tkinter as tk
from pathlib import Path

# Add project source to path
project_root = Path(__file__).parent
src_path = project_root / "src"
sys.path.insert(0, str(src_path))

def main():
    """Launch the Excel Explorer GUI application"""
    try:
        from src.gui.excel_explorer_gui import ExcelExplorerApp
        
        # Create and run the application
        root = tk.Tk()
        app = ExcelExplorerApp(root)
        root.mainloop()
        
    except ImportError as e:
        print(f"Failed to import GUI module: {e}")
        print("Please ensure all dependencies are installed: pip install -r requirements.txt")
        return 1
    except Exception as e:
        print(f"Application failed to start: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
