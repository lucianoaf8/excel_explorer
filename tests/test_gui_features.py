#!/usr/bin/env python3
"""
Test script to verify new GUI features:
1. Checkbox for auto-generating reports
2. Button to open reports folder
"""

import tkinter as tk
from tkinter import ttk
from pathlib import Path
import sys

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from src.gui.excel_explorer_gui import ExcelExplorerApp


def test_gui_features():
    """Test the new GUI features"""
    print("Testing Excel Explorer GUI features...")
    print("1. Check that 'Generate reports upon completion' checkbox appears in File Selection")
    print("2. Check that 'Open Reports Folder' button appears in action buttons")
    print("3. Test the checkbox functionality by running an analysis")
    print("4. Test the Open Reports Folder button")
    
    # Create and run the GUI
    root = tk.Tk()
    app = ExcelExplorerApp(root)
    
    # Print initial state
    print(f"\nInitial checkbox state: {app.auto_generate_reports.get()}")
    
    # Set up a callback to monitor checkbox changes
    def on_checkbox_change(*args):
        print(f"Checkbox changed to: {app.auto_generate_reports.get()}")
    
    app.auto_generate_reports.trace('w', on_checkbox_change)
    
    # Run the GUI
    root.mainloop()


if __name__ == "__main__":
    test_gui_features()