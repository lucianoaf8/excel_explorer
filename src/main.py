#!/usr/bin/env python3
"""
Excel Explorer v2.0 - Unified Entry Point
Supports both GUI and CLI modes
"""

import argparse
import sys
from pathlib import Path


def main():
    """Unified entry point with mode selection"""
    parser = argparse.ArgumentParser(
        description='Excel Explorer - Advanced Excel File Analysis Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py                                    # Launch GUI (default)
  python main.py --mode gui                         # Launch GUI explicitly
  python main.py --mode cli --file data.xlsx       # CLI analysis with HTML report
  python main.py --mode cli --file data.xlsx --format json --output ./reports
  python main.py --mode cli --file data.xlsx --format text --config custom.yaml
        """
    )
    
    parser.add_argument('--mode', choices=['gui', 'cli'], default='gui', 
                       help='Execution mode (default: gui)')
    parser.add_argument('--file', type=str, 
                       help='Excel file to analyze (required for CLI mode)')
    parser.add_argument('--output', type=str, 
                       help='Output directory (default: ./reports)')
    parser.add_argument('--format', choices=['html', 'json', 'text', 'markdown'], 
                       default='html', help='Report format (default: html)')
    parser.add_argument('--config', type=str, default='config.yaml', 
                       help='Configuration file path (default: config.yaml)')
    parser.add_argument('--verbose', '-v', action='store_true', 
                       help='Enable verbose output')
    
    args = parser.parse_args()
    
    try:
        if args.mode == 'gui':
            return _launch_gui()
        else:
            if not args.file:
                print("Error: --file is required for CLI mode")
                parser.print_help()
                return 1
            
            from cli.cli_runner import run_cli_analysis
            return run_cli_analysis(
                file_path=args.file,
                output_dir=args.output,
                format_type=args.format,
                config_path=args.config,
                verbose=args.verbose
            )
            
    except ImportError as e:
        print(f"Import failed: {e}")
        print("Install required dependencies: pip install -r requirements.txt")
        return 1
    except Exception as e:
        print(f"Startup failed: {e}")
        if args.verbose if hasattr(args, 'verbose') else False:
            import traceback
            traceback.print_exc()
        return 1


def _launch_gui():
    """Launch the GUI application"""
    try:
        import tkinter as tk
        from gui.excel_explorer_gui import ExcelExplorerApp
        
        root = tk.Tk()
        app = ExcelExplorerApp(root)
        root.mainloop()
        return 0
        
    except ImportError as e:
        if 'tkinter' in str(e).lower():
            print("GUI not available: tkinter not installed")
            print("On Linux: sudo apt-get install python3-tk")
            print("On macOS: tkinter should be included with Python")
            print("On Windows: tkinter should be included with Python")
        else:
            print(f"GUI import failed: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
