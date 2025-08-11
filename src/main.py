#!/usr/bin/env python3
"""
Excel Explorer v2.0 - Unified Entry Point with Anonymizer
Supports both GUI and CLI modes with data anonymization
"""

import argparse
import sys
from pathlib import Path


def main():
    """Unified entry point with mode selection and anonymization"""
    parser = argparse.ArgumentParser(
        description='Excel Explorer - Advanced Excel File Analysis Tool with Anonymization',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # GUI Mode
  python main.py                                    # Launch GUI (default)
  python main.py --mode gui                         # Launch GUI explicitly
  
  # CLI Analysis
  python main.py --mode cli --file data.xlsx       # CLI analysis with HTML report
  python main.py --mode cli --file data.xlsx --format json --output ./reports
  python main.py --mode cli --file data.xlsx --format text --config custom.yaml
  
  # Anonymization
  python main.py --mode cli --file data.xlsx --anonymize
  python main.py --mode cli --file data.xlsx --anonymize --anonymize-columns "Sheet1:Name" "Sheet1:Company"
  python main.py --mode cli --file data.xlsx --anonymize --mapping-file mappings.json
  
  # Reverse Anonymization
  python main.py --mode cli --file anonymized.xlsx --reverse mappings.json
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
    parser.add_argument('--screenshots', action='store_true',
                       help='Enable screenshot capture of Excel sheets')
    
    # Anonymization arguments
    anon_group = parser.add_argument_group('Anonymization Options')
    anon_group.add_argument('--anonymize', action='store_true',
                          help='Enable data anonymization before analysis')
    anon_group.add_argument('--anonymize-columns', nargs='+', metavar='SHEET:COLUMN',
                          help='Specific columns to anonymize (e.g., "Sheet1:B" or "Sheet1:Name")')
    anon_group.add_argument('--mapping-file', type=str, metavar='PATH',
                          help='Path for mapping dictionary file (default: auto-generated)')
    anon_group.add_argument('--mapping-format', choices=['json', 'excel'], default='json',
                          help='Format for mapping file (default: json)')
    anon_group.add_argument('--reverse', type=str, metavar='MAPPING_FILE',
                          help='Reverse anonymization using the specified mapping file')
    anon_group.add_argument('--anonymized-output', type=str, metavar='PATH',
                          help='Output path for anonymized file (default: adds _anonymized)')
    
    args = parser.parse_args()
    
    try:
        if args.mode == 'gui':
            # Check if anonymization was requested in GUI mode
            if args.anonymize or args.reverse:
                print("Note: Anonymization features are currently CLI-only")
                print("Use --mode cli for anonymization")
                return 1
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
                verbose=args.verbose,
                enable_screenshots=args.screenshots,
                # Anonymizer parameters
                anonymize=args.anonymize,
                anonymize_columns=args.anonymize_columns,
                mapping_file=args.mapping_file,
                mapping_format=args.mapping_format,
                reverse=args.reverse,
                anonymized_output=args.anonymized_output
            )
            
    except ImportError as e:
        print(f"Import failed: {e}")
        if 'faker' in str(e).lower():
            print("Anonymizer requires faker library: pip install faker")
        else:
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
