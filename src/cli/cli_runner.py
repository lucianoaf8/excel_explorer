#!/usr/bin/env python3
"""
CLI Runner for Excel Explorer
Handles command-line analysis execution with progress feedback
"""

import sys
from pathlib import Path
from typing import Optional
from datetime import datetime

from core.analysis_service import AnalysisService
from reports.report_adapter import ReportService


class CLIProgressCallback:
    """Simple progress callback for CLI mode"""
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        self.last_message = None
        
    def __call__(self, message: str, progress: float):
        """Handle progress updates from AnalysisService"""
        # Only show different messages to avoid spam
        if message != self.last_message:
            if self.verbose:
                print(f"[{progress*100:.0f}%] {message}")
            elif "complete" in message.lower() or progress >= 1.0:
                print(f"✓ {message}")
            elif progress == 0.0 or "starting" in message.lower():
                print(f"→ {message}")
            
            self.last_message = message


def run_cli_analysis(
    file_path: str,
    output_dir: Optional[str] = None,
    format_type: str = 'html',
    config_path: str = 'config.yaml',
    verbose: bool = False
) -> int:
    """
    Execute CLI-based analysis with comprehensive error handling
    
    Args:
        file_path: Path to Excel file to analyze
        output_dir: Output directory (default: ./reports)
        format_type: Report format (html, json, text, markdown)
        config_path: Configuration file path
        verbose: Enable detailed progress output
        
    Returns:
        Exit code (0 = success, 1 = error)
    """
    try:
        # Validate input file
        input_file = Path(file_path)
        if not input_file.exists():
            print(f"Error: File not found: {file_path}")
            return 1
            
        if not input_file.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
            print(f"Error: Unsupported file type: {input_file.suffix}")
            print("Supported formats: .xlsx, .xls, .xlsm")
            return 1
        
        # Setup output directory
        if not output_dir:
            output_dir = Path("reports")
        else:
            output_dir = Path(output_dir)
        
        output_dir.mkdir(exist_ok=True)
        
        # Initialize analysis service
        analysis_service = AnalysisService(config_path)
        
        if verbose:
            print(f"Configuration loaded from: {config_path}")
            print(f"Output directory: {output_dir.absolute()}")
            print(f"Report format: {format_type}")
            print(f"Available modules: {analysis_service.get_available_modules()}")
            print(f"Analyzing: {input_file.name}\n")
        else:
            print(f"Analyzing: {input_file.name}")
        
        # Setup progress callback
        progress_callback = CLIProgressCallback(verbose)
        
        # Validate file first
        validation = analysis_service.validate_file(str(input_file))
        if not validation['is_valid']:
            print(f"File validation failed:")
            for error in validation['errors']:
                print(f"  Error: {error}")
            for warning in validation['warnings']:
                print(f"  Warning: {warning}")
            return 1
        
        # Show warnings if any
        if validation['warnings'] and verbose:
            for warning in validation['warnings']:
                print(f"  Warning: {warning}")
        
        # Run analysis
        start_time = datetime.now()
        results = analysis_service.analyze_file(str(input_file), progress_callback=progress_callback)
        analysis_time = (datetime.now() - start_time).total_seconds()
        
        # Generate timestamp for output files
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"{input_file.stem}_{timestamp}"
        
        # Determine output file extension
        ext_map = {'html': 'html', 'json': 'json', 'text': 'txt', 'markdown': 'md'}
        if format_type not in ext_map:
            print(f"Error: Unsupported format type: {format_type}")
            return 1
        
        output_file = output_dir / f"{base_name}.{ext_map[format_type]}"
        
        # Generate report using ReportService
        report_service = ReportService()
        try:
            final_output_path = report_service.generate_report(results, format_type, str(output_file))
            output_file = Path(final_output_path)
        except Exception as e:
            print(f"Error generating report: {e}")
            return 1
        
        # Success summary
        file_size_mb = output_file.stat().st_size / (1024 * 1024)
        print(f"\nAnalysis completed successfully!")
        print(f"Analysis time: {analysis_time:.1f}s")
        print(f"Report generated: {output_file}")
        print(f"Report size: {file_size_mb:.2f} MB")
        
        # Display summary metrics
        if verbose:
            _print_analysis_summary(results)
        
        return 0
        
    except KeyboardInterrupt:
        print("\nAnalysis cancelled by user")
        return 1
        
    except Exception as e:
        print(f"\nAnalysis failed: {e}")
        if verbose:
            import traceback
            print("\nDetailed error information:")
            traceback.print_exc()
        return 1


def _print_analysis_summary(results: dict):
    """Print detailed analysis summary for verbose mode"""
    try:
        print("\n" + "="*50)
        print("ANALYSIS SUMMARY")
        print("="*50)
        
        # File info
        file_info = results.get('file_info', {})
        print(f"File: {file_info.get('name', 'Unknown')}")
        print(f"Size: {file_info.get('size_mb', 0):.2f} MB")
        print(f"Sheets: {file_info.get('sheet_count', 0)}")
        
        # Analysis metadata
        metadata = results.get('analysis_metadata', {})
        print(f"Quality Score: {metadata.get('quality_score', 0):.1%}")
        print(f"Security Score: {metadata.get('security_score', 0):.1%}")
        print(f"Success Rate: {metadata.get('success_rate', 0):.1%}")
        
        # Module execution
        exec_summary = results.get('execution_summary', {})
        print(f"Modules: {exec_summary.get('successful_modules', 0)}/{exec_summary.get('total_modules', 0)} successful")
        
        # Data metrics
        module_results = results.get('module_results', {})
        data_profiler = module_results.get('data_profiler', {})
        if data_profiler:
            print(f"Total Cells: {data_profiler.get('total_cells', 0):,}")
            print(f"Data Density: {data_profiler.get('overall_data_density', 0):.1%}")
        
        print("="*50)
        
    except Exception as e:
        print(f"Warning: Could not display summary: {e}")


def validate_cli_environment() -> bool:
    """Validate that CLI environment has required dependencies"""
    try:
        import openpyxl
        import yaml
        return True
    except ImportError as e:
        print(f"Missing required dependencies: {e}")
        print("Install with: pip install -r requirements.txt")
        return False


if __name__ == "__main__":
    # Direct CLI usage for testing
    if len(sys.argv) < 2:
        print("Usage: python cli_runner.py <excel_file> [output_dir] [format]")
        sys.exit(1)
    
    file_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    format_type = sys.argv[3] if len(sys.argv) > 3 else 'html'
    
    sys.exit(run_cli_analysis(file_path, output_dir, format_type, verbose=True))
