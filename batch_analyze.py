#!/usr/bin/env python3
"""
Batch Excel Analyzer
Processes multiple Excel files and generates individual reports for each
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime
from typing import List, Tuple

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from cli.cli_runner import run_cli_analysis


def find_files(file_paths: List[str], directory: str = None) -> List[Path]:
    """
    Find Excel files from paths, glob patterns, or directory

    Args:
        file_paths: List of file paths or glob patterns
        directory: Optional directory to search in

    Returns:
        List of resolved file paths
    """
    files = []

    if directory:
        # Search directory for Excel files
        dir_path = Path(directory)
        if not dir_path.exists():
            print(f"Warning: Directory not found: {directory}")
            return []

        # Find all Excel files in directory
        for ext in ['*.xlsx', '*.xls', '*.xlsm']:
            files.extend(dir_path.glob(ext))

    # Process individual file paths/patterns
    for file_path in file_paths:
        path = Path(file_path)

        # Check if it's a glob pattern
        if '*' in file_path or '?' in file_path:
            # Expand glob pattern
            if path.parent.exists():
                files.extend(path.parent.glob(path.name))
            else:
                print(f"Warning: Pattern base path not found: {file_path}")
        else:
            # Direct file path
            if path.exists():
                files.append(path)
            else:
                print(f"Warning: File not found: {file_path}")

    # Remove duplicates and sort
    files = sorted(set(files))

    # Filter for Excel files only
    excel_extensions = {'.xlsx', '.xls', '.xlsm'}
    files = [f for f in files if f.suffix.lower() in excel_extensions]

    # Filter out Excel temporary files (starting with ~$)
    files = [f for f in files if not f.name.startswith('~$')]

    return files


def batch_analyze(
    files: List[Path],
    output_dir: str = None,
    format_type: str = 'html',
    config_path: str = 'config/config.yaml',
    verbose: bool = False,
    enable_screenshots: bool = False,
    continue_on_error: bool = True
) -> Tuple[List[Tuple[Path, str]], List[Tuple[Path, str]]]:
    """
    Analyze multiple Excel files in batch

    Args:
        files: List of file paths to analyze
        output_dir: Output directory for reports
        format_type: Report format (html, json, text, markdown)
        config_path: Configuration file path
        verbose: Enable detailed output
        enable_screenshots: Enable screenshot capture
        continue_on_error: Continue processing if a file fails

    Returns:
        Tuple of (successful_files, failed_files)
    """
    if not files:
        print("No files to process")
        return [], []

    # Setup output directory
    if not output_dir:
        output_dir = Path("reports/batch") / datetime.now().strftime("%Y%m%d_%H%M%S")
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    print("="*70)
    print(f"BATCH ANALYSIS - {len(files)} files")
    print("="*70)
    print(f"Output directory: {output_dir.absolute()}")
    print(f"Report format: {format_type}")
    print(f"Screenshots: {'enabled' if enable_screenshots else 'disabled'}")
    print("="*70)
    print()

    successful = []
    failed = []

    for idx, file_path in enumerate(files, 1):
        print(f"\n[{idx}/{len(files)}] Processing: {file_path.name}")
        print("-" * 70)

        try:
            # Run analysis for this file
            result = run_cli_analysis(
                file_path=str(file_path),
                output_dir=str(output_dir),
                format_type=format_type,
                config_path=config_path,
                verbose=verbose,
                enable_screenshots=enable_screenshots
            )

            if result == 0:
                successful.append((file_path, "Success"))
                print(f"✓ {file_path.name} - SUCCESS")
            else:
                failed.append((file_path, f"Exit code: {result}"))
                print(f"✗ {file_path.name} - FAILED (exit code: {result})")

                if not continue_on_error:
                    print("\nStopping batch processing due to error")
                    break

        except KeyboardInterrupt:
            print("\n\nBatch processing cancelled by user")
            failed.append((file_path, "Cancelled"))
            break

        except Exception as e:
            error_msg = str(e)
            failed.append((file_path, error_msg))
            print(f"✗ {file_path.name} - ERROR: {error_msg}")

            if verbose:
                import traceback
                traceback.print_exc()

            if not continue_on_error:
                print("\nStopping batch processing due to error")
                break

    # Print summary
    print("\n" + "="*70)
    print("BATCH ANALYSIS SUMMARY")
    print("="*70)
    print(f"Total files: {len(files)}")
    print(f"Successful: {len(successful)} ({len(successful)/len(files)*100:.1f}%)")
    print(f"Failed: {len(failed)} ({len(failed)/len(files)*100:.1f}%)")

    if successful:
        print(f"\nSuccessful files:")
        for file_path, _ in successful:
            print(f"  ✓ {file_path.name}")

    if failed:
        print(f"\nFailed files:")
        for file_path, reason in failed:
            print(f"  ✗ {file_path.name} - {reason}")

    print(f"\nReports saved to: {output_dir.absolute()}")
    print("="*70)

    return successful, failed


def main():
    """Main entry point for batch analyzer"""
    parser = argparse.ArgumentParser(
        description='Batch Excel File Analyzer - Process multiple files at once',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Analyze all files in a directory
  python batch_analyze.py --directory C:\\Data\\Reports

  # Analyze specific files
  python batch_analyze.py --files file1.xlsx file2.xlsx file3.xlsx

  # Use glob patterns
  python batch_analyze.py --files "C:\\Data\\*.xlsx"

  # Combine directory and specific files
  python batch_analyze.py --directory C:\\Data --files extra1.xlsx extra2.xlsx

  # With custom output and format
  python batch_analyze.py --directory C:\\Data --output C:\\Reports --format json

  # With screenshots (Windows only)
  python batch_analyze.py --directory C:\\Data --screenshots --verbose
        """
    )

    parser.add_argument('--files', nargs='+', default=[],
                       help='Excel files to analyze (supports glob patterns)')
    parser.add_argument('--directory', '--dir', '-d', type=str,
                       help='Directory containing Excel files to analyze')
    parser.add_argument('--output', '-o', type=str,
                       help='Output directory for reports (default: reports/batch/TIMESTAMP)')
    parser.add_argument('--format', '-f', choices=['html', 'json', 'text', 'markdown'],
                       default='html', help='Report format (default: html)')
    parser.add_argument('--config', type=str, default='config/config.yaml',
                       help='Configuration file path (default: config/config.yaml)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose output')
    parser.add_argument('--screenshots', action='store_true',
                       help='Enable screenshot capture (Windows only)')
    parser.add_argument('--stop-on-error', action='store_true',
                       help='Stop processing if any file fails (default: continue)')

    args = parser.parse_args()

    # Validate arguments
    if not args.files and not args.directory:
        parser.error("Either --files or --directory must be specified")

    # Find files to process
    print("Searching for Excel files...")
    files = find_files(args.files, args.directory)

    if not files:
        print("No Excel files found to process")
        return 1

    # Run batch analysis
    successful, failed = batch_analyze(
        files=files,
        output_dir=args.output,
        format_type=args.format,
        config_path=args.config,
        verbose=args.verbose,
        enable_screenshots=args.screenshots,
        continue_on_error=not args.stop_on_error
    )

    # Return exit code based on results
    if failed:
        return 1 if not successful else 2  # 1 = all failed, 2 = partial success
    return 0


if __name__ == "__main__":
    sys.exit(main())
