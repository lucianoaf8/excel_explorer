#!/usr/bin/env python3
"""
Excel Explorer - Main Analysis Entry Point
Simple command-line interface for analyzing Excel files.
"""

import sys
import argparse
import json
from pathlib import Path
from typing import Optional

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.orchestrator import ExcelExplorer, create_cli_parser


def main():
    """Main entry point for Excel Explorer analysis"""
    try:
        # Parse command line arguments
        parser = create_cli_parser()
        parser.prog = "analyze.py"
        args = parser.parse_args()
        
        # Validate input file
        excel_file = Path(args.excel_file)
        if not excel_file.exists():
            print(f"Error: File not found: {excel_file}", file=sys.stderr)
            return 1
        
        if not excel_file.suffix.lower() in ['.xlsx', '.xlsm', '.xls']:
            print(f"Error: Unsupported file type: {excel_file.suffix}", file=sys.stderr)
            print("Supported formats: .xlsx, .xlsm, .xls", file=sys.stderr)
            return 1
        
        # Initialize explorer with configuration
        print(f"Initializing Excel Explorer...")
        explorer = ExcelExplorer(config_path=args.config, log_file=args.logfile)
        
        # Apply CLI overrides
        if args.memory_limit:
            explorer.analysis_config.max_memory_mb = args.memory_limit
            print(f"Memory limit set to {args.memory_limit}MB")
        
        if args.deep_analysis:
            explorer.analysis_config.deep_analysis = True
            print("Deep analysis mode enabled")
        
        if args.parallel:
            explorer.analysis_config.parallel_processing = True
            print("Parallel processing enabled")
        
        # Execute analysis
        print(f"Analyzing: {excel_file.name}")
        print("=" * 50)
        
        results = explorer.analyze_file(str(excel_file))
        
        # Output results
        if args.output:
            output_path = Path(args.output)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, default=str)
            
            print(f"\n✅ Analysis complete! Results saved to: {output_path}")
            
            # Print summary to console
            print("\n" + "=" * 50)
            print("ANALYSIS SUMMARY")
            print("=" * 50)
            
            if 'analysis_metadata' in results:
                metadata = results['analysis_metadata']
                print(f"Success Rate: {metadata.get('success_rate', 0):.1%}")
                print(f"Processing Time: {metadata.get('total_duration_seconds', 0):.1f}s")
                print(f"Quality Score: {metadata.get('quality_score', 0):.2f}")
                print(f"Modules Executed: {', '.join(metadata.get('modules_executed', []))}")
            
            if 'file_info' in results:
                file_info = results['file_info']
                print(f"File: {file_info.get('name', 'unknown')} ({file_info.get('size_mb', 0):.1f}MB)")
            
            if 'recommendations' in results and results['recommendations']:
                print(f"\nRecommendations:")
                for i, rec in enumerate(results['recommendations'][:3], 1):
                    print(f"  {i}. {rec}")
        
        else:
            # Print to stdout
            print(json.dumps(results, indent=2, default=str))
        
        return 0
        
    except KeyboardInterrupt:
        print("\n❌ Analysis interrupted by user", file=sys.stderr)
        return 130
    
    except Exception as e:
        print(f"❌ Analysis failed: {e}", file=sys.stderr)
        if args.logfile:
            print(f"Check log file for details: {args.logfile}", file=sys.stderr)
        return 1


def print_usage_examples():
    """Print usage examples"""
    print("""
Examples:
  python analyze.py sample.xlsx
  python analyze.py data.xlsx --output results.json
  python analyze.py large_file.xlsx --memory-limit 8192 --logfile analysis.log
  python analyze.py complex.xlsx --deep-analysis --config custom_config.yaml
    """)


if __name__ == "__main__":
    # Add help for examples
    if len(sys.argv) == 1:
        print("Excel Explorer - Comprehensive Excel File Analysis")
        print("Usage: python analyze.py <excel_file> [options]")
        print("\nFor detailed help: python analyze.py --help")
        print_usage_examples()
        sys.exit(0)
    
    sys.exit(main())
