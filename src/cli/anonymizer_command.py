#!/usr/bin/env python3
"""
CLI command handler for anonymizer functionality
"""

from pathlib import Path
from typing import Optional, List
import sys

def run_anonymizer(
    file_path: str,
    anonymize: bool = False,
    anonymize_columns: Optional[List[str]] = None,
    mapping_file: Optional[str] = None,
    mapping_format: str = 'json',
    reverse: Optional[str] = None,
    output_path: Optional[str] = None,
    verbose: bool = False
) -> int:
    """
    Execute anonymization or reversal based on CLI arguments
    
    Args:
        file_path: Path to Excel file
        anonymize: Whether to perform anonymization
        anonymize_columns: Specific columns to anonymize (format: "Sheet:Column")
        mapping_file: Path for mapping file output/input
        mapping_format: Format for mapping file ('json' or 'excel')
        reverse: Path to mapping file for reversal
        output_path: Output path for anonymized/restored file
        verbose: Enable verbose output
        
    Returns:
        Exit code (0 for success, 1 for error)
    """
    try:
        from services.anonymizer_service import AnonymizerService
        
        # Create progress callback for CLI
        def cli_progress(message: str, progress: float):
            if verbose:
                print(f"[{progress*100:.0f}%] {message}")
            elif "complete" in message.lower() or progress >= 1.0:
                print(f"✓ {message}")
            elif progress == 0.0 or "starting" in message.lower():
                print(f"→ {message}")
        
        service = AnonymizerService(progress_callback=cli_progress)
        
        # Handle reversal mode
        if reverse:
            print(f"Reversing anonymization using mappings: {reverse}")
            try:
                restored_path = service.reverse_anonymization(file_path, reverse, output_path)
                print(f"✓ Successfully restored file: {restored_path}")
                return 0
            except Exception as e:
                print(f"✗ Failed to restore file: {e}")
                return 1
        
        # Handle anonymization mode
        if anonymize:
            print(f"Anonymizing file: {file_path}")
            
            # Parse specific columns if provided
            columns_dict = None
            if anonymize_columns:
                columns_dict = {}
                for col_spec in anonymize_columns:
                    if ':' in col_spec:
                        sheet, col = col_spec.split(':', 1)
                        # Try to detect type from column name
                        col_type = 'name'  # default
                        col_lower = col.lower()
                        if 'company' in col_lower or 'org' in col_lower:
                            col_type = 'company'
                        elif 'email' in col_lower:
                            col_type = 'email'
                        elif 'phone' in col_lower or 'mobile' in col_lower:
                            col_type = 'phone'
                        elif 'address' in col_lower or 'street' in col_lower:
                            col_type = 'address'
                        
                        if sheet not in columns_dict:
                            columns_dict[sheet] = []
                        columns_dict[sheet].append((col, col_type))
                    else:
                        print(f"Warning: Invalid column format '{col_spec}'. Use 'Sheet:Column'")
            
            # Perform anonymization using service
            try:
                anon_path, map_path, stats = service.anonymize_excel_file(
                    file_path=file_path,
                    output_path=output_path,
                    mapping_path=mapping_file,
                    columns=columns_dict,
                    auto_detect=(columns_dict is None),
                    mapping_format=mapping_format
                )
                
                if anon_path and map_path:
                    print(f"✓ Anonymization complete!")
                    print(f"  Anonymized file: {anon_path}")
                    print(f"  Mapping file: {map_path}")
                    print(f"  Total values anonymized: {sum(stats.values())}")
                    return 0
                else:
                    print("✗ No columns were anonymized")
                    return 1
            except Exception as e:
                print(f"✗ Anonymization failed: {e}")
                return 1
        
        # If neither anonymize nor reverse, just analyze for sensitive columns
        if not anonymize and not reverse:
            print(f"Analyzing file for sensitive columns: {file_path}")
            try:
                sensitive = service.detect_sensitive_columns(file_path)
                
                if sensitive:
                    print("\nSensitive columns detected:")
                    for sheet, columns in sensitive.items():
                        print(f"\nSheet: {sheet}")
                        for col, col_type in columns:
                            print(f"  - Column {col}: {col_type}")
                    
                    print("\nTo anonymize these columns, run with --anonymize flag")
                else:
                    print("No sensitive columns detected")
                
                return 0
            except Exception as e:
                print(f"✗ Detection failed: {e}")
                return 1
            
    except ImportError as e:
        print(f"Error: Missing dependency - {e}")
        print("Install with: pip install faker")
        return 1
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return 1
    except Exception as e:
        print(f"Error: {e}")
        if verbose:
            import traceback
            traceback.print_exc()
        return 1


def add_anonymizer_arguments(parser):
    """
    Add anonymizer-specific arguments to an argument parser
    
    Args:
        parser: ArgumentParser instance to add arguments to
    """
    anon_group = parser.add_argument_group('Anonymization Options')
    
    anon_group.add_argument(
        '--anonymize',
        action='store_true',
        help='Enable data anonymization'
    )
    
    anon_group.add_argument(
        '--anonymize-columns',
        nargs='+',
        metavar='SHEET:COLUMN',
        help='Specific columns to anonymize (format: "Sheet1:A" or "Sheet1:Name")'
    )
    
    anon_group.add_argument(
        '--mapping-file',
        type=str,
        metavar='PATH',
        help='Path for mapping dictionary file (default: auto-generated)'
    )
    
    anon_group.add_argument(
        '--mapping-format',
        choices=['json', 'excel'],
        default='json',
        help='Format for mapping file (default: json)'
    )
    
    anon_group.add_argument(
        '--reverse',
        type=str,
        metavar='MAPPING_FILE',
        help='Reverse anonymization using the specified mapping file'
    )
    
    anon_group.add_argument(
        '--anonymized-output',
        type=str,
        metavar='PATH',
        help='Output path for anonymized file (default: adds _anonymized to filename)'
    )