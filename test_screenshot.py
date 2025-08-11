#!/usr/bin/env python3
"""
Test Script for Screenshot Feature
Demonstrates how to use the new screenshot functionality
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from core.analyzers.screenshot import ScreenshotUtility

def main():
    """Test screenshot functionality"""
    # Example usage
    test_file = "testing_files/sample.xlsx"
    
    if not Path(test_file).exists():
        print(f"Test file not found: {test_file}")
        print("Please provide a valid Excel file path")
        return 1
    
    print(f"Testing screenshot capture for: {test_file}")
    
    # Capture screenshots
    result = ScreenshotUtility.capture_excel_file(
        file_path=test_file,
        output_dir="output/test_screenshots",
        show_excel=True  # Show Excel for testing
    )
    
    if result.get('status') == 'success':
        print(f"✓ Successfully captured {result.get('total_sheets_captured', 0)} sheets")
        print(f"  Output directory: {result.get('output_directory')}")
        
        for screenshot in result.get('screenshots', []):
            print(f"  - {screenshot['sheet_name']}: {screenshot['file_path']}")
            print(f"    Size: {screenshot['width']}x{screenshot['height']}")
            
    else:
        print(f"✗ Screenshot capture failed:")
        print(f"  Status: {result.get('status')}")
        print(f"  Error: {result.get('error')}")

if __name__ == "__main__":
    sys.exit(main())