#!/usr/bin/env python3
"""Basic functionality test"""

import sys
from pathlib import Path
from src.main import main

def test_basic_functionality():
    """Test with a simple Excel file"""
    # This should be run with: python test_basic.py sample.xlsx
    if len(sys.argv) != 2:
        print("Usage: python test_basic.py <excel_file>")
        return False
  
    try:
        # Override sys.argv for main function
        original_argv = sys.argv
        sys.argv = ['src.main', sys.argv[1]]
      
        result = main()
      
        sys.argv = original_argv
      
        if result == 0:
            print("✅ Basic functionality test PASSED")
            return True
        else:
            print("❌ Basic functionality test FAILED")
            return False
          
    except Exception as e:
        print(f"❌ Test failed with exception: {e}")
        return False

if __name__ == "__main__":
    success = test_basic_functionality()
    sys.exit(0 if success else 1)