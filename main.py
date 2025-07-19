#!/usr/bin/env python3
"""
Excel Explorer v2.0 - Entry Point
Launch script for the reorganized package structure
"""

import sys
from pathlib import Path

# Add src to Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

# Import and run main
from src.main import main

if __name__ == "__main__":
    sys.exit(main())
