#!/usr/bin/env python3
import sys
import json
from pathlib import Path
from src.core.orchestrator import ExcelExplorer

def main():
    if len(sys.argv) != 2:
        print("Usage: python -m src.main <excel_file>")
        sys.exit(1)
  
    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(f"File not found: {file_path}")
        sys.exit(1)
  
    try:
        explorer = ExcelExplorer()
        results = explorer.analyze_file(str(file_path))
        print(json.dumps(results, indent=2, default=str))
        return 0
    except Exception as e:
        print(f"Analysis failed: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())