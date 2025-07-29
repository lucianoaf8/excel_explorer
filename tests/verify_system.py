#!/usr/bin/env python3
"""Verify Excel Explorer system configuration"""
import os
import sys
import importlib
import subprocess
from pathlib import Path

def verify_system():
    """Comprehensive system verification"""
    print("🔧 EXCEL EXPLORER SYSTEM VERIFICATION")
    print("=" * 60)
    
    # Python environment
    print("\n📍 Python Environment:")
    print(f"  Python: {sys.version}")
    print(f"  Executable: {sys.executable}")
    print(f"  PYTHONDONTWRITEBYTECODE: {os.environ.get('PYTHONDONTWRITEBYTECODE', 'Not set')}")
    print(f"  Cache disabled: {sys.dont_write_bytecode}")
    
    # Check for cache files
    print("\n🗃️ Cache Status:")
    pyc_files = list(Path('.').rglob('*.pyc'))
    pycache_dirs = list(Path('.').rglob('__pycache__'))
    print(f"  .pyc files: {len(pyc_files)}")
    print(f"  __pycache__ dirs: {len(pycache_dirs)}")
    
    # Module imports
    print("\n📦 Module Import Verification:")
    modules = [
        'src.reports.comprehensive_text_report',
        'src.gui.excel_explorer_gui',
        'src.core.analyzer'
    ]
    
    for module_name in modules:
        try:
            module = importlib.import_module(module_name)
            print(f"  ✅ {module_name}")
            
            # Check for LLM methods in comprehensive_text_report
            if module_name == 'src.reports.comprehensive_text_report':
                cls = getattr(module, 'ComprehensiveTextReportGenerator', None)
                if cls:
                    methods = ['generate_text_report', 'generate_markdown_report', 
                              '_generate_data_quality_issues', '_generate_llm_automation_guide']
                    for method in methods:
                        if hasattr(cls, method):
                            print(f"     ✓ {method}")
                        else:
                            print(f"     ✗ {method} MISSING!")
                            
        except ImportError as e:
            print(f"  ❌ {module_name}: {e}")
    
    # File structure
    print("\n📁 Report Directory Structure:")
    report_dir = Path("output/reports")
    if report_dir.exists():
        recent_reports = sorted(report_dir.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)[:3]
        print(f"  Directory exists: ✅")
        print(f"  Recent markdown reports:")
        for report in recent_reports:
            print(f"    - {report.name}")
    else:
        print(f"  Directory exists: ❌")
    
    print("\n✅ Verification complete!")

if __name__ == "__main__":
    # Set environment
    os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
    verify_system()