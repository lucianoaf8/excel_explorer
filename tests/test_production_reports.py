#!/usr/bin/env python3
"""Production report generation test script"""
import os
import sys
import subprocess
import time
from pathlib import Path

# Ensure no caching
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'

def test_production_reports():
    """Test report generation in production mode"""
    print("🧪 EXCEL EXPLORER PRODUCTION TEST")
    print("=" * 50)
    
    # Test file
    test_file = "testing_files/All Accounts - Forecasting.xlsx"
    if not Path(test_file).exists():
        print(f"❌ Test file not found: {test_file}")
        return False
    
    # Clean previous reports
    report_dir = Path("output/reports")
    if report_dir.exists():
        for f in report_dir.glob("All Accounts - Forecasting_*.md"):
            f.unlink()
            print(f"🗑️ Cleaned: {f.name}")
    
    print(f"\n📊 Testing CLI report generation...")
    print("-" * 50)
    
    # Run CLI test
    cmd = [
        sys.executable,
        "main.py",
        "--mode", "cli",
        "--file", test_file,
        "--format", "markdown",
        "--output", "./output/reports"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"❌ CLI test failed: {result.stderr}")
        return False
    
    print(result.stdout)
    
    # Find generated report
    reports = list(report_dir.glob("All Accounts - Forecasting_*.md"))
    if not reports:
        print("❌ No markdown report generated!")
        return False
    
    latest_report = max(reports, key=lambda p: p.stat().st_mtime)
    print(f"\n✅ Report generated: {latest_report}")
    
    # Verify LLM enhancements
    print(f"\n🔍 Verifying LLM enhancements...")
    print("-" * 50)
    
    content = latest_report.read_text(encoding='utf-8')
    
    checks = [
        ("Task 1: 10 Row Samples", "Sample Data (First 10 Rows)"),
        ("Task 2: Sample Values Column", "| Sample Values |"),
        ("Task 3: Data Quality Section", "## 🔍 Data Quality Issues"),
        ("Task 4: LLM Automation Guide", "## 🤖 LLM Automation Guide"),
    ]
    
    all_passed = True
    for task, pattern in checks:
        if pattern in content:
            print(f"✅ {task}: FOUND")
        else:
            print(f"❌ {task}: MISSING")
            all_passed = False
    
    if all_passed:
        print(f"\n🎉 ALL LLM ENHANCEMENTS VERIFIED!")
        print(f"📄 Report location: {latest_report.absolute()}")
    else:
        print(f"\n⚠️ SOME ENHANCEMENTS MISSING!")
        print(f"Check report: {latest_report.absolute()}")
    
    return all_passed

if __name__ == "__main__":
    success = test_production_reports()
    sys.exit(0 if success else 1)