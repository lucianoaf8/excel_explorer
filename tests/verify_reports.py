#!/usr/bin/env python3
"""Verify report generation from command line"""
import sys
import argparse
from pathlib import Path
import re

def verify_markdown_report(file_path: str) -> bool:
    """Verify markdown report has all LLM enhancements"""
    path = Path(file_path)
    
    if not path.exists():
        print(f"❌ Report not found: {file_path}")
        return False
    
    content = path.read_text(encoding='utf-8')
    
    print(f"\n📄 Verifying: {path.name}")
    print("=" * 60)
    
    # Define checks with descriptions
    checks = [
        ("10 Row Samples", r"Sample Data \(First 10 Rows\)", "Task 1"),
        ("Sample Values Column", r"\| Sample Values \|", "Task 2"),
        ("Data Quality Section", r"##\s*[🔍📊]*\s*Data Quality Issues", "Task 3"),
        ("LLM Automation Guide", r"##\s*[🤖]*\s*LLM Automation Guide", "Task 4"),
        ("Quality Subsections", r"###\s*\d+\.\d+\s+\w+\s+Quality Issues", "Task 3 Detail"),
        ("LLM Step Instructions", r"###\s*Step\s+\d+:", "Task 4 Detail"),
        ("Column Analysis Table", r"\| Column \| Header \| Type \| Fill Rate \| Unique Values \| Sample Values \|", "Enhanced Table"),
        ("Error Details", r"-\s*\*\*Error Type\*\*:", "Quality Details"),
        ("Sheet Summaries", r"Sheet:\s+.+\s+\(\d+\s+rows\s+×\s+\d+\s+columns\)", "Sheet Info"),
        ("Automation Prompts", r"When analyzing this Excel file", "LLM Prompts")
    ]
    
    passed = 0
    failed = 0
    
    for name, pattern, category in checks:
        matches = len(re.findall(pattern, content, re.IGNORECASE | re.MULTILINE))
        if matches > 0:
            print(f"✅ [{category}] {name}: Found {matches} instance(s)")
            passed += 1
        else:
            print(f"❌ [{category}] {name}: NOT FOUND")
            failed += 1
            
            # Show context where it should appear
            if "10 Row" in name:
                sample_matches = re.findall(r"Sample Data.*?:", content)
                if sample_matches:
                    print(f"   Found instead: {sample_matches[0]}")
    
    print(f"\n📊 Summary: {passed}/{len(checks)} checks passed")
    
    if failed == 0:
        print("🎉 ALL LLM ENHANCEMENTS VERIFIED!")
    else:
        print(f"⚠️  {failed} ENHANCEMENTS MISSING!")
        
        # Additional debugging
        print("\n🔍 Debug Info:")
        print(f"  File size: {path.stat().st_size:,} bytes")
        print(f"  Line count: {len(content.splitlines()):,}")
        print(f"  Has 'First 3 Rows': {'First 3 Rows' in content}")
        print(f"  Has 'First 10 Rows': {'First 10 Rows' in content}")
    
    return failed == 0

def main():
    parser = argparse.ArgumentParser(description="Verify Excel Explorer markdown reports")
    parser.add_argument("report", nargs="?", help="Path to markdown report")
    parser.add_argument("--latest", action="store_true", help="Verify latest report")
    
    args = parser.parse_args()
    
    if args.latest or not args.report:
        # Find latest report
        report_dir = Path("output/reports")
        if not report_dir.exists():
            print("❌ Report directory not found!")
            return 1
            
        reports = list(report_dir.glob("*.md"))
        if not reports:
            print("❌ No markdown reports found!")
            return 1
            
        latest = max(reports, key=lambda p: p.stat().st_mtime)
        report_path = str(latest)
    else:
        report_path = args.report
    
    success = verify_markdown_report(report_path)
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())