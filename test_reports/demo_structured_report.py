#!/usr/bin/env python3
"""
Demonstration of the new structured text report functionality
"""

import os
import sys
import tempfile
from pathlib import Path

# Add the project directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from structured_text_report import StructuredTextReportGenerator

def create_demo_results():
    """Create sample analysis results for demonstration"""
    from datetime import datetime
    
    return {
        'file_info': {
            'name': 'demo_file.xlsx',
            'size_mb': 2.45,
            'path': '/path/to/demo_file.xlsx',
            'created': '2025-07-15 10:00:00',
            'modified': '2025-07-15 10:30:00',
            'excel_version': '2016',
            'compression_ratio': 65.2,
            'file_signature': 'Excel Workbook',
            'sheets': ['Sales Data', 'Customer Info', 'Products']
        },
        'analysis_metadata': {
            'quality_score': 85.5,
            'completeness_score': 92.3,
            'consistency_score': 78.9,
            'accuracy_score': 88.7,
            'validity_score': 84.1,
            'data_density': 67.8,
            'total_duration_seconds': 3.2,
            'timestamp': datetime.now().timestamp()
        },
        'sheets': [
            {
                'name': 'Sales Data',
                'type': 'Worksheet',
                'max_row': 1000,
                'max_column': 8,
                'active': True,
                'headers': ['Date', 'Product', 'Customer', 'Quantity', 'Price', 'Total', 'Region', 'Salesperson'],
                'sample_data': [
                    ['2025-01-01', 'Laptop', 'Acme Corp', 5, 999.99, 4999.95, 'North', 'John Smith'],
                    ['2025-01-02', 'Mouse', 'Beta LLC', 20, 25.50, 510.00, 'South', 'Jane Doe'],
                    ['2025-01-03', 'Keyboard', 'Gamma Inc', 15, 75.00, 1125.00, 'East', 'Bob Johnson']
                ],
                'columns': {
                    'A': {'name': 'Date', 'type': 'Date', 'fill_rate': 98.5, 'unique_count': 365},
                    'B': {'name': 'Product', 'type': 'Text', 'fill_rate': 100.0, 'unique_count': 25},
                    'C': {'name': 'Customer', 'type': 'Text', 'fill_rate': 99.8, 'unique_count': 150},
                    'D': {'name': 'Quantity', 'type': 'Number', 'fill_rate': 100.0, 'unique_count': 50},
                    'E': {'name': 'Price', 'type': 'Currency', 'fill_rate': 100.0, 'unique_count': 45},
                    'F': {'name': 'Total', 'type': 'Currency', 'fill_rate': 100.0, 'unique_count': 800},
                    'G': {'name': 'Region', 'type': 'Text', 'fill_rate': 100.0, 'unique_count': 4},
                    'H': {'name': 'Salesperson', 'type': 'Text', 'fill_rate': 95.2, 'unique_count': 12}
                },
                'freeze_panes': 'A2',
                'protection': False,
                'comment_count': 3,
                'hyperlink_count': 0
            },
            {
                'name': 'Customer Info',
                'type': 'Worksheet',
                'max_row': 200,
                'max_column': 6,
                'active': False,
                'headers': ['ID', 'Company', 'Contact', 'Email', 'Phone', 'Address'],
                'sample_data': [
                    ['C001', 'Acme Corp', 'John Wilson', 'john@acme.com', '555-1234', '123 Main St'],
                    ['C002', 'Beta LLC', 'Sarah Davis', 'sarah@beta.com', '555-5678', '456 Oak Ave'],
                    ['C003', 'Gamma Inc', 'Mike Brown', 'mike@gamma.com', '555-9012', '789 Pine Rd']
                ],
                'columns': {
                    'A': {'name': 'ID', 'type': 'Text', 'fill_rate': 100.0, 'unique_count': 200},
                    'B': {'name': 'Company', 'type': 'Text', 'fill_rate': 100.0, 'unique_count': 150},
                    'C': {'name': 'Contact', 'type': 'Text', 'fill_rate': 98.5, 'unique_count': 195},
                    'D': {'name': 'Email', 'type': 'Text', 'fill_rate': 95.0, 'unique_count': 190},
                    'E': {'name': 'Phone', 'type': 'Text', 'fill_rate': 88.5, 'unique_count': 177},
                    'F': {'name': 'Address', 'type': 'Text', 'fill_rate': 92.0, 'unique_count': 184}
                },
                'freeze_panes': None,
                'protection': True,
                'comment_count': 0,
                'hyperlink_count': 5
            }
        ],
        'module_results': {
            'health_checker': {'status': 'success'},
            'structure_mapper': {
                'features': {
                    'named_ranges': 5,
                    'tables': 2,
                    'charts': 3,
                    'pivot_tables': 1,
                    'connections': 0,
                    'protection': True
                }
            },
            'data_profiler': {
                'data_distribution': {
                    'Text': 450,
                    'Number': 320,
                    'Date': 180,
                    'Currency': 150,
                    'Boolean': 25,
                    'Formula': 75
                }
            },
            'security_inspector': {
                'security_score': 75.5,
                'patterns_found': {
                    'email_addresses': ['john@acme.com', 'sarah@beta.com'],
                    'phone_numbers': ['555-1234', '555-5678'],
                    'credit_cards': []
                },
                'recommendations': [
                    'Consider masking or encrypting email addresses',
                    'Review phone number data classification',
                    'Implement data access controls'
                ]
            }
        },
        'execution_summary': {
            'total_modules': 9,
            'successful_modules': 8,
            'failed_modules': 1,
            'module_statuses': {
                'health_checker': 'success',
                'structure_mapper': 'success',
                'data_profiler': 'success',
                'formula_analyzer': 'success',
                'visual_cataloger': 'success',
                'security_inspector': 'success',
                'dependency_mapper': 'success',
                'relationship_analyzer': 'success',
                'performance_monitor': 'failed'
            },
            'total_time': 3.2
        }
    }

def main():
    """Demonstrate the structured text report functionality"""
    print("🚀 Structured Text Report Demo")
    print("=" * 50)
    
    # Create demo results
    demo_results = create_demo_results()
    
    # Generate structured text report
    print("📊 Generating structured text report...")
    report_generator = StructuredTextReportGenerator()
    report_text = report_generator.generate_report(demo_results)
    
    print(f"✅ Generated report ({len(report_text)} characters)\n")
    
    # Display first part of the report
    print("📄 Report Preview:")
    print("-" * 50)
    lines = report_text.split('\n')
    for i, line in enumerate(lines[:40]):  # Show first 40 lines
        print(line)
    
    if len(lines) > 40:
        print("...")
        print(f"[Report continues for {len(lines) - 40} more lines]")
    
    print("\n" + "-" * 50)
    
    # Test export functionality
    print("💾 Testing export functionality...")
    
    # Ensure reports directory exists
    reports_dir = Path("demo_reports")
    reports_dir.mkdir(exist_ok=True)
    
    # Export as text
    text_file = reports_dir / "demo_report.txt"
    report_generator.export_to_file(report_text, str(text_file), 'txt')
    print(f"✅ Text report exported to: {text_file}")
    
    # Export as markdown
    markdown_file = reports_dir / "demo_report.md"
    report_generator.export_to_file(report_text, str(markdown_file), 'md')
    print(f"✅ Markdown report exported to: {markdown_file}")
    
    print("\n🎉 Demo completed successfully!")
    print(f"📁 Check the '{reports_dir}' directory for exported reports")

if __name__ == "__main__":
    main()