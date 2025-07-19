#!/usr/bin/env python3
"""
Report Consistency Validation Script
Ensures all report formats contain the same core data and metrics
"""

import json
import sys
from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime

from core import SimpleExcelAnalyzer
from reports import ReportGenerator, ReportValidator, ReportDataModel
from reports.structured_text_report import StructuredTextReportGenerator


class ReportConsistencyValidator:
    """Validates consistency across all report formats"""
    
    def __init__(self, config_path: str = 'config.yaml'):
        self.config_path = config_path
        self.temp_dir = Path("validation_temp")
        self.temp_dir.mkdir(exist_ok=True)
    
    def validate_report_consistency(self, file_path: str) -> Dict[str, Any]:
        """
        Generate all report formats and validate consistency
        
        Args:
            file_path: Path to Excel file to analyze
            
        Returns:
            Validation results with any discrepancies found
        """
        print(f"🔍 Validating report consistency for: {file_path}")
        
        try:
            # Run analysis once
            analyzer = SimpleExcelAnalyzer(self.config_path)
            results = analyzer.analyze(file_path)
            
            # Generate all report formats
            generated_reports = self._generate_all_formats(results, file_path)
            
            # Validate consistency using the base validator
            validation = ReportValidator.validate_consistency(results, generated_reports)
            
            # Add detailed format-specific checks
            detailed_validation = self._perform_detailed_validation(results, generated_reports)
            
            # Combine results
            validation.update(detailed_validation)
            
            return validation
            
        except Exception as e:
            return {
                'consistent': False,
                'error': str(e),
                'discrepancies': [f"Validation failed: {e}"]
            }
    
    def _generate_all_formats(self, results: Dict[str, Any], file_path: str) -> Dict[str, str]:
        """Generate all report formats"""
        base_name = Path(file_path).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        generated_reports = {}
        
        try:
            # HTML Report
            html_path = self.temp_dir / f"{base_name}_{timestamp}.html"
            generator = ReportGenerator()
            generator.generate_html_report(results, str(html_path))
            generated_reports['html'] = str(html_path)
            print("✅ HTML report generated")
            
            # JSON Report
            json_path = self.temp_dir / f"{base_name}_{timestamp}.json"
            generator.generate_json_report(results, str(json_path))
            generated_reports['json'] = str(json_path)
            print("✅ JSON report generated")
            
            # Text Report
            text_path = self.temp_dir / f"{base_name}_{timestamp}.txt"
            text_generator = StructuredTextReportGenerator()
            report_text = text_generator.generate_report(results)
            text_generator.export_to_file(report_text, str(text_path), 'text')
            generated_reports['text'] = str(text_path)
            print("✅ Text report generated")
            
            # Markdown Report
            md_path = self.temp_dir / f"{base_name}_{timestamp}.md"
            text_generator.export_to_file(report_text, str(md_path), 'markdown')
            generated_reports['markdown'] = str(md_path)
            print("✅ Markdown report generated")
            
        except Exception as e:
            print(f"❌ Error generating reports: {e}")
            raise
        
        return generated_reports
    
    def _perform_detailed_validation(self, results: Dict[str, Any], 
                                   generated_reports: Dict[str, str]) -> Dict[str, Any]:
        """Perform detailed validation of generated reports"""
        validation_results = {
            'format_validations': {},
            'size_metrics': {},
            'content_checks': {},
            'detailed_discrepancies': []
        }
        
        # Extract standardized data for comparison
        data_model = ReportDataModel(results)
        standardized_data = data_model.get_standardized_data()
        
        # Core metrics that must be identical across all formats
        core_metrics = {
            'file_name': standardized_data['file_summary']['name'],
            'file_size_mb': standardized_data['file_summary']['size_mb'],
            'sheet_count': standardized_data['file_summary']['sheet_count'],
            'quality_score': standardized_data['quality_metrics']['overall_quality_score'],
            'security_score': standardized_data['quality_metrics']['security_score'],
            'total_cells': standardized_data['quality_metrics']['total_cells']
        }
        
        validation_results['expected_core_metrics'] = core_metrics
        
        # Validate each format
        for format_name, report_path in generated_reports.items():
            try:
                format_validation = self._validate_format(
                    format_name, report_path, core_metrics
                )
                validation_results['format_validations'][format_name] = format_validation
                
                # Record file size
                file_size = Path(report_path).stat().st_size
                validation_results['size_metrics'][format_name] = {
                    'size_bytes': file_size,
                    'size_mb': file_size / (1024 * 1024)
                }
                
            except Exception as e:
                validation_results['format_validations'][format_name] = {
                    'valid': False,
                    'error': str(e)
                }
        
        return validation_results
    
    def _validate_format(self, format_name: str, report_path: str, 
                        expected_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Validate a specific report format"""
        validation = {
            'valid': True,
            'file_exists': Path(report_path).exists(),
            'metrics_found': {},
            'missing_metrics': [],
            'format_specific_checks': {}
        }
        
        if not validation['file_exists']:
            validation['valid'] = False
            return validation
        
        try:
            if format_name == 'json':
                validation.update(self._validate_json_format(report_path, expected_metrics))
            elif format_name == 'html':
                validation.update(self._validate_html_format(report_path, expected_metrics))
            elif format_name in ['text', 'markdown']:
                validation.update(self._validate_text_format(report_path, expected_metrics))
                
        except Exception as e:
            validation['valid'] = False
            validation['error'] = str(e)
        
        return validation
    
    def _validate_json_format(self, json_path: str, 
                            expected_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Validate JSON report format"""
        validation = {'format_specific_checks': {}}
        
        with open(json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        # Check JSON structure
        validation['format_specific_checks']['valid_json'] = True
        validation['format_specific_checks']['has_file_info'] = 'file_info' in json_data
        validation['format_specific_checks']['has_analysis_metadata'] = 'analysis_metadata' in json_data
        
        # Validate core metrics in JSON
        file_info = json_data.get('file_info', {})
        metadata = json_data.get('analysis_metadata', {})
        
        validation['metrics_found'] = {
            'file_name': file_info.get('name'),
            'file_size_mb': file_info.get('size_mb'),
            'sheet_count': file_info.get('sheet_count'),
            'quality_score': metadata.get('quality_score'),
            'security_score': metadata.get('security_score')
        }
        
        return validation
    
    def _validate_html_format(self, html_path: str, 
                            expected_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Validate HTML report format"""
        validation = {'format_specific_checks': {}}
        
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Basic HTML structure checks
        validation['format_specific_checks']['has_html_tags'] = '<html>' in html_content
        validation['format_specific_checks']['has_body'] = '<body>' in html_content
        validation['format_specific_checks']['has_title'] = '<title>' in html_content
        
        # Look for key metrics in HTML content (basic text search)
        validation['metrics_found'] = {}
        for metric_name, expected_value in expected_metrics.items():
            if str(expected_value) in html_content:
                validation['metrics_found'][metric_name] = expected_value
        
        return validation
    
    def _validate_text_format(self, text_path: str, 
                            expected_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Validate text/markdown report format"""
        validation = {'format_specific_checks': {}}
        
        with open(text_path, 'r', encoding='utf-8') as f:
            text_content = f.read()
        
        # Basic content checks
        validation['format_specific_checks']['has_content'] = len(text_content.strip()) > 0
        validation['format_specific_checks']['has_sections'] = '=' in text_content or '#' in text_content
        
        # Look for key metrics in text content
        validation['metrics_found'] = {}
        for metric_name, expected_value in expected_metrics.items():
            if str(expected_value) in text_content:
                validation['metrics_found'][metric_name] = expected_value
        
        return validation
    
    def cleanup(self):
        """Clean up temporary files"""
        try:
            import shutil
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
                print(f"🧹 Cleaned up temporary directory: {self.temp_dir}")
        except Exception as e:
            print(f"⚠️  Warning: Could not clean up temp directory: {e}")


def main():
    """CLI interface for validation script"""
    if len(sys.argv) < 2:
        print("Usage: python validate_reports.py <excel_file> [config_file]")
        print("Example: python validate_reports.py test.xlsx config.yaml")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    config_file = sys.argv[2] if len(sys.argv) > 2 else 'config.yaml'
    
    # Validate input file
    if not Path(excel_file).exists():
        print(f"❌ Error: File not found: {excel_file}")
        sys.exit(1)
    
    # Run validation
    validator = ReportConsistencyValidator(config_file)
    
    try:
        print("🔍 Starting report consistency validation...")
        results = validator.validate_report_consistency(excel_file)
        
        # Display results
        print("\n" + "="*60)
        print("📊 VALIDATION RESULTS")
        print("="*60)
        
        if results.get('consistent', False):
            print("✅ All reports are CONSISTENT")
        else:
            print("❌ INCONSISTENCIES FOUND")
        
        # Show core metrics
        if 'core_metrics' in results:
            print("\n📋 Core Metrics:")
            for metric, value in results['core_metrics'].items():
                print(f"  {metric}: {value}")
        
        # Show format validations
        if 'format_validations' in results:
            print("\n📄 Format Validations:")
            for format_name, validation in results['format_validations'].items():
                status = "✅" if validation.get('valid', False) else "❌"
                print(f"  {status} {format_name.upper()}: {'Valid' if validation.get('valid') else 'Invalid'}")
        
        # Show file sizes
        if 'size_metrics' in results:
            print("\n📏 Report Sizes:")
            for format_name, size_info in results['size_metrics'].items():
                print(f"  {format_name}: {size_info['size_mb']:.2f} MB")
        
        # Show discrepancies
        if results.get('discrepancies'):
            print("\n⚠️  Discrepancies Found:")
            for discrepancy in results['discrepancies']:
                print(f"  • {discrepancy}")
        
        print("="*60)
        
        # Return appropriate exit code
        sys.exit(0 if results.get('consistent', False) else 1)
        
    except KeyboardInterrupt:
        print("\n❌ Validation cancelled by user")
        sys.exit(1)
    finally:
        validator.cleanup()


if __name__ == "__main__":
    main()
