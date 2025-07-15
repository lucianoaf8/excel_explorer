"""
Structured Text Report Generator for Excel Explorer GUI Integration

Generates AI-friendly, structured text reports that display directly in the GUI
while maintaining all data points from the comprehensive HTML analysis.
"""

from typing import Dict, Any, List, Optional, Union
from datetime import datetime
import json
import math


class StructuredTextReportGenerator:
    """Generates structured text reports optimized for GUI display and AI parsing"""
    
    def __init__(self):
        self.report_sections: List[str] = []
        self.current_indent: int = 0
        self.section_separators = {
            'major': '=' * 80,
            'section': '=' * 60,
            'subsection': '-' * 40,
            'minor': '.' * 30
        }
        
    def generate_report(self, analysis_results: Dict[str, Any]) -> str:
        """Generate complete structured text report with error handling"""
        self.report_sections = []
        
        try:
            # Report header
            self._safe_add_section("Report Header", 
                                 lambda: self._add_report_header(analysis_results))
            
            # Main sections with error handling
            self._safe_add_section("Executive Summary", 
                                 lambda: self._add_executive_summary(analysis_results))
            self._safe_add_section("File Information", 
                                 lambda: self._add_file_information(analysis_results))
            self._safe_add_section("Structure Analysis", 
                                 lambda: self._add_structure_analysis(analysis_results))
            self._safe_add_section("Data Quality Analysis", 
                                 lambda: self._add_data_quality_analysis(analysis_results))
            self._safe_add_section("Detailed Sheet Analysis", 
                                 lambda: self._add_detailed_sheet_analysis(analysis_results))
            self._safe_add_section("Security Analysis", 
                                 lambda: self._add_security_analysis(analysis_results))
            self._safe_add_section("Recommendations", 
                                 lambda: self._add_recommendations(analysis_results))
            self._safe_add_section("Module Execution Summary", 
                                 lambda: self._add_module_execution_summary(analysis_results))
            
        except Exception as e:
            self.report_sections.append(f"CRITICAL ERROR: Report generation failed - {str(e)}")
            self.report_sections.append("")
            self.report_sections.append("Partial report data available above.")
        
        return '\n'.join(self.report_sections)
    
    def _safe_add_section(self, section_name: str, section_func: callable) -> None:
        """Safely add a section with error handling"""
        try:
            section_func()
        except Exception as e:
            self.report_sections.append(f"ERROR in {section_name}: {str(e)}")
            self.report_sections.append("")
    
    def _add_report_header(self, results: Dict[str, Any]) -> None:
        """Add report header with basic file information"""
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        
        self.report_sections.append(self.section_separators['major'])
        self.report_sections.append('EXCEL ANALYSIS REPORT'.center(80))
        self.report_sections.append(self.section_separators['major'])
        
        # File summary line
        file_name = file_info.get('name', 'Unknown')
        file_size = file_info.get('size_mb', 0)
        generated_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        summary_line = f"File: {file_name} | Size: {file_size:.2f} MB | Generated: {generated_time}"
        self.report_sections.append(summary_line.center(80))
        self.report_sections.append('')
    
    def _add_executive_summary(self, results: Dict[str, Any]) -> None:
        """Add executive summary section"""
        self._add_section_header('EXECUTIVE SUMMARY', 'section')
        
        # Extract key metrics
        metadata = results.get('analysis_metadata', {})
        file_info = results.get('file_info', {})
        execution_summary = results.get('execution_summary', {})
        
        # Calculate success rate
        module_results = results.get('module_results', {})
        total_modules = len(module_results)
        successful_modules = sum(1 for result in module_results.values() 
                               if not isinstance(result, dict) or 'error' not in result)
        success_rate = (successful_modules / total_modules * 100) if total_modules > 0 else 0
        
        summary_items = [
            f"Analysis Success Rate: {success_rate:.1f}%",
            f"Quality Score: {metadata.get('quality_score', 0):.1f}%",
            f"Processing Time: {execution_summary.get('total_time', 0):.1f}s",
            f"Total Sheets: {len(results.get('sheets', []))}",
            f"Data Density: {metadata.get('data_density', 0):.1f}%"
        ]
        
        for item in summary_items:
            self.report_sections.append(f"• {item}")
        
        self.report_sections.append('')
    
    def _add_file_information(self, results: Dict[str, Any]) -> None:
        """Add detailed file information section"""
        self._add_section_header('FILE INFORMATION', 'section')
        
        file_info = results.get('file_info', {})
        
        # Basic file details
        file_details = [
            ('Name', file_info.get('name', 'Unknown')),
            ('Size', f"{file_info.get('size_mb', 0):.2f} MB"),
            ('Path', file_info.get('path', 'Unknown')),
            ('Created', file_info.get('created', 'Unknown')),
            ('Modified', file_info.get('modified', 'Unknown')),
            ('Excel Version', file_info.get('excel_version', 'Unknown')),
            ('Compression Ratio', f"{file_info.get('compression_ratio', 0):.1f}%"),
            ('File Signature', file_info.get('file_signature', 'Unknown'))
        ]
        
        for label, value in file_details:
            self.report_sections.append(f"  {label:<20}: {value}")
        
        self.report_sections.append('')
    
    def _add_structure_analysis(self, results: Dict[str, Any]) -> None:
        """Add structure analysis section with sheet inventory"""
        self._add_section_header('STRUCTURE ANALYSIS', 'section')
        
        # Try to get sheet data from different sources
        sheets = results.get('sheets', [])
        module_results = results.get('module_results', {})
        structure_data = module_results.get('structure_mapper', {})
        
        # If no sheets in results, try to extract from structure_data
        if not sheets and isinstance(structure_data, dict) and 'sheets' in structure_data:
            sheets = structure_data['sheets']
        
        # Sheet inventory table
        if sheets:
            self._add_subsection_header('Sheet Inventory')
            
            # Table headers
            headers = ['Sheet Name', 'Type', 'Rows', 'Cols', 'Status']
            table_data = []
            
            for sheet in sheets:
                sheet_name = sheet.get('name', 'Unknown')
                sheet_type = sheet.get('type', 'Worksheet')
                max_row = sheet.get('max_row', 0)
                max_col = sheet.get('max_column', 0)
                status = 'Active' if sheet.get('active', False) else 'Inactive'
                
                table_data.append([sheet_name, sheet_type, str(max_row), str(max_col), status])
            
            self._add_table(headers, table_data)
        else:
            # Try to get basic sheet info from file_info or other sources
            file_info = results.get('file_info', {})
            if 'sheets' in file_info:
                sheet_names = file_info['sheets']
                if sheet_names:
                    self._add_subsection_header('Sheet Names')
                    for sheet_name in sheet_names:
                        self.report_sections.append(f"  • {sheet_name}")
        
        # Workbook features
        features = structure_data.get('features', {})
        if features:
            self._add_subsection_header('Workbook Features')
            
            feature_items = [
                f"Named Ranges: {features.get('named_ranges', 0)}",
                f"Tables: {features.get('tables', 0)}",
                f"Charts: {features.get('charts', 0)}",
                f"Pivot Tables: {features.get('pivot_tables', 0)}",
                f"External Connections: {features.get('connections', 0)}",
                f"Protection: {'Yes' if features.get('protection', False) else 'No'}"
            ]
            
            for item in feature_items:
                self.report_sections.append(f"  • {item}")
        
        self.report_sections.append('')
    
    def _add_data_quality_analysis(self, results: Dict[str, Any]) -> None:
        """Add data quality analysis section"""
        self._add_section_header('DATA QUALITY ANALYSIS', 'section')
        
        metadata = results.get('analysis_metadata', {})
        module_results = results.get('module_results', {})
        data_profiler = module_results.get('data_profiler', {})
        
        # Quality metrics
        quality_metrics = [
            ('Overall Quality Score', f"{metadata.get('quality_score', 0):.1f}%"),
            ('Data Completeness', f"{metadata.get('completeness_score', 0):.1f}%"),
            ('Data Consistency', f"{metadata.get('consistency_score', 0):.1f}%"),
            ('Data Accuracy', f"{metadata.get('accuracy_score', 0):.1f}%"),
            ('Data Validity', f"{metadata.get('validity_score', 0):.1f}%")
        ]
        
        for label, value in quality_metrics:
            self.report_sections.append(f"  {label:<25}: {value}")
        
        # Data distribution
        if 'data_distribution' in data_profiler:
            self._add_subsection_header('Data Type Distribution')
            
            distribution = data_profiler['data_distribution']
            for data_type, count in distribution.items():
                percentage = (count / sum(distribution.values()) * 100) if distribution.values() else 0
                self.report_sections.append(f"  {data_type:<15}: {count:>6} ({percentage:.1f}%)")
        
        self.report_sections.append('')
    
    def _add_detailed_sheet_analysis(self, results: Dict[str, Any]) -> None:
        """Add detailed sheet-by-sheet analysis"""
        self._add_section_header('DETAILED SHEET ANALYSIS', 'section')
        
        # Try to get sheet data from different sources
        sheets = results.get('sheets', [])
        module_results = results.get('module_results', {})
        
        # If no sheets in results, try to extract from various modules
        if not sheets:
            # Try structure_mapper
            structure_data = module_results.get('structure_mapper', {})
            if isinstance(structure_data, dict) and 'sheets' in structure_data:
                sheets = structure_data['sheets']
            
            # Try data_profiler
            if not sheets:
                data_profiler = module_results.get('data_profiler', {})
                if isinstance(data_profiler, dict) and 'sheets' in data_profiler:
                    sheets = data_profiler['sheets']
        
        if not sheets:
            # Try to get basic sheet info from file_info
            file_info = results.get('file_info', {})
            if 'sheets' in file_info:
                sheet_names = file_info['sheets']
                if sheet_names:
                    self.report_sections.append("Available sheets:")
                    for sheet_name in sheet_names:
                        self.report_sections.append(f"  • {sheet_name}")
                        self.report_sections.append("    Details unavailable due to analysis limitations")
                    self.report_sections.append("")
                    return
            
            # No sheet data available at all
            self.report_sections.append("No detailed sheet analysis available.")
            self.report_sections.append("This may be due to analysis module failures or data access issues.")
            self.report_sections.append("")
            return
        
        for sheet in sheets:
            sheet_name = sheet.get('name', 'Unknown')
            self._add_subsection_header(f"Sheet: {sheet_name}")
            
            # Sheet dimensions
            max_row = sheet.get('max_row', 0)
            max_col = sheet.get('max_column', 0)
            self.report_sections.append(f"  Dimensions: {max_row} rows × {max_col} columns")
            
            # Column headers
            headers = sheet.get('headers', [])
            if headers:
                self.report_sections.append(f"  Headers: {' | '.join(headers[:10])}")  # First 10 headers
                if len(headers) > 10:
                    self.report_sections.append(f"  ... and {len(headers) - 10} more columns")
            
            # Sample data
            sample_data = sheet.get('sample_data', [])
            if sample_data:
                self.report_sections.append("  Sample Data:")
                for i, row in enumerate(sample_data[:3]):  # First 3 rows
                    row_str = ' | '.join(str(cell)[:20] for cell in row[:5])  # First 5 columns, max 20 chars
                    self.report_sections.append(f"    Row {i+1}: {row_str}")
            
            # Column analysis
            columns = sheet.get('columns', {})
            if columns:
                self.report_sections.append("  Column Analysis:")
                for col_letter, col_data in list(columns.items())[:10]:  # First 10 columns
                    col_name = col_data.get('name', col_letter)
                    col_type = col_data.get('type', 'Unknown')
                    fill_rate = col_data.get('fill_rate', 0)
                    unique_count = col_data.get('unique_count', 0)
                    
                    self.report_sections.append(
                        f"    {col_letter} ({col_name}): {col_type}, {fill_rate:.1f}% filled, {unique_count} unique"
                    )
            
            # Sheet properties
            properties = []
            if sheet.get('freeze_panes'):
                properties.append(f"Freeze Panes: {sheet.get('freeze_panes')}")
            if sheet.get('protection'):
                properties.append("Protection: Yes")
            else:
                properties.append("Protection: No")
            
            comment_count = sheet.get('comment_count', 0)
            hyperlink_count = sheet.get('hyperlink_count', 0)
            properties.extend([
                f"Comments: {comment_count}",
                f"Hyperlinks: {hyperlink_count}"
            ])
            
            self.report_sections.append("  Properties:")
            for prop in properties:
                self.report_sections.append(f"    {prop}")
            
            self.report_sections.append('')
    
    def _add_security_analysis(self, results: Dict[str, Any]) -> None:
        """Add security analysis section"""
        self._add_section_header('SECURITY ANALYSIS', 'section')
        
        module_results = results.get('module_results', {})
        security_data = module_results.get('security_inspector', {})
        
        if 'error' in security_data:
            self.report_sections.append("  Security analysis failed to complete")
            self.report_sections.append(f"  Error: {security_data['error']}")
        else:
            # Security score
            security_score = security_data.get('security_score', 0)
            self.report_sections.append(f"  Overall Security Score: {security_score:.1f}%")
            
            # Threat detection
            patterns_found = security_data.get('patterns_found', {})
            if patterns_found:
                self.report_sections.append("  Sensitive Data Detected:")
                for pattern_type, locations in patterns_found.items():
                    self.report_sections.append(f"    {pattern_type}: {len(locations)} instances")
            
            # Recommendations
            recommendations = security_data.get('recommendations', [])
            if recommendations:
                self.report_sections.append("  Security Recommendations:")
                for rec in recommendations[:5]:  # First 5 recommendations
                    self.report_sections.append(f"    • {rec}")
        
        self.report_sections.append('')
    
    def _add_recommendations(self, results: Dict[str, Any]) -> None:
        """Add recommendations section"""
        self._add_section_header('RECOMMENDATIONS', 'section')
        
        # Collect recommendations from various modules
        all_recommendations = []
        
        # Add quality-based recommendations
        metadata = results.get('analysis_metadata', {})
        quality_score = metadata.get('quality_score', 100)
        
        if quality_score < 80:
            all_recommendations.append(('High', 'Data Quality', 'Review data quality issues and implement data validation'))
        
        if quality_score < 60:
            all_recommendations.append(('High', 'Data Cleanup', 'Perform comprehensive data cleanup and standardization'))
        
        # Add security recommendations
        module_results = results.get('module_results', {})
        security_data = module_results.get('security_inspector', {})
        if security_data.get('security_score', 100) < 80:
            all_recommendations.append(('Medium', 'Security', 'Review and address security vulnerabilities'))
        
        # Add structure recommendations
        sheets = results.get('sheets', [])
        if len(sheets) > 10:
            all_recommendations.append(('Medium', 'Structure', 'Consider consolidating sheets to improve maintainability'))
        
        # Display recommendations by priority
        for priority in ['High', 'Medium', 'Low']:
            priority_recs = [rec for rec in all_recommendations if rec[0] == priority]
            if priority_recs:
                self.report_sections.append(f"  {priority} Priority:")
                for _, category, recommendation in priority_recs:
                    self.report_sections.append(f"    • [{category}] {recommendation}")
        
        self.report_sections.append('')
    
    def _add_module_execution_summary(self, results: Dict[str, Any]) -> None:
        """Add module execution summary"""
        self._add_section_header('MODULE EXECUTION SUMMARY', 'section')
        
        module_results = results.get('module_results', {})
        
        # Create status table
        headers = ['Module', 'Status', 'Details']
        table_data = []
        
        for module_name, result in module_results.items():
            if isinstance(result, dict) and 'error' in result:
                status = 'FAILED'
                details = result['error'][:50] + '...' if len(result['error']) > 50 else result['error']
            else:
                status = 'SUCCESS'
                details = 'Completed successfully'
            
            table_data.append([module_name, status, details])
        
        self._add_table(headers, table_data)
        self.report_sections.append('')
    
    def _add_section_header(self, title: str, level: str = 'section') -> None:
        """Add a formatted section header"""
        separator = self.section_separators.get(level, self.section_separators['section'])
        self.report_sections.append(separator)
        self.report_sections.append(f" {title}")
        self.report_sections.append(separator)
        self.report_sections.append('')
    
    def _add_subsection_header(self, title: str) -> None:
        """Add a formatted subsection header"""
        self.report_sections.append(f"--- {title} ---")
        self.report_sections.append('')
    
    def _add_table(self, headers: List[str], rows: List[List[str]]) -> None:
        """Add an ASCII-formatted table"""
        if not headers or not rows:
            return
        
        # Calculate column widths
        col_widths = []
        for i, header in enumerate(headers):
            max_width = len(header)
            for row in rows:
                if i < len(row):
                    max_width = max(max_width, len(str(row[i])))
            col_widths.append(min(max_width, 30))  # Max width of 30 characters
        
        # Create table separator
        separator = '+' + '+'.join('-' * (width + 2) for width in col_widths) + '+'
        
        # Add table header
        self.report_sections.append(separator)
        header_row = '|'
        for i, header in enumerate(headers):
            header_row += f' {header.ljust(col_widths[i])} |'
        self.report_sections.append(header_row)
        self.report_sections.append(separator)
        
        # Add table rows
        for row in rows:
            row_str = '|'
            for i, cell in enumerate(row):
                if i < len(col_widths):
                    cell_str = str(cell)[:col_widths[i]]  # Truncate if too long
                    row_str += f' {cell_str.ljust(col_widths[i])} |'
            self.report_sections.append(row_str)
        
        self.report_sections.append(separator)
        self.report_sections.append('')
    
    def _add_metric(self, label: str, value: Union[str, int, float], unit: str = "") -> None:
        """Add a formatted metric line"""
        value_str = f"{value}{unit}" if unit else str(value)
        self.report_sections.append(f"  {label:<25}: {value_str}")
    
    def _add_list(self, items: List[str], numbered: bool = False) -> None:
        """Add a formatted list"""
        for i, item in enumerate(items):
            if numbered:
                self.report_sections.append(f"  {i+1}. {item}")
            else:
                self.report_sections.append(f"  • {item}")
        self.report_sections.append('')
    
    def export_to_file(self, report_content: str, file_path: str, format_type: str = 'txt') -> None:
        """Export report to file in specified format"""
        if format_type.lower() == 'txt':
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
        elif format_type.lower() == 'md':
            # Convert to markdown format
            markdown_content = self._convert_to_markdown(report_content)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
        else:
            raise ValueError(f"Unsupported format: {format_type}")
    
    def _convert_to_markdown(self, content: str) -> str:
        """Convert structured text to markdown format"""
        lines = content.split('\n')
        markdown_lines = []
        
        for line in lines:
            # Convert section headers
            if line.startswith('=') and line.endswith('='):
                continue  # Skip separator lines
            elif line.strip() and not line.startswith(' '):
                if 'EXCEL ANALYSIS REPORT' in line:
                    markdown_lines.append(f"# {line.strip()}")
                else:
                    markdown_lines.append(f"## {line.strip()}")
            elif line.startswith('---') and line.endswith('---'):
                section_name = line.strip('- ')
                markdown_lines.append(f"### {section_name}")
            elif line.startswith('  •'):
                markdown_lines.append(f"- {line.strip('• ')}")
            else:
                markdown_lines.append(line)
        
        return '\n'.join(markdown_lines)