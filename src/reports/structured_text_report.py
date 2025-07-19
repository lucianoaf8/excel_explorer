"""
Enhanced Structured Text Report Generator for Excel Explorer
"""

from typing import Dict, Any, List
from datetime import datetime


class StructuredTextReportGenerator:
    """Generates structured text and markdown reports"""
    
    def __init__(self):
        self.report_sections: List[str] = []
        
    def generate_report(self, analysis_results: Dict[str, Any]) -> str:
        """Generate text report in traditional format"""
        self.report_sections = []
        
        try:
            # Header
            self._add_header(analysis_results)
            
            # Main sections
            self._add_file_info(analysis_results)
            self._add_quality_metrics(analysis_results)
            self._add_structure_analysis(analysis_results)
            self._add_data_analysis(analysis_results)
            self._add_security_analysis(analysis_results)
            self._add_recommendations(analysis_results)
            self._add_execution_summary(analysis_results)
            
            return '\n'.join(self.report_sections)
            
        except Exception as e:
            return f"Error generating text report: {str(e)}"
    
    def generate_markdown_report(self, analysis_results: Dict[str, Any], title: str = "Excel Analysis Report") -> str:
        """Generate markdown report with manual formatting"""
        try:
            # Build markdown manually
            lines = []
            lines.append(f"# {title}")
            lines.append("")
            
            # Add header info
            file_info = analysis_results.get('file_info', {})
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            lines.append(f"**File:** {file_info.get('name', 'Unknown')}")
            lines.append(f"**Generated:** {timestamp}")
            lines.append("")
            
            # File Information
            lines.append("## 📁 File Information")
            lines.append("")
            lines.append("| Property | Value |")
            lines.append("|----------|-------|")
            lines.append(f"| Name | {file_info.get('name', 'Unknown')} |")
            lines.append(f"| Size | {file_info.get('size_mb', 0):.2f} MB |")
            lines.append(f"| Sheets | {file_info.get('sheet_count', 0)} |")
            lines.append(f"| Created | {file_info.get('created', 'Unknown')} |")
            lines.append(f"| Modified | {file_info.get('modified', 'Unknown')} |")
            lines.append("")
            
            # Quality Metrics
            metadata = analysis_results.get('analysis_metadata', {})
            lines.append("## 📊 Quality Metrics")
            lines.append("")
            lines.append("| Metric | Value |")
            lines.append("|--------|-------|")
            lines.append(f"| Quality Score | {metadata.get('quality_score', 0):.1%} |")
            lines.append(f"| Security Score | {metadata.get('security_score', 0):.1f}/10 |")
            lines.append(f"| Success Rate | {metadata.get('success_rate', 0):.1%} |")
            lines.append("")
            
            # Security Analysis
            security = analysis_results.get('module_results', {}).get('security_inspector', {})
            lines.append("## 🔒 Security Analysis")
            lines.append("")
            lines.append(f"**Security Score:** {security.get('overall_score', 0):.1f}/10")
            lines.append(f"**Risk Level:** {security.get('risk_level', 'Unknown')}")
            lines.append("")
            
            threats = security.get('threats', [])
            if threats:
                lines.append("### Threats Detected")
                for threat in threats:
                    lines.append(f"- {threat}")
                lines.append("")
            else:
                lines.append("✅ No security threats detected.")
                lines.append("")
            
            # Recommendations
            recommendations = analysis_results.get('recommendations', [])
            lines.append("## 💡 Recommendations")
            lines.append("")
            if recommendations:
                for i, rec in enumerate(recommendations, 1):
                    lines.append(f"{i}. {rec}")
                lines.append("")
            else:
                lines.append("✅ No specific recommendations - file appears well-structured.")
                lines.append("")
            
            return '\n'.join(lines)
            
        except Exception as e:
            return f"# Error\n\nError generating markdown report: {str(e)}"
    
    def _add_header(self, results: Dict[str, Any]):
        """Add text report header"""
        self.report_sections.append("=" * 80)
        self.report_sections.append("EXCEL EXPLORER ANALYSIS REPORT")
        self.report_sections.append("=" * 80)
        
        file_info = results.get('file_info', {})
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        self.report_sections.append(f"File: {file_info.get('name', 'Unknown')}")
        self.report_sections.append(f"Generated: {timestamp}")
        self.report_sections.append("")
    
    def _add_file_info(self, results: Dict[str, Any]):
        """Add file information section"""
        self.report_sections.append("FILE INFORMATION")
        self.report_sections.append("=" * 60)
        
        file_info = results.get('file_info', {})
        
        info_items = [
            ("Name", file_info.get('name', 'Unknown')),
            ("Size", f"{file_info.get('size_mb', 0):.2f} MB"),
            ("Sheets", str(file_info.get('sheet_count', 0))),
            ("Created", file_info.get('created', 'Unknown')),
            ("Modified", file_info.get('modified', 'Unknown')),
            ("Excel Version", file_info.get('excel_version', 'Unknown'))
        ]
        
        for key, value in info_items:
            self.report_sections.append(f"{key:15}: {value}")
        
        self.report_sections.append("")
    
    def _add_quality_metrics(self, results: Dict[str, Any]):
        """Add quality metrics section"""
        self.report_sections.append("QUALITY METRICS")
        self.report_sections.append("=" * 60)
        
        metadata = results.get('analysis_metadata', {})
        
        metrics = [
            ("Quality Score", f"{metadata.get('quality_score', 0):.1%}"),
            ("Security Score", f"{metadata.get('security_score', 0):.1f}/10"),
            ("Success Rate", f"{metadata.get('success_rate', 0):.1%}"),
            ("Processing Time", f"{metadata.get('total_duration_seconds', 0):.1f}s")
        ]
        
        for metric, value in metrics:
            self.report_sections.append(f"{metric:15}: {value}")
        
        self.report_sections.append("")
    
    def _add_structure_analysis(self, results: Dict[str, Any]):
        """Add structure analysis section"""
        self.report_sections.append("STRUCTURE ANALYSIS")
        self.report_sections.append("=" * 60)
        
        structure = results.get('module_results', {}).get('structure_mapper', {})
        
        self.report_sections.append(f"Total Sheets: {structure.get('total_sheets', 0)}")
        self.report_sections.append(f"Visible Sheets: {len(structure.get('visible_sheets', []))}")
        self.report_sections.append(f"Hidden Sheets: {len(structure.get('hidden_sheets', []))}")
        self.report_sections.append(f"Named Ranges: {structure.get('named_ranges_count', 0)}")
        self.report_sections.append("")
    
    def _add_data_analysis(self, results: Dict[str, Any]):
        """Add data analysis section"""
        self.report_sections.append("DATA ANALYSIS")
        self.report_sections.append("=" * 60)
        
        data_profiler = results.get('module_results', {}).get('data_profiler', {})
        
        self.report_sections.append(f"Total Cells: {data_profiler.get('total_cells', 0):,}")
        self.report_sections.append(f"Data Density: {data_profiler.get('overall_data_density', 0):.1%}")
        self.report_sections.append(f"Quality Score: {data_profiler.get('data_quality_score', 0):.1%}")
        self.report_sections.append("")
    
    def _add_security_analysis(self, results: Dict[str, Any]):
        """Add security analysis section"""
        self.report_sections.append("SECURITY ANALYSIS")
        self.report_sections.append("=" * 60)
        
        security = results.get('module_results', {}).get('security_inspector', {})
        
        self.report_sections.append(f"Security Score: {security.get('overall_score', 0):.1f}/10")
        self.report_sections.append(f"Risk Level: {security.get('risk_level', 'Unknown')}")
        
        threats = security.get('threats', [])
        if threats:
            self.report_sections.append("\nThreats Detected:")
            for threat in threats:
                self.report_sections.append(f"  - {threat}")
        else:
            self.report_sections.append("No security threats detected.")
        
        self.report_sections.append("")
    
    def _add_recommendations(self, results: Dict[str, Any]):
        """Add recommendations section"""
        self.report_sections.append("RECOMMENDATIONS")
        self.report_sections.append("=" * 60)
        
        recommendations = results.get('recommendations', [])
        
        if recommendations:
            for i, rec in enumerate(recommendations, 1):
                self.report_sections.append(f"{i}. {rec}")
        else:
            self.report_sections.append("No specific recommendations - file appears well-structured.")
        
        self.report_sections.append("")
    
    def _add_execution_summary(self, results: Dict[str, Any]):
        """Add execution summary section"""
        self.report_sections.append("EXECUTION SUMMARY")
        self.report_sections.append("=" * 60)
        
        exec_summary = results.get('execution_summary', {})
        
        self.report_sections.append(f"Total Modules: {exec_summary.get('total_modules', 0)}")
        self.report_sections.append(f"Successful: {exec_summary.get('successful_modules', 0)}")
        self.report_sections.append(f"Failed: {exec_summary.get('failed_modules', 0)}")
        self.report_sections.append(f"Success Rate: {exec_summary.get('success_rate', 0):.1%}")
    
    def export_to_file(self, content: str, file_path: str, format_type: str = 'text'):
        """Export report content to file"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"Error exporting {format_type} report: {e}")
            return False


# Legacy function for backward compatibility
def generate_structured_text_report(analysis_results: Dict[str, Any]) -> str:
    """Generate structured text report (legacy function)"""
    generator = StructuredTextReportGenerator()
    return generator.generate_report(analysis_results)
