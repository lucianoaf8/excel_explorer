"""
Comprehensive Text and Markdown Report Generator
Provides the same detailed information as the HTML report
"""

from typing import Dict, Any, List, Optional
from datetime import datetime
from pathlib import Path


class ComprehensiveTextReportGenerator:
    """Generates comprehensive text and markdown reports with full detail"""
    
    def __init__(self):
        self.width = 80  # Terminal width for text reports
        
    def generate_text_report(self, analysis_results: Dict[str, Any], output_path: str) -> str:
        """Generate comprehensive text report"""
        try:
            content = self._create_text_content(analysis_results)
            
            # Write to file
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            return str(output_file)
            
        except Exception as e:
            print(f"Error generating text report: {e}")
            return self._create_fallback_text_report(analysis_results, output_path)
    
    def generate_markdown_report(self, analysis_results: Dict[str, Any], output_path: str) -> str:
        """Generate comprehensive markdown report"""
        try:
            content = self._create_markdown_content(analysis_results)
            
            # Write to file
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            return str(output_file)
            
        except Exception as e:
            print(f"Error generating markdown report: {e}")
            return self._create_fallback_markdown_report(analysis_results, output_path)
    
    def _create_text_content(self, results: Dict[str, Any]) -> str:
        """Create comprehensive text report content"""
        lines = []
        
        # Extract all data sections
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        modules = results.get('module_results', {})
        exec_summary = results.get('execution_summary', {})
        recommendations = results.get('recommendations', [])
        
        # Header
        lines.append("=" * self.width)
        lines.append("EXCEL EXPLORER COMPREHENSIVE ANALYSIS REPORT".center(self.width))
        lines.append("=" * self.width)
        lines.append("")
        
        timestamp = datetime.fromtimestamp(metadata.get('timestamp', datetime.now().timestamp()))
        lines.append(f"File: {file_info.get('name', 'Unknown')}")
        lines.append(f"Generated: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Size: {file_info.get('size_mb', 0):.2f} MB")
        lines.append("")
        
        # 1. OVERVIEW SECTION
        lines.append("=" * self.width)
        lines.append("1. OVERVIEW")
        lines.append("=" * self.width)
        lines.extend(self._create_text_overview(file_info, metadata, modules, exec_summary))
        
        # 2. STRUCTURE ANALYSIS
        lines.append("=" * self.width)
        lines.append("2. STRUCTURE ANALYSIS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_structure(modules.get('structure_mapper', {})))
        
        # 3. DATA QUALITY
        lines.append("=" * self.width)
        lines.append("3. DATA QUALITY ANALYSIS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_data_quality(modules.get('data_profiler', {})))
        
        # 4. SHEET-BY-SHEET ANALYSIS
        lines.append("=" * self.width)
        lines.append("4. DETAILED SHEET ANALYSIS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_sheet_analysis(modules.get('data_profiler', {})))
        
        # 5. CROSS-SHEET RELATIONSHIPS
        lines.append("=" * self.width)
        lines.append("5. CROSS-SHEET RELATIONSHIPS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_relationships(modules.get('relationship_analyzer', {})))
        
        # 6. SECURITY ANALYSIS
        lines.append("=" * self.width)
        lines.append("6. SECURITY ANALYSIS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_security(modules.get('security_inspector', {})))
        
        # 7. RECOMMENDATIONS
        lines.append("=" * self.width)
        lines.append("7. RECOMMENDATIONS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_recommendations(recommendations))
        
        # 8. MODULE EXECUTION STATUS
        lines.append("=" * self.width)
        lines.append("8. MODULE EXECUTION STATUS")
        lines.append("=" * self.width)
        lines.extend(self._create_text_execution_status(exec_summary))
        
        return '\n'.join(lines)
    
    def _create_text_overview(self, file_info: Dict, metadata: Dict, modules: Dict, exec_summary: Dict) -> List[str]:
        """Create overview section for text report"""
        lines = []
        
        # Analysis Summary
        lines.append("Analysis Summary:")
        lines.append("-" * 40)
        lines.append(f"  Success Rate:     {exec_summary.get('success_rate', 0) * 100:.1f}%")
        lines.append(f"  Quality Score:    {metadata.get('quality_score', 0) * 100:.1f}%")
        lines.append(f"  Security Score:   {metadata.get('security_score', 0) * 10:.1f}/10")
        lines.append(f"  Processing Time:  {metadata.get('total_duration_seconds', 0):.1f}s")
        lines.append(f"  Modules Executed: {exec_summary.get('successful_modules', 0)}/{exec_summary.get('total_modules', 0)}")
        lines.append("")
        
        # File Structure
        structure = modules.get('structure_mapper', {})
        lines.append("File Structure:")
        lines.append("-" * 40)
        lines.append(f"  Total Sheets:     {structure.get('total_sheets', 0)}")
        lines.append(f"  Visible Sheets:   {len(structure.get('visible_sheets', []))}")
        lines.append(f"  Hidden Sheets:    {len(structure.get('hidden_sheets', []))}")
        lines.append(f"  Named Ranges:     {structure.get('named_ranges_count', 0)}")
        lines.append(f"  Tables:           {structure.get('table_count', 0)}")
        lines.append("")
        
        # Data Metrics
        data_profiler = modules.get('data_profiler', {})
        lines.append("Data Metrics:")
        lines.append("-" * 40)
        lines.append(f"  Total Cells:      {data_profiler.get('total_cells', 0):,}")
        lines.append(f"  Data Cells:       {data_profiler.get('total_data_cells', 0):,}")
        lines.append(f"  Empty Cells:      {data_profiler.get('total_cells', 0) - data_profiler.get('total_data_cells', 0):,}")
        lines.append(f"  Data Density:     {data_profiler.get('overall_data_density', 0) * 100:.1f}%")
        lines.append("")
        
        # Formulas & Features
        formulas = modules.get('formula_analyzer', {})
        visuals = modules.get('visual_cataloger', {})
        lines.append("Formulas & Features:")
        lines.append("-" * 40)
        lines.append(f"  Total Formulas:   {formulas.get('total_formulas', 0)}")
        lines.append(f"  Complex Formulas: {len(formulas.get('complex_formulas', []))}")
        lines.append(f"  External Refs:    {'Yes' if formulas.get('has_external_refs', False) else 'No'}")
        lines.append(f"  Charts:           {visuals.get('total_charts', 0)}")
        lines.append(f"  Images:           {visuals.get('total_images', 0)}")
        lines.append("")
        
        return lines
    
    def _create_text_structure(self, structure: Dict) -> List[str]:
        """Create structure analysis section for text report"""
        lines = []
        
        # Sheet Overview Table
        lines.append("Sheet Overview:")
        lines.append("-" * self.width)
        
        # Table header
        header = f"{'Sheet Name':<30} {'Rows':>10} {'Columns':>10} {'Status':>15}"
        lines.append(header)
        lines.append("-" * self.width)
        
        # Sheet details
        for sheet in structure.get('sheet_details', []):
            name = sheet.get('name', 'Unknown')[:30]
            rows = sheet.get('max_row', 0)
            cols = sheet.get('max_column', 0)
            status = sheet.get('status', 'Unknown')
            
            line = f"{name:<30} {rows:>10,} {cols:>10} {status:>15}"
            lines.append(line)
        
        lines.append("")
        
        # Workbook Features
        features = structure.get('workbook_features', {})
        lines.append("Workbook Features:")
        lines.append("-" * 40)
        lines.append(f"  Protection:       {'Yes' if structure.get('protection_info', {}).get('has_protection', False) else 'No'}")
        lines.append(f"  Macros:           {'Present' if features.get('has_macros', False) else 'Not Present'}")
        lines.append(f"  External Connections: {features.get('has_external_connections', 0)}")
        lines.append(f"  Pivot Tables:     {features.get('has_pivot_tables', 0)}")
        lines.append(f"  Data Validation:  {features.get('data_validation_rules', 0)} rules")
        lines.append(f"  Conditional Formatting: {features.get('conditional_formatting_rules', 0)} rules")
        lines.append("")
        
        return lines
    
    def _create_text_data_quality(self, data_profiler: Dict) -> List[str]:
        """Create data quality section for text report"""
        lines = []
        
        # Overall metrics
        overall_metrics = data_profiler.get('overall_metrics', {})
        lines.append("Overall Data Quality:")
        lines.append("-" * 40)
        lines.append(f"  Data Quality Score: {overall_metrics.get('quality_score', 0) * 100:.1f}%")
        lines.append(f"  Data Density:       {overall_metrics.get('data_density', 0) * 100:.1f}%")
        lines.append(f"  Data Variety Score: {overall_metrics.get('data_variety_score', 0) * 100:.1f}%")
        lines.append("")
        
        # Data type distribution
        type_dist = data_profiler.get('data_type_distribution', {})
        total_values = sum(type_dist.values()) if type_dist else 1
        
        lines.append("Data Type Distribution:")
        lines.append("-" * 40)
        for data_type, count in type_dist.items():
            percentage = (count / total_values) * 100
            lines.append(f"  {data_type.capitalize():<10}: {percentage:>6.1f}% ({count:,} cells)")
        lines.append("")
        
        return lines
    
    def _create_text_sheet_analysis(self, data_profiler: Dict) -> List[str]:
        """Create detailed sheet analysis section for text report"""
        lines = []
        
        sheet_analysis = data_profiler.get('sheet_analysis', {})
        
        for sheet_name, sheet_data in sheet_analysis.items():
            lines.append(f"\nSheet: {sheet_name}")
            lines.append("=" * 60)
            
            # Basic info
            lines.append(f"Dimensions: {sheet_data.get('dimensions', 'N/A')}")
            lines.append(f"Data Cells: {sheet_data.get('estimated_data_cells', 0):,}")
            lines.append(f"Data Density: {sheet_data.get('data_density', 0) * 100:.1f}%")
            lines.append("")
            
            # Headers preview
            columns = sheet_data.get('columns', [])
            if columns:
                headers = [col.get('header', f"Column {col.get('letter', '')}") for col in columns[:8]]
                lines.append("Headers Preview:")
                lines.append("  " + " | ".join(headers))
                lines.append("")
                
                # Sample data (first 3 rows)
                lines.append("Sample Data (First 3 Rows):")
                for row_idx in range(3):
                    row_data = []
                    for col in columns[:5]:  # First 5 columns
                        sample_values = col.get('sample_values', [])
                        if row_idx < len(sample_values):
                            value = str(sample_values[row_idx])[:15]  # Truncate
                            row_data.append(value)
                        else:
                            row_data.append("")
                    
                    if any(row_data):
                        lines.append("  " + " | ".join(f"{v:<15}" for v in row_data))
                lines.append("")
                
                # Column analysis table
                lines.append("Column Analysis:")
                lines.append(f"  {'Column':<8} {'Header':<20} {'Type':<10} {'Fill Rate':>10} {'Unique':>10}")
                lines.append("  " + "-" * 60)
                
                for col in columns[:10]:  # First 10 columns
                    letter = col.get('letter', '')
                    header = col.get('header', f'Column {letter}')[:20]  # Truncate long headers
                    data_type = col.get('data_type', 'unknown')
                    fill_rate = col.get('fill_rate', 0) * 100
                    unique = col.get('unique_values', 0)
                    
                    lines.append(f"  {letter:<8} {header:<20} {data_type:<10} {fill_rate:>9.1f}% {unique:>10,}")
            
            lines.append("")
        
        return lines
    
    def _create_text_relationships(self, relationships: Dict) -> List[str]:
        """Create cross-sheet relationships section for text report"""
        lines = []
        
        if relationships.get('skipped', False):
            lines.append("Cross-sheet analysis was skipped.")
            return lines
        
        relationships_found = relationships.get('relationships_found', [])
        
        if not relationships_found:
            lines.append("No cross-sheet relationships found.")
            return lines
        
        # Table header
        lines.append(f"{'Type':<15} {'Source Sheet':<25} {'Target Sheet':<25} {'Match':<8}")
        lines.append("-" * self.width)
        
        # Relationships
        for rel in relationships_found:
            rel_type = rel.get('relationship_type', 'unknown')[:15]
            source = rel.get('source_sheet', '')[:25]
            target = rel.get('target_sheet', '')[:25]
            match_rate = rel.get('match_rate', 0) * 100
            
            lines.append(f"{rel_type:<15} {source:<25} {target:<25} {match_rate:>6.1f}%")
            
            # Key columns (indented)
            key_columns = rel.get('key_columns', [])
            if key_columns:
                key_str = ', '.join(key_columns)
                # Wrap long key column lists
                wrapped = self._wrap_text(f"Key columns: {key_str}", self.width - 4, indent=2)
                lines.extend(wrapped)
            lines.append("")
        
        return lines
    
    def _create_text_security(self, security: Dict) -> List[str]:
        """Create security analysis section for text report"""
        lines = []
        
        if not security:
            lines.append("Security analysis not available.")
            return lines
        
        lines.append(f"Overall Security Score: {security.get('overall_score', 0):.1f}/10")
        lines.append(f"Risk Level: {security.get('risk_level', 'Unknown')}")
        lines.append("")
        
        # Threats
        threats = security.get('threats', [])
        if threats:
            lines.append("Threats Detected:")
            for threat in threats:
                lines.append(f"  - {threat}")
            lines.append("")
        else:
            lines.append("✓ No security threats detected")
            lines.append("")
        
        # Security recommendations
        security_recs = security.get('recommendations', [])
        if security_recs:
            lines.append("Security Recommendations:")
            for rec in security_recs:
                lines.append(f"  - {rec}")
            lines.append("")
        
        return lines
    
    def _create_text_recommendations(self, recommendations: List[str]) -> List[str]:
        """Create recommendations section for text report"""
        lines = []
        
        if recommendations:
            for i, rec in enumerate(recommendations, 1):
                lines.append(f"{i}. {rec}")
        else:
            lines.append("No specific recommendations - file appears well-structured.")
        
        lines.append("")
        return lines
    
    def _create_text_execution_status(self, exec_summary: Dict) -> List[str]:
        """Create module execution status section for text report"""
        lines = []
        
        lines.append(f"Total Modules:    {exec_summary.get('total_modules', 0)}")
        lines.append(f"Successful:       {exec_summary.get('successful_modules', 0)}")
        lines.append(f"Failed:           {exec_summary.get('failed_modules', 0)}")
        lines.append(f"Success Rate:     {exec_summary.get('success_rate', 0) * 100:.1f}%")
        lines.append("")
        
        # Module status table
        module_statuses = exec_summary.get('module_statuses', {})
        if module_statuses:
            lines.append("Module Status Details:")
            lines.append(f"{'Module':<30} {'Status':<15}")
            lines.append("-" * 45)
            
            for module, status in module_statuses.items():
                module_name = module.replace('_', ' ').title()[:30]
                lines.append(f"{module_name:<30} {status.title():<15}")
        
        return lines
    
    def _wrap_text(self, text: str, width: int, indent: int = 0) -> List[str]:
        """Wrap text to specified width with optional indentation"""
        words = text.split()
        lines = []
        current_line = []
        current_length = indent
        
        for word in words:
            if current_length + len(word) + 1 > width:
                lines.append(" " * indent + " ".join(current_line))
                current_line = [word]
                current_length = indent + len(word)
            else:
                current_line.append(word)
                current_length += len(word) + 1
        
        if current_line:
            lines.append(" " * indent + " ".join(current_line))
        
        return lines
    
    def _create_markdown_content(self, results: Dict[str, Any]) -> str:
        """Create comprehensive markdown report content"""
        lines = []
        
        # Extract all data sections
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        modules = results.get('module_results', {})
        exec_summary = results.get('execution_summary', {})
        recommendations = results.get('recommendations', [])
        
        # Title and header
        lines.append("# 📊 Excel Analysis Report")
        lines.append("")
        
        timestamp = datetime.fromtimestamp(metadata.get('timestamp', datetime.now().timestamp()))
        lines.append(f"**File:** {file_info.get('name', 'Unknown')}")
        lines.append(f"**Generated:** {timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"**Size:** {file_info.get('size_mb', 0):.2f} MB")
        lines.append("")
        
        # Table of Contents
        lines.append("## 📋 Table of Contents")
        lines.append("")
        lines.append("1. [Overview](#overview)")
        lines.append("2. [Structure Analysis](#structure-analysis)")
        lines.append("3. [Data Quality](#data-quality)")
        lines.append("4. [Sheet Analysis](#sheet-analysis)")
        lines.append("5. [Cross-Sheet Relationships](#cross-sheet-relationships)")
        lines.append("6. [Security Analysis](#security-analysis)")
        lines.append("7. [Recommendations](#recommendations)")
        lines.append("8. [Module Execution Status](#module-execution-status)")
        lines.append("")
        
        # 1. Overview
        lines.append("## 📈 Overview")
        lines.append("")
        lines.extend(self._create_markdown_overview(file_info, metadata, modules, exec_summary))
        
        # 2. Structure Analysis
        lines.append("## 🏗️ Structure Analysis")
        lines.append("")
        lines.extend(self._create_markdown_structure(modules.get('structure_mapper', {})))
        
        # 3. Data Quality
        lines.append("## 🔍 Data Quality")
        lines.append("")
        lines.extend(self._create_markdown_data_quality(modules.get('data_profiler', {})))
        
        # 4. Sheet Analysis
        lines.append("## 📄 Sheet Analysis")
        lines.append("")
        lines.extend(self._create_markdown_sheet_analysis(modules.get('data_profiler', {})))
        
        # 5. Cross-Sheet Relationships
        lines.append("## 🔗 Cross-Sheet Relationships")
        lines.append("")
        lines.extend(self._create_markdown_relationships(modules.get('relationship_analyzer', {})))
        
        # 6. Security Analysis
        lines.append("## 🔒 Security Analysis")
        lines.append("")
        lines.extend(self._create_markdown_security(modules.get('security_inspector', {})))
        
        # 7. Recommendations
        lines.append("## 💡 Recommendations")
        lines.append("")
        lines.extend(self._create_markdown_recommendations(recommendations))
        
        # 8. Module Execution Status
        lines.append("## 🔧 Module Execution Status")
        lines.append("")
        lines.extend(self._create_markdown_execution_status(exec_summary))
        
        return '\n'.join(lines)
    
    def _create_markdown_overview(self, file_info: Dict, metadata: Dict, modules: Dict, exec_summary: Dict) -> List[str]:
        """Create overview section for markdown report"""
        lines = []
        
        # Create a summary table
        lines.append("### Analysis Summary")
        lines.append("")
        lines.append("| Metric | Value |")
        lines.append("|--------|-------|")
        lines.append(f"| Success Rate | {exec_summary.get('success_rate', 0) * 100:.1f}% |")
        lines.append(f"| Quality Score | {metadata.get('quality_score', 0) * 100:.1f}% |")
        lines.append(f"| Security Score | {metadata.get('security_score', 0) * 10:.1f}/10 |")
        lines.append(f"| Processing Time | {metadata.get('total_duration_seconds', 0):.1f}s |")
        lines.append(f"| Modules Executed | {exec_summary.get('successful_modules', 0)}/{exec_summary.get('total_modules', 0)} |")
        lines.append("")
        
        # File structure metrics
        structure = modules.get('structure_mapper', {})
        lines.append("### File Structure")
        lines.append("")
        lines.append("| Property | Value |")
        lines.append("|----------|-------|")
        lines.append(f"| Total Sheets | {structure.get('total_sheets', 0)} |")
        lines.append(f"| Visible Sheets | {len(structure.get('visible_sheets', []))} |")
        lines.append(f"| Hidden Sheets | {len(structure.get('hidden_sheets', []))} |")
        lines.append(f"| Named Ranges | {structure.get('named_ranges_count', 0)} |")
        lines.append(f"| Tables | {structure.get('table_count', 0)} |")
        lines.append("")
        
        # Data metrics
        data_profiler = modules.get('data_profiler', {})
        lines.append("### Data Metrics")
        lines.append("")
        lines.append("| Metric | Value |")
        lines.append("|--------|-------|")
        lines.append(f"| Total Cells | {data_profiler.get('total_cells', 0):,} |")
        lines.append(f"| Data Cells | {data_profiler.get('total_data_cells', 0):,} |")
        lines.append(f"| Empty Cells | {data_profiler.get('total_cells', 0) - data_profiler.get('total_data_cells', 0):,} |")
        lines.append(f"| Data Density | {data_profiler.get('overall_data_density', 0) * 100:.1f}% |")
        lines.append("")
        
        return lines
    
    def _create_markdown_structure(self, structure: Dict) -> List[str]:
        """Create structure analysis section for markdown report"""
        lines = []
        
        # Sheet overview table
        lines.append("### Sheet Overview")
        lines.append("")
        lines.append("| Sheet Name | Rows | Columns | Status |")
        lines.append("|------------|------|---------|--------|")
        
        for sheet in structure.get('sheet_details', []):
            name = sheet.get('name', 'Unknown')
            rows = sheet.get('max_row', 0)
            cols = sheet.get('max_column', 0)
            status = sheet.get('status', 'Unknown')
            lines.append(f"| {name} | {rows:,} | {cols} | {status} |")
        
        lines.append("")
        
        # Workbook features
        features = structure.get('workbook_features', {})
        lines.append("### Workbook Features")
        lines.append("")
        lines.append("| Feature | Status |")
        lines.append("|---------|--------|")
        lines.append(f"| Protection | {'✅ Protected' if structure.get('protection_info', {}).get('has_protection', False) else '❌ None'} |")
        lines.append(f"| Macros | {'⚠️ Present' if features.get('has_macros', False) else '✅ Not Present'} |")
        lines.append(f"| External Connections | {features.get('has_external_connections', 0)} |")
        lines.append(f"| Pivot Tables | {features.get('has_pivot_tables', 0)} |")
        lines.append(f"| Data Validation Rules | {features.get('data_validation_rules', 0)} |")
        lines.append(f"| Conditional Formatting | {features.get('conditional_formatting_rules', 0)} |")
        lines.append("")
        
        return lines
    
    def _create_markdown_data_quality(self, data_profiler: Dict) -> List[str]:
        """Create data quality section for markdown report"""
        lines = []
        
        # Overall metrics
        overall_metrics = data_profiler.get('overall_metrics', {})
        lines.append("### Overall Data Quality Metrics")
        lines.append("")
        lines.append(f"- **Data Quality Score:** {overall_metrics.get('quality_score', 0) * 100:.1f}%")
        lines.append(f"- **Data Density:** {overall_metrics.get('data_density', 0) * 100:.1f}%")
        lines.append(f"- **Data Variety Score:** {overall_metrics.get('data_variety_score', 0) * 100:.1f}%")
        lines.append("")
        
        # Data type distribution
        type_dist = data_profiler.get('data_type_distribution', {})
        if type_dist:
            total_values = sum(type_dist.values())
            lines.append("### Data Type Distribution")
            lines.append("")
            lines.append("| Data Type | Percentage | Count |")
            lines.append("|-----------|------------|-------|")
            
            for data_type, count in type_dist.items():
                percentage = (count / total_values) * 100 if total_values > 0 else 0
                lines.append(f"| {data_type.capitalize()} | {percentage:.1f}% | {count:,} |")
            
            lines.append("")
        
        return lines
    
    def _create_markdown_sheet_analysis(self, data_profiler: Dict) -> List[str]:
        """Create detailed sheet analysis section for markdown report"""
        lines = []
        
        sheet_analysis = data_profiler.get('sheet_analysis', {})
        
        for sheet_name, sheet_data in sheet_analysis.items():
            lines.append(f"### 📄 {sheet_name}")
            lines.append("")
            
            # Basic info
            lines.append("**Sheet Properties:**")
            lines.append(f"- Dimensions: {sheet_data.get('dimensions', 'N/A')}")
            lines.append(f"- Data Cells: {sheet_data.get('estimated_data_cells', 0):,}")
            lines.append(f"- Data Density: {sheet_data.get('data_density', 0) * 100:.1f}%")
            lines.append("")
            
            # Headers preview
            columns = sheet_data.get('columns', [])
            if columns:
                headers = [col.get('header', f"Column {col.get('letter', '')}") for col in columns[:8]]
                lines.append("**Headers Preview:**")
                lines.append("```")
                lines.append(" | ".join(headers))
                lines.append("```")
                lines.append("")
                
                # Sample data
                lines.append("**Sample Data (First 3 Rows):**")
                lines.append("")
                
                # Create sample data table
                header_row = [col.get('header', f"Col {col.get('letter', '')}")[:15] for col in columns[:5]]
                lines.append("| " + " | ".join(header_row) + " |")
                lines.append("|" + "|".join(["-" * 17 for _ in header_row]) + "|")
                
                for row_idx in range(3):
                    row_data = []
                    for col in columns[:5]:
                        sample_values = col.get('sample_values', [])
                        if row_idx < len(sample_values):
                            value = str(sample_values[row_idx])[:15]
                            row_data.append(value)
                        else:
                            row_data.append("")
                    
                    if any(row_data):
                        lines.append("| " + " | ".join(row_data) + " |")
                
                lines.append("")
                
                # Column analysis
                lines.append("**Column Analysis:**")
                lines.append("")
                lines.append("| Column | Header | Type | Fill Rate | Unique Values |")
                lines.append("|--------|--------|------|-----------|---------------|")
                
                for col in columns[:10]:  # First 10 columns
                    letter = col.get('letter', '')
                    header = col.get('header', f'Column {letter}')
                    data_type = col.get('data_type', 'unknown')
                    fill_rate = col.get('fill_rate', 0) * 100
                    unique = col.get('unique_values', 0)
                    
                    lines.append(f"| {letter} | {header} | {data_type} | {fill_rate:.1f}% | {unique:,} |")
                
                lines.append("")
        
        return lines
    
    def _create_markdown_relationships(self, relationships: Dict) -> List[str]:
        """Create cross-sheet relationships section for markdown report"""
        lines = []
        
        if relationships.get('skipped', False):
            lines.append("*Cross-sheet analysis was skipped.*")
            return lines
        
        relationships_found = relationships.get('relationships_found', [])
        
        if not relationships_found:
            lines.append("*No cross-sheet relationships found.*")
            return lines
        
        lines.append("| Relationship | Source Sheet | Target Sheet | Match Rate |")
        lines.append("|--------------|--------------|--------------|------------|")
        
        for rel in relationships_found:
            rel_type = rel.get('relationship_type', 'unknown')
            source = rel.get('source_sheet', '')
            target = rel.get('target_sheet', '')
            match_rate = rel.get('match_rate', 0) * 100
            
            lines.append(f"| {rel_type} | {source} | {target} | {match_rate:.1f}% |")
        
        lines.append("")
        
        # Key columns details
        lines.append("**Key Columns for Relationships:**")
        lines.append("")
        
        for rel in relationships_found:
            source = rel.get('source_sheet', '')
            target = rel.get('target_sheet', '')
            key_columns = rel.get('key_columns', [])
            
            if key_columns:
                lines.append(f"- **{source} → {target}:** {', '.join(key_columns)}")
        
        lines.append("")
        
        return lines
    
    def _create_markdown_security(self, security: Dict) -> List[str]:
        """Create security analysis section for markdown report"""
        lines = []
        
        if not security:
            lines.append("*Security analysis not available.*")
            return lines
        
        # Security score
        score = security.get('overall_score', 0)
        risk_level = security.get('risk_level', 'Unknown')
        
        lines.append(f"**Overall Security Score:** {score:.1f}/10")
        lines.append(f"**Risk Level:** {risk_level}")
        lines.append("")
        
        # Security status
        if score >= 8:
            lines.append("✅ **Security Status: Good**")
        elif score >= 6:
            lines.append("⚠️ **Security Status: Fair**")
        else:
            lines.append("❌ **Security Status: Poor**")
        lines.append("")
        
        # Threats
        threats = security.get('threats', [])
        if threats:
            lines.append("### Threats Detected")
            lines.append("")
            for threat in threats:
                lines.append(f"- ⚠️ {threat}")
            lines.append("")
        else:
            lines.append("✅ No security threats detected.")
            lines.append("")
        
        # Sensitive data patterns
        patterns = security.get('patterns_detected', {})
        if patterns.get('patterns_found', False):
            lines.append("### Sensitive Data Patterns")
            lines.append("")
            pattern_counts = patterns.get('pattern_counts', {})
            for pattern, count in pattern_counts.items():
                pattern_name = pattern.replace('_', ' ').title()
                lines.append(f"- {pattern_name}: {count} instances")
            lines.append("")
        
        # Security recommendations
        security_recs = security.get('recommendations', [])
        if security_recs:
            lines.append("### Security Recommendations")
            lines.append("")
            for rec in security_recs:
                lines.append(f"- {rec}")
            lines.append("")
        
        return lines
    
    def _create_markdown_recommendations(self, recommendations: List[str]) -> List[str]:
        """Create recommendations section for markdown report"""
        lines = []
        
        if recommendations:
            for i, rec in enumerate(recommendations, 1):
                lines.append(f"{i}. {rec}")
        else:
            lines.append("✅ No specific recommendations - file appears well-structured.")
        
        lines.append("")
        return lines
    
    def _create_markdown_execution_status(self, exec_summary: Dict) -> List[str]:
        """Create module execution status section for markdown report"""
        lines = []
        
        # Summary
        lines.append(f"**Total Modules:** {exec_summary.get('total_modules', 0)}")
        lines.append(f"**Successful:** {exec_summary.get('successful_modules', 0)}")
        lines.append(f"**Failed:** {exec_summary.get('failed_modules', 0)}")
        lines.append(f"**Success Rate:** {exec_summary.get('success_rate', 0) * 100:.1f}%")
        lines.append("")
        
        # Module status table
        module_statuses = exec_summary.get('module_statuses', {})
        if module_statuses:
            lines.append("### Module Status Details")
            lines.append("")
            lines.append("| Module | Status |")
            lines.append("|--------|--------|")
            
            for module, status in module_statuses.items():
                module_name = module.replace('_', ' ').title()
                status_icon = "✅" if status == "success" else "❌"
                lines.append(f"| {module_name} | {status_icon} {status.title()} |")
        
        lines.append("")
        return lines
    
    def _create_fallback_text_report(self, results: Dict, output_path: str) -> str:
        """Create a basic fallback text report if comprehensive generation fails"""
        content = "EXCEL ANALYSIS REPORT - FALLBACK MODE\n"
        content += "=" * 50 + "\n"
        content += f"File: {results.get('file_info', {}).get('name', 'Unknown')}\n"
        content += "Report generation encountered an issue.\n"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return output_path
    
    def _create_fallback_markdown_report(self, results: Dict, output_path: str) -> str:
        """Create a basic fallback markdown report if comprehensive generation fails"""
        content = "# Excel Analysis Report - Fallback Mode\n\n"
        content += f"**File:** {results.get('file_info', {}).get('name', 'Unknown')}\n\n"
        content += "Report generation encountered an issue.\n"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return output_path