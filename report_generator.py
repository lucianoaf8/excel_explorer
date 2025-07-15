"""
HTML and JSON report generation
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Any


class ReportGenerator:
    """Generate HTML and JSON reports from analysis results"""
    
    def generate_html_report(self, results: Dict[str, Any], output_path: str) -> str:
        """Generate complete HTML report"""
        html_content = self._create_html_template(results)
        
        output_file = Path(output_path)
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return str(output_file)
    
    def generate_json_report(self, results: Dict[str, Any], output_path: str) -> str:
        """Generate JSON report"""
        output_file = Path(output_path)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, default=str)
        
        return str(output_file)
    
    def _create_html_template(self, results: Dict[str, Any]) -> str:
        """Create HTML report template"""
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        exec_summary = results.get('execution_summary', {})
        modules = results.get('module_results', {})
        
        timestamp = datetime.fromtimestamp(metadata.get('timestamp', 0))
        # Render helper HTML elements
        sheet_list_html = "".join(f"<li>{s}</li>" for s in file_info.get('sheets', []))
        
        return f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analysis - {file_info.get('name', 'Unknown')}</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; margin: -30px -30px 30px -30px; border-radius: 8px 8px 0 0; }}
        .header h1 {{ margin: 0 0 10px 0; font-size: 2.5em; }}
        .subtitle {{ opacity: 0.9; font-size: 1.1em; }}
        .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 30px 0; }}
        .card {{ background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 6px; padding: 20px; }}
        .card h3 {{ margin: 0 0 15px 0; color: #495057; }}
        .metric {{ display: flex; justify-content: space-between; margin: 10px 0; }}
        .metric-value {{ font-weight: bold; color: #28a745; }}
        .section {{ margin: 30px 0; }}
        .section h2 {{ color: #495057; border-bottom: 2px solid #667eea; padding-bottom: 10px; }}
        .module {{ background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 6px; border-left: 4px solid #28a745; }}
        .module.failed {{ border-left-color: #dc3545; }}
        .module h4 {{ margin: 0 0 10px 0; color: #495057; }}
        .status-success {{ color: #28a745; font-weight: bold; }}
        .status-failed {{ color: #dc3545; font-weight: bold; }}
        pre {{ background: #f8f9fa; padding: 15px; border-radius: 4px; overflow-x: auto; }}
        .recommendations {{ background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 6px; padding: 20px; margin: 20px 0; }}
        .rec-item {{ margin: 10px 0; padding: 10px; background: white; border-radius: 4px; }}
        .sheet-table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
        .sheet-table th, .sheet-table td {{ border: 1px solid #dee2e6; padding: 6px 8px; text-align: left; }}
        .sheet-table th {{ background: #e9ecef; }}
        details {{ margin: 10px 0; padding: 10px; background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; }}
        details summary {{ cursor: pointer; font-weight: 600; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Excel Analysis Report</h1>
            <div class="subtitle">
                File: {file_info.get('name', 'Unknown')} | 
                Size: {file_info.get('size_mb', 0):.2f} MB | 
                Generated: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}
            </div>
        </div>
        
        <div class="section">
            <h2>📑 File & Workbook Overview</h2>
            <div class="card">
                <div class="metric"><span>Full Path:</span><span class="metric-value">{file_info.get('path','')}</span></div>
                <div class="metric"><span>Created:</span><span class="metric-value">{file_info.get('created','')}</span></div>
                <div class="metric"><span>Modified:</span><span class="metric-value">{file_info.get('modified','')}</span></div>
            </div>
            <details>
                <summary>Sheet List ({file_info.get('sheet_count',0)})</summary>
                <ol>
                    {sheet_list_html}
                </ol>
            </details>
        </div>
        
        <div class="grid">
            <div class="card">
                <h3>📈 Analysis Summary</h3>
                <div class="metric">
                    <span>Success Rate:</span>
                    <span class="metric-value">{metadata.get('success_rate', 0):.1%}</span>
                </div>
                <div class="metric">
                    <span>Quality Score:</span>
                    <span class="metric-value">{metadata.get('quality_score', 0):.1%}</span>
                </div>
                <div class="metric">
                    <span>Duration:</span>
                    <span class="metric-value">{metadata.get('total_duration_seconds', 0):.1f}s</span>
                </div>
            </div>
            
            <div class="card">
                <h3>📁 File Structure</h3>
                <div class="metric">
                    <span>Total Sheets:</span>
                    <span class="metric-value">{modules.get('structure_mapper', {}).get('total_sheets', 0)}</span>
                </div>
                <div class="metric">
                    <span>Named Ranges:</span>
                    <span class="metric-value">{modules.get('structure_mapper', {}).get('named_ranges_count', 0)}</span>
                </div>
                <div class="metric">
                    <span>Tables:</span>
                    <span class="metric-value">{modules.get('structure_mapper', {}).get('table_count', 0)}</span>
                </div>
            </div>
            
            <div class="card">
                <h3>🔢 Data Analysis</h3>
                <div class="metric">
                    <span>Total Cells:</span>
                    <span class="metric-value">{modules.get('data_profiler', {}).get('total_cells', 0):,}</span>
                </div>
                <div class="metric">
                    <span>Data Cells:</span>
                    <span class="metric-value">{modules.get('data_profiler', {}).get('total_data_cells', 0):,}</span>
                </div>
                <div class="metric">
                    <span>Data Quality:</span>
                    <span class="metric-value">{modules.get('data_profiler', {}).get('data_quality_score', 0):.1%}</span>
                </div>
            </div>
            
            <div class="card">
                <h3>⚡ Formulas & Visuals</h3>
                <div class="metric">
                    <span>Total Formulas:</span>
                    <span class="metric-value">{modules.get('formula_analyzer', {}).get('total_formulas', 0)}</span>
                </div>
                <div class="metric">
                    <span>Charts:</span>
                    <span class="metric-value">{modules.get('visual_cataloger', {}).get('total_charts', 0)}</span>
                </div>
                <div class="metric">
                    <span>Images:</span>
                    <span class="metric-value">{modules.get('visual_cataloger', {}).get('total_images', 0)}</span>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2>🔧 Module Execution</h2>
            {self._generate_module_sections(exec_summary, modules)}
        </div>
        
        <div class="section">
            <h2>📄 Sheet Details</h2>
            {self._generate_sheet_details(modules)}
        </div>
        
        <div class="recommendations">
            <h2>💡 Recommendations</h2>
            {self._generate_recommendations(results.get('recommendations', []))}
        </div>
        
        
    </div>
</body>
</html>
"""
    
    def _generate_module_sections(self, exec_summary: Dict, modules: Dict) -> str:
        """Generate module status sections"""
        statuses = exec_summary.get('module_statuses', {})
        sections = []
        
        for module, status in statuses.items():
            icon = "✅" if status == "success" else "❌"
            css_class = "module" if status == "success" else "module failed"
            status_class = "status-success" if status == "success" else "status-failed"
            
            sections.append(f"""
            <div class="{css_class}">
                <h4>{icon} {module.replace('_', ' ').title()}</h4>
                <div class="{status_class}">Status: {status.title()}</div>
            </div>
            """)
        
        return "".join(sections)
    
    def _generate_sheet_details(self, modules: Dict) -> str:
        """Generate per-sheet column detail tables"""
        sheet_info = modules.get('data_profiler', {}).get('sheet_analysis', {}) if modules else {}
        if not sheet_info:
            return "<p>No sheet details available.</p>"
        sections = []
        for sheet, info in sheet_info.items():
            columns = info.get('columns', [])
            if not columns:
                continue
            rows_html = "".join(
                f"<tr><td>{col['letter']}</td><td>{col['range']}</td><td>{col['data_type'].title()}</td></tr>"
                for col in columns
            )
            summary_text = (
                f"{sheet} • Range: {info.get('used_range')} • "
                f"Data: {info.get('estimated_data_cells', 0):,} • "
                f"Empty: {info.get('empty_cells', 0):,}"
            )
            section_html = f"""
            <details>
                <summary>{summary_text}</summary>
                <table class=\"sheet-table\">
                    <tr><th>Column</th><th>Range</th><th>Data Type</th></tr>
                    {rows_html}
                </table>
            </details>
            """
            sections.append(section_html)
        return "".join(sections)

    def _generate_recommendations(self, recommendations: list) -> str:
        """Generate recommendations section"""
        if not recommendations:
            return "<div class='rec-item'>No specific recommendations - file appears well-structured.</div>"
        
        items = []
        for i, rec in enumerate(recommendations, 1):
            items.append(f"<div class='rec-item'><strong>{i}.</strong> {rec}</div>")
        
        return "".join(items)
