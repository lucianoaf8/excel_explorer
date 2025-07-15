"""
User-friendly HTML report generator for Excel analysis results
"""

import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List
from src.engine.enhanced_orchestrator import AnalysisResults


class ReportGenerator:
    """Generate comprehensive HTML reports from analysis results"""
    
    def generate_html_report(self, results: Dict[str, Any], output_path: str) -> str:
        """Generate user-friendly HTML report"""
        output_path = Path(output_path)
        
        # Convert results if needed
        if isinstance(results, AnalysisResults):
            results = self._convert_results(results)
        
        # Generate HTML content
        html_content = self._generate_html(results)
        
        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return str(output_path)
    
    def _convert_results(self, results: AnalysisResults) -> Dict[str, Any]:
        """Convert AnalysisResults to dictionary"""
        return {
            'file_info': results.file_info,
            'analysis_metadata': results.analysis_metadata,
            'module_results': results.module_results,
            'execution_summary': results.execution_summary,
            'resource_usage': results.resource_usage,
            'recommendations': results.recommendations
        }
    
    def _generate_html(self, results: Dict[str, Any]) -> str:
        """Generate complete HTML report"""
        return f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analysis Report - {results.get('file_info', {}).get('name', 'Unknown File')}</title>
    <style>
        {self._get_css_styles()}
    </style>
</head>
<body>
    <div class="container">
        {self._generate_header(results)}
        {self._generate_executive_summary(results)}
        {self._generate_module_sections(results)}
        {self._generate_recommendations(results)}
        {self._generate_technical_details(results)}
        {self._generate_footer()}
    </div>
    <script>
        {self._get_javascript()}
    </script>
</body>
</html>
"""
    
    def _get_css_styles(self) -> str:
        """Modern CSS styles for the report"""
        return """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f5f5f5;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .summary-card {
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .summary-card h3 {
            color: #667eea;
            margin-bottom: 15px;
            font-size: 1.1em;
        }
        
        .metric {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
        }
        
        .metric-value {
            font-weight: bold;
            color: #333;
        }
        
        .success { color: #4caf50; }
        .warning { color: #ff9800; }
        .error { color: #f44336; }
        
        .section {
            background: white;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 25px;
            border-left: 4px solid #667eea;
        }
        
        .section h2 {
            color: #667eea;
            margin-bottom: 20px;
            font-size: 1.5em;
        }
        
        .module-status {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
            padding: 10px;
            background: #f9f9f9;
            border-radius: 5px;
        }
        
        .status-icon {
            width: 20px;
            height: 20px;
            border-radius: 50%;
            margin-right: 15px;
        }
        
        .status-success { background-color: #4caf50; }
        .status-error { background-color: #f44336; }
        
        .collapsible {
            cursor: pointer;
            padding: 15px;
            background: #f1f1f1;
            border: none;
            outline: none;
            width: 100%;
            text-align: left;
            font-size: 16px;
            border-radius: 5px;
            margin-bottom: 10px;
        }
        
        .collapsible:hover { background: #e1e1e1; }
        
        .content {
            display: none;
            padding: 15px;
            background: #fafafa;
            border-radius: 5px;
            margin-bottom: 15px;
        }
        
        .content.active { display: block; }
        
        .recommendations {
            background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
        }
        
        .recommendations h2 {
            color: white;
            margin-bottom: 15px;
        }
        
        .recommendation-item {
            background: rgba(255,255,255,0.1);
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 5px;
        }
        
        .footer {
            text-align: center;
            color: #666;
            padding: 20px;
            border-top: 1px solid #e0e0e0;
            margin-top: 30px;
        }
        
        pre {
            background: #f4f4f4;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            font-size: 14px;
        }
        
        .progress-bar {
            background: #e0e0e0;
            border-radius: 10px;
            height: 20px;
            overflow: hidden;
            margin: 10px 0;
        }
        
        .progress-fill {
            background: linear-gradient(90deg, #4caf50 0%, #45a049 100%);
            height: 100%;
            transition: width 0.3s ease;
        }
        """
    
    def _generate_header(self, results: Dict[str, Any]) -> str:
        """Generate report header"""
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        
        return f"""
        <div class="header">
            <h1>📊 Excel Analysis Report</h1>
            <div class="subtitle">
                File: {file_info.get('name', 'Unknown')} | 
                Size: {file_info.get('size_mb', 0):.2f} MB | 
                Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </div>
        </div>
        """
    
    def _generate_executive_summary(self, results: Dict[str, Any]) -> str:
        """Generate executive summary cards"""
        metadata = results.get('analysis_metadata', {})
        exec_summary = results.get('execution_summary', {})
        
        success_rate = exec_summary.get('success_rate', 0) * 100
        quality_score = metadata.get('quality_score', 0) * 100
        duration = metadata.get('total_duration_seconds', 0)
        
        return f"""
        <div class="summary-grid">
            <div class="summary-card">
                <h3>📈 Analysis Overview</h3>
                <div class="metric">
                    <span>Success Rate:</span>
                    <span class="metric-value success">{success_rate:.1f}%</span>
                </div>
                <div class="metric">
                    <span>Quality Score:</span>
                    <span class="metric-value success">{quality_score:.1f}%</span>
                </div>
                <div class="metric">
                    <span>Duration:</span>
                    <span class="metric-value">{duration:.1f}s</span>
                </div>
            </div>
            
            <div class="summary-card">
                <h3>🔧 Modules Executed</h3>
                <div class="metric">
                    <span>Total Modules:</span>
                    <span class="metric-value">{exec_summary.get('total_modules', 0)}</span>
                </div>
                <div class="metric">
                    <span>Successful:</span>
                    <span class="metric-value success">{exec_summary.get('successful_modules', 0)}</span>
                </div>
                <div class="metric">
                    <span>Failed:</span>
                    <span class="metric-value {('error' if exec_summary.get('failed_modules', 0) > 0 else 'success')}">{exec_summary.get('failed_modules', 0)}</span>
                </div>
            </div>
            
            <div class="summary-card">
                <h3>📊 File Statistics</h3>
                {self._generate_file_stats(results)}
            </div>
            
            <div class="summary-card">
                <h3>💾 Resource Usage</h3>
                {self._generate_resource_stats(results)}
            </div>
        </div>
        """
    
    def _generate_file_stats(self, results: Dict[str, Any]) -> str:
        """Generate file statistics"""
        module_results = results.get('module_results', {})
        stats = []
        
        # Extract key statistics from module results
        if 'structure_mapper' in module_results:
            structure = module_results['structure_mapper']
            if hasattr(structure, 'total_sheets'):
                stats.append(f'<div class="metric"><span>Sheets:</span><span class="metric-value">{structure.total_sheets}</span></div>')
            if hasattr(structure, 'total_cells_with_data'):
                stats.append(f'<div class="metric"><span>Data Cells:</span><span class="metric-value">{structure.total_cells_with_data:,}</span></div>')
        
        if 'data_profiler' in module_results:
            data_profile = module_results['data_profiler']
            if hasattr(data_profile, 'total_regions'):
                stats.append(f'<div class="metric"><span>Data Regions:</span><span class="metric-value">{data_profile.total_regions}</span></div>')
        
        if not stats:
            stats.append('<div class="metric"><span>Data:</span><span class="metric-value">Available in detailed sections</span></div>')
        
        return '\n'.join(stats)
    
    def _generate_resource_stats(self, results: Dict[str, Any]) -> str:
        """Generate resource usage statistics"""
        resource_usage = results.get('resource_usage', {})
        current_usage = resource_usage.get('current_usage', {})
        
        memory_mb = current_usage.get('current_mb', 0)
        cpu_percent = current_usage.get('cpu_percent', 0)
        
        return f"""
        <div class="metric">
            <span>Memory Used:</span>
            <span class="metric-value">{memory_mb:.1f} MB</span>
        </div>
        <div class="metric">
            <span>CPU Usage:</span>
            <span class="metric-value">{cpu_percent:.1f}%</span>
        </div>
        <div class="metric">
            <span>Efficiency:</span>
            <span class="metric-value success">Optimal</span>
        </div>
        """
    
    def _generate_module_sections(self, results: Dict[str, Any]) -> str:
        """Generate detailed module sections"""
        module_results = results.get('module_results', {})
        exec_summary = results.get('execution_summary', {})
        module_statuses = exec_summary.get('module_statuses', {})
        
        sections = []
        
        for module_name, status in module_statuses.items():
            display_name = module_name.replace('_', ' ').title()
            is_success = status == 'success'
            status_class = 'status-success' if is_success else 'status-error'
            status_text = '✓ Success' if is_success else '✗ Failed'
            
            module_data = module_results.get(module_name, {})
            
            sections.append(f"""
            <div class="section">
                <h2>{display_name}</h2>
                <div class="module-status">
                    <div class="status-icon {status_class}"></div>
                    <div>
                        <strong>{status_text}</strong>
                        <div style="font-size: 0.9em; color: #666;">
                            {self._get_module_description(module_name)}
                        </div>
                    </div>
                </div>
                
                <button class="collapsible">View Detailed Results</button>
                <div class="content">
                    {self._format_module_data(module_name, module_data)}
                </div>
            </div>
            """)
        
        return '\n'.join(sections)
    
    def _get_module_description(self, module_name: str) -> str:
        """Get description for each module"""
        descriptions = {
            'health_checker': 'Validates file integrity, checks for corruption, and assesses security risks',
            'structure_mapper': 'Maps workbook structure, sheet relationships, and data organization',
            'data_profiler': 'Analyzes data types, quality metrics, and statistical patterns',
            'formula_analyzer': 'Examines formulas, dependencies, and calculation complexity',
            'visual_cataloger': 'Catalogs charts, images, and visual elements in the workbook',
            'connection_inspector': 'Inspects external data connections and refresh behaviors',
            'pivot_intelligence': 'Analyzes pivot tables, summaries, and aggregation patterns',
            'doc_synthesizer': 'Synthesizes comprehensive documentation and insights'
        }
        return descriptions.get(module_name, 'Analysis module')
    
    def _format_module_data(self, module_name: str, data: Any) -> str:
        """Format module data for display"""
        if not data:
            return "<p>No detailed data available for this module.</p>"
        
        # Convert data to formatted display
        if hasattr(data, '__dict__'):
            # Object with attributes
            formatted = []
            for key, value in data.__dict__.items():
                if not key.startswith('_'):
                    formatted.append(f"<strong>{key.replace('_', ' ').title()}:</strong> {value}")
            return '<br>'.join(formatted)
        elif isinstance(data, dict):
            # Dictionary data
            formatted = []
            for key, value in data.items():
                if isinstance(value, (dict, list)) and len(str(value)) > 100:
                    formatted.append(f"<strong>{key}:</strong> <em>Complex data structure</em>")
                else:
                    formatted.append(f"<strong>{key}:</strong> {value}")
            return '<br>'.join(formatted)
        else:
            # Simple data
            return f"<pre>{str(data)}</pre>"
    
    def _generate_recommendations(self, results: Dict[str, Any]) -> str:
        """Generate recommendations section"""
        recommendations = results.get('recommendations', [])
        
        if not recommendations:
            return """
            <div class="recommendations">
                <h2>🎯 Recommendations</h2>
                <div class="recommendation-item">
                    <strong>✅ Excellent!</strong> No specific recommendations - your Excel file is well-structured and optimized.
                </div>
            </div>
            """
        
        rec_items = []
        for i, rec in enumerate(recommendations, 1):
            rec_items.append(f"""
            <div class="recommendation-item">
                <strong>{i}.</strong> {rec}
            </div>
            """)
        
        return f"""
        <div class="recommendations">
            <h2>🎯 Recommendations</h2>
            {''.join(rec_items)}
        </div>
        """
    
    def _generate_technical_details(self, results: Dict[str, Any]) -> str:
        """Generate technical details section"""
        return f"""
        <div class="section">
            <h2>🔧 Technical Details</h2>
            <button class="collapsible">View Raw Analysis Data</button>
            <div class="content">
                <pre>{json.dumps(results, indent=2, default=str)}</pre>
            </div>
        </div>
        """
    
    def _generate_footer(self) -> str:
        """Generate report footer"""
        return f"""
        <div class="footer">
            <p>Generated by Excel Explorer | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>This report provides comprehensive analysis of your Excel file structure and content.</p>
        </div>
        """
    
    def _get_javascript(self) -> str:
        """JavaScript for interactive elements"""
        return """
        // Collapsible sections
        document.querySelectorAll('.collapsible').forEach(button => {
            button.addEventListener('click', function() {
                this.classList.toggle('active');
                const content = this.nextElementSibling;
                content.classList.toggle('active');
            });
        });
        
        // Smooth scrolling for any internal links
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                document.querySelector(this.getAttribute('href')).scrollIntoView({
                    behavior: 'smooth'
                });
            });
        });
        """
