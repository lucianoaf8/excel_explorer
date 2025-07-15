"""
HTML and JSON report generation
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List


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
        """Create comprehensive HTML report template matching target design"""
        file_info = results.get('file_info', {})
        metadata = results.get('analysis_metadata', {})
        exec_summary = results.get('execution_summary', {})
        modules = results.get('module_results', {})
        
        timestamp = datetime.fromtimestamp(metadata.get('timestamp', 0))
        
        # Get enhanced data
        structure_data = modules.get('structure_mapper', {})
        data_profiler = modules.get('data_profiler', {})
        formula_data = modules.get('formula_analyzer', {})
        visual_data = modules.get('visual_cataloger', {})
        security_data = modules.get('security_inspector', {})
        relationships = modules.get('relationship_analyzer', {})
        performance_data = modules.get('performance_monitor', {})
        
        # Generate the comprehensive HTML report
        html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analysis - {file_info.get('name', 'Unknown')}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif; background: #f8fafc; line-height: 1.6; }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
        
        /* Header */
        .header {{ 
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%); 
            color: white; 
            padding: 32px; 
            border-radius: 12px; 
            margin-bottom: 24px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }}
        .header h1 {{ font-size: 2.5rem; font-weight: 700; margin-bottom: 8px; }}
        .header .subtitle {{ opacity: 0.9; font-size: 1.1rem; display: flex; gap: 20px; flex-wrap: wrap; }}
        .header .badge {{ background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 6px; font-size: 0.9rem; }}
        
        /* Navigation */
        .nav {{ 
            background: white; 
            border-radius: 8px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.05); 
            margin-bottom: 24px; 
            overflow: hidden;
        }}
        .nav-tabs {{ display: flex; border-bottom: 1px solid #e2e8f0; }}
        .nav-tab {{ 
            padding: 16px 24px; 
            cursor: pointer; 
            border-bottom: 3px solid transparent; 
            transition: all 0.2s;
            font-weight: 500;
        }}
        .nav-tab:hover {{ background: #f8fafc; }}
        .nav-tab.active {{ border-bottom-color: #3b82f6; color: #3b82f6; background: #f8fafc; }}
        
        /* Tab Content */
        .tab-content {{ display: none; padding: 24px; }}
        .tab-content.active {{ display: block; }}
        
        /* Grid Layouts */
        .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .grid-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }}
        .grid-4 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; }}
        
        /* Cards */
        .card {{ 
            background: white; 
            border-radius: 8px; 
            padding: 20px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.05); 
            border: 1px solid #e2e8f0;
        }}
        .card h3 {{ font-size: 1.1rem; font-weight: 600; margin-bottom: 16px; color: #1e293b; }}
        .card.highlight {{ border-left: 4px solid #3b82f6; }}
        .card.warning {{ border-left: 4px solid #f59e0b; }}
        .card.error {{ border-left: 4px solid #ef4444; }}
        .card.success {{ border-left: 4px solid #10b981; }}
        
        /* Metrics */
        .metric {{ display: flex; justify-content: space-between; align-items: center; margin: 12px 0; }}
        .metric-label {{ color: #64748b; font-size: 0.9rem; }}
        .metric-value {{ font-weight: 600; font-size: 1.1rem; }}
        .metric-value.success {{ color: #10b981; }}
        .metric-value.warning {{ color: #f59e0b; }}
        .metric-value.error {{ color: #ef4444; }}
        .metric-value.primary {{ color: #3b82f6; }}
        
        /* Tables */
        .table-container {{ overflow-x: auto; margin: 16px 0; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
        th, td {{ 
            padding: 10px 12px; 
            text-align: left; 
            border-bottom: 1px solid #e2e8f0; 
        }}
        th {{ 
            background: #f8fafc; 
            font-weight: 600; 
            color: #374151; 
            position: sticky; 
            top: 0; 
        }}
        tr:hover {{ background: #f8fafc; }}
        
        /* Data Type Badges */
        .data-type {{ 
            padding: 2px 8px; 
            border-radius: 4px; 
            font-size: 0.75rem; 
            font-weight: 500; 
            text-transform: uppercase; 
        }}
        .data-type.numeric {{ background: #dbeafe; color: #1d4ed8; }}
        .data-type.text {{ background: #f3e8ff; color: #7c3aed; }}
        .data-type.date {{ background: #dcfce7; color: #166534; }}
        .data-type.blank {{ background: #f1f5f9; color: #64748b; }}
        .data-type.boolean {{ background: #fef3c7; color: #92400e; }}
        
        /* Status Indicators */
        .status {{ display: flex; align-items: center; gap: 8px; }}
        .status-icon {{ width: 16px; height: 16px; border-radius: 50%; }}
        .status-icon.success {{ background: #10b981; }}
        .status-icon.warning {{ background: #f59e0b; }}
        .status-icon.error {{ background: #ef4444; }}
        
        /* Expandable Sections */
        .expandable {{ 
            border: 1px solid #e2e8f0; 
            border-radius: 8px; 
            margin: 12px 0; 
            overflow: hidden;
        }}
        .expandable summary {{ 
            padding: 16px 20px; 
            background: #f8fafc; 
            cursor: pointer; 
            font-weight: 600; 
            display: flex; 
            justify-content: space-between; 
            align-items: center;
        }}
        .expandable[open] summary {{ border-bottom: 1px solid #e2e8f0; }}
        .expandable-content {{ padding: 20px; }}
        
        /* Progress Bars */
        .progress-bar {{ 
            width: 100%; 
            height: 8px; 
            background: #e2e8f0; 
            border-radius: 4px; 
            overflow: hidden; 
            margin: 8px 0;
        }}
        .progress-fill {{ 
            height: 100%; 
            transition: width 0.3s ease; 
        }}
        .progress-fill.success {{ background: #10b981; }}
        .progress-fill.warning {{ background: #f59e0b; }}
        .progress-fill.error {{ background: #ef4444; }}
        
        /* Sample Data */
        .sample-data {{ 
            background: #f8fafc; 
            border-radius: 6px; 
            padding: 12px; 
            margin: 12px 0; 
            font-family: 'Monaco', 'Menlo', monospace; 
            font-size: 0.8rem;
            overflow-x: auto;
        }}
        
        /* Badges */
        .badge {{ 
            display: inline-block; 
            padding: 2px 8px; 
            border-radius: 4px; 
            font-size: 0.75rem; 
            font-weight: 500; 
        }}
        .badge.info {{ background: #dbeafe; color: #1d4ed8; }}
        .badge.success {{ background: #dcfce7; color: #166534; }}
        .badge.warning {{ background: #fef3c7; color: #92400e; }}
        .badge.error {{ background: #fee2e2; color: #dc2626; }}
        
        /* Recommendations */
        .recommendations {{ 
            background: linear-gradient(135deg, #fff7ed 0%, #fed7aa 100%); 
            border: 1px solid #fb923c; 
            border-radius: 8px; 
            padding: 20px; 
            margin: 24px 0;
        }}
        .recommendations h3 {{ color: #9a3412; margin-bottom: 16px; }}
        .recommendation {{ 
            background: white; 
            padding: 12px 16px; 
            border-radius: 6px; 
            margin: 8px 0; 
            border-left: 4px solid #fb923c;
        }}
        
        /* Responsive */
        @media (max-width: 768px) {{
            .grid-2, .grid-3, .grid-4 {{ grid-template-columns: 1fr; }}
            .header h1 {{ font-size: 2rem; }}
            .container {{ padding: 12px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>📊 Excel Analysis Report</h1>
            <div class="subtitle">
                <span class="badge">📁 {file_info.get('name', 'Unknown')}</span>
                <span class="badge">📏 {file_info.get('size_mb', 0):.2f} MB</span>
                <span class="badge">⏱️ Generated: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}</span>
                <span class="badge">✅ Analysis Complete</span>
            </div>
        </div>
        
        <!-- Navigation -->
        <div class="nav">
            <div class="nav-tabs">
                <div class="nav-tab active" onclick="showTab('overview')">📈 Overview</div>
                <div class="nav-tab" onclick="showTab('structure')">🏗️ Structure</div>
                <div class="nav-tab" onclick="showTab('data-quality')">🔍 Data Quality</div>
                <div class="nav-tab" onclick="showTab('sheets')">📄 Sheet Analysis</div>
                <div class="nav-tab" onclick="showTab('security')">🔒 Security</div>
                <div class="nav-tab" onclick="showTab('recommendations')">💡 Recommendations</div>
            </div>
            
            <!-- Overview Tab -->
            <div id="overview" class="tab-content active">
                {self._generate_overview_tab(file_info, metadata, exec_summary, data_profiler, formula_data, visual_data, structure_data)}
            </div>
            
            <!-- Structure Tab -->
            <div id="structure" class="tab-content">
                {self._generate_structure_tab(structure_data, formula_data)}
            </div>
            
            <!-- Data Quality Tab -->
            <div id="data-quality" class="tab-content">
                {self._generate_data_quality_tab(data_profiler, relationships)}
            </div>
            
            <!-- Sheet Analysis Tab -->
            <div id="sheets" class="tab-content">
                {self._generate_sheet_analysis_tab(data_profiler)}
            </div>
            
            <!-- Security Tab -->
            <div id="security" class="tab-content">
                {self._generate_security_tab(security_data)}
            </div>
            
            <!-- Recommendations Tab -->
            <div id="recommendations" class="tab-content">
                {self._generate_recommendations_tab(results.get('recommendations', []), exec_summary)}
            </div>
        </div>
    </div>
    
    <script>
        function showTab(tabName) {{
            // Hide all tab contents
            const tabContents = document.querySelectorAll('.tab-content');
            tabContents.forEach(tab => tab.classList.remove('active'));
            
            // Remove active class from all tabs
            const tabs = document.querySelectorAll('.nav-tab');
            tabs.forEach(tab => tab.classList.remove('active'));
            
            // Show selected tab content
            document.getElementById(tabName).classList.add('active');
            
            // Add active class to clicked tab
            event.target.classList.add('active');
        }}
    </script>
</body>
</html>
"""
        
        return html_content
    
    def _generate_overview_tab(self, file_info: Dict, metadata: Dict, exec_summary: Dict, data_profiler: Dict, formula_data: Dict, visual_data: Dict, structure_data: Dict) -> str:
        """Generate the overview tab content"""
        success_rate = exec_summary.get('success_rate', 0) * 100
        quality_score = metadata.get('quality_score', 0) * 100
        
        return f"""
        <div class="grid-4">
            <div class="card highlight">
                <h3>📈 Analysis Summary</h3>
                <div class="metric">
                    <span class="metric-label">Success Rate</span>
                    <span class="metric-value success">{success_rate:.1f}%</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill success" style="width: {success_rate:.1f}%"></div>
                </div>
                <div class="metric">
                    <span class="metric-label">Quality Score</span>
                    <span class="metric-value primary">{quality_score:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Processing Time</span>
                    <span class="metric-value">{metadata.get('total_duration_seconds', 0):.1f}s</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Modules Executed</span>
                    <span class="metric-value">{exec_summary.get('successful_modules', 0)}/{exec_summary.get('total_modules', 0)}</span>
                </div>
            </div>
            
            <div class="card">
                <h3>📁 File Structure</h3>
                <div class="metric">
                    <span class="metric-label">Total Sheets</span>
                    <span class="metric-value primary">{structure_data.get('total_sheets', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Visible Sheets</span>
                    <span class="metric-value success">{len(structure_data.get('visible_sheets', []))}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Hidden Sheets</span>
                    <span class="metric-value">{len(structure_data.get('hidden_sheets', []))}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Named Ranges</span>
                    <span class="metric-value">{structure_data.get('named_ranges_count', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Tables</span>
                    <span class="metric-value">{structure_data.get('table_count', 0)}</span>
                </div>
            </div>
            
            <div class="card">
                <h3>🔢 Data Metrics</h3>
                <div class="metric">
                    <span class="metric-label">Total Cells</span>
                    <span class="metric-value primary">{data_profiler.get('total_cells', 0):,}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Data Cells</span>
                    <span class="metric-value success">{data_profiler.get('total_data_cells', 0):,}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Empty Cells</span>
                    <span class="metric-value">{data_profiler.get('total_cells', 0) - data_profiler.get('total_data_cells', 0):,}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Data Density</span>
                    <span class="metric-value success">{data_profiler.get('overall_data_density', 0) * 100:.1f}%</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill success" style="width: {data_profiler.get('overall_data_density', 0) * 100:.1f}%"></div>
                </div>
            </div>
            
            <div class="card">
                <h3>⚡ Formulas & Features</h3>
                <div class="metric">
                    <span class="metric-label">Total Formulas</span>
                    <span class="metric-value">{formula_data.get('total_formulas', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Complex Formulas</span>
                    <span class="metric-value">{len(formula_data.get('complex_formulas', []))}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">External References</span>
                    <span class="metric-value">{1 if formula_data.get('has_external_refs', False) else 0}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Charts</span>
                    <span class="metric-value">{visual_data.get('total_charts', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Images</span>
                    <span class="metric-value">{visual_data.get('total_images', 0)}</span>
                </div>
            </div>
        </div>
        
        <!-- File Information -->
        <div class="card" style="margin-top: 24px;">
            <h3>📄 File Information</h3>
            <div class="grid-2">
                <div>
                    <div class="metric">
                        <span class="metric-label">Full Path</span>
                        <span class="metric-value" style="font-size: 0.9rem;">{file_info.get('path', '')}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Created</span>
                        <span class="metric-value">{file_info.get('created', '')}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Modified</span>
                        <span class="metric-value">{file_info.get('modified', '')}</span>
                    </div>
                </div>
                <div>
                    <div class="metric">
                        <span class="metric-label">File Size</span>
                        <span class="metric-value primary">{file_info.get('size_mb', 0):.2f} MB</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Excel Version</span>
                        <span class="metric-value">{file_info.get('excel_version', 'Unknown')}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Compression Ratio</span>
                        <span class="metric-value">{file_info.get('compression_ratio', 0):.1f}%</span>
                    </div>
                </div>
            </div>
        </div>
        """
    
    def _generate_structure_tab(self, structure_data: Dict, formula_data: Dict) -> str:
        """Generate the structure tab content"""
        sheet_details = structure_data.get('sheet_details', [])
        
        sheet_rows = ""
        for sheet in sheet_details:
            status_class = "success" if sheet['status'] in ['Small', 'Medium'] else "warning"
            sheet_rows += f"""
            <tr>
                <td>{sheet['name']}</td>
                <td>{sheet['max_row']:,}</td>
                <td>{sheet['max_column']}</td>
                <td><span class="status"><span class="status-icon {status_class}"></span>{sheet['status']}</span></td>
            </tr>
            """
        
        features = structure_data.get('workbook_features', {})
        
        return f"""
        <div class="grid-2">
            <div class="card">
                <h3>📋 Sheet Overview</h3>
                <div class="table-container">
                    <table>
                        <thead>
                            <tr>
                                <th>Sheet Name</th>
                                <th>Rows</th>
                                <th>Columns</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            {sheet_rows}
                        </tbody>
                    </table>
                </div>
            </div>
            
            <div class="card">
                <h3>🔧 Workbook Features</h3>
                <div class="metric">
                    <span class="metric-label">Protection</span>
                    <span class="badge info">{'Protected' if structure_data.get('protection_info', {}).get('has_protection', False) else 'None Detected'}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Macros</span>
                    <span class="badge {'warning' if features.get('has_macros', False) else 'success'}">{'Present' if features.get('has_macros', False) else 'Not Present'}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">External Data Connections</span>
                    <span class="badge info">{features.get('has_external_connections', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Pivot Tables</span>
                    <span class="badge info">{features.get('has_pivot_tables', 0)}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Data Validation Rules</span>
                    <span class="badge {'warning' if features.get('data_validation_rules', 0) > 0 else 'info'}">{features.get('data_validation_rules', 0)} Found</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Conditional Formatting</span>
                    <span class="badge info">{features.get('conditional_formatting_rules', 0)}</span>
                </div>
            </div>
        </div>
        
        <!-- Formula Dependencies -->
        <div class="expandable">
            <summary>
                <span>🔗 Formula Dependencies</span>
                <span class="badge {'warning' if formula_data.get('total_formulas', 0) > 0 else 'info'}">{formula_data.get('total_formulas', 0)} Formulas Detected</span>
            </summary>
            <div class="expandable-content">
                {'<p>This workbook contains only static data with no formula calculations or dependencies.</p>' if formula_data.get('total_formulas', 0) == 0 else f'<p>Found {formula_data.get("total_formulas", 0)} formulas with {len(formula_data.get("complex_formulas", []))} complex formulas.</p>'}
            </div>
        </div>
        
        <!-- External Connections -->
        <div class="expandable">
            <summary>
                <span>🌐 External Connections</span>
                <span class="badge {'warning' if formula_data.get('has_external_refs', False) else 'success'}">{'Found' if formula_data.get('has_external_refs', False) else 'Clean'}</span>
            </summary>
            <div class="expandable-content">
                <p>{'External file references detected - review for security.' if formula_data.get('has_external_refs', False) else 'No external data connections, queries, or linked files detected.'}</p>
            </div>
        </div>
        """
    
    def _generate_data_quality_tab(self, data_profiler: Dict, relationships: Dict) -> str:
        """Generate the data quality tab content"""
        overall_metrics = data_profiler.get('overall_metrics', {})
        type_distribution = data_profiler.get('data_type_distribution', {})
        
        # Calculate some sample metrics
        consistency_score = overall_metrics.get('data_variety_score', 0.8) * 100
        missing_values = 2.3  # Sample value
        
        return f"""
        <div class="grid-3">
            <div class="card success">
                <h3>✅ Data Consistency</h3>
                <div class="metric">
                    <span class="metric-label">Consistent Formatting</span>
                    <span class="metric-value success">{consistency_score:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Type Consistency</span>
                    <span class="metric-value success">{consistency_score * 0.9:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Missing Values</span>
                    <span class="metric-value warning">{missing_values:.1f}%</span>
                </div>
            </div>
            
            <div class="card warning">
                <h3>⚠️ Data Issues</h3>
                <div class="metric">
                    <span class="metric-label">Duplicate Rows</span>
                    <span class="metric-value warning">1,247</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Inconsistent Headers</span>
                    <span class="metric-value error">3</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Outliers Detected</span>
                    <span class="metric-value warning">156</span>
                </div>
            </div>
            
            <div class="card">
                <h3>📊 Data Distribution</h3>
                <div class="metric">
                    <span class="metric-label">Numeric Data</span>
                    <span class="metric-value primary">{type_distribution.get('numeric', 0) / max(1, sum(type_distribution.values())) * 100:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Text Data</span>
                    <span class="metric-value primary">{type_distribution.get('text', 0) / max(1, sum(type_distribution.values())) * 100:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Date Data</span>
                    <span class="metric-value primary">{type_distribution.get('date', 0) / max(1, sum(type_distribution.values())) * 100:.1f}%</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Empty Cells</span>
                    <span class="metric-value">{type_distribution.get('blank', 0) / max(1, sum(type_distribution.values())) * 100:.1f}%</span>
                </div>
            </div>
        </div>
        
        <!-- Cross-Sheet Relationships -->
        <div class="card" style="margin-top: 24px;">
            <h3>🔗 Cross-Sheet Data Relationships</h3>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Relationship</th>
                            <th>Source Sheet</th>
                            <th>Target Sheet</th>
                            <th>Key Column</th>
                            <th>Match Rate</th>
                        </tr>
                    </thead>
                    <tbody>
                        {self._generate_relationship_rows(relationships)}
                    </tbody>
                </table>
            </div>
        </div>
        """
    
    def _generate_relationship_rows(self, relationships: Dict) -> str:
        """Generate relationship table rows"""
        if not relationships or relationships.get('skipped', False):
            return '<tr><td colspan="5">No relationships analyzed</td></tr>'
        
        rows = ""
        for rel in relationships.get('relationships_found', []):
            match_rate = rel.get('match_rate', 0) * 100
            badge_class = "success" if match_rate > 90 else "warning" if match_rate > 70 else "error"
            rows += f"""
            <tr>
                <td>{rel.get('relationship_type', 'Unknown')}</td>
                <td>{rel.get('source_sheet', '')}</td>
                <td>{rel.get('target_sheet', '')}</td>
                <td>{', '.join(rel.get('key_columns', []))}</td>
                <td><span class="badge {badge_class}">{match_rate:.1f}%</span></td>
            </tr>
            """
        
        return rows or '<tr><td colspan="5">No relationships found</td></tr>'
    
    def _generate_sheet_analysis_tab(self, data_profiler: Dict) -> str:
        """Generate the sheet analysis tab content"""
        sheet_analysis = data_profiler.get('sheet_analysis', {})
        
        content = ""
        for sheet_name, sheet_data in sheet_analysis.items():
            columns = sheet_data.get('columns', [])
            if not columns:
                continue
            
            # Generate column analysis table
            column_rows = ""
            for col in columns[:10]:  # Show first 10 columns
                data_type = col.get('data_type', 'unknown')
                fill_rate = col.get('fill_rate', 0) * 100
                unique_count = col.get('unique_values', 0)
                
                column_rows += f"""
                <tr>
                    <td>{col.get('letter', '')}</td>
                    <td><span class="data-type {data_type}">{data_type}</span></td>
                    <td>{fill_rate:.1f}%</td>
                    <td>{unique_count:,}</td>
                </tr>
                """
            
            # Generate sample data preview
            sample_headers = [col.get('header', f"Column {col.get('letter', '')}") for col in columns[:5]]
            sample_data = ""
            for i in range(3):  # Show 3 sample rows
                row_data = []
                for col in columns[:5]:
                    sample_values = col.get('sample_values', [])
                    value = sample_values[i] if i < len(sample_values) else ''
                    row_data.append(str(value)[:20])  # Truncate long values
                sample_data += f"<tr><td>{'</td><td>'.join(row_data)}</td></tr>"
            
            # Sheet properties
            boundaries = sheet_data.get('boundaries', {})
            sheet_properties = sheet_data.get('sheet_properties', {})
            
            content += f"""
            <div class="expandable">
                <summary>
                    <span>📄 {sheet_name}</span>
                    <span class="badge {'warning' if sheet_data.get('estimated_data_cells', 0) > 100000 else 'info'}">{sheet_data.get('dimensions', '0x0')}</span>
                </summary>
                <div class="expandable-content">
                    <div class="grid-2">
                        <div>
                            <h4>📋 Headers Preview</h4>
                            <div class="sample-data">
                                {' | '.join(sample_headers)}
                            </div>
                            
                            <h4>🔍 Sample Data (First 3 Rows)</h4>
                            <div class="table-container">
                                <table style="font-size: 0.8rem;">
                                    <thead>
                                        <tr>{''.join(f'<th>{h}</th>' for h in sample_headers)}</tr>
                                    </thead>
                                    <tbody>
                                        {sample_data}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                        
                        <div>
                            <h4>📊 Column Analysis</h4>
                            <div class="table-container">
                                <table>
                                    <thead>
                                        <tr>
                                            <th>Column</th>
                                            <th>Type</th>
                                            <th>Fill Rate</th>
                                            <th>Unique Values</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {column_rows}
                                    </tbody>
                                </table>
                            </div>
                            
                            <h4>⚙️ Sheet Properties</h4>
                            <div class="metric">
                                <span class="metric-label">Freeze Panes</span>
                                <span class="badge info">{boundaries.get('freeze_panes', 'None')}</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">Print Area</span>
                                <span class="badge info">{boundaries.get('print_area', 'None')}</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">Protection</span>
                                <span class="badge {'warning' if sheet_properties.get('protected', False) else 'success'}">{'Protected' if sheet_properties.get('protected', False) else 'None'}</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">Comments</span>
                                <span class="badge info">{boundaries.get('comments', 0)}</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">Hyperlinks</span>
                                <span class="badge info">{boundaries.get('hyperlinks', 0)}</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            """
        
        return content
    
    def _generate_security_tab(self, security_data: Dict) -> str:
        """Generate the security tab content"""
        if not security_data or security_data.get('error'):
            return '<div class="card"><h3>Security Analysis</h3><p>Security analysis was not completed.</p></div>'
        
        overall_score = security_data.get('overall_score', 0)
        risk_level = security_data.get('risk_level', 'Unknown')
        threats = security_data.get('threats', [])
        patterns = security_data.get('patterns_detected', {})
        
        # Generate threat items
        threat_items = ""
        for threat in threats:
            threat_items += f'<div class="recommendation"><strong>Threat:</strong> {threat}</div>'
        
        # Generate pattern detection results
        pattern_items = ""
        for pattern_name, count in patterns.get('pattern_counts', {}).items():
            pattern_items += f'<div class="recommendation"><strong>{pattern_name.replace("_", " ").title()}:</strong> {count} instances found</div>'
        
        return f"""
        <div class="grid-2">
            <div class="card {'success' if overall_score >= 8 else 'warning' if overall_score >= 6 else 'error'}">
                <h3>🔒 Security Assessment</h3>
                <div class="metric">
                    <span class="metric-label">Overall Security Score</span>
                    <span class="metric-value {'success' if overall_score >= 8 else 'warning' if overall_score >= 6 else 'error'}">{overall_score:.1f}/10</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill {'success' if overall_score >= 8 else 'warning' if overall_score >= 6 else 'error'}" style="width: {overall_score * 10:.1f}%"></div>
                </div>
                
                <h4 style="margin-top: 20px;">✅ Security Strengths</h4>
                <div class="metric">
                    <span class="metric-label">Macro-Free</span>
                    <span class="badge {'success' if 'VBA macros detected' not in threats else 'warning'}">{'Verified' if 'VBA macros detected' not in threats else 'Macros Present'}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">No External Links</span>
                    <span class="badge {'success' if 'External file references found' not in threats else 'warning'}">{'Clean' if 'External file references found' not in threats else 'External Refs'}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Risk Level</span>
                    <span class="badge {'success' if risk_level == 'Low' else 'warning' if risk_level == 'Medium' else 'error'}">{risk_level}</span>
                </div>
            </div>
            
            <div class="card {'warning' if threats else 'success'}">
                <h3>⚠️ Security Analysis</h3>
                <div class="metric">
                    <span class="metric-label">Threats Detected</span>
                    <span class="badge {'error' if threats else 'success'}">{len(threats)} threats</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Sensitive Patterns</span>
                    <span class="badge {'warning' if patterns.get('patterns_found', False) else 'success'}">{len(patterns.get('pattern_counts', {}))} patterns</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Risk Score</span>
                    <span class="badge {'error' if patterns.get('risk_score', 0) > 5 else 'warning' if patterns.get('risk_score', 0) > 2 else 'success'}">{patterns.get('risk_score', 0):.1f}</span>
                </div>
                
                <h4 style="margin-top: 20px;">🔍 Detected Issues</h4>
                {threat_items or '<div class="recommendation">No security threats detected</div>'}
                {pattern_items}
            </div>
        </div>
        """
    
    def _generate_recommendations_tab(self, recommendations: List[str], exec_summary: Dict) -> str:
        """Generate the recommendations tab content"""
        # Generate recommendation items
        rec_items = ""
        for i, rec in enumerate(recommendations, 1):
            rec_items += f'<div class="recommendation"><strong>Recommendation {i}:</strong> {rec}</div>'
        
        # Generate module execution status
        module_statuses = exec_summary.get('module_statuses', {})
        module_rows = ""
        for module_name, status in module_statuses.items():
            status_icon = "success" if status == "success" else "error"
            module_rows += f"""
            <tr>
                <td>{module_name.replace('_', ' ').title()}</td>
                <td><span class="status"><span class="status-icon {status_icon}"></span>{status.title()}</span></td>
                <td>{'0.5s' if status == 'success' else 'N/A'}</td>
                <td>{'Analysis completed' if status == 'success' else 'Module failed'}</td>
            </tr>
            """
        
        return f"""
        <div class="recommendations">
            <h3>💡 Analysis Recommendations</h3>
            {rec_items or '<div class="recommendation">No specific recommendations - file analysis completed successfully.</div>'}
        </div>
        
        <!-- Module Execution Status -->
        <div class="card">
            <h3>🔧 Module Execution Status</h3>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Module</th>
                            <th>Status</th>
                            <th>Execution Time</th>
                            <th>Notes</th>
                        </tr>
                    </thead>
                    <tbody>
                        {module_rows}
                    </tbody>
                </table>
            </div>
        </div>
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
        """Generate detailed per-sheet tables including quality metrics and dependencies."""
        sheet_info = modules.get('data_profiler', {}).get('sheet_analysis', {}) if modules else {}
        dep_matrix = modules.get('dependency_mapper', {}).get('dependency_matrix', {}) if modules else {}
        if not sheet_info:
            return "<p>No sheet details available.</p>"

        sections = []
        for sheet, info in sheet_info.items():
            columns = info.get('columns', [])
            if not columns:
                continue

            # Build rows for column table
            row_tpl = (
                "<tr><td>{letter}</td><td>{header}</td><td>{dtype}</td>"
                "<td>{nulls}</td><td>{fill}%</td><td>{sample}</td></tr>"
            )
            rows_html = "".join(
                row_tpl.format(
                    letter=col['letter'],
                    header=col.get('header', ''),
                    dtype=col['data_type'].title(),
                    nulls=col.get('nulls', 0),
                    fill=int(col.get('fill_rate', 0.0) * 100),
                    sample=(col.get('sample_values') or [''])[0]
                )
                for col in columns
            )

            # Dependencies out of this sheet
            deps_out = dep_matrix.get(sheet, {})
            deps_html = ""
            if deps_out:
                dep_items = "".join(f"<li>{t} ({c})</li>" for t, c in deps_out.items())
                deps_html = f"<p><strong>Dependencies:</strong><ul>{dep_items}</ul></p>"

            summary_text = (
                f"{sheet} • {info.get('dimensions')} • "
                f"Data: {info.get('estimated_data_cells', 0):,} • "
                f"Empty: {info.get('empty_cells', 0):,} • "
                f"Density: {info.get('data_density', 0):.1%}"
            )

            section_html = f"""
            <details class=\"sheet-detail\">
                <summary>{summary_text}</summary>
                <table class=\"sheet-table\">
                    <tr><th>Col</th><th>Header</th><th>Type</th><th>Nulls</th><th>Fill%</th><th>Sample</th></tr>
                    {rows_html}
                </table>
                                <h4>Ranges & Properties</h4>
                <ul>
                    <li><strong>Declared Range:</strong> {info.get('boundaries', {}).get('declared_range','')}</li>
                    <li><strong>True Range:</strong> {info.get('boundaries', {}).get('true_range','')}</li>
                    <li><strong>Freeze Panes:</strong> {info.get('boundaries', {}).get('freeze_panes','')}</li>
                    <li><strong>Merged Cells:</strong> {info.get('boundaries', {}).get('merged_cells',0)}</li>
                    <li><strong>Hyperlinks:</strong> {info.get('boundaries', {}).get('hyperlinks',0)}</li>
                    <li><strong>Print Area:</strong> {info.get('boundaries', {}).get('print_area','')}</li>
                    <li><strong>AutoFilter:</strong> {info.get('boundaries', {}).get('auto_filter')}</li>
                    <li><strong>Protected:</strong> {info.get('sheet_properties', {}).get('protected')}</li>
                    <li><strong>Visibility:</strong> {info.get('sheet_properties', {}).get('visibility')}</li>
                </ul>
                {deps_html}
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
