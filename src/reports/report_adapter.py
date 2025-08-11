"""
Report Adapter - Bridge between new AnalysisService results and existing ReportDataModel
"""

from typing import Dict, Any
from datetime import datetime
from reports.report_base import ReportDataModel


class AnalysisResultsAdapter:
    """Adapts AnalysisService results to work with existing ReportDataModel"""
    
    @staticmethod
    def adapt_results(analysis_results: Dict[str, Any]) -> Dict[str, Any]:
        """
        Transform AnalysisService results into format expected by ReportDataModel
        
        Args:
            analysis_results: Results from AnalysisService.analyze_file()
            
        Returns:
            Dictionary formatted for ReportDataModel
        """
        # Extract components from AnalysisService results
        file_info = analysis_results.get('file_info', {})
        analysis_metadata = analysis_results.get('analysis_metadata', {})
        module_results = analysis_results.get('module_results', {})
        summary = analysis_results.get('summary', {})
        recommendations = analysis_results.get('recommendations', [])
        
        # Adapt to ReportDataModel expected format
        adapted_results = {
            'file_info': AnalysisResultsAdapter._adapt_file_info(file_info),
            'analysis_metadata': AnalysisResultsAdapter._adapt_metadata(analysis_metadata),
            'module_results': AnalysisResultsAdapter._adapt_module_results(module_results),
            'execution_summary': AnalysisResultsAdapter._adapt_execution_summary(analysis_metadata, summary),
            'recommendations': recommendations
        }
        
        return adapted_results
    
    @staticmethod
    def _adapt_file_info(file_info: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt file information"""
        return {
            'name': file_info.get('file_name', 'Unknown'),
            'size_mb': file_info.get('size_mb', 0.0),
            'size_bytes': file_info.get('file_size_bytes', 0),
            'path': file_info.get('file_path', ''),
            'created': '',  # Not available from new analysis
            'modified': datetime.fromtimestamp(file_info.get('modified_time', 0)).strftime('%Y-%m-%d %H:%M:%S') if file_info.get('modified_time') else '',
            'excel_version': 'Unknown',  # Not analyzed in new system
            'compression_ratio': 0.0,  # Not analyzed in new system
            'sheet_count': file_info.get('total_sheets', 0),
            'sheets': file_info.get('sheet_names', []),
            'file_signature_valid': True,  # Assume valid since file opened successfully
            'active_sheet': file_info.get('active_sheet', '')
        }
    
    @staticmethod
    def _adapt_metadata(metadata: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt analysis metadata"""
        return {
            'timestamp': metadata.get('timestamp', datetime.now().timestamp()),
            'total_duration_seconds': metadata.get('total_duration_seconds', 0.0),
            'success_rate': metadata.get('success_rate', 0.0),
            'quality_score': metadata.get('quality_score', 0.0),
            'security_score': 0.0,  # Not implemented in new system yet
            'successful_modules': metadata.get('successful_modules', []),
            'failed_modules': metadata.get('failed_modules', []),
            'module_execution_times': metadata.get('module_execution_times', {})
        }
    
    @staticmethod
    def _adapt_module_results(module_results: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt module results to expected format"""
        adapted = {}
        
        # Map new module names to old expected names
        module_mapping = {
            'structure': 'structure_mapper',
            'data': 'data_profiler', 
            'formula': 'formula_analyzer'
        }
        
        for new_name, result in module_results.items():
            old_name = module_mapping.get(new_name, new_name)
            adapted[old_name] = AnalysisResultsAdapter._adapt_individual_module(old_name, result)
        
        # Add missing expected modules with empty results
        expected_modules = [
            'health_checker', 'structure_mapper', 'data_profiler', 
            'formula_analyzer', 'visual_cataloger', 'security_inspector',
            'dependency_mapper', 'relationship_analyzer', 'performance_monitor',
            'connection_inspector', 'pivot_intelligence', 'doc_synthesizer'
        ]
        
        for module in expected_modules:
            if module not in adapted:
                adapted[module] = {'status': 'not_run', 'execution_time': 0.0}
        
        return adapted
    
    @staticmethod
    def _adapt_individual_module(module_name: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt individual module result"""
        if 'error' in result:
            return {
                'status': 'failed',
                'error': result['error'],
                'execution_time': result.get('analysis_duration', 0.0)
            }
        
        # Module-specific adaptations
        if module_name == 'structure_mapper':
            return AnalysisResultsAdapter._adapt_structure_result(result)
        elif module_name == 'data_profiler':
            return AnalysisResultsAdapter._adapt_data_result(result)
        elif module_name == 'formula_analyzer':
            return AnalysisResultsAdapter._adapt_formula_result(result)
        else:
            # Generic adaptation
            adapted = result.copy()
            adapted['status'] = 'success'
            return adapted
    
    @staticmethod
    def _adapt_structure_result(result: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt structure analysis result"""
        return {
            'status': 'success',
            'execution_time': result.get('analysis_duration', 0.0),
            'total_sheets': result.get('total_sheets', 0),
            'visible_sheets': result.get('visible_sheets', []),
            'hidden_sheets': result.get('hidden_sheets', []),
            'sheet_details': result.get('sheet_details', []),
            'named_ranges_count': result.get('named_ranges_count', 0),
            'named_ranges_list': result.get('named_ranges_list', []),
            'table_count': result.get('table_count', 0),
            'table_details': result.get('table_details', []),
            'has_hidden_content': result.get('has_hidden_content', False),
            'workbook_features': result.get('workbook_features', {}),
            'protection_info': result.get('protection_info', {})
        }
    
    @staticmethod
    def _adapt_data_result(result: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt data analysis result"""
        return {
            'status': 'success',
            'execution_time': result.get('analysis_duration', 0.0),
            'sheet_analysis': result.get('sheet_analysis', {}),
            'total_cells': result.get('total_cells', 0),
            'total_data_cells': result.get('total_data_cells', 0),
            'overall_data_density': result.get('overall_data_density', 0.0),
            'data_quality_score': result.get('data_quality_score', 0.0),
            'data_type_distribution': result.get('data_type_distribution', {}),
            'cross_sheet_analysis': result.get('cross_sheet_analysis', {}),
            'overall_metrics': {
                'quality_score': result.get('data_quality_score', 0.0),
                'data_variety_score': min(1.0, len(result.get('data_type_distribution', {})) / 6.0)
            }
        }
    
    @staticmethod
    def _adapt_formula_result(result: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt formula analysis result"""
        return {
            'status': 'success',
            'execution_time': result.get('analysis_duration', 0.0),
            'total_formulas': result.get('total_formulas', 0),
            'complex_formulas': result.get('complex_formulas', []),
            'has_external_refs': result.get('has_external_refs', False),
            'formula_complexity_score': result.get('formula_complexity_score', 0.0),
            'function_usage': result.get('function_usage', {}),
            'function_diversity': result.get('function_diversity', 0),
            'most_used_functions': result.get('most_used_functions', []),
            'sheet_statistics': result.get('sheet_statistics', {}),
            'circular_references': result.get('circular_references', []),
            'performance_impact': result.get('performance_impact', {})
        }
    
    @staticmethod
    def _adapt_execution_summary(metadata: Dict[str, Any], summary: Dict[str, Any]) -> Dict[str, Any]:
        """Adapt execution summary"""
        successful_modules = len(metadata.get('successful_modules', []))
        failed_modules = len(metadata.get('failed_modules', []))
        total_modules = successful_modules + failed_modules
        
        return {
            'total_modules': total_modules,
            'successful_modules': successful_modules,
            'failed_modules': failed_modules,
            'success_rate': metadata.get('success_rate', 0.0),
            'module_statuses': {
                name: 'success' for name in metadata.get('successful_modules', [])
            } | {
                name: 'failed' for name in metadata.get('failed_modules', [])
            }
        }


class ReportService:
    """Service for generating reports from AnalysisService results"""
    
    def __init__(self):
        self.adapter = AnalysisResultsAdapter()
    
    def create_report_model(self, analysis_results: Dict[str, Any]) -> ReportDataModel:
        """
        Create a ReportDataModel from AnalysisService results
        
        Args:
            analysis_results: Results from AnalysisService.analyze_file()
            
        Returns:
            ReportDataModel instance ready for report generation
        """
        adapted_results = self.adapter.adapt_results(analysis_results)
        return ReportDataModel(adapted_results)
    
    def generate_report(self, analysis_results: Dict[str, Any], format_type: str, output_path: str) -> str:
        """
        Generate a report in the specified format
        
        Args:
            analysis_results: Results from AnalysisService.analyze_file()
            format_type: Report format (html, json, text, markdown)
            output_path: Path for output file
            
        Returns:
            Path to generated report file
        """
        # Import generators here to avoid circular imports
        from reports.report_generator import ReportGenerator
        from reports.comprehensive_text_report import ComprehensiveTextReportGenerator
        
        # Create report model
        report_model = self.create_report_model(analysis_results)
        standardized_data = report_model.get_standardized_data()
        
        if format_type == 'html':
            generator = ReportGenerator()
            generator.generate_html_report(analysis_results, output_path)  # Use original results for HTML
            return output_path
            
        elif format_type == 'json':
            generator = ReportGenerator()
            generator.generate_json_report(standardized_data, output_path)
            return output_path
            
        elif format_type in ['text', 'markdown']:
            generator = ComprehensiveTextReportGenerator()
            if format_type == 'markdown':
                return generator.generate_markdown_report(analysis_results, output_path)
            else:
                return generator.generate_text_report(analysis_results, output_path)
        
        else:
            raise ValueError(f"Unsupported format type: {format_type}")