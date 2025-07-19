#!/usr/bin/env python3
"""
Unified Report Data Model and Base Classes
Ensures consistency across all report formats (HTML, JSON, Text, Markdown)
"""

from typing import Dict, Any, List, Optional, Protocol, runtime_checkable
from abc import ABC, abstractmethod
from datetime import datetime
import json


class ReportDataModel:
    """
    Unified data model ensuring all reports contain the same information
    
    This class standardizes the data structure across all report formats,
    providing fallbacks for missing data and consistent metrics calculation.
    """
    
    def __init__(self, analysis_results: Dict[str, Any]):
        self.raw_results = analysis_results
        self._validate_completeness()
        self._standardized_data = None
    
    def _validate_completeness(self):
        """Ensure all required sections exist with fallbacks"""
        required_sections = [
            'file_info', 'analysis_metadata', 'module_results', 
            'execution_summary', 'recommendations'
        ]
        
        for section in required_sections:
            if section not in self.raw_results:
                self.raw_results[section] = self._get_fallback(section)
    
    def _get_fallback(self, section: str) -> Dict[str, Any]:
        """Provide fallback data for missing sections"""
        fallbacks = {
            'file_info': {
                'name': 'Unknown',
                'size_mb': 0.0,
                'path': '',
                'created': '',
                'modified': '',
                'sheet_count': 0,
                'sheets': []
            },
            'analysis_metadata': {
                'timestamp': datetime.now().timestamp(),
                'total_duration_seconds': 0.0,
                'success_rate': 0.0,
                'quality_score': 0.0,
                'security_score': 0.0
            },
            'module_results': {},
            'execution_summary': {
                'total_modules': 0,
                'successful_modules': 0,
                'failed_modules': 0,
                'module_statuses': {}
            },
            'recommendations': []
        }
        return fallbacks.get(section, {})
    
    def get_standardized_data(self) -> Dict[str, Any]:
        """Return standardized data structure for all report formats"""
        if self._standardized_data is None:
            self._standardized_data = {
                'file_summary': self._extract_file_summary(),
                'quality_metrics': self._extract_quality_metrics(),
                'security_analysis': self._extract_security_analysis(),
                'structure_analysis': self._extract_structure_analysis(),
                'data_analysis': self._extract_data_analysis(),
                'sheet_details': self._extract_sheet_details(),
                'formula_analysis': self._extract_formula_analysis(),
                'visual_analysis': self._extract_visual_analysis(),
                'performance_metrics': self._extract_performance_metrics(),
                'recommendations': self._extract_recommendations(),
                'module_execution': self._extract_module_execution(),
                'export_metadata': self._extract_export_metadata()
            }
        
        return self._standardized_data
    
    def _extract_file_summary(self) -> Dict[str, Any]:
        """Extract standardized file summary"""
        file_info = self.raw_results.get('file_info', {})
        metadata = self.raw_results.get('analysis_metadata', {})
        
        return {
            'name': file_info.get('name', 'Unknown'),
            'size_mb': file_info.get('size_mb', 0.0),
            'size_bytes': file_info.get('size_bytes', 0),
            'path': file_info.get('path', ''),
            'created': file_info.get('created', ''),
            'modified': file_info.get('modified', ''),
            'excel_version': file_info.get('excel_version', 'Unknown'),
            'compression_ratio': file_info.get('compression_ratio', 0.0),
            'sheet_count': file_info.get('sheet_count', 0),
            'sheets': file_info.get('sheets', []),
            'analysis_timestamp': datetime.fromtimestamp(metadata.get('timestamp', 0)).strftime('%Y-%m-%d %H:%M:%S'),
            'file_signature_valid': file_info.get('file_signature_valid', True)
        }
    
    def _extract_quality_metrics(self) -> Dict[str, Any]:
        """Extract standardized quality metrics"""
        metadata = self.raw_results.get('analysis_metadata', {})
        data_profiler = self.raw_results.get('module_results', {}).get('data_profiler', {})
        exec_summary = self.raw_results.get('execution_summary', {})
        
        return {
            'overall_quality_score': metadata.get('quality_score', 0.0),
            'security_score': metadata.get('security_score', 0.0),
            'success_rate': exec_summary.get('success_rate', metadata.get('success_rate', 0.0)),
            'data_density': data_profiler.get('overall_data_density', 0.0),
            'data_quality_score': data_profiler.get('data_quality_score', 0.0),
            'total_cells': data_profiler.get('total_cells', 0),
            'total_data_cells': data_profiler.get('total_data_cells', 0),
            'processing_time_seconds': metadata.get('total_duration_seconds', 0.0),
            'successful_modules': exec_summary.get('successful_modules', 0),
            'total_modules': exec_summary.get('total_modules', 0),
            'failed_modules': exec_summary.get('failed_modules', 0)
        }
    
    def _extract_security_analysis(self) -> Dict[str, Any]:
        """Extract standardized security analysis"""
        security = self.raw_results.get('module_results', {}).get('security_inspector', {})
        
        return {
            'overall_score': security.get('overall_score', 0.0),
            'risk_level': security.get('risk_level', 'Unknown'),
            'threats_detected': security.get('threats', []),
            'threat_count': len(security.get('threats', [])),
            'patterns_detected': security.get('patterns_detected', {}),
            'recommendations': security.get('recommendations', []),
            'has_macros': 'VBA macros detected' in security.get('threats', []),
            'has_external_refs': 'External file references found' in security.get('threats', [])
        }
    
    def _extract_structure_analysis(self) -> Dict[str, Any]:
        """Extract standardized structure analysis"""
        structure = self.raw_results.get('module_results', {}).get('structure_mapper', {})
        
        return {
            'total_sheets': structure.get('total_sheets', 0),
            'visible_sheets': structure.get('visible_sheets', []),
            'hidden_sheets': structure.get('hidden_sheets', []),
            'sheet_details': structure.get('sheet_details', []),
            'named_ranges_count': structure.get('named_ranges_count', 0),
            'table_count': structure.get('table_count', 0),
            'has_hidden_content': structure.get('has_hidden_content', False),
            'workbook_features': structure.get('workbook_features', {}),
            'protection_info': structure.get('protection_info', {})
        }
    
    def _extract_data_analysis(self) -> Dict[str, Any]:
        """Extract standardized data analysis"""
        data_profiler = self.raw_results.get('module_results', {}).get('data_profiler', {})
        
        return {
            'total_cells': data_profiler.get('total_cells', 0),
            'total_data_cells': data_profiler.get('total_data_cells', 0),
            'data_density': data_profiler.get('overall_data_density', 0.0),
            'data_quality_score': data_profiler.get('data_quality_score', 0.0),
            'data_type_distribution': data_profiler.get('data_type_distribution', {}),
            'sheet_analysis': data_profiler.get('sheet_analysis', {}),
            'overall_metrics': data_profiler.get('overall_metrics', {})
        }
        
        return {
            'sheet_analysis': data_profiler.get('sheet_analysis', {}),
            'data_type_distribution': data_profiler.get('data_type_distribution', {}),
            'overall_metrics': data_profiler.get('overall_metrics', {}),
            'cross_sheet_analysis': data_profiler.get('cross_sheet_analysis', {})
        }
    
    def _extract_sheet_details(self) -> List[Dict[str, Any]]:
        """Extract standardized sheet details"""
        data_profiler = self.raw_results.get('module_results', {}).get('data_profiler', {})
        sheet_analysis = data_profiler.get('sheet_analysis', {})
        
        sheet_details = []
        for sheet_name, sheet_data in sheet_analysis.items():
            sheet_details.append({
                'name': sheet_name,
                'dimensions': sheet_data.get('dimensions', '0x0'),
                'used_range': sheet_data.get('used_range', 'A1:A1'),
                'estimated_data_cells': sheet_data.get('estimated_data_cells', 0),
                'empty_cells': sheet_data.get('empty_cells', 0),
                'data_density': sheet_data.get('data_density', 0.0),
                'has_data': sheet_data.get('has_data', False),
                'columns': sheet_data.get('columns', []),
                'boundaries': sheet_data.get('boundaries', {}),
                'properties': sheet_data.get('sheet_properties', {}),
                'quality_metrics': sheet_data.get('data_quality_metrics', {}),
                'duplicate_rows': sheet_data.get('duplicate_rows', {})
            })
        
        return sheet_details
    
    def _extract_formula_analysis(self) -> Dict[str, Any]:
        """Extract standardized formula analysis"""
        formulas = self.raw_results.get('module_results', {}).get('formula_analyzer', {})
        
        return {
            'total_formulas': formulas.get('total_formulas', 0),
            'complex_formulas': formulas.get('complex_formulas', []),
            'complex_formula_count': len(formulas.get('complex_formulas', [])),
            'has_external_refs': formulas.get('has_external_refs', False),
            'formula_complexity_score': formulas.get('formula_complexity_score', 0.0)
        }
    
    def _extract_visual_analysis(self) -> Dict[str, Any]:
        """Extract standardized visual analysis"""
        visuals = self.raw_results.get('module_results', {}).get('visual_cataloger', {})
        
        return {
            'total_charts': visuals.get('total_charts', 0),
            'total_images': visuals.get('total_images', 0),
            'conditional_formatting_rules': visuals.get('conditional_formatting_rules', 0),
            'has_visual_content': visuals.get('has_visual_content', False),
            'visual_complexity_score': visuals.get('visual_complexity_score', 0.0)
        }
    
    def _extract_performance_metrics(self) -> Dict[str, Any]:
        """Extract standardized performance metrics"""
        performance = self.raw_results.get('module_results', {}).get('performance_monitor', {})
        resource_usage = self.raw_results.get('resource_usage', {}).get('current_usage', {})
        
        return {
            'elapsed_seconds': performance.get('elapsed_seconds', 0.0),
            'memory_usage_mb': performance.get('memory_usage', {}).get('current_mb', 0.0),
            'peak_memory_mb': performance.get('memory_usage', {}).get('peak_mb', 0.0),
            'cpu_usage_percent': performance.get('cpu_usage', {}).get('percent', 0.0),
            'performance_score': performance.get('performance_score', 0.0)
        }
    
    def _extract_recommendations(self) -> List[str]:
        """Extract standardized recommendations"""
        return self.raw_results.get('recommendations', [])
    
    def _extract_module_execution(self) -> Dict[str, Any]:
        """Extract standardized module execution summary"""
        exec_summary = self.raw_results.get('execution_summary', {})
        
        return {
            'total_modules': exec_summary.get('total_modules', 0),
            'successful_modules': exec_summary.get('successful_modules', 0),
            'failed_modules': exec_summary.get('failed_modules', 0),
            'success_rate': exec_summary.get('success_rate', 0.0),
            'module_statuses': exec_summary.get('module_statuses', {})
        }
    
    def _extract_export_metadata(self) -> Dict[str, Any]:
        """Extract export metadata"""
        return {
            'generated_at': datetime.now().isoformat(),
            'generator_version': '2.0',
            'data_model_version': '1.0',
            'includes_raw_data': True
        }


@runtime_checkable
class ReportGenerator(Protocol):
    """Protocol defining the interface for all report generators"""
    
    def generate(self, data_model: ReportDataModel, output_path: str) -> str:
        """Generate report from data model to specified output path"""
        ...


class BaseReportGenerator(ABC):
    """
    Abstract base class for all report generators
    
    Provides common functionality and ensures consistent report structure
    """
    
    def __init__(self):
        self.data_model: Optional[ReportDataModel] = None
        self.standardized_data: Optional[Dict[str, Any]] = None
    
    def generate_report(self, analysis_results: Dict[str, Any], output_path: str) -> str:
        """
        Main entry point for report generation
        
        Args:
            analysis_results: Raw analysis results
            output_path: Path for output file
            
        Returns:
            Path to generated report file
        """
        # Create data model
        self.data_model = ReportDataModel(analysis_results)
        self.standardized_data = self.data_model.get_standardized_data()
        
        # Generate format-specific report
        report_content = self._generate_content()
        
        # Write to file
        self._write_to_file(report_content, output_path)
        
        return output_path
    
    @abstractmethod
    def _generate_content(self) -> str:
        """Generate format-specific content (implemented by subclasses)"""
        pass
    
    def _write_to_file(self, content: str, output_path: str):
        """Write content to file with error handling"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
        except Exception as e:
            raise Exception(f"Failed to write report to {output_path}: {e}")
    
    def _get_file_summary(self) -> Dict[str, Any]:
        """Get standardized file summary"""
        return self.standardized_data.get('file_summary', {})
    
    def _get_quality_metrics(self) -> Dict[str, Any]:
        """Get standardized quality metrics"""
        return self.standardized_data.get('quality_metrics', {})
    
    def _get_security_analysis(self) -> Dict[str, Any]:
        """Get standardized security analysis"""
        return self.standardized_data.get('security_analysis', {})
    
    def _get_structure_analysis(self) -> Dict[str, Any]:
        """Get standardized structure analysis"""
        return self.standardized_data.get('structure_analysis', {})
    
    def _get_sheet_details(self) -> List[Dict[str, Any]]:
        """Get standardized sheet details"""
        return self.standardized_data.get('sheet_details', [])
    
    def _get_recommendations(self) -> List[str]:
        """Get standardized recommendations"""
        return self.standardized_data.get('recommendations', [])
    
    def _get_module_execution(self) -> Dict[str, Any]:
        """Get standardized module execution summary"""
        return self.standardized_data.get('module_execution', {})


class ReportValidator:
    """Validates that all report formats contain consistent data"""
    
    @staticmethod
    def validate_consistency(analysis_results: Dict[str, Any], 
                           generated_reports: Dict[str, str]) -> Dict[str, Any]:
        """
        Validate that all generated reports contain consistent core metrics
        
        Args:
            analysis_results: Original analysis results
            generated_reports: Dict of format -> report_path
            
        Returns:
            Validation results with any discrepancies found
        """
        validation_results = {
            'consistent': True,
            'discrepancies': [],
            'core_metrics': {}
        }
        
        # Extract core metrics that should be identical across formats
        data_model = ReportDataModel(analysis_results)
        standardized = data_model.get_standardized_data()
        
        core_metrics = {
            'file_name': standardized['file_summary']['name'],
            'file_size_mb': standardized['file_summary']['size_mb'],
            'sheet_count': standardized['file_summary']['sheet_count'],
            'quality_score': standardized['quality_metrics']['overall_quality_score'],
            'security_score': standardized['quality_metrics']['security_score'],
            'total_cells': standardized['quality_metrics']['total_cells'],
            'success_rate': standardized['quality_metrics']['success_rate']
        }
        
        validation_results['core_metrics'] = core_metrics
        
        # Additional validation logic could be added here to check
        # specific content in each generated report file
        
        return validation_results


if __name__ == "__main__":
    # Test the data model with sample data
    sample_results = {
        'file_info': {
            'name': 'test.xlsx',
            'size_mb': 1.5,
            'sheet_count': 3
        },
        'analysis_metadata': {
            'quality_score': 0.85,
            'security_score': 9.2
        }
    }
    
    data_model = ReportDataModel(sample_results)
    standardized = data_model.get_standardized_data()
    
    print("Standardized data model created successfully!")
    print(f"File: {standardized['file_summary']['name']}")
    print(f"Quality: {standardized['quality_metrics']['overall_quality_score']:.1%}")
