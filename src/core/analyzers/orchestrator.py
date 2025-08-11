"""
Analyzer Orchestrator - Coordinates execution of all analyzer modules
"""

import time
from typing import Dict, Any, List, Optional
import openpyxl
from ..config import load_config
from .structure import StructureAnalyzer
from .data import DataAnalyzer
from .formula import FormulaAnalyzer


class AnalyzerOrchestrator:
    """Orchestrates the execution of all analysis modules"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the orchestrator with configuration
        
        Args:
            config: Configuration dictionary, loads default if None
        """
        self.config = config or load_config()
        self.analyzers = self._initialize_analyzers()
        self.results = {}
        self.execution_times = {}
        self.errors = {}
        
    def _initialize_analyzers(self) -> Dict[str, Any]:
        """Initialize all analyzer instances"""
        return {
            'structure': StructureAnalyzer(self.config),
            'data': DataAnalyzer(self.config),
            'formula': FormulaAnalyzer(self.config)
        }
    
    def analyze_workbook(self, workbook: openpyxl.Workbook, 
                        modules: Optional[List[str]] = None,
                        progress_callback: Optional[callable] = None) -> Dict[str, Any]:
        """
        Perform comprehensive analysis using all or selected modules
        
        Args:
            workbook: Loaded openpyxl workbook
            modules: List of modules to run, runs all if None
            progress_callback: Optional callback for progress updates
            
        Returns:
            Dictionary containing all analysis results
        """
        start_time = time.time()
        
        # Determine which modules to run
        modules_to_run = modules or list(self.analyzers.keys())
        
        # Reset results for this analysis
        self.results = {}
        self.execution_times = {}
        self.errors = {}
        
        total_modules = len(modules_to_run)
        
        for i, module_name in enumerate(modules_to_run):
            if progress_callback:
                progress_callback(f"Running {module_name} analysis...", i / total_modules)
            
            self._execute_module(module_name, workbook)
        
        if progress_callback:
            progress_callback("Compiling results...", 1.0)
        
        # Compile final results
        final_results = self._compile_results(workbook, start_time)
        
        return final_results
    
    def _execute_module(self, module_name: str, workbook: openpyxl.Workbook):
        """Execute a single analyzer module with error handling"""
        if module_name not in self.analyzers:
            self.errors[module_name] = f"Unknown module: {module_name}"
            return
        
        analyzer = self.analyzers[module_name]
        module_start = time.time()
        
        try:
            # Set timeout if configured
            timeout = self.config.get('analysis', {}).get('module_timeout_seconds', 60)
            
            # Execute the analysis
            result = analyzer.analyze(workbook)
            
            # Store successful result
            self.results[module_name] = result
            self.execution_times[module_name] = time.time() - module_start
            
        except Exception as e:
            # Log error but continue with other modules
            error_msg = f"Error in {module_name}: {str(e)}"
            self.errors[module_name] = error_msg
            self.execution_times[module_name] = time.time() - module_start
            
            # Store partial result to maintain structure
            self.results[module_name] = {
                'error': error_msg,
                'analysis_duration': self.execution_times[module_name],
                'status': 'failed'
            }
    
    def _compile_results(self, workbook: openpyxl.Workbook, start_time: float) -> Dict[str, Any]:
        """Compile all results into final analysis report"""
        total_duration = time.time() - start_time
        
        # Calculate success metrics
        successful_modules = [name for name in self.results.keys() if name not in self.errors]
        failed_modules = list(self.errors.keys())
        success_rate = len(successful_modules) / len(self.analyzers) if self.analyzers else 0
        
        # Extract key metrics from each module
        structure_data = self.results.get('structure', {})
        data_analysis = self.results.get('data', {})
        formula_analysis = self.results.get('formula', {})
        
        # Calculate overall quality score
        quality_score = self._calculate_overall_quality(structure_data, data_analysis, formula_analysis)
        
        # Generate recommendations
        recommendations = self._generate_recommendations(structure_data, data_analysis, formula_analysis)
        
        return {
            'analysis_metadata': {
                'timestamp': time.time(),
                'total_duration_seconds': total_duration,
                'success_rate': success_rate,
                'successful_modules': successful_modules,
                'failed_modules': failed_modules,
                'module_execution_times': self.execution_times,
                'quality_score': quality_score
            },
            'file_info': self._extract_file_info(workbook),
            'module_results': self.results,
            'errors': self.errors if self.errors else None,
            'summary': {
                'total_sheets': structure_data.get('total_sheets', 0),
                'total_data_cells': data_analysis.get('total_data_cells', 0),
                'total_formulas': formula_analysis.get('total_formulas', 0),
                'data_quality_score': data_analysis.get('data_quality_score', 0),
                'formula_complexity_score': formula_analysis.get('formula_complexity_score', 0),
                'has_hidden_content': structure_data.get('has_hidden_content', False),
                'has_external_refs': formula_analysis.get('has_external_refs', False)
            },
            'recommendations': recommendations
        }
    
    def _extract_file_info(self, workbook: openpyxl.Workbook) -> Dict[str, Any]:
        """Extract basic file information"""
        return {
            'total_sheets': len(workbook.sheetnames),
            'sheet_names': workbook.sheetnames,
            'active_sheet': workbook.active.title if workbook.active else None,
            'has_properties': bool(getattr(workbook, 'properties', None))
        }
    
    def _calculate_overall_quality(self, structure: Dict, data: Dict, formula: Dict) -> float:
        """Calculate overall workbook quality score"""
        scores = []
        
        # Data quality component (40%)
        if 'data_quality_score' in data:
            scores.append(data['data_quality_score'] * 0.4)
        
        # Structure quality component (30%)
        structure_score = 1.0
        if structure.get('has_hidden_content'):
            structure_score -= 0.2
        if len(structure.get('sheet_details', [])) > 10:  # Too many sheets
            structure_score -= 0.1
        scores.append(structure_score * 0.3)
        
        # Formula quality component (30%)
        formula_score = 1.0 - formula.get('formula_complexity_score', 0)
        if formula.get('has_external_refs'):
            formula_score -= 0.1
        scores.append(formula_score * 0.3)
        
        return max(0.0, min(1.0, sum(scores)))
    
    def _generate_recommendations(self, structure: Dict, data: Dict, formula: Dict) -> List[str]:
        """Generate actionable recommendations based on analysis"""
        recommendations = []
        
        # Data quality recommendations
        data_quality = data.get('data_quality_score', 1.0)
        if data_quality < 0.5:
            recommendations.append("Low data quality detected - consider data cleanup and validation")
        elif data_quality < 0.7:
            recommendations.append("Moderate data quality issues found - review data consistency")
        
        # Structure recommendations
        if structure.get('has_hidden_content'):
            recommendations.append("Hidden sheets detected - review for sensitive information")
        
        if len(structure.get('sheet_details', [])) > 15:
            recommendations.append("Many sheets detected - consider consolidating or organizing data")
        
        # Formula recommendations
        formula_complexity = formula.get('formula_complexity_score', 0)
        if formula_complexity > 0.7:
            recommendations.append("High formula complexity detected - consider simplification for maintainability")
        
        if formula.get('has_external_refs'):
            recommendations.append("External references found - verify linked files are available")
        
        # Performance recommendations
        total_formulas = formula.get('total_formulas', 0)
        if total_formulas > 1000:
            recommendations.append("High formula count may impact performance - consider optimization")
        
        # Data density recommendations
        data_density = data.get('overall_data_density', 0)
        if data_density < 0.1:
            recommendations.append("Low data density - consider removing empty regions to reduce file size")
        
        # Default recommendation if none found
        if not recommendations:
            recommendations.append("Workbook structure and content appear well-optimized")
        
        return recommendations
    
    def get_module_status(self, module_name: str) -> Dict[str, Any]:
        """Get detailed status for a specific module"""
        if module_name not in self.analyzers:
            return {'status': 'unknown', 'error': 'Module not found'}
        
        if module_name in self.errors:
            return {
                'status': 'failed',
                'error': self.errors[module_name],
                'execution_time': self.execution_times.get(module_name, 0)
            }
        elif module_name in self.results:
            return {
                'status': 'success',
                'execution_time': self.execution_times.get(module_name, 0),
                'result_keys': list(self.results[module_name].keys())
            }
        else:
            return {'status': 'pending'}
    
    def get_available_modules(self) -> List[str]:
        """Get list of all available analyzer modules"""
        return list(self.analyzers.keys())