"""
Enhanced Excel Explorer Orchestrator
Simplified, reliable orchestrator that ensures all modules execute properly.
"""

import time
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable, Union
from dataclasses import dataclass, field

from src.core.analysis_context import AnalysisContext, AnalysisConfig
from src.core.base_analyzer import BaseAnalyzer
from src.utils.config_loader import load_config
from src.utils.error_handler import initialize_error_handler
from src.utils.memory_manager import initialize_memory_manager

# Import all modules
from src.modules.health_checker import HealthChecker
from src.modules.structure_mapper import StructureMapper
from src.modules.data_profiler import DataProfiler
from src.modules.formula_analyzer import FormulaAnalyzer
from src.modules.visual_cataloger import VisualCataloger
from src.modules.connection_inspector import ConnectionInspector
from src.modules.pivot_intelligence import PivotIntelligence
from src.modules.doc_synthesizer import DocSynthesizer


@dataclass
class ModuleExecutionResult:
    """Result of a single module execution"""
    module_name: str
    success: bool
    data: Any = None
    error_message: str = ""
    duration_seconds: float = 0.0
    memory_delta_mb: float = 0.0
    warnings: List[str] = field(default_factory=list)


@dataclass
class AnalysisResults:
    """Complete analysis results"""
    file_info: Dict[str, Any]
    analysis_metadata: Dict[str, Any]
    module_results: Dict[str, Any]
    execution_summary: Dict[str, Any]
    resource_usage: Dict[str, Any]
    recommendations: List[str]
    success: bool = True
    error_message: str = ""


class EnhancedExcelExplorer:
    """Simplified, reliable Excel analysis orchestrator"""
    
    # Module registry with execution order
    MODULES = [
        ('health_checker', HealthChecker, "Checking file integrity and security"),
        ('structure_mapper', StructureMapper, "Mapping workbook structure and sheets"),
        ('data_profiler', DataProfiler, "Profiling data types and quality"),
        ('formula_analyzer', FormulaAnalyzer, "Analyzing formulas and dependencies"),
        ('visual_cataloger', VisualCataloger, "Cataloging charts and visual elements"),
        ('connection_inspector', ConnectionInspector, "Inspecting external connections"),
        ('pivot_intelligence', PivotIntelligence, "Analyzing pivot tables and summaries"),
        ('doc_synthesizer', DocSynthesizer, "Synthesizing comprehensive documentation")
    ]
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize the enhanced orchestrator"""
        # Load configuration
        self.config_data = load_config(config_path)
        self.analysis_config = AnalysisConfig(**self.config_data.get('analysis', {}))
        
        # Initialize systems
        self.error_handler = initialize_error_handler()
        self.memory_manager = initialize_memory_manager(
            self.analysis_config.max_memory_mb,
            self.analysis_config.warning_threshold
        )
        
        # Initialize logger
        self.logger = logging.getLogger(__name__)
        
        # Progress callback
        self.progress_callback: Optional[Callable] = None
        
    def analyze_file(self, file_path: str, progress_callback: Optional[Callable] = None) -> AnalysisResults:
        """
        Execute comprehensive analysis of Excel file
        
        Args:
            file_path: Path to Excel file
            progress_callback: Optional callback for progress updates
            
        Returns:
            Complete analysis results
        """
        self.progress_callback = progress_callback
        file_path = Path(file_path)
        
        start_time = time.time()
        
        try:
            # Initialize analysis context
            with AnalysisContext(file_path, self.analysis_config) as context:
                self.logger.info(f"Starting analysis of {file_path.name}")
                
                # Execute all modules
                module_results = self._execute_all_modules(context)
                
                # Calculate totals
                total_duration = time.time() - start_time
                successful_modules = sum(1 for result in module_results.values() if result.success)
                total_modules = len(module_results)
                
                # Generate final results as dictionary for GUI compatibility
                return self._generate_results_dict(context, module_results, total_duration, successful_modules, total_modules)
                
        except Exception as e:
            self.logger.error(f"Analysis failed: {e}")
            return self._generate_error_results(file_path, str(e), time.time() - start_time)
    
    def _execute_all_modules(self, context: AnalysisContext) -> Dict[str, ModuleExecutionResult]:
        """Execute all modules in sequence, continuing even if some fail"""
        results = {}
        
        for module_name, module_class, description in self.MODULES:
            try:
                # Signal module start
                if self.progress_callback:
                    self.progress_callback(module_name, "starting", description)
                
                # Execute module
                result = self._execute_single_module(module_name, module_class, context)
                results[module_name] = result
                
                # Signal completion
                if self.progress_callback:
                    if result.success:
                        self.progress_callback(module_name, "complete", "")
                    else:
                        error_detail = result.error_message or "Module execution failed"
                        self.progress_callback(module_name, "error", error_detail)
                
            except Exception as e:
                # Record failure but continue
                error_msg = str(e)
                self.logger.error(f"Module {module_name} failed: {error_msg}")
                
                results[module_name] = ModuleExecutionResult(
                    module_name=module_name,
                    success=False,
                    error_message=error_msg
                )
                
                if self.progress_callback:
                    self.progress_callback(module_name, "error", error_msg)
        
        return results
    
    def _execute_single_module(self, module_name: str, module_class: type, context: AnalysisContext) -> ModuleExecutionResult:
        """Execute a single analysis module with granular progress"""
        start_time = time.time()
        start_memory = self.memory_manager.get_current_usage()['current_mb']
        
        try:
            # Get module configuration
            module_config = self.config_data.get(module_name, {})
            
            # Skip if explicitly disabled
            if not module_config.get('enabled', True):
                return ModuleExecutionResult(
                    module_name=module_name,
                    success=True,
                    data="Module disabled by configuration",
                    duration_seconds=0.0
                )
            
            # Progress: Initialization
            if self.progress_callback:
                self.progress_callback(module_name, "step", "Initializing module...")
            
            # Initialize and configure module
            module = module_class(module_name, [])
            module.configure(module_config)
            
            # Progress: Configuration
            if self.progress_callback:
                self.progress_callback(module_name, "step", "Module configured, starting analysis...")
            
            # Add timeout for data_profiler
            if module_name == 'data_profiler':
                module_config = module_config.copy()
                module_config['timeout_seconds'] = 120  # 2 minute timeout
                module_config['quick_mode'] = True  # Enable quick mode for large files
                module.configure(module_config)
            
            # Execute analysis with progress tracking
            result = module.analyze(context)
            
            # Progress: Completion
            if self.progress_callback:
                self.progress_callback(module_name, "step", "Analysis complete, processing results...")
            
            # Calculate metrics
            duration = time.time() - start_time
            end_memory = self.memory_manager.get_current_usage()['current_mb']
            memory_delta = end_memory - start_memory
            
            # Check if result is valid
            if result is None:
                return ModuleExecutionResult(
                    module_name=module_name,
                    success=False,
                    error_message="Module returned None result",
                    duration_seconds=duration
                )
            
            success = result.is_successful if hasattr(result, 'is_successful') else bool(result)
            data = result.data if hasattr(result, 'data') and result.has_data else None
            warnings = getattr(result, 'warnings', []) if result else []
            
            return ModuleExecutionResult(
                module_name=module_name,
                success=success,
                data=data,
                duration_seconds=duration,
                memory_delta_mb=memory_delta,
                warnings=warnings
            )
            
        except Exception as e:
            duration = time.time() - start_time
            error_msg = f"{type(e).__name__}: {str(e)}"
            self.logger.error(f"Module {module_name} failed with {error_msg}")
            
            if self.progress_callback:
                self.progress_callback(module_name, "step", f"Error: {error_msg}")
                
            return ModuleExecutionResult(
                module_name=module_name,
                success=False,
                error_message=error_msg,
                duration_seconds=duration
            )
    
    def _generate_results(self, context: AnalysisContext, module_results: Dict[str, ModuleExecutionResult], 
                         total_duration: float, successful_modules: int, total_modules: int) -> AnalysisResults:
        """Generate comprehensive analysis results"""
        
        # File information
        file_info = {
            'name': context.file_metadata.file_path.name,
            'size_mb': context.file_metadata.file_size_mb,
            'path': str(context.file_metadata.file_path)
        }
        
        # Analysis metadata
        analysis_metadata = {
            'timestamp': time.time(),
            'success_rate': successful_modules / total_modules if total_modules > 0 else 0.0,
            'total_duration_seconds': total_duration,
            'modules_executed': [name for name in module_results.keys()],
            'quality_score': self._calculate_quality_score(module_results)
        }
        
        # Module results data
        module_data = {}
        for name, result in module_results.items():
            if result.success and result.data:
                module_data[name] = result.data
        
        # Execution summary
        execution_summary = {
            'total_modules': total_modules,
            'successful_modules': successful_modules,
            'failed_modules': total_modules - successful_modules,
            'success_rate': successful_modules / total_modules if total_modules > 0 else 0.0,
            'average_quality_score': analysis_metadata['quality_score'],
            'total_duration_seconds': total_duration,
            'execution_order': list(module_results.keys()),
            'module_statuses': {name: "success" if result.success else "failed" 
                              for name, result in module_results.items()}
        }
        
        # Resource usage
        resource_usage = self.memory_manager.get_resource_report()
        
        # Recommendations
        recommendations = self._generate_recommendations(module_results, context)
        
        return AnalysisResults(
            file_info=file_info,
            analysis_metadata=analysis_metadata,
            module_results=module_data,
            execution_summary=execution_summary,
            resource_usage=resource_usage,
            recommendations=recommendations,
            success=successful_modules > 0
        )
    
    def _generate_results_dict(self, context: AnalysisContext, module_results: Dict[str, ModuleExecutionResult], 
                             total_duration: float, successful_modules: int, total_modules: int) -> Dict[str, Any]:
        """Generate results as dictionary for GUI compatibility"""
        
        # File information
        file_info = {
            'name': context.file_metadata.file_path.name,
            'size_mb': context.file_metadata.file_size_mb,
            'path': str(context.file_metadata.file_path)
        }
        
        # Analysis metadata
        analysis_metadata = {
            'timestamp': time.time(),
            'success_rate': successful_modules / total_modules if total_modules > 0 else 0.0,
            'total_duration_seconds': total_duration,
            'modules_executed': [name for name in module_results.keys()],
            'quality_score': self._calculate_quality_score(module_results)
        }
        
        # Module results data
        module_data = {}
        for name, result in module_results.items():
            if result.success and result.data:
                module_data[name] = result.data
        
        # Execution summary
        execution_summary = {
            'total_modules': total_modules,
            'successful_modules': successful_modules,
            'failed_modules': total_modules - successful_modules,
            'success_rate': successful_modules / total_modules if total_modules > 0 else 0.0,
            'average_quality_score': analysis_metadata['quality_score'],
            'total_duration_seconds': total_duration,
            'execution_order': list(module_results.keys()),
            'module_statuses': {name: "success" if result.success else "failed" 
                              for name, result in module_results.items()}
        }
        
        # Resource usage
        resource_usage = self.memory_manager.get_resource_report()
        
        # Recommendations
        recommendations = self._generate_recommendations(module_results, context)
        
        return {
            'file_info': file_info,
            'analysis_metadata': analysis_metadata,
            'module_results': module_data,
            'execution_summary': execution_summary,
            'resource_usage': resource_usage,
            'recommendations': recommendations,
            'success': successful_modules > 0
        }
    
    def _generate_error_results(self, file_path: Path, error_message: str, duration: float) -> AnalysisResults:
        """Generate results for failed analysis"""
        return AnalysisResults(
            file_info={'name': file_path.name, 'path': str(file_path)},
            analysis_metadata={
                'timestamp': time.time(),
                'success_rate': 0.0,
                'total_duration_seconds': duration,
                'modules_executed': [],
                'quality_score': 0.0
            },
            module_results={},
            execution_summary={
                'total_modules': 0,
                'successful_modules': 0,
                'failed_modules': 0,
                'success_rate': 0.0,
                'total_duration_seconds': duration
            },
            resource_usage={},
            recommendations=[],
            success=False,
            error_message=error_message
        )
    
    def _calculate_quality_score(self, module_results: Dict[str, ModuleExecutionResult]) -> float:
        """Calculate overall quality score based on module success and warnings"""
        if not module_results:
            return 0.0
        
        total_score = 0.0
        for result in module_results.values():
            if result.success:
                # Base score for success
                score = 1.0
                # Reduce score for warnings
                if result.warnings:
                    score -= 0.1 * len(result.warnings)
                total_score += max(score, 0.0)
        
        return total_score / len(module_results)
    
    def _generate_recommendations(self, module_results: Dict[str, ModuleExecutionResult], 
                                context: AnalysisContext) -> List[str]:
        """Generate actionable recommendations"""
        recommendations = []
        
        # Check for failed modules
        failed_modules = [name for name, result in module_results.items() if not result.success]
        if failed_modules:
            recommendations.append(f"Some analysis modules failed: {', '.join(failed_modules)}")
        
        # Performance recommendations
        total_duration = sum(result.duration_seconds for result in module_results.values())
        if total_duration > 300:  # 5 minutes
            recommendations.append("Analysis took a long time - consider optimizing file size or structure")
        
        # Memory recommendations
        total_memory = sum(result.memory_delta_mb for result in module_results.values())
        if total_memory > 1000:  # 1GB
            recommendations.append("High memory usage detected - consider processing in smaller chunks")
        
        # File size recommendations
        if context.file_metadata.file_size_mb > 100:
            recommendations.append("Large file detected - performance may be improved by splitting into smaller files")
        
        return recommendations


def create_analysis_dict(results: AnalysisResults) -> Dict[str, Any]:
    """Convert AnalysisResults to dictionary for JSON serialization"""
    return {
        'file_info': results.file_info,
        'analysis_metadata': results.analysis_metadata,
        'module_results': results.module_results,
        'execution_summary': results.execution_summary,
        'resource_usage': results.resource_usage,
        'recommendations': results.recommendations,
        'success': results.success,
        'error_message': results.error_message
    }
