"""
Enhanced Excel Explorer Orchestrator
Coordinates full analysis pipeline with dependency management and resource monitoring.
"""

import argparse
import json
import time
from pathlib import Path
from typing import Dict, List, Any, Optional, Type
import logging

from .analysis_context import AnalysisContext, AnalysisConfig, AnalysisPhase
from .base_analyzer import BaseAnalyzer
from .module_result import ResultCollector, ModuleResult
from ..utils.config_loader import load_config
from ..utils.error_handler import (
    initialize_error_handler, ExcelAnalysisError, ErrorSeverity
)
from ..utils.memory_manager import initialize_memory_manager

# Import all available modules
from ..modules.health_checker import HealthChecker
from ..modules.structure_mapper import StructureMapper
from ..modules.data_profiler import DataProfiler
from ..modules.formula_analyzer import FormulaAnalyzer
from ..modules.visual_cataloger import VisualCataloger
from ..modules.connection_inspector import ConnectionInspector
from ..modules.pivot_intelligence import PivotIntelligence
from ..modules.doc_synthesizer import DocSynthesizer


class ExcelExplorer:
    """Enhanced orchestrator with dependency-driven execution and resource management"""
    
    # Module execution order with dependencies
    MODULE_REGISTRY = {
        'health_checker': {
            'class': HealthChecker,
            'dependencies': [],
            'critical': True,
            'phase': AnalysisPhase.HEALTH_CHECK
        },
        'structure_mapper': {
            'class': StructureMapper,
            'dependencies': ['health_checker'],
            'critical': True,
            'phase': AnalysisPhase.STRUCTURE_MAPPING
        },
        'data_profiler': {
            'class': DataProfiler,
            'dependencies': ['health_checker', 'structure_mapper'],
            'critical': False,
            'phase': AnalysisPhase.DATA_PROFILING
        },
        'formula_analyzer': {
            'class': FormulaAnalyzer,
            'dependencies': ['structure_mapper'],
            'critical': False,
            'phase': AnalysisPhase.FORMULA_ANALYSIS
        },
        'visual_cataloger': {
            'class': VisualCataloger,
            'dependencies': ['structure_mapper'],
            'critical': False,
            'phase': AnalysisPhase.VISUAL_ANALYSIS
        },
        'connection_inspector': {
            'class': ConnectionInspector,
            'dependencies': ['health_checker'],
            'critical': False,
            'phase': AnalysisPhase.CONNECTION_ANALYSIS
        },
        'pivot_intelligence': {
            'class': PivotIntelligence,
            'dependencies': ['structure_mapper', 'data_profiler'],
            'critical': False,
            'phase': AnalysisPhase.PIVOT_ANALYSIS
        },
        'doc_synthesizer': {
            'class': DocSynthesizer,
            'dependencies': ['structure_mapper', 'data_profiler'],
            'critical': False,
            'phase': AnalysisPhase.SYNTHESIS
        }
    }
    
    def __init__(self, config_path: Optional[str] = None, log_file: Optional[str] = None):
        # Load configuration
        self.config_data = load_config(config_path)
        self.analysis_config = AnalysisConfig(**self.config_data.get('analysis', {}))
        
        # Initialize core systems
        self.error_handler = initialize_error_handler(
            Path(log_file) if log_file else None
        )
        self.memory_manager = initialize_memory_manager(
            self.analysis_config.max_memory_mb,
            self.analysis_config.warning_threshold
        )
        
        # Initialize modules
        self.modules: Dict[str, BaseAnalyzer] = {}
        self.result_collector = ResultCollector()
        self.logger = logging.getLogger(__name__)
        
        self._initialize_modules()
    
    def _initialize_modules(self) -> None:
        """Initialize all registered modules with configuration"""
        for module_name, module_info in self.MODULE_REGISTRY.items():
            try:
                module_config = self.config_data.get(module_name, {})
                module_class = module_info['class']
                dependencies = module_info['dependencies']
                
                # Create module instance
                module = module_class(module_name, dependencies)
                module.configure(module_config)
                
                self.modules[module_name] = module
                self.logger.debug(f"Initialized module: {module_name}")
                
            except Exception as e:
                self.logger.error(f"Failed to initialize module {module_name}: {e}")
                raise ExcelAnalysisError(
                    f"Module initialization failed: {module_name}",
                    severity=ErrorSeverity.CRITICAL,
                    module_name=module_name
                )
    
    def _calculate_execution_order(self, enabled_modules: List[str]) -> List[str]:
        """Calculate optimal execution order based on dependencies"""
        ordered = []
        remaining = set(enabled_modules)
        
        while remaining:
            # Find modules with satisfied dependencies
            ready = []
            for module_name in remaining:
                deps = self.MODULE_REGISTRY[module_name]['dependencies']
                if all(dep in ordered or dep not in enabled_modules for dep in deps):
                    ready.append(module_name)
            
            if not ready:
                # Circular dependency or missing dependency
                remaining_deps = {}
                for module_name in remaining:
                    unsatisfied = [
                        dep for dep in self.MODULE_REGISTRY[module_name]['dependencies']
                        if dep not in ordered and dep in enabled_modules
                    ]
                    if unsatisfied:
                        remaining_deps[module_name] = unsatisfied
                
                raise ExcelAnalysisError(
                    f"Circular or missing dependencies: {remaining_deps}",
                    severity=ErrorSeverity.CRITICAL,
                    module_name="orchestrator"
                )
            
            # Sort ready modules by phase order
            ready.sort(key=lambda x: list(AnalysisPhase).index(
                self.MODULE_REGISTRY[x]['phase']
            ))
            
            ordered.extend(ready)
            remaining -= set(ready)
        
        return ordered
    
    def _filter_modules_by_config(self) -> List[str]:
        """Filter modules based on configuration settings"""
        enabled = []
        
        for module_name in self.MODULE_REGISTRY.keys():
            # Check if module is explicitly disabled
            module_config = self.config_data.get(module_name, {})
            if not module_config.get('enabled', True):
                continue
            
            # Check feature flags
            if module_name == 'visual_cataloger' and not self.analysis_config.include_charts:
                continue
            if module_name == 'pivot_intelligence' and not self.analysis_config.include_pivots:
                continue
            if module_name == 'connection_inspector' and not self.analysis_config.include_connections:
                continue
            
            enabled.append(module_name)
        
        return enabled
    
    def analyze_file(self, file_path: str) -> Dict[str, Any]:
        """Execute full analysis pipeline with dependency management
        
        Args:
            file_path: Path to Excel file to analyze
            
        Returns:
            Comprehensive analysis results
        """
        file_path = Path(file_path)
        
        try:
            with AnalysisContext(file_path, self.analysis_config) as context:
                self.logger.info(f"Starting analysis of {file_path.name}")
                
                # Determine execution plan
                enabled_modules = self._filter_modules_by_config()
                execution_order = self._calculate_execution_order(enabled_modules)
                
                self.logger.info(f"Execution plan: {' -> '.join(execution_order)}")
                
                # Execute modules in dependency order
                for module_name in execution_order:
                    try:
                        self._execute_module(module_name, context)
                        
                        # Check for critical failures
                        if self.result_collector.has_critical_failures():
                            self.logger.error("Critical module failure detected")
                            break
                        
                        # Resource pressure check
                        if context.should_reduce_complexity():
                            self.logger.warning("Resource pressure detected, may skip optional modules")
                    
                    except Exception as e:
                        self.logger.error(f"Module execution failed: {module_name}: {e}")
                        # Continue with other modules unless critical
                        if self.MODULE_REGISTRY[module_name]['critical']:
                            raise
                
                # Generate final results
                return self._generate_results(context)
                
        except Exception as e:
            self.logger.error(f"Analysis failed: {e}")
            self.error_handler.handle_error(e)
            return self._generate_error_results(str(e))
    
    def _execute_module(self, module_name: str, context: AnalysisContext) -> None:
        """Execute single module with error handling
        
        Args:
            module_name: Name of module to execute
            context: Analysis context
        """
        module = self.modules[module_name]
        module_info = self.MODULE_REGISTRY[module_name]
        
        # Set analysis phase
        context.set_phase(module_info['phase'])
        
        # Execute module
        self.logger.info(f"Executing module: {module_name}")
        result = module.analyze(context)
        
        # Collect result
        self.result_collector.add_result(result)
        
        # Log execution summary
        if result.metrics:
            self.logger.info(
                f"Module {module_name} completed: "
                f"{result.metrics.duration_seconds:.1f}s, "
                f"quality={result.quality_score:.2f}"
            )
        
        # Handle failures
        if not result.is_successful:
            if module_info['critical']:
                raise ExcelAnalysisError(
                    f"Critical module failed: {module_name}",
                    severity=ErrorSeverity.CRITICAL,
                    module_name=module_name
                )
            else:
                self.logger.warning(f"Optional module failed: {module_name}")
    
    def _generate_results(self, context: AnalysisContext) -> Dict[str, Any]:
        """Generate comprehensive analysis results
        
        Args:
            context: Analysis context
            
        Returns:
            Complete analysis results
        """
        execution_summary = self.result_collector.get_execution_summary()
        analysis_summary = context.get_analysis_summary()
        resource_report = self.memory_manager.get_resource_report()
        error_summary = self.error_handler.logger.get_error_summary()
        
        # Collect module data
        module_data = {}
        for name, result in self.result_collector.get_successful_results().items():
            if result.has_data:
                module_data[name] = result.data
        
        return {
            'file_info': analysis_summary['file_info'],
            'analysis_metadata': {
                'timestamp': time.time(),
                'success_rate': execution_summary['success_rate'],
                'total_duration_seconds': execution_summary['total_duration_seconds'],
                'modules_executed': execution_summary['execution_order'],
                'quality_score': execution_summary['average_quality_score']
            },
            'module_results': module_data,
            'execution_summary': execution_summary,
            'resource_usage': resource_report,
            'error_summary': error_summary,
            'recommendations': self._generate_recommendations(context)
        }
    
    def _generate_error_results(self, error_message: str) -> Dict[str, Any]:
        """Generate results structure for failed analysis
        
        Args:
            error_message: Primary error message
            
        Returns:
            Error results structure
        """
        return {
            'analysis_status': 'failed',
            'error_message': error_message,
            'partial_results': self.result_collector.get_successful_results(),
            'execution_summary': self.result_collector.get_execution_summary(),
            'resource_usage': self.memory_manager.get_resource_report()
        }
    
    def _generate_recommendations(self, context: AnalysisContext) -> List[str]:
        """Generate actionable recommendations based on analysis
        
        Args:
            context: Analysis context
            
        Returns:
            List of recommendations
        """
        recommendations = []
        
        # Performance recommendations
        resource_usage = self.memory_manager.get_current_usage()
        if resource_usage['usage_ratio'] > 0.8:
            recommendations.append("Consider analyzing smaller file sections or increasing memory limits")
        
        # Quality recommendations  
        failed_results = self.result_collector.get_failed_results()
        if failed_results:
            recommendations.append(f"Some analysis modules failed: {list(failed_results.keys())}")
        
        # File-specific recommendations
        if context.file_metadata.file_size_mb > 100:
            recommendations.append("Large file detected - consider enabling parallel processing")
        
        return recommendations


def create_cli_parser() -> argparse.ArgumentParser:
    """Create command-line interface parser"""
    parser = argparse.ArgumentParser(
        description="Excel Explorer - Comprehensive Excel file analysis",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  excel_explorer file.xlsx
  excel_explorer file.xlsx --config custom.yaml --output results.json
  excel_explorer file.xlsx --deep-analysis --memory-limit 8192
        """
    )
    
    parser.add_argument("excel_file", help="Path to Excel file to analyze")
    parser.add_argument("--config", help="Path to YAML configuration file")
    parser.add_argument("--output", help="Path to save JSON results")
    parser.add_argument("--logfile", help="Path to write detailed log")
    parser.add_argument("--memory-limit", type=int, help="Memory limit in MB")
    parser.add_argument("--deep-analysis", action="store_true", 
                       help="Enable deep analysis mode")
    parser.add_argument("--parallel", action="store_true",
                       help="Enable parallel processing")
    
    return parser


def main():
    """Main CLI entry point"""
    parser = create_cli_parser()
    args = parser.parse_args()
    
    try:
        # Initialize explorer
        explorer = ExcelExplorer(config_path=args.config, log_file=args.logfile)
        
        # Apply CLI overrides
        if args.memory_limit:
            explorer.analysis_config.max_memory_mb = args.memory_limit
        if args.deep_analysis:
            explorer.analysis_config.deep_analysis = True
        if args.parallel:
            explorer.analysis_config.parallel_processing = True
        
        # Execute analysis
        results = explorer.analyze_file(args.excel_file)
        
        # Output results
        if args.output:
            output_path = Path(args.output)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, default=str)
            print(f"Results saved to: {output_path}")
        else:
            print(json.dumps(results, indent=2, default=str))
            
    except Exception as e:
        print(f"Analysis failed: {e}", file=sys.stderr)
        return 1
    
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
