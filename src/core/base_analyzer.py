"""
Enhanced base class for all analyzer modules
Integrates with ModuleResult framework, AnalysisContext, and resource management.
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional, Type, TypeVar
import time
import logging

from .analysis_context import AnalysisContext
from .module_result import ModuleResult, ResultStatus, ExecutionMetrics, ValidationResult
from ..utils.memory_manager import ResourceMonitor, get_memory_manager
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, error_boundary

T = TypeVar('T')


class BaseAnalyzer(ABC):
    """Enhanced base class for all Excel analysis modules"""
    
    def __init__(self, name: str, dependencies: Optional[List[str]] = None):
        self.name = name
        self.dependencies = dependencies or []
        self.logger = logging.getLogger(f"excel_explorer.{name}")
        self.memory_manager = get_memory_manager()
        
        # Module configuration
        self.config: Dict[str, Any] = {}
        self.result: Optional[ModuleResult] = None
        
        # State tracking
        self._execution_start: Optional[float] = None
        self._memory_start: Optional[float] = None
        
    @abstractmethod
    def _perform_analysis(self, context: AnalysisContext) -> T:
        """Perform the core analysis logic - must be implemented by subclasses
        
        Args:
            context: AnalysisContext with workbook access and shared state
            
        Returns:
            Analysis-specific data structure
        """
        pass
    
    @abstractmethod
    def _validate_result(self, data: T, context: AnalysisContext) -> ValidationResult:
        """Validate analysis results - must be implemented by subclasses
        
        Args:
            data: The analysis data to validate
            context: AnalysisContext for additional validation context
            
        Returns:
            ValidationResult with quality metrics
        """
        pass
    
    def get_data_type(self) -> Type[T]:
        """Get the expected data type for this analyzer - override if needed"""
        return dict  # Default to dict for generic analyzers
    
    def configure(self, config: Dict[str, Any]) -> None:
        """Update module configuration"""
        self.config.update(config)
        self.logger.debug(f"Configuration updated: {config}")
    
    def check_dependencies(self, context: AnalysisContext) -> bool:
        """Check if all required dependencies are satisfied
        
        Args:
            context: AnalysisContext to check module completion status
            
        Returns:
            bool: True if all dependencies are met
        """
        if not self.dependencies:
            return True
            
        missing_deps = []
        failed_deps = []
        
        for dep in self.dependencies:
            if not context.is_module_completed(dep):
                if dep in context.failed_modules:
                    failed_deps.append(dep)
                else:
                    missing_deps.append(dep)
        
        if failed_deps:
            self.logger.error(f"Dependencies failed: {failed_deps}")
            return False
            
        if missing_deps:
            self.logger.warning(f"Dependencies not yet completed: {missing_deps}")
            return False
        
        self.logger.debug(f"All dependencies satisfied: {self.dependencies}")
        return True
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate processing complexity (1.0 = baseline) - override for specific modules
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier (1.0 = baseline)
        """
        # Base complexity estimation from file size
        size_mb = context.file_metadata.file_size_mb
        
        if size_mb < 1:
            return 0.5
        elif size_mb < 10:
            return 1.0
        elif size_mb < 50:
            return 2.0
        else:
            return 3.0
    
    def should_skip_analysis(self, context: AnalysisContext) -> bool:
        """Determine if analysis should be skipped due to constraints
        
        Args:
            context: AnalysisContext to check resource constraints
            
        Returns:
            bool: True if analysis should be skipped
        """
        # Skip if dependencies failed critically
        if context.has_dependency_failed(self.dependencies):
            return True
            
        # Skip if resource constraints are too high
        if context.should_reduce_complexity():
            complexity = self.estimate_complexity(context)
            if complexity > 2.0:  # Skip high-complexity modules under pressure
                self.logger.warning(f"Skipping {self.name} due to resource constraints")
                return True
        
        return False
    
    def analyze(self, context: AnalysisContext) -> ModuleResult[T]:
        """Main analysis entry point with full error handling and resource tracking
        
        Args:
            context: AnalysisContext with workbook access and shared state
            
        Returns:
            ModuleResult with analysis data and execution metrics
        """
        # Initialize result
        self.result = ModuleResult[T](
            module_name=self.name,
            status=ResultStatus.FAILED,
            dependencies_met=self.dependencies.copy()
        )
        
        try:
            # Pre-execution checks
            if not self.check_dependencies(context):
                self.result.status = ResultStatus.SKIPPED
                self.result.add_error("Dependencies not satisfied")
                return self.result
            
            if self.should_skip_analysis(context):
                self.result.status = ResultStatus.SKIPPED
                self.result.add_warning("Analysis skipped due to constraints")
                return self.result
            
            # Execute with resource monitoring
            complexity = self.estimate_complexity(context)
            estimated_time = self.memory_manager.estimate_processing_time(
                context.file_metadata.file_size_mb, complexity
            )
            
            self.logger.info(f"Starting {self.name} analysis (estimated: {estimated_time:.1f}s)")
            
            with ResourceMonitor(self.memory_manager, self.name, context.file_metadata.file_size_mb) as monitor:
                with error_boundary(self.name, ErrorSeverity.MEDIUM):
                    # Track execution start
                    self._execution_start = time.time()
                    self._memory_start = self.memory_manager.get_current_usage()['current_mb']
                    
                    # Perform analysis
                    data = self._perform_analysis(context)
                    
                    # Validate results
                    validation = self._validate_result(data, context)
                    
                    # Create execution metrics
                    current_usage = self.memory_manager.get_current_usage()
                    metrics = ExecutionMetrics(
                        start_time=self._execution_start,
                        end_time=time.time(),
                        memory_start_mb=self._memory_start,
                        memory_end_mb=current_usage['current_mb'],
                        peak_memory_mb=current_usage['peak_mb'],
                        cpu_percent=current_usage['cpu_percent'],
                        items_processed=self._count_processed_items(data)
                    )
                    
                    # Update result with success
                    self.result.data = data
                    self.result.validation = validation
                    self.result.metrics = metrics
                    self.result.status = ResultStatus.SUCCESS
                    
                    # Add metadata
                    self.result.metadata.update({
                        'complexity_factor': complexity,
                        'estimated_time_seconds': estimated_time,
                        'file_size_mb': context.file_metadata.file_size_mb
                    })
                    
                    self.logger.info(f"Completed {self.name} successfully")
                    
        except ExcelAnalysisError as e:
            self._handle_analysis_error(e)
        except Exception as e:
            # Wrap unexpected errors
            excel_error = ExcelAnalysisError(
                f"Unexpected error in {self.name}: {e}",
                severity=ErrorSeverity.HIGH,
                module_name=self.name
            )
            self._handle_analysis_error(excel_error)
        
        finally:
            # Register result with context
            context.register_module_result(
                self.name, 
                self.result, 
                success=self.result.is_successful
            )
        
        return self.result
    
    def _handle_analysis_error(self, error: ExcelAnalysisError) -> None:
        """Handle analysis errors and update result"""
        self.result.add_error(str(error))
        
        if error.severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL]:
            self.result.status = ResultStatus.FAILED
        else:
            self.result.status = ResultStatus.PARTIAL
        
        # Add error context to metadata
        self.result.metadata['last_error'] = {
            'message': str(error),
            'severity': error.severity.value,
            'category': error.category.value
        }
        
        self.logger.error(f"Analysis error in {self.name}: {error}")
    
    def _count_processed_items(self, data: T) -> int:
        """Count processed items for metrics - override for specific counting logic
        
        Args:
            data: Analysis result data
            
        Returns:
            int: Number of items processed
        """
        if isinstance(data, dict):
            return len(data)
        elif isinstance(data, list):
            return len(data)
        elif hasattr(data, '__len__'):
            return len(data)
        else:
            return 1
    
    def get_cache_key(self, context: AnalysisContext, suffix: str = "") -> str:
        """Generate cache key for this module's data
        
        Args:
            context: AnalysisContext for file identification
            suffix: Optional suffix for specific cache items
            
        Returns:
            str: Cache key
        """
        file_hash = str(hash(str(context.file_path)))
        base_key = f"{self.name}_{file_hash}"
        return f"{base_key}_{suffix}" if suffix else base_key
    
    def cache_intermediate_result(self, context: AnalysisContext, key: str, data: Any) -> None:
        """Cache intermediate analysis data
        
        Args:
            context: AnalysisContext for cache access
            key: Cache key
            data: Data to cache
        """
        full_key = self.get_cache_key(context, key)
        context.cache_set(full_key, data)
        self.logger.debug(f"Cached intermediate result: {full_key}")
    
    def get_cached_result(self, context: AnalysisContext, key: str) -> Optional[Any]:
        """Retrieve cached intermediate data
        
        Args:
            context: AnalysisContext for cache access
            key: Cache key
            
        Returns:
            Cached data or None
        """
        full_key = self.get_cache_key(context, key)
        return context.cache_get(full_key)
    
    @property
    def execution_time(self) -> Optional[float]:
        """Get last execution time in seconds"""
        if self.result and self.result.metrics:
            return self.result.metrics.duration_seconds
        return None
    
    @property
    def memory_usage(self) -> Optional[float]:
        """Get last memory delta in MB"""
        if self.result and self.result.metrics:
            return self.result.metrics.memory_delta_mb
        return None
    
    def __repr__(self) -> str:
        """String representation"""
        status = self.result.status.value if self.result else "not_executed"
        return f"{self.__class__.__name__}(name='{self.name}', status='{status}')"
