"""
Enhanced exception management and structured logging for Excel Explorer
Provides hierarchical error handling with recovery mechanisms.
"""

import logging
import traceback
import json
import time
from typing import Optional, Dict, Any, List, Union
from enum import Enum
from pathlib import Path
from dataclasses import dataclass, asdict
from contextlib import contextmanager


class ErrorSeverity(Enum):
    """Error severity levels"""
    LOW = "low"           # Minor issues, analysis continues
    MEDIUM = "medium"     # Module-level failures, partial results
    HIGH = "high"         # Major component failures, significant impact
    CRITICAL = "critical" # Pipeline-stopping errors


class ErrorCategory(Enum):
    """Error categorization for better handling"""
    FILE_ACCESS = "file_access"
    MEMORY_LIMIT = "memory_limit"
    DATA_CORRUPTION = "data_corruption"
    FORMULA_PARSING = "formula_parsing"
    DEPENDENCY_FAILURE = "dependency_failure"
    CONFIGURATION = "configuration"
    TIMEOUT = "timeout"
    UNEXPECTED = "unexpected"


@dataclass
class ErrorContext:
    """Structured error context information"""
    module_name: str
    error_category: ErrorCategory
    severity: ErrorSeverity
    file_path: Optional[str] = None
    sheet_name: Optional[str] = None
    cell_reference: Optional[str] = None
    row_index: Optional[int] = None
    memory_usage_mb: Optional[float] = None
    processing_time_seconds: Optional[float] = None
    additional_data: Optional[Dict[str, Any]] = None


class ExcelAnalysisError(Exception):
    """Base exception for Excel analysis errors with rich context"""
    
    def __init__(
        self, 
        message: str, 
        severity: ErrorSeverity = ErrorSeverity.MEDIUM,
        category: ErrorCategory = ErrorCategory.UNEXPECTED,
        module_name: str = "unknown",
        file_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        cell_reference: Optional[str] = None,
        recovery_suggestion: Optional[str] = None,
        **kwargs
    ):
        super().__init__(message)
        self.message = message
        self.severity = severity
        self.category = category
        self.module_name = module_name
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.cell_reference = cell_reference
        self.recovery_suggestion = recovery_suggestion
        self.timestamp = time.time()
        self.additional_data = kwargs
        
    def to_dict(self) -> Dict[str, Any]:
        """Convert error to structured dictionary"""
        return {
            'message': self.message,
            'severity': self.severity.value,
            'category': self.category.value,
            'module_name': self.module_name,
            'file_path': self.file_path,
            'sheet_name': self.sheet_name,
            'cell_reference': self.cell_reference,
            'recovery_suggestion': self.recovery_suggestion,
            'timestamp': self.timestamp,
            'additional_data': self.additional_data
        }


class MemoryLimitError(ExcelAnalysisError):
    """Memory usage exceeded configured limits"""
    def __init__(self, current_mb: float, limit_mb: float, **kwargs):
        super().__init__(
            f"Memory limit exceeded: {current_mb:.1f}MB > {limit_mb:.1f}MB",
            severity=ErrorSeverity.HIGH,
            category=ErrorCategory.MEMORY_LIMIT,
            **kwargs
        )
        self.current_mb = current_mb
        self.limit_mb = limit_mb


class FileAccessError(ExcelAnalysisError):
    """File access or format issues"""
    def __init__(self, file_path: str, original_error: Exception, **kwargs):
        super().__init__(
            f"File access failed: {original_error}",
            severity=ErrorSeverity.CRITICAL,
            category=ErrorCategory.FILE_ACCESS,
            file_path=file_path,
            **kwargs
        )
        self.original_error = original_error


class DataCorruptionError(ExcelAnalysisError):
    """Corrupted or invalid data encountered"""
    def __init__(self, location: str, issue_description: str, **kwargs):
        super().__init__(
            f"Data corruption at {location}: {issue_description}",
            severity=ErrorSeverity.MEDIUM,
            category=ErrorCategory.DATA_CORRUPTION,
            **kwargs
        )
        self.location = location


class ModuleDependencyError(ExcelAnalysisError):
    """Module dependency requirements not met"""
    def __init__(self, module_name: str, missing_dependencies: List[str], **kwargs):
        dependencies_str = ", ".join(missing_dependencies)
        super().__init__(
            f"Module {module_name} missing dependencies: {dependencies_str}",
            severity=ErrorSeverity.HIGH,
            category=ErrorCategory.DEPENDENCY_FAILURE,
            module_name=module_name,
            **kwargs
        )
        self.missing_dependencies = missing_dependencies


class ProcessingTimeoutError(ExcelAnalysisError):
    """Processing exceeded time limits"""
    def __init__(self, module_name: str, timeout_seconds: float, **kwargs):
        super().__init__(
            f"Module {module_name} timed out after {timeout_seconds:.1f}s",
            severity=ErrorSeverity.HIGH,
            category=ErrorCategory.TIMEOUT,
            module_name=module_name,
            **kwargs
        )
        self.timeout_seconds = timeout_seconds


class StructuredLogger:
    """Enhanced logger with structured error tracking"""
    
    def __init__(self, name: str, log_file: Optional[Path] = None):
        self.logger = logging.getLogger(name)
        self.error_history: List[Dict[str, Any]] = []
        self.module_error_counts: Dict[str, int] = {}
        
        if log_file and not self.logger.handlers:
            self._configure_handlers(log_file)
    
    def _configure_handlers(self, log_file: Path) -> None:
        """Configure logging handlers with structured format"""
        formatter = logging.Formatter(
            '%(asctime)s [%(levelname)8s] %(name)s: %(message)s'
        )
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)
        
        # File handler for detailed logs
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.DEBUG)
        
        self.logger.addHandler(console_handler)
        self.logger.addHandler(file_handler)
        self.logger.setLevel(logging.DEBUG)
    
    def log_error(self, error: ExcelAnalysisError, include_traceback: bool = True) -> None:
        """Log structured error with context"""
        error_dict = error.to_dict()
        self.error_history.append(error_dict)
        
        # Track module error frequency
        module = error.module_name
        self.module_error_counts[module] = self.module_error_counts.get(module, 0) + 1
        
        # Format log message
        context_parts = []
        if error.file_path:
            context_parts.append(f"file={Path(error.file_path).name}")
        if error.sheet_name:
            context_parts.append(f"sheet={error.sheet_name}")
        if error.cell_reference:
            context_parts.append(f"cell={error.cell_reference}")
        
        context_str = f" [{', '.join(context_parts)}]" if context_parts else ""
        
        log_message = f"{error.category.value.upper()}: {error.message}{context_str}"
        
        # Add recovery suggestion if available
        if error.recovery_suggestion:
            log_message += f" | Recovery: {error.recovery_suggestion}"
        
        # Log at appropriate level
        if error.severity == ErrorSeverity.CRITICAL:
            self.logger.error(log_message)
        elif error.severity == ErrorSeverity.HIGH:
            self.logger.warning(log_message)
        else:
            self.logger.info(log_message)
        
        # Include traceback for high severity errors
        if include_traceback and error.severity.value in ['high', 'critical']:
            self.logger.debug(f"Traceback for {error.module_name}:", exc_info=True)
    
    def log_module_completion(self, module_name: str, duration_seconds: float, 
                            memory_delta_mb: float) -> None:
        """Log successful module completion"""
        self.logger.info(
            f"Module {module_name} completed: {duration_seconds:.1f}s, "
            f"{memory_delta_mb:+.1f}MB"
        )
    
    def log_analysis_summary(self, errors_by_severity: Dict[str, int], 
                           successful_modules: int, total_modules: int) -> None:
        """Log final analysis summary"""
        success_rate = (successful_modules / total_modules * 100) if total_modules > 0 else 0
        
        summary_parts = [f"Analysis complete: {success_rate:.1f}% success rate"]
        if errors_by_severity:
            error_summary = ", ".join([
                f"{count} {severity}" for severity, count in errors_by_severity.items()
            ])
            summary_parts.append(f"Errors: {error_summary}")
        
        self.logger.info(" | ".join(summary_parts))
    
    def get_error_summary(self) -> Dict[str, Any]:
        """Get comprehensive error statistics"""
        severity_counts = {}
        category_counts = {}
        
        for error in self.error_history:
            severity = error['severity']
            category = error['category']
            
            severity_counts[severity] = severity_counts.get(severity, 0) + 1
            category_counts[category] = category_counts.get(category, 0) + 1
        
        return {
            'total_errors': len(self.error_history),
            'by_severity': severity_counts,
            'by_category': category_counts,
            'by_module': self.module_error_counts.copy(),
            'recent_errors': self.error_history[-10:] if self.error_history else []
        }


class ErrorHandler:
    """Central error handling with recovery mechanisms"""
    
    def __init__(self, log_file: Optional[Path] = None):
        self.logger = StructuredLogger("excel_explorer", log_file)
        self.recovery_strategies: Dict[ErrorCategory, callable] = {
            ErrorCategory.MEMORY_LIMIT: self._recover_memory_limit,
            ErrorCategory.DATA_CORRUPTION: self._recover_data_corruption,
            ErrorCategory.DEPENDENCY_FAILURE: self._recover_dependency_failure,
            ErrorCategory.TIMEOUT: self._recover_timeout
        }
    
    def handle_error(self, error: Union[ExcelAnalysisError, Exception], 
                    context: Optional[ErrorContext] = None) -> bool:
        """Handle error with optional recovery attempt
        
        Returns:
            bool: True if error was recovered, False if fatal
        """
        # Convert generic exceptions to ExcelAnalysisError
        if not isinstance(error, ExcelAnalysisError):
            error = ExcelAnalysisError(
                str(error),
                severity=ErrorSeverity.MEDIUM,
                category=ErrorCategory.UNEXPECTED,
                module_name=context.module_name if context else "unknown"
            )
        
        # Log the error
        self.logger.log_error(error)
        
        # Attempt recovery for non-critical errors
        if error.severity != ErrorSeverity.CRITICAL:
            recovery_func = self.recovery_strategies.get(error.category)
            if recovery_func:
                try:
                    return recovery_func(error, context)
                except Exception as recovery_error:
                    self.logger.log_error(ExcelAnalysisError(
                        f"Recovery failed: {recovery_error}",
                        severity=ErrorSeverity.HIGH,
                        category=ErrorCategory.UNEXPECTED,
                        module_name=error.module_name
                    ))
        
        return False
    
    def _recover_memory_limit(self, error: ExcelAnalysisError, 
                            context: Optional[ErrorContext]) -> bool:
        """Attempt memory limit recovery"""
        # Could implement cache clearing, chunk size reduction, etc.
        self.logger.logger.info(f"Attempting memory recovery for {error.module_name}")
        return False  # Placeholder - implement actual recovery logic
    
    def _recover_data_corruption(self, error: ExcelAnalysisError, 
                               context: Optional[ErrorContext]) -> bool:
        """Attempt data corruption recovery"""
        self.logger.logger.info(f"Attempting data corruption recovery for {error.module_name}")
        return False  # Placeholder
    
    def _recover_dependency_failure(self, error: ExcelAnalysisError, 
                                  context: Optional[ErrorContext]) -> bool:
        """Attempt dependency failure recovery"""
        self.logger.logger.info(f"Attempting dependency recovery for {error.module_name}")
        return False  # Placeholder
    
    def _recover_timeout(self, error: ExcelAnalysisError, 
                        context: Optional[ErrorContext]) -> bool:
        """Attempt timeout recovery"""
        self.logger.logger.info(f"Attempting timeout recovery for {error.module_name}")
        return False  # Placeholder


# Global error handler instance
_global_error_handler: Optional[ErrorHandler] = None


def get_error_handler() -> ErrorHandler:
    """Get or create global error handler"""
    global _global_error_handler
    if _global_error_handler is None:
        _global_error_handler = ErrorHandler()
    return _global_error_handler


def initialize_error_handler(log_file: Optional[Path] = None) -> ErrorHandler:
    """Initialize global error handler with configuration"""
    global _global_error_handler
    _global_error_handler = ErrorHandler(log_file)
    return _global_error_handler


@contextmanager
def error_boundary(module_name: str, severity: ErrorSeverity = ErrorSeverity.MEDIUM):
    """Context manager for automatic error handling"""
    error_handler = get_error_handler()
    try:
        yield
    except Exception as e:
        if isinstance(e, ExcelAnalysisError):
            error_handler.handle_error(e)
        else:
            excel_error = ExcelAnalysisError(
                str(e),
                severity=severity,
                module_name=module_name,
                category=ErrorCategory.UNEXPECTED
            )
            error_handler.handle_error(excel_error)
        raise


# Legacy compatibility
ExplorerError = ExcelAnalysisError
LOG_FORMAT = "%(asctime)s [%(levelname)8s] %(name)s: %(message)s"


def configure_logging(level: int = logging.INFO, log_file: Optional[str] = None) -> None:
    """Legacy logging configuration function"""
    log_path = Path(log_file) if log_file else None
    initialize_error_handler(log_path)


def log_exception(exc: Exception) -> None:
    """Legacy exception logging function"""
    error_handler = get_error_handler()
    if isinstance(exc, ExcelAnalysisError):
        error_handler.handle_error(exc)
    else:
        excel_error = ExcelAnalysisError(
            str(exc),
            severity=ErrorSeverity.HIGH,
            category=ErrorCategory.UNEXPECTED
        )
        error_handler.handle_error(excel_error)
