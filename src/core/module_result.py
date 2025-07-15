"""
Module Result Framework - Standardized result structures for analysis modules
Provides consistent interfaces for module outputs and success metrics.
"""

from typing import Dict, Any, List, Optional, Union, Generic, TypeVar
from dataclasses import dataclass, field
from enum import Enum
from abc import ABC, abstractmethod
import time
import json
from pathlib import Path

T = TypeVar('T')


class ResultStatus(Enum):
    """Module execution status"""
    SUCCESS = "success"
    PARTIAL = "partial"
    FAILED = "failed"
    SKIPPED = "skipped"


class ConfidenceLevel(Enum):
    """Confidence in analysis results"""
    HIGH = "high"        # >95% confidence
    MEDIUM = "medium"    # 80-95% confidence  
    LOW = "low"          # 60-80% confidence
    UNCERTAIN = "uncertain"  # <60% confidence


@dataclass
class ExecutionMetrics:
    """Standardized execution performance metrics"""
    start_time: float
    end_time: float
    duration_seconds: float = field(init=False)
    memory_start_mb: float
    memory_end_mb: float
    memory_delta_mb: float = field(init=False)
    peak_memory_mb: float
    cpu_percent: float
    items_processed: int = 0
    processing_rate: float = field(init=False, default=0.0)
    
    def __post_init__(self):
        self.duration_seconds = self.end_time - self.start_time
        self.memory_delta_mb = self.memory_end_mb - self.memory_start_mb
        if self.duration_seconds > 0:
            self.processing_rate = self.items_processed / self.duration_seconds


@dataclass
class ValidationResult:
    """Result validation metrics"""
    completeness_score: float  # 0.0-1.0
    accuracy_score: float      # 0.0-1.0  
    consistency_score: float   # 0.0-1.0
    confidence_level: ConfidenceLevel
    validation_notes: List[str] = field(default_factory=list)
    
    @property
    def overall_score(self) -> float:
        """Weighted overall quality score"""
        return (
            self.completeness_score * 0.4 + 
            self.accuracy_score * 0.4 + 
            self.consistency_score * 0.2
        )


@dataclass
class ModuleResult(Generic[T]):
    """Base result structure for all analysis modules"""
    module_name: str
    status: ResultStatus
    data: Optional[T] = None
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    metrics: Optional[ExecutionMetrics] = None
    validation: Optional[ValidationResult] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    dependencies_met: List[str] = field(default_factory=list)
    dependencies_failed: List[str] = field(default_factory=list)
    
    @property
    def is_successful(self) -> bool:
        """Check if module executed successfully"""
        return self.status in [ResultStatus.SUCCESS, ResultStatus.PARTIAL]
    
    @property
    def has_data(self) -> bool:
        """Check if module produced usable data"""
        return self.data is not None
    
    @property
    def quality_score(self) -> float:
        """Get overall quality score (0.0-1.0)"""
        if not self.validation:
            return 1.0 if self.is_successful else 0.0
        return self.validation.overall_score
    
    def add_error(self, error: str) -> None:
        """Add error message"""
        self.errors.append(error)
        if self.status == ResultStatus.SUCCESS:
            self.status = ResultStatus.PARTIAL
    
    def add_warning(self, warning: str) -> None:
        """Add warning message"""
        self.warnings.append(warning)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert result to dictionary for serialization"""
        return {
            'module_name': self.module_name,
            'status': self.status.value,
            'has_data': self.has_data,
            'data': self.data,
            'errors': self.errors,
            'warnings': self.warnings,
            'metrics': self.metrics.__dict__ if self.metrics else None,
            'validation': self.validation.__dict__ if self.validation else None,
            'metadata': self.metadata,
            'dependencies_met': self.dependencies_met,
            'dependencies_failed': self.dependencies_failed,
            'quality_score': self.quality_score
        }


# Specialized result types for specific modules

@dataclass
class HealthCheckData:
    """Health checker module result data"""
    file_accessible: bool
    file_size_mb: float
    corruption_detected: bool
    security_issues: List[str]
    excel_version: Optional[str]
    macro_enabled: bool
    password_protected: bool
    repair_suggestions: List[str] = field(default_factory=list)


@dataclass
class StructureData:
    """Structure mapper module result data"""
    worksheet_count: int
    worksheet_names: List[str]
    named_ranges: Dict[str, str]
    sheet_relationships: Dict[str, List[str]]
    hidden_sheets: List[str]
    chart_count: int
    pivot_table_count: int
    total_cells: int
    non_empty_cells: int


@dataclass
class DataProfileData:
    """Data profiler module result data"""
    sheet_profiles: Dict[str, Dict[str, Any]]
    column_statistics: Dict[str, Dict[str, Any]]
    data_quality_score: float
    null_percentages: Dict[str, float]
    data_types: Dict[str, str]
    outliers_detected: int
    duplicate_rows: int
    patterns_found: List[Dict[str, Any]]


@dataclass
class FormulaAnalysisData:
    """Formula analyzer module result data"""
    total_formulas: int
    formula_complexity_score: float
    circular_references: List[str]
    external_references: List[str]
    formula_errors: List[Dict[str, str]]
    dependency_chains: List[List[str]]
    volatile_formulas: int
    array_formulas: int


@dataclass
class VisualCatalogData:
    """Visual cataloger module result data"""
    charts: List[Dict[str, Any]]
    images: List[Dict[str, Any]]
    shapes: List[Dict[str, Any]]
    conditional_formatting: List[Dict[str, Any]]
    data_validation: List[Dict[str, Any]]
    visual_complexity_score: float


@dataclass
class ConnectionData:
    """Connection inspector module result data"""
    external_connections: List[Dict[str, Any]]
    linked_workbooks: List[str]
    database_connections: List[Dict[str, Any]]
    web_queries: List[Dict[str, Any]]
    refresh_settings: Dict[str, Any]
    security_assessment: Dict[str, Any]


@dataclass
class PivotAnalysisData:
    """Pivot intelligence module result data"""
    pivot_tables: List[Dict[str, Any]]
    pivot_charts: List[Dict[str, Any]]
    data_sources: List[str]
    calculated_fields: List[Dict[str, Any]]
    slicers: List[Dict[str, Any]]
    refresh_metadata: Dict[str, Any]

@dataclass
class DocumentationData:
    """Documentation synthesizer module result data"""
    file_overview: Dict[str, Any]
    executive_summary: str
    detailed_analysis: Dict[str, Any]
    recommendations: List[str]
    ai_navigation_guide: Dict[str, Any]
    metadata: Dict[str, Any]


class ResultCollector:
    """Collects and manages module results throughout analysis pipeline"""
    
    def __init__(self):
        self.results: Dict[str, ModuleResult] = {}
        self.execution_order: List[str] = []
        self.start_time = time.time()
    
    def add_result(self, result: ModuleResult) -> None:
        """Add module result to collection"""
        self.results[result.module_name] = result
        if result.module_name not in self.execution_order:
            self.execution_order.append(result.module_name)
    
    def get_result(self, module_name: str) -> Optional[ModuleResult]:
        """Get result for specific module"""
        return self.results.get(module_name)
    
    def get_successful_results(self) -> Dict[str, ModuleResult]:
        """Get all successful module results"""
        return {
            name: result for name, result in self.results.items()
            if result.is_successful
        }
    
    def get_failed_results(self) -> Dict[str, ModuleResult]:
        """Get all failed module results"""
        return {
            name: result for name, result in self.results.items()
            if not result.is_successful
        }
    
    def calculate_overall_success_rate(self) -> float:
        """Calculate overall pipeline success rate"""
        if not self.results:
            return 0.0
        
        successful = sum(1 for result in self.results.values() if result.is_successful)
        return successful / len(self.results)
    
    def calculate_average_quality_score(self) -> float:
        """Calculate average quality score across all modules"""
        if not self.results:
            return 0.0
        
        scores = [result.quality_score for result in self.results.values()]
        return sum(scores) / len(scores)
    
    def get_execution_summary(self) -> Dict[str, Any]:
        """Get comprehensive execution summary"""
        total_duration = time.time() - self.start_time
        
        # Performance metrics
        total_items_processed = sum(
            result.metrics.items_processed if result.metrics else 0
            for result in self.results.values()
        )
        
        total_memory_delta = sum(
            result.metrics.memory_delta_mb if result.metrics else 0
            for result in self.results.values()
        )
        
        # Error aggregation
        all_errors = []
        all_warnings = []
        for result in self.results.values():
            all_errors.extend(result.errors)
            all_warnings.extend(result.warnings)
        
        return {
            'total_modules': len(self.results),
            'successful_modules': len(self.get_successful_results()),
            'failed_modules': len(self.get_failed_results()),
            'success_rate': self.calculate_overall_success_rate(),
            'average_quality_score': self.calculate_average_quality_score(),
            'total_duration_seconds': total_duration,
            'total_items_processed': total_items_processed,
            'total_memory_delta_mb': total_memory_delta,
            'total_errors': len(all_errors),
            'total_warnings': len(all_warnings),
            'execution_order': self.execution_order.copy(),
            'module_statuses': {
                name: result.status.value for name, result in self.results.items()
            }
        }
    
    def export_results(self, output_path: Path, include_data: bool = True) -> None:
        """Export all results to JSON file"""
        export_data = {
            'summary': self.get_execution_summary(),
            'results': {}
        }
        
        for name, result in self.results.items():
            result_dict = result.to_dict()
            if not include_data:
                result_dict.pop('data', None)
            export_data['results'][name] = result_dict
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2, default=str)
    
    def has_critical_failures(self) -> bool:
        """Check if any critical modules failed"""
        critical_modules = ['health_checker', 'structure_mapper']
        return any(
            self.results.get(module, ModuleResult(module, ResultStatus.FAILED)).status == ResultStatus.FAILED
            for module in critical_modules
        )
    
    def get_dependency_chain_status(self, target_module: str) -> Dict[str, bool]:
        """Get status of all dependencies for a target module"""
        result = self.get_result(target_module)
        if not result:
            return {}
        
        dependency_status = {}
        for dep in result.dependencies_met + result.dependencies_failed:
            dep_result = self.get_result(dep)
            dependency_status[dep] = dep_result.is_successful if dep_result else False
        
        return dependency_status


class ResultValidator:
    """Validates and scores module results"""
    
    @staticmethod
    def validate_health_check(result: ModuleResult[HealthCheckData]) -> ValidationResult:
        """Validate health check results"""
        if not result.data:
            return ValidationResult(0.0, 0.0, 0.0, ConfidenceLevel.UNCERTAIN)
        
        completeness = 1.0  # Health check is binary
        accuracy = 0.9 if result.data.file_accessible else 0.1
        consistency = 0.9 if not result.data.corruption_detected else 0.5
        
        confidence = ConfidenceLevel.HIGH if accuracy > 0.8 else ConfidenceLevel.LOW
        
        return ValidationResult(completeness, accuracy, consistency, confidence)
    
    @staticmethod
    def validate_structure_mapping(result: ModuleResult[StructureData]) -> ValidationResult:
        """Validate structure mapping results"""
        if not result.data:
            return ValidationResult(0.0, 0.0, 0.0, ConfidenceLevel.UNCERTAIN)
        
        # Completeness: Did we map all expected elements?
        expected_elements = ['worksheet_names', 'named_ranges', 'chart_count']
        mapped_elements = sum(1 for elem in expected_elements if getattr(result.data, elem, None) is not None)
        completeness = mapped_elements / len(expected_elements)
        
        # Accuracy: Basic sanity checks
        accuracy = 1.0
        if result.data.worksheet_count != len(result.data.worksheet_names):
            accuracy -= 0.2
        if result.data.total_cells < result.data.non_empty_cells:
            accuracy -= 0.3
        
        consistency = 0.9  # Assume consistent unless proven otherwise
        confidence = ConfidenceLevel.HIGH if accuracy > 0.8 else ConfidenceLevel.MEDIUM
        
        return ValidationResult(completeness, max(0.0, accuracy), consistency, confidence)
    
    @staticmethod
    def validate_data_profiling(result: ModuleResult[DataProfileData]) -> ValidationResult:
        """Validate data profiling results"""
        if not result.data:
            return ValidationResult(0.0, 0.0, 0.0, ConfidenceLevel.UNCERTAIN)
        
        # Check if all sheets were profiled
        completeness = 0.8 if result.data.sheet_profiles else 0.2
        
        # Accuracy based on data quality score
        accuracy = result.data.data_quality_score if result.data.data_quality_score <= 1.0 else 0.5
        
        # Consistency: Check for logical relationships
        consistency = 0.9
        if result.data.duplicate_rows < 0:
            consistency -= 0.3
        
        confidence = ConfidenceLevel.HIGH if accuracy > 0.7 else ConfidenceLevel.MEDIUM
        
        return ValidationResult(completeness, accuracy, consistency, confidence)
