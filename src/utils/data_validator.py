"""
Data Validation Framework - Comprehensive data validation and quality assessment
Provides business logic validation, consistency checking, and anomaly detection.
"""

import logging
from typing import List, Dict, Any, Optional, Set, Tuple, Union, Callable
from dataclasses import dataclass, field
from enum import Enum
import pandas as pd
import numpy as np
import re
from datetime import datetime, date
from collections import Counter, defaultdict

from .error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class ValidationLevel(Enum):
    """Validation severity levels"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


class ValidationType(Enum):
    """Types of validation checks"""
    DATA_TYPE = "data_type"
    RANGE = "range"
    FORMAT = "format"
    BUSINESS_RULE = "business_rule"
    CONSISTENCY = "consistency"
    COMPLETENESS = "completeness"
    UNIQUENESS = "uniqueness"
    RELATIONSHIP = "relationship"


@dataclass
class ValidationRule:
    """Individual validation rule definition"""
    name: str
    validation_type: ValidationType
    level: ValidationLevel
    condition: Callable[[Any], bool]
    message: str
    column_names: Optional[List[str]] = None
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ValidationResult:
    """Result of a validation check"""
    rule_name: str
    validation_type: ValidationType
    level: ValidationLevel
    passed: bool
    message: str
    affected_rows: List[int] = field(default_factory=list)
    affected_columns: List[str] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ValidationSummary:
    """Summary of all validation results"""
    total_rules: int
    passed_rules: int
    failed_rules: int
    critical_failures: int
    error_failures: int
    warning_failures: int
    info_items: int
    overall_score: float
    results: List[ValidationResult] = field(default_factory=list)


class DataValidator:
    """Comprehensive data validation framework with configurable rules"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.rules: List[ValidationRule] = []
        self.custom_validators: Dict[str, Callable] = {}
        self._initialize_default_rules()
    
    def add_rule(self, rule: ValidationRule) -> None:
        """Add a validation rule to the framework"""
        self.rules.append(rule)
        self.logger.debug(f"Added validation rule: {rule.name}")
    
    def add_custom_validator(self, name: str, validator_func: Callable) -> None:
        """Add a custom validator function"""
        self.custom_validators[name] = validator_func
        self.logger.debug(f"Added custom validator: {name}")
    
    def validate_dataframe(self, df: pd.DataFrame, 
                          column_metadata: Optional[Dict[str, Dict[str, Any]]] = None) -> ValidationSummary:
        """Validate a pandas DataFrame against all configured rules
        
        Args:
            df: DataFrame to validate
            column_metadata: Optional metadata about columns for enhanced validation
            
        Returns:
            ValidationSummary with all results
        """
        results = []
        
        self.logger.info(f"Starting validation of DataFrame with {len(df)} rows and {len(df.columns)} columns")
        
        for rule in self.rules:
            try:
                result = self._apply_rule(df, rule, column_metadata)
                results.append(result)
                
                if not result.passed and result.level in [ValidationLevel.ERROR, ValidationLevel.CRITICAL]:
                    self.logger.warning(f"Validation failed: {rule.name} - {result.message}")
                
            except Exception as e:
                self.logger.error(f"Error applying rule {rule.name}: {e}")
                results.append(ValidationResult(
                    rule_name=rule.name,
                    validation_type=rule.validation_type,
                    level=ValidationLevel.ERROR,
                    passed=False,
                    message=f"Rule execution failed: {e}",
                    details={'exception': str(e)}
                ))
        
        return self._create_summary(results)
    
    def validate_column_data_types(self, df: pd.DataFrame, 
                                  expected_types: Dict[str, str]) -> List[ValidationResult]:
        """Validate column data types against expected types
        
        Args:
            df: DataFrame to validate
            expected_types: Dict mapping column names to expected type names
            
        Returns:
            List of validation results
        """
        results = []
        
        for col_name, expected_type in expected_types.items():
            if col_name not in df.columns:
                results.append(ValidationResult(
                    rule_name=f"column_exists_{col_name}",
                    validation_type=ValidationType.DATA_TYPE,
                    level=ValidationLevel.ERROR,
                    passed=False,
                    message=f"Expected column '{col_name}' not found",
                    affected_columns=[col_name]
                ))
                continue
            
            # Check data type consistency
            series = df[col_name].dropna()
            if series.empty:
                continue
                
            type_consistency = self._check_type_consistency(series, expected_type)
            
            results.append(ValidationResult(
                rule_name=f"data_type_{col_name}",
                validation_type=ValidationType.DATA_TYPE,
                level=ValidationLevel.WARNING if type_consistency['consistency_ratio'] > 0.8 else ValidationLevel.ERROR,
                passed=type_consistency['consistency_ratio'] > 0.9,
                message=f"Column '{col_name}' type consistency: {type_consistency['consistency_ratio']:.1%}",
                affected_columns=[col_name],
                details=type_consistency
            ))
        
        return results
    
    def validate_business_rules(self, df: pd.DataFrame,
                               business_rules: List[Dict[str, Any]]) -> List[ValidationResult]:
        """Validate against business-specific rules
        
        Args:
            df: DataFrame to validate
            business_rules: List of business rule definitions
            
        Returns:
            List of validation results
        """
        results = []
        
        for rule_def in business_rules:
            try:
                result = self._apply_business_rule(df, rule_def)
                results.append(result)
            except Exception as e:
                self.logger.error(f"Error applying business rule {rule_def.get('name', 'unknown')}: {e}")
                results.append(ValidationResult(
                    rule_name=rule_def.get('name', 'unknown_business_rule'),
                    validation_type=ValidationType.BUSINESS_RULE,
                    level=ValidationLevel.ERROR,
                    passed=False,
                    message=f"Business rule execution failed: {e}"
                ))
        
        return results
    
    def validate_data_consistency(self, df: pd.DataFrame) -> List[ValidationResult]:
        """Validate data consistency across columns and rows
        
        Args:
            df: DataFrame to validate
            
        Returns:
            List of validation results
        """
        results = []
        
        # Check for consistent data patterns within columns
        for col in df.columns:
            series = df[col].dropna()
            if len(series) < 10:  # Skip small datasets
                continue
            
            # Pattern consistency check
            pattern_consistency = self._check_pattern_consistency(series)
            
            if pattern_consistency['consistency_score'] < 0.8:
                results.append(ValidationResult(
                    rule_name=f"pattern_consistency_{col}",
                    validation_type=ValidationType.CONSISTENCY,
                    level=ValidationLevel.WARNING,
                    passed=False,
                    message=f"Column '{col}' has inconsistent patterns (score: {pattern_consistency['consistency_score']:.2f})",
                    affected_columns=[col],
                    details=pattern_consistency
                ))
        
        # Check cross-column relationships
        cross_column_results = self._validate_cross_column_consistency(df)
        results.extend(cross_column_results)
        
        return results
    
    def detect_anomalies(self, df: pd.DataFrame, 
                        sensitivity: float = 0.95) -> List[ValidationResult]:
        """Detect anomalies in the dataset
        
        Args:
            df: DataFrame to analyze
            sensitivity: Anomaly detection sensitivity (0.0 to 1.0)
            
        Returns:
            List of validation results for detected anomalies
        """
        results = []
        
        for col in df.columns:
            series = df[col].dropna()
            
            if pd.api.types.is_numeric_dtype(series) and len(series) > 10:
                # Statistical outlier detection
                outliers = self._detect_statistical_outliers(series, sensitivity)
                
                if outliers['outlier_count'] > 0:
                    severity = ValidationLevel.WARNING if outliers['outlier_ratio'] < 0.05 else ValidationLevel.ERROR
                    
                    results.append(ValidationResult(
                        rule_name=f"statistical_outliers_{col}",
                        validation_type=ValidationType.RANGE,
                        level=severity,
                        passed=outliers['outlier_ratio'] < 0.1,
                        message=f"Column '{col}' has {outliers['outlier_count']} statistical outliers ({outliers['outlier_ratio']:.1%})",
                        affected_rows=outliers['outlier_indices'],
                        affected_columns=[col],
                        details=outliers
                    ))
            
            elif pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
                # Text anomaly detection
                text_anomalies = self._detect_text_anomalies(series)
                
                if text_anomalies['anomaly_count'] > 0:
                    results.append(ValidationResult(
                        rule_name=f"text_anomalies_{col}",
                        validation_type=ValidationType.FORMAT,
                        level=ValidationLevel.WARNING,
                        passed=text_anomalies['anomaly_ratio'] < 0.05,
                        message=f"Column '{col}' has {text_anomalies['anomaly_count']} text anomalies",
                        affected_rows=text_anomalies['anomaly_indices'],
                        affected_columns=[col],
                        details=text_anomalies
                    ))
        
        return results
    
    def validate_completeness(self, df: pd.DataFrame,
                            required_columns: Optional[List[str]] = None,
                            completeness_threshold: float = 0.95) -> List[ValidationResult]:
        """Validate data completeness
        
        Args:
            df: DataFrame to validate
            required_columns: List of columns that must be complete
            completeness_threshold: Minimum acceptable completeness ratio
            
        Returns:
            List of validation results
        """
        results = []
        
        # Overall completeness
        total_cells = len(df) * len(df.columns)
        non_null_cells = total_cells - df.isnull().sum().sum()
        overall_completeness = non_null_cells / total_cells if total_cells > 0 else 0.0
        
        results.append(ValidationResult(
            rule_name="overall_completeness",
            validation_type=ValidationType.COMPLETENESS,
            level=ValidationLevel.WARNING if overall_completeness < completeness_threshold else ValidationLevel.INFO,
            passed=overall_completeness >= completeness_threshold,
            message=f"Overall data completeness: {overall_completeness:.1%}",
            details={'completeness_ratio': overall_completeness, 'threshold': completeness_threshold}
        ))
        
        # Column-specific completeness
        for col in df.columns:
            null_count = df[col].isnull().sum()
            completeness = 1.0 - (null_count / len(df))
            
            is_required = required_columns and col in required_columns
            threshold = 1.0 if is_required else completeness_threshold
            
            level = ValidationLevel.ERROR if is_required and completeness < 1.0 else (
                ValidationLevel.WARNING if completeness < threshold else ValidationLevel.INFO
            )
            
            results.append(ValidationResult(
                rule_name=f"completeness_{col}",
                validation_type=ValidationType.COMPLETENESS,
                level=level,
                passed=completeness >= threshold,
                message=f"Column '{col}' completeness: {completeness:.1%}",
                affected_columns=[col],
                details={'completeness_ratio': completeness, 'null_count': null_count, 'is_required': is_required}
            ))
        
        return results
    
    def _initialize_default_rules(self) -> None:
        """Initialize default validation rules"""
        # Null value checks
        self.add_rule(ValidationRule(
            name="no_all_null_columns",
            validation_type=ValidationType.COMPLETENESS,
            level=ValidationLevel.WARNING,
            condition=lambda df: not any(df[col].isnull().all() for col in df.columns),
            message="Dataset contains columns with all null values"
        ))
        
        # Duplicate row detection
        self.add_rule(ValidationRule(
            name="no_duplicate_rows",
            validation_type=ValidationType.UNIQUENESS,
            level=ValidationLevel.WARNING,
            condition=lambda df: not df.duplicated().any(),
            message="Dataset contains duplicate rows"
        ))
        
        # Basic data type consistency
        self.add_rule(ValidationRule(
            name="consistent_numeric_types",
            validation_type=ValidationType.DATA_TYPE,
            level=ValidationLevel.WARNING,
            condition=lambda df: self._check_numeric_consistency(df),
            message="Numeric columns have inconsistent data types"
        ))
    
    def _apply_rule(self, df: pd.DataFrame, rule: ValidationRule,
                   column_metadata: Optional[Dict[str, Dict[str, Any]]] = None) -> ValidationResult:
        """Apply a single validation rule to the DataFrame"""
        try:
            # Filter DataFrame to relevant columns if specified
            target_df = df
            if rule.column_names:
                available_cols = [col for col in rule.column_names if col in df.columns]
                if available_cols:
                    target_df = df[available_cols]
                else:
                    return ValidationResult(
                        rule_name=rule.name,
                        validation_type=rule.validation_type,
                        level=ValidationLevel.ERROR,
                        passed=False,
                        message=f"None of the required columns {rule.column_names} found in dataset",
                        affected_columns=rule.column_names or []
                    )
            
            # Apply the rule condition
            passed = rule.condition(target_df)
            
            # Determine affected rows/columns for failed rules
            affected_rows = []
            affected_columns = rule.column_names or []
            
            if not passed and rule.validation_type in [ValidationType.UNIQUENESS, ValidationType.CONSISTENCY]:
                affected_rows = self._identify_affected_rows(df, rule)
            
            return ValidationResult(
                rule_name=rule.name,
                validation_type=rule.validation_type,
                level=rule.level,
                passed=passed,
                message=rule.message,
                affected_rows=affected_rows,
                affected_columns=affected_columns,
                details=rule.metadata
            )
            
        except Exception as e:
            return ValidationResult(
                rule_name=rule.name,
                validation_type=rule.validation_type,
                level=ValidationLevel.ERROR,
                passed=False,
                message=f"Rule execution failed: {e}",
                details={'exception': str(e)}
            )
    
    def _apply_business_rule(self, df: pd.DataFrame, rule_def: Dict[str, Any]) -> ValidationResult:
        """Apply a business rule defined in configuration"""
        rule_name = rule_def.get('name', 'unknown_rule')
        rule_type = rule_def.get('type', 'custom')
        level = ValidationLevel(rule_def.get('level', 'warning'))
        
        if rule_type == 'range_check':
            return self._apply_range_check(df, rule_def)
        elif rule_type == 'format_check':
            return self._apply_format_check(df, rule_def)
        elif rule_type == 'relationship_check':
            return self._apply_relationship_check(df, rule_def)
        elif rule_type == 'custom' and 'expression' in rule_def:
            return self._apply_expression_rule(df, rule_def)
        else:
            return ValidationResult(
                rule_name=rule_name,
                validation_type=ValidationType.BUSINESS_RULE,
                level=ValidationLevel.ERROR,
                passed=False,
                message=f"Unknown business rule type: {rule_type}"
            )
    
    def _apply_range_check(self, df: pd.DataFrame, rule_def: Dict[str, Any]) -> ValidationResult:
        """Apply numeric range validation"""
        column = rule_def['column']
        min_val = rule_def.get('min_value')
        max_val = rule_def.get('max_value')
        
        if column not in df.columns:
            return ValidationResult(
                rule_name=rule_def['name'],
                validation_type=ValidationType.RANGE,
                level=ValidationLevel.ERROR,
                passed=False,
                message=f"Column '{column}' not found for range check"
            )
        
        series = pd.to_numeric(df[column], errors='coerce').dropna()
        violations = []
        
        if min_val is not None:
            violations.extend(series[series < min_val].index.tolist())
        if max_val is not None:
            violations.extend(series[series > max_val].index.tolist())
        
        violations = list(set(violations))  # Remove duplicates
        
        return ValidationResult(
            rule_name=rule_def['name'],
            validation_type=ValidationType.RANGE,
            level=ValidationLevel(rule_def.get('level', 'warning')),
            passed=len(violations) == 0,
            message=f"Column '{column}' range validation: {len(violations)} violations found",
            affected_rows=violations,
            affected_columns=[column]
        )
    
    def _apply_format_check(self, df: pd.DataFrame, rule_def: Dict[str, Any]) -> ValidationResult:
        """Apply format validation using regex"""
        column = rule_def['column']
        pattern = rule_def['pattern']
        
        if column not in df.columns:
            return ValidationResult(
                rule_name=rule_def['name'],
                validation_type=ValidationType.FORMAT,
                level=ValidationLevel.ERROR,
                passed=False,
                message=f"Column '{column}' not found for format check"
            )
        
        series = df[column].astype(str).dropna()
        regex = re.compile(pattern)
        violations = []
        
        for idx, value in series.items():
            if not regex.match(value):
                violations.append(idx)
        
        return ValidationResult(
            rule_name=rule_def['name'],
            validation_type=ValidationType.FORMAT,
            level=ValidationLevel(rule_def.get('level', 'warning')),
            passed=len(violations) == 0,
            message=f"Column '{column}' format validation: {len(violations)} violations found",
            affected_rows=violations,
            affected_columns=[column]
        )
    
    def _apply_relationship_check(self, df: pd.DataFrame, rule_def: Dict[str, Any]) -> ValidationResult:
        """Apply relationship validation between columns"""
        # Implement relationship checks (e.g., foreign key, conditional logic)
        # This is a placeholder for more complex relationship validation
        return ValidationResult(
            rule_name=rule_def['name'],
            validation_type=ValidationType.RELATIONSHIP,
            level=ValidationLevel.INFO,
            passed=True,
            message="Relationship check not yet implemented"
        )
    
    def _apply_expression_rule(self, df: pd.DataFrame, rule_def: Dict[str, Any]) -> ValidationResult:
        """Apply custom expression-based rule"""
        try:
            expression = rule_def['expression']
            # Safe evaluation of pandas expressions
            result = df.eval(expression)
            violations = df[~result].index.tolist() if hasattr(result, 'index') else []
            
            return ValidationResult(
                rule_name=rule_def['name'],
                validation_type=ValidationType.BUSINESS_RULE,
                level=ValidationLevel(rule_def.get('level', 'warning')),
                passed=len(violations) == 0,
                message=f"Expression rule validation: {len(violations)} violations found",
                affected_rows=violations
            )
        except Exception as e:
            return ValidationResult(
                rule_name=rule_def['name'],
                validation_type=ValidationType.BUSINESS_RULE,
                level=ValidationLevel.ERROR,
                passed=False,
                message=f"Expression rule failed: {e}"
            )
    
    def _check_type_consistency(self, series: pd.Series, expected_type: str) -> Dict[str, Any]:
        """Check data type consistency within a series"""
        type_mapping = {
            'integer': lambda x: pd.api.types.is_integer_dtype(x),
            'float': lambda x: pd.api.types.is_float_dtype(x),
            'numeric': lambda x: pd.api.types.is_numeric_dtype(x),
            'string': lambda x: pd.api.types.is_string_dtype(x),
            'datetime': lambda x: pd.api.types.is_datetime64_any_dtype(x),
            'boolean': lambda x: pd.api.types.is_bool_dtype(x)
        }
        
        if expected_type not in type_mapping:
            return {'consistency_ratio': 0.0, 'error': f'Unknown type: {expected_type}'}
        
        # Convert and check consistency
        try:
            if expected_type == 'integer':
                converted = pd.to_numeric(series, errors='coerce')
                valid_count = (~converted.isna()).sum()
                int_count = (converted == converted.astype('Int64', errors='ignore')).sum()
                consistency = int_count / len(series) if len(series) > 0 else 0
            elif expected_type in ['float', 'numeric']:
                converted = pd.to_numeric(series, errors='coerce')
                consistency = (~converted.isna()).sum() / len(series) if len(series) > 0 else 0
            elif expected_type == 'datetime':
                converted = pd.to_datetime(series, errors='coerce')
                consistency = (~converted.isna()).sum() / len(series) if len(series) > 0 else 0
            else:
                # For string and boolean, check if conversion is possible
                consistency = 1.0  # Default for string types
            
            return {
                'consistency_ratio': consistency,
                'expected_type': expected_type,
                'actual_dtype': str(series.dtype)
            }
            
        except Exception as e:
            return {'consistency_ratio': 0.0, 'error': str(e)}
    
    def _check_pattern_consistency(self, series: pd.Series) -> Dict[str, Any]:
        """Check pattern consistency within a series"""
        str_series = series.astype(str)
        
        # Analyze common patterns
        patterns = defaultdict(int)
        
        for value in str_series:
            # Simple pattern analysis
            pattern = re.sub(r'\d', 'D', value)  # Replace digits with D
            pattern = re.sub(r'[a-zA-Z]', 'A', pattern)  # Replace letters with A
            pattern = re.sub(r'[^\w]', 'S', pattern)  # Replace special chars with S
            patterns[pattern] += 1
        
        if not patterns:
            return {'consistency_score': 1.0, 'dominant_pattern': None}
        
        total_count = len(str_series)
        dominant_pattern = max(patterns, key=patterns.get)
        dominant_count = patterns[dominant_pattern]
        consistency_score = dominant_count / total_count
        
        return {
            'consistency_score': consistency_score,
            'dominant_pattern': dominant_pattern,
            'pattern_distribution': dict(patterns),
            'unique_patterns': len(patterns)
        }
    
    def _validate_cross_column_consistency(self, df: pd.DataFrame) -> List[ValidationResult]:
        """Validate consistency across related columns"""
        results = []
        
        # Look for columns that might be related (similar names, types)
        potential_relationships = self._identify_column_relationships(df)
        
        for relationship in potential_relationships:
            try:
                consistency_check = self._check_relationship_consistency(df, relationship)
                
                if not consistency_check['consistent']:
                    results.append(ValidationResult(
                        rule_name=f"cross_column_consistency_{relationship['type']}",
                        validation_type=ValidationType.CONSISTENCY,
                        level=ValidationLevel.WARNING,
                        passed=False,
                        message=consistency_check['message'],
                        affected_columns=relationship['columns'],
                        details=consistency_check
                    ))
            except Exception as e:
                self.logger.warning(f"Error checking cross-column consistency: {e}")
        
        return results
    
    def _identify_column_relationships(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Identify potential relationships between columns"""
        relationships = []
        
        # Date/time relationships
        date_cols = [col for col in df.columns if 'date' in col.lower() or 'time' in col.lower()]
        if len(date_cols) >= 2:
            relationships.append({
                'type': 'temporal',
                'columns': date_cols,
                'description': 'Temporal relationship between date/time columns'
            })
        
        # Numeric relationships (totals, subtotals)
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        total_cols = [col for col in numeric_cols if 'total' in col.lower() or 'sum' in col.lower()]
        if total_cols and len(numeric_cols) > len(total_cols):
            relationships.append({
                'type': 'summation',
                'columns': numeric_cols,
                'total_columns': total_cols,
                'description': 'Potential summation relationship'
            })
        
        return relationships
    
    def _check_relationship_consistency(self, df: pd.DataFrame, relationship: Dict[str, Any]) -> Dict[str, Any]:
        """Check consistency of a specific relationship"""
        rel_type = relationship['type']
        
        if rel_type == 'temporal':
            return self._check_temporal_consistency(df, relationship['columns'])
        elif rel_type == 'summation':
            return self._check_summation_consistency(df, relationship)
        else:
            return {'consistent': True, 'message': 'Unknown relationship type'}
    
    def _check_temporal_consistency(self, df: pd.DataFrame, date_columns: List[str]) -> Dict[str, Any]:
        """Check temporal consistency between date columns"""
        try:
            date_data = {}
            for col in date_columns:
                if col in df.columns:
                    date_data[col] = pd.to_datetime(df[col], errors='coerce')
            
            if len(date_data) < 2:
                return {'consistent': True, 'message': 'Insufficient date columns for comparison'}
            
            # Check for logical temporal relationships
            inconsistencies = 0
            for i, row in df.iterrows():
                dates = [date_data[col].iloc[i] for col in date_data.keys() if pd.notna(date_data[col].iloc[i])]
                if len(dates) >= 2:
                    # Check if dates are in logical order (this is a simple heuristic)
                    if not all(dates[i] <= dates[i+1] for i in range(len(dates)-1)):
                        inconsistencies += 1
            
            consistency_ratio = 1.0 - (inconsistencies / len(df)) if len(df) > 0 else 1.0
            
            return {
                'consistent': consistency_ratio > 0.95,
                'consistency_ratio': consistency_ratio,
                'inconsistencies': inconsistencies,
                'message': f"Temporal consistency: {consistency_ratio:.1%}"
            }
            
        except Exception as e:
            return {'consistent': False, 'message': f'Temporal consistency check failed: {e}'}
    
    def _check_summation_consistency(self, df: pd.DataFrame, relationship: Dict[str, Any]) -> Dict[str, Any]:
        """Check summation consistency between numeric columns"""
        try:
            numeric_cols = relationship['columns']
            total_cols = relationship.get('total_columns', [])
            
            if not total_cols:
                return {'consistent': True, 'message': 'No total columns identified'}
            
            # Simple check: see if any total column approximately equals sum of others
            component_cols = [col for col in numeric_cols if col not in total_cols]
            
            inconsistencies = 0
            for total_col in total_cols:
                if total_col in df.columns and component_cols:
                    total_values = pd.to_numeric(df[total_col], errors='coerce')
                    component_sum = df[component_cols].select_dtypes(include=[np.number]).sum(axis=1)
                    
                    # Allow for small rounding differences
                    differences = abs(total_values - component_sum)
                    tolerance = abs(component_sum * 0.01)  # 1% tolerance
                    inconsistent_rows = (differences > tolerance).sum()
                    inconsistencies += inconsistent_rows
            
            consistency_ratio = 1.0 - (inconsistencies / (len(df) * len(total_cols))) if len(df) > 0 else 1.0
            
            return {
                'consistent': consistency_ratio > 0.9,
                'consistency_ratio': consistency_ratio,
                'inconsistencies': inconsistencies,
                'message': f"Summation consistency: {consistency_ratio:.1%}"
            }
            
        except Exception as e:
            return {'consistent': False, 'message': f'Summation consistency check failed: {e}'}
    
    def _detect_statistical_outliers(self, series: pd.Series, sensitivity: float) -> Dict[str, Any]:
        """Detect statistical outliers using IQR and Z-score methods"""
        numeric_series = pd.to_numeric(series, errors='coerce').dropna()
        
        if len(numeric_series) < 10:
            return {'outlier_count': 0, 'outlier_ratio': 0.0, 'outlier_indices': []}
        
        # IQR method
        Q1 = numeric_series.quantile(0.25)
        Q3 = numeric_series.quantile(0.75)
        IQR = Q3 - Q1
        
        # Adjust multiplier based on sensitivity
        multiplier = 1.5 + (1.0 - sensitivity) * 2.0  # Range: 1.5 to 3.5
        
        lower_bound = Q1 - multiplier * IQR
        upper_bound = Q3 + multiplier * IQR
        
        iqr_outliers = numeric_series[(numeric_series < lower_bound) | (numeric_series > upper_bound)]
        
        # Z-score method for additional validation
        z_threshold = 2.0 + (1.0 - sensitivity) * 1.0  # Range: 2.0 to 3.0
        z_scores = np.abs((numeric_series - numeric_series.mean()) / numeric_series.std())
        z_outliers = numeric_series[z_scores > z_threshold]
        
        # Combine both methods
        combined_outliers = set(iqr_outliers.index) | set(z_outliers.index)
        
        return {
            'outlier_count': len(combined_outliers),
            'outlier_ratio': len(combined_outliers) / len(series),
            'outlier_indices': list(combined_outliers),
            'iqr_outliers': len(iqr_outliers),
            'z_score_outliers': len(z_outliers),
            'bounds': {'lower': lower_bound, 'upper': upper_bound},
            'z_threshold': z_threshold
        }
    
    def _detect_text_anomalies(self, series: pd.Series) -> Dict[str, Any]:
        """Detect anomalies in text data"""
        str_series = series.astype(str).dropna()
        
        if len(str_series) < 10:
            return {'anomaly_count': 0, 'anomaly_ratio': 0.0, 'anomaly_indices': []}
        
        # Length-based anomalies
        lengths = str_series.str.len()
        length_q1 = lengths.quantile(0.25)
        length_q3 = lengths.quantile(0.75)
        length_iqr = length_q3 - length_q1
        
        length_outliers = lengths[(lengths < length_q1 - 2 * length_iqr) | 
                                 (lengths > length_q3 + 2 * length_iqr)]
        
        # Character composition anomalies
        anomaly_indices = set(length_outliers.index)
        
        # Check for unusual character patterns
        for idx, value in str_series.items():
            if self._has_unusual_characters(value):
                anomaly_indices.add(idx)
        
        return {
            'anomaly_count': len(anomaly_indices),
            'anomaly_ratio': len(anomaly_indices) / len(series),
            'anomaly_indices': list(anomaly_indices),
            'length_outliers': len(length_outliers)
        }
    
    def _has_unusual_characters(self, text: str) -> bool:
        """Check if text contains unusual character patterns"""
        # Check for excessive special characters
        special_char_ratio = len(re.findall(r'[^a-zA-Z0-9\s]', text)) / len(text) if text else 0
        
        # Check for mixed scripts (simplified)
        has_ascii = bool(re.search(r'[a-zA-Z]', text))
        has_non_ascii = bool(re.search(r'[^\x00-\x7F]', text))
        
        return special_char_ratio > 0.5 or (has_ascii and has_non_ascii and len(text) < 50)
    
    def _check_numeric_consistency(self, df: pd.DataFrame) -> bool:
        """Check consistency of numeric types across the dataset"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) == 0:
            return True
        
        # Check for mixed integer/float types that might indicate inconsistency
        inconsistent_cols = 0
        
        for col in numeric_cols:
            series = df[col].dropna()
            if len(series) == 0:
                continue
            
            # Check if column has mix of integers and floats
            if pd.api.types.is_float_dtype(series):
                # Check if all values are actually integers
                is_all_int = (series == series.astype(int)).all()
                if not is_all_int:
                    # Check ratio of integer vs float values
                    int_count = (series == series.astype(int)).sum()
                    if 0.1 < int_count / len(series) < 0.9:  # Mixed types
                        inconsistent_cols += 1
        
        return inconsistent_cols == 0
    
    def _identify_affected_rows(self, df: pd.DataFrame, rule: ValidationRule) -> List[int]:
        """Identify rows affected by a failed validation rule"""
        # This is a simplified implementation
        # In practice, this would need rule-specific logic
        
        if rule.validation_type == ValidationType.UNIQUENESS:
            return df[df.duplicated(keep=False)].index.tolist()
        else:
            return []
    
    def _create_summary(self, results: List[ValidationResult]) -> ValidationSummary:
        """Create a validation summary from individual results"""
        total_rules = len(results)
        passed_rules = sum(1 for r in results if r.passed)
        failed_rules = total_rules - passed_rules
        
        # Count by severity
        critical_failures = sum(1 for r in results if not r.passed and r.level == ValidationLevel.CRITICAL)
        error_failures = sum(1 for r in results if not r.passed and r.level == ValidationLevel.ERROR)
        warning_failures = sum(1 for r in results if not r.passed and r.level == ValidationLevel.WARNING)
        info_items = sum(1 for r in results if r.level == ValidationLevel.INFO)
        
        # Calculate overall score
        if total_rules == 0:
            overall_score = 1.0
        else:
            # Weight different severities
            score = passed_rules
            score -= critical_failures * 1.0
            score -= error_failures * 0.7
            score -= warning_failures * 0.3
            
            overall_score = max(0.0, score / total_rules)
        
        return ValidationSummary(
            total_rules=total_rules,
            passed_rules=passed_rules,
            failed_rules=failed_rules,
            critical_failures=critical_failures,
            error_failures=error_failures,
            warning_failures=warning_failures,
            info_items=info_items,
            overall_score=overall_score,
            results=results
        )


def create_validator_with_business_rules(business_rules: List[Dict[str, Any]]) -> DataValidator:
    """Create a data validator with pre-configured business rules"""
    validator = DataValidator()
    
    # Add business-specific validation rules
    for rule_def in business_rules:
        if rule_def.get('type') == 'completeness':
            validator.add_rule(ValidationRule(
                name=rule_def['name'],
                validation_type=ValidationType.COMPLETENESS,
                level=ValidationLevel(rule_def.get('level', 'warning')),
                condition=lambda df, threshold=rule_def.get('threshold', 0.95): 
                    (1 - df.isnull().sum().sum() / (len(df) * len(df.columns))) >= threshold,
                message=rule_def.get('message', 'Completeness check failed'),
                metadata=rule_def
            ))
    
    return validator
