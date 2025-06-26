"""
Data quality and patterns analysis module
Comprehensive data profiling with type inference and quality metrics.
"""

from typing import List, Optional, Dict, Any, Tuple, Set
import pandas as pd
import numpy as np
from collections import Counter, defaultdict
import re
from datetime import datetime

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import DataProfileData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class DataProfiler(BaseAnalyzer):
    """Enhanced data profiler with comprehensive quality analysis"""
    
    def __init__(self, name: str = "data_profiler", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["health_checker", "structure_mapper"])
    
    def _perform_analysis(self, context: AnalysisContext) -> DataProfileData:
        """Perform comprehensive data profiling across all worksheets
        
        Args:
            context: AnalysisContext with workbook access
            
        Returns:
            DataProfileData with complete profiling results
        """
        try:
            # Get structure information from dependency
            structure_result = context.get_module_result("structure_mapper")
            if not structure_result or not structure_result.data:
                raise ExcelAnalysisError(
                    "Structure mapping data not available",
                    severity=ErrorSeverity.HIGH,
                    category=ErrorCategory.DEPENDENCY_FAILURE,
                    module_name=self.name
                )
            
            structure_data = structure_result.data
            sheet_names = structure_data.worksheet_names
            
            # Profile each worksheet
            sheet_profiles = {}
            column_statistics = {}
            all_null_percentages = {}
            all_data_types = {}
            total_outliers = 0
            total_duplicates = 0
            patterns_found = []
            
            sample_limit = self.config.get("sample_size_limit", 10000)
            
            with context.get_workbook_access().get_workbook() as wb:
                for sheet_name in sheet_names:
                    try:
                        # Profile individual sheet
                        sheet_profile = self._profile_sheet(
                            wb[sheet_name], sheet_name, sample_limit, context
                        )
                        
                        sheet_profiles[sheet_name] = sheet_profile
                        
                        # Aggregate statistics
                        if 'column_stats' in sheet_profile:
                            column_statistics[sheet_name] = sheet_profile['column_stats']
                        
                        if 'null_percentages' in sheet_profile:
                            all_null_percentages[sheet_name] = sheet_profile['null_percentages']
                        
                        if 'data_types' in sheet_profile:
                            all_data_types[sheet_name] = sheet_profile['data_types']
                        
                        total_outliers += sheet_profile.get('outliers_count', 0)
                        total_duplicates += sheet_profile.get('duplicate_rows', 0)
                        
                        if 'patterns' in sheet_profile:
                            patterns_found.extend(sheet_profile['patterns'])
                    
                    except Exception as e:
                        self.logger.warning(f"Failed to profile sheet {sheet_name}: {e}")
                        sheet_profiles[sheet_name] = {
                            'error': str(e),
                            'status': 'failed'
                        }
            
            # Calculate overall data quality score
            data_quality_score = self._calculate_overall_quality_score(sheet_profiles)
            
            return DataProfileData(
                sheet_profiles=sheet_profiles,
                column_statistics=column_statistics,
                data_quality_score=data_quality_score,
                null_percentages=all_null_percentages,
                data_types=all_data_types,
                outliers_detected=total_outliers,
                duplicate_rows=total_duplicates,
                patterns_found=patterns_found
            )
            
        except Exception as e:
            raise ExcelAnalysisError(
                f"Data profiling failed: {e}",
                severity=ErrorSeverity.HIGH,
                category=ErrorCategory.DATA_CORRUPTION,
                module_name=self.name,
                file_path=str(context.file_path)
            )
    
    def _validate_result(self, data: DataProfileData, context: AnalysisContext) -> ValidationResult:
        """Validate data profiling results
        
        Args:
            data: DataProfileData to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness: Did we profile all sheets?
        structure_result = context.get_module_result("structure_mapper")
        if structure_result and structure_result.data:
            expected_sheets = len(structure_result.data.worksheet_names)
            profiled_sheets = len([
                sheet for sheet, profile in data.sheet_profiles.items()
                if 'error' not in profile
            ])
            completeness = profiled_sheets / expected_sheets if expected_sheets > 0 else 0.0
        else:
            completeness = 0.5  # Assume partial if no structure data
        
        # Accuracy based on data quality score
        accuracy = data.data_quality_score
        
        # Consistency checks
        consistency = 0.9
        
        # Check for logical inconsistencies
        if data.duplicate_rows < 0:
            consistency -= 0.3
            validation_notes.append("Negative duplicate row count")
        
        if data.outliers_detected < 0:
            consistency -= 0.2
            validation_notes.append("Negative outlier count")
        
        if not (0.0 <= data.data_quality_score <= 1.0):
            consistency -= 0.4
            validation_notes.append("Data quality score out of range")
        
        # Confidence assessment
        if accuracy > 0.8 and completeness > 0.8:
            confidence = ConfidenceLevel.HIGH
        elif accuracy > 0.6 and completeness > 0.6:
            confidence = ConfidenceLevel.MEDIUM
        else:
            confidence = ConfidenceLevel.LOW
        
        if len(data.sheet_profiles) > 10:
            validation_notes.append("Large number of sheets profiled")
        if data.outliers_detected > 100:
            validation_notes.append("High number of outliers detected")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _profile_sheet(self, worksheet, sheet_name: str, sample_limit: int, 
                      context: AnalysisContext) -> Dict[str, Any]:
        """Profile individual worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the worksheet
            sample_limit: Maximum rows to sample
            context: AnalysisContext for caching
            
        Returns:
            Dict with sheet profiling results
        """
        try:
            # Check cache first
            cache_key = f"sheet_profile_{sheet_name}"
            cached = self.get_cached_result(context, cache_key)
            if cached:
                return cached
            
            # Extract data region
            data_region = self._detect_data_region(worksheet)
            if not data_region:
                return {
                    'status': 'no_data',
                    'message': 'No data region detected'
                }
            
            # Convert to pandas for analysis
            df = self._worksheet_to_dataframe(
                worksheet, data_region, sample_limit
            )
            
            if df.empty:
                return {
                    'status': 'empty',
                    'message': 'No data extracted'
                }
            
            # Perform profiling
            profile = {
                'status': 'success',
                'data_region': data_region,
                'row_count': len(df),
                'column_count': len(df.columns),
                'headers': self._detect_headers(df),
                'data_types': self._infer_data_types(df),
                'null_percentages': self._calculate_null_percentages(df),
                'column_stats': self._calculate_column_statistics(df),
                'duplicate_rows': self._count_duplicate_rows(df),
                'outliers_count': self._detect_outliers(df),
                'patterns': self._detect_patterns(df),
                'quality_score': self._calculate_sheet_quality_score(df)
            }
            
            # Cache result
            self.cache_intermediate_result(context, cache_key, profile)
            
            return profile
            
        except Exception as e:
            self.logger.warning(f"Error profiling sheet {sheet_name}: {e}")
            return {
                'status': 'error',
                'error': str(e)
            }
    
    def _detect_data_region(self, worksheet) -> Optional[Dict[str, int]]:
        """Detect the primary data region in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            
        Returns:
            Dict with min_row, max_row, min_col, max_col or None
        """
        if not worksheet.max_row or not worksheet.max_column:
            return None
        
        # Find first and last non-empty rows/columns
        min_row, max_row = None, None
        min_col, max_col = None, None
        
        # Scan for data boundaries (sample approach for performance)
        sample_step = max(1, worksheet.max_row // 100)  # Sample every nth row
        
        for row_idx in range(1, worksheet.max_row + 1, sample_step):
            row = list(worksheet.iter_rows(
                min_row=row_idx, max_row=row_idx, values_only=True
            ))[0]
            
            if any(cell is not None and str(cell).strip() for cell in row):
                if min_row is None:
                    min_row = row_idx
                max_row = row_idx
        
        # Find column boundaries
        if min_row and max_row:
            for col_idx in range(1, worksheet.max_column + 1):
                col_values = [
                    worksheet.cell(row=r, column=col_idx).value
                    for r in range(min_row, min(min_row + 10, max_row + 1))
                ]
                
                if any(cell is not None and str(cell).strip() for cell in col_values):
                    if min_col is None:
                        min_col = col_idx
                    max_col = col_idx
        
        if all(v is not None for v in [min_row, max_row, min_col, max_col]):
            return {
                'min_row': min_row,
                'max_row': max_row,
                'min_col': min_col,
                'max_col': max_col
            }
        
        return None
    
    def _worksheet_to_dataframe(self, worksheet, data_region: Dict[str, int], 
                               sample_limit: int) -> pd.DataFrame:
        """Convert worksheet region to pandas DataFrame
        
        Args:
            worksheet: openpyxl worksheet object
            data_region: Dict with row/column boundaries
            sample_limit: Maximum rows to extract
            
        Returns:
            pandas DataFrame
        """
        try:
            # Calculate actual range to extract
            start_row = data_region['min_row']
            end_row = min(
                data_region['max_row'],
                start_row + sample_limit - 1
            )
            start_col = data_region['min_col']
            end_col = data_region['max_col']
            
            # Extract data
            data = []
            for row in worksheet.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=start_col,
                max_col=end_col,
                values_only=True
            ):
                data.append(list(row))
            
            if not data:
                return pd.DataFrame()
            
            # Create DataFrame with generic column names
            num_cols = len(data[0]) if data else 0
            df = pd.DataFrame(data, columns=[f'col_{i}' for i in range(num_cols)])
            
            return df
            
        except Exception as e:
            self.logger.warning(f"Error converting worksheet to DataFrame: {e}")
            return pd.DataFrame()
    
    def _detect_headers(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Detect and analyze headers in DataFrame
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Dict with header analysis results
        """
        if df.empty:
            return {'detected': False}
        
        # Simple header detection: check if first row looks like headers
        first_row = df.iloc[0] if len(df) > 0 else None
        
        if first_row is None:
            return {'detected': False}
        
        # Heuristics for header detection
        header_score = 0.0
        
        # Check for text values in first row
        text_count = sum(1 for val in first_row if isinstance(val, str) and val.strip())
        if text_count > len(first_row) * 0.7:
            header_score += 0.4
        
        # Check for uniqueness
        unique_count = len(set(str(val).strip() for val in first_row if val is not None))
        if unique_count == len(first_row):
            header_score += 0.3
        
        # Check if subsequent rows are different type
        if len(df) > 1:
            second_row = df.iloc[1]
            type_diff = sum(
                1 for i, (h, s) in enumerate(zip(first_row, second_row))
                if type(h) != type(s)
            )
            if type_diff > len(first_row) * 0.5:
                header_score += 0.3
        
        threshold = self.config.get("header_detection_threshold", 0.8)
        detected = header_score >= threshold
        
        return {
            'detected': detected,
            'confidence': header_score,
            'row_index': 0 if detected else None,
            'headers': list(first_row) if detected else None
        }
    
    def _infer_data_types(self, df: pd.DataFrame) -> Dict[str, str]:
        """Infer data types for each column
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Dict mapping column names to inferred types
        """
        data_types = {}
        
        for col in df.columns:
            series = df[col].dropna()
            if series.empty:
                data_types[col] = 'empty'
                continue
            
            # Sample for performance
            sample_size = min(1000, len(series))
            sample = series.sample(n=sample_size) if len(series) > sample_size else series
            
            # Type inference logic
            inferred_type = self._infer_column_type(sample)
            data_types[col] = inferred_type
        
        return data_types
    
    def _infer_column_type(self, series: pd.Series) -> str:
        """Infer data type for a single column
        
        Args:
            series: pandas Series
            
        Returns:
            String representing inferred type
        """
        if series.empty:
            return 'empty'
        
        # Convert to string and check patterns
        str_values = series.astype(str).str.strip()
        non_empty = str_values[str_values != '']
        
        if non_empty.empty:
            return 'empty'
        
        # Numeric patterns
        numeric_pattern = re.compile(r'^-?\d*\.?\d+$')
        integer_pattern = re.compile(r'^-?\d+$')
        
        numeric_count = sum(1 for val in non_empty if numeric_pattern.match(val))
        integer_count = sum(1 for val in non_empty if integer_pattern.match(val))
        
        # Date patterns
        date_patterns = [
            r'\d{1,2}/\d{1,2}/\d{2,4}',  # MM/DD/YYYY or DD/MM/YYYY
            r'\d{4}-\d{1,2}-\d{1,2}',    # YYYY-MM-DD
            r'\d{1,2}-\d{1,2}-\d{2,4}',  # MM-DD-YYYY
        ]
        
        date_count = 0
        for pattern in date_patterns:
            date_count += sum(1 for val in non_empty if re.match(pattern, val))
        
        # Email pattern
        email_pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
        email_count = sum(1 for val in non_empty if email_pattern.match(val))
        
        # URL pattern
        url_pattern = re.compile(r'^https?://[^\s]+$')
        url_count = sum(1 for val in non_empty if url_pattern.match(val))
        
        # Phone pattern
        phone_pattern = re.compile(r'^[\+]?[\d\s\-\(\)]{10,}$')
        phone_count = sum(1 for val in non_empty if phone_pattern.match(val))
        
        total_count = len(non_empty)
        threshold = 0.8
        
        # Determine type based on highest match percentage
        if integer_count / total_count >= threshold:
            return 'integer'
        elif numeric_count / total_count >= threshold:
            return 'numeric'
        elif date_count / total_count >= threshold:
            return 'date'
        elif email_count / total_count >= threshold:
            return 'email'
        elif url_count / total_count >= threshold:
            return 'url'
        elif phone_count / total_count >= threshold:
            return 'phone'
        else:
            return 'text'
    
    def _calculate_null_percentages(self, df: pd.DataFrame) -> Dict[str, float]:
        """Calculate null percentage for each column
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Dict mapping column names to null percentages
        """
        null_percentages = {}
        
        for col in df.columns:
            null_count = df[col].isna().sum()
            total_count = len(df)
            null_percentage = (null_count / total_count) if total_count > 0 else 0.0
            null_percentages[col] = null_percentage
        
        return null_percentages
    
    def _calculate_column_statistics(self, df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        """Calculate basic statistics for each column
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Dict mapping column names to statistics
        """
        column_stats = {}
        
        for col in df.columns:
            series = df[col].dropna()
            
            stats = {
                'count': len(series),
                'unique_count': series.nunique(),
                'null_count': df[col].isna().sum(),
            }
            
            # Numeric statistics
            if pd.api.types.is_numeric_dtype(series):
                try:
                    stats.update({
                        'mean': float(series.mean()),
                        'median': float(series.median()),
                        'std': float(series.std()),
                        'min': float(series.min()),
                        'max': float(series.max()),
                    })
                except Exception:
                    pass
            else:
                # Text statistics
                try:
                    str_lengths = series.astype(str).str.len()
                    stats.update({
                        'avg_length': float(str_lengths.mean()),
                        'min_length': int(str_lengths.min()),
                        'max_length': int(str_lengths.max()),
                    })
                except Exception:
                    pass
            
            column_stats[col] = stats
        
        return column_stats
    
    def _count_duplicate_rows(self, df: pd.DataFrame) -> int:
        """Count duplicate rows in DataFrame
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Number of duplicate rows
        """
        try:
            return int(df.duplicated().sum())
        except Exception:
            return 0
    
    def _detect_outliers(self, df: pd.DataFrame) -> int:
        """Detect outliers in numeric columns using IQR method
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Total number of outliers detected
        """
        outlier_count = 0
        
        for col in df.columns:
            series = df[col].dropna()
            
            # Only check numeric columns
            if pd.api.types.is_numeric_dtype(series) and len(series) > 4:
                try:
                    Q1 = series.quantile(0.25)
                    Q3 = series.quantile(0.75)
                    IQR = Q3 - Q1
                    
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    
                    outliers = series[(series < lower_bound) | (series > upper_bound)]
                    outlier_count += len(outliers)
                    
                except Exception:
                    continue
        
        return outlier_count
    
    def _detect_patterns(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Detect common data patterns
        
        Args:
            df: pandas DataFrame
            
        Returns:
            List of detected patterns
        """
        patterns = []
        
        for col in df.columns:
            series = df[col].dropna()
            if series.empty:
                continue
            
            # Sample for performance
            sample_size = min(100, len(series))
            sample = series.sample(n=sample_size) if len(series) > sample_size else series
            
            # Check for sequential patterns
            if pd.api.types.is_numeric_dtype(sample):
                try:
                    sorted_vals = sorted(sample.head(20))
                    if len(sorted_vals) > 2:
                        diffs = [sorted_vals[i+1] - sorted_vals[i] for i in range(len(sorted_vals)-1)]
                        if len(set(diffs)) == 1 and diffs[0] > 0:
                            patterns.append({
                                'column': col,
                                'type': 'sequential',
                                'pattern': f'Incremental sequence with step {diffs[0]}'
                            })
                except Exception:
                    pass
            
            # Check for repeated values
            value_counts = sample.value_counts()
            if len(value_counts) > 0:
                most_common_ratio = value_counts.iloc[0] / len(sample)
                if most_common_ratio > 0.8:
                    patterns.append({
                        'column': col,
                        'type': 'repeated_value',
                        'pattern': f'Mostly same value: {value_counts.index[0]} ({most_common_ratio:.1%})'
                    })
        
        return patterns
    
    def _calculate_sheet_quality_score(self, df: pd.DataFrame) -> float:
        """Calculate overall quality score for a sheet
        
        Args:
            df: pandas DataFrame
            
        Returns:
            Quality score between 0.0 and 1.0
        """
        if df.empty:
            return 0.0
        
        factors = []
        
        # Completeness factor
        total_cells = len(df) * len(df.columns)
        non_null_cells = total_cells - df.isna().sum().sum()
        completeness = non_null_cells / total_cells if total_cells > 0 else 0.0
        factors.append(completeness * 0.4)
        
        # Consistency factor (low null variance across columns)
        null_percentages = [df[col].isna().mean() for col in df.columns]
        if null_percentages:
            null_variance = np.var(null_percentages)
            consistency = max(0.0, 1.0 - null_variance)
            factors.append(consistency * 0.3)
        
        # Uniqueness factor (avoid too many duplicates)
        duplicate_ratio = df.duplicated().mean()
        uniqueness = 1.0 - duplicate_ratio
        factors.append(uniqueness * 0.3)
        
        return sum(factors)
    
    def _calculate_overall_quality_score(self, sheet_profiles: Dict[str, Any]) -> float:
        """Calculate overall data quality score across all sheets
        
        Args:
            sheet_profiles: Dict of sheet profiling results
            
        Returns:
            Overall quality score between 0.0 and 1.0
        """
        successful_profiles = [
            profile for profile in sheet_profiles.values()
            if isinstance(profile, dict) and profile.get('status') == 'success'
        ]
        
        if not successful_profiles:
            return 0.0
        
        quality_scores = [
            profile.get('quality_score', 0.0) for profile in successful_profiles
        ]
        
        return sum(quality_scores) / len(quality_scores) if quality_scores else 0.0
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity based on file size and sheet count
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier
        """
        base_complexity = super().estimate_complexity(context)
        
        # Data profiling is more intensive than structure mapping
        return base_complexity * 2.0
    
    def _count_processed_items(self, data: DataProfileData) -> int:
        """Count items processed during analysis
        
        Args:
            data: DataProfileData result
            
        Returns:
            int: Number of processed items
        """
        total_rows = 0
        for sheet_profile in data.sheet_profiles.values():
            if isinstance(sheet_profile, dict) and 'row_count' in sheet_profile:
                total_rows += sheet_profile['row_count']
        
        return total_rows


# Legacy compatibility
def create_data_profiler(config: dict = None) -> DataProfiler:
    """Factory function for backward compatibility"""
    profiler = DataProfiler()
    if config:
        profiler.configure(config)
    return profiler
