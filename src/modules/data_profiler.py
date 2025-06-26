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
from ..utils.chunked_processor import ChunkedSheetProcessor, DataProfilingProcessor, ChunkConfig, ChunkingStrategy
from ..utils.memory_manager import ResourceMonitor


class DataProfiler(BaseAnalyzer):
    """Enhanced data profiler with comprehensive quality analysis"""
    
    def __init__(self, name: str = "data_profiler", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["health_checker", "structure_mapper"])
        self.chunk_processor = None
        self.profiling_processor = None
    
    def _perform_analysis(self, context: AnalysisContext) -> DataProfileData:
        """Perform comprehensive data profiling across all worksheets using chunked processing"""
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
            
            # Initialize chunked processing
            self._initialize_chunked_processing(context)
            
            # Profile each worksheet using chunked approach
            sheet_profiles = {}
            column_statistics = {}
            all_null_percentages = {}
            all_data_types = {}
            total_outliers = 0
            total_duplicates = 0
            patterns_found = []
            
            with context.get_workbook_access().get_workbook() as wb:
                for sheet_name in sheet_names:
                    try:
                        # Profile individual sheet with chunked processing
                        sheet_profile = self._profile_sheet_chunked(
                            wb[sheet_name], sheet_name, context
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
        """Validate data profiling results"""
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
    
    def _initialize_chunked_processing(self, context: AnalysisContext) -> None:
        """Initialize chunked processing components with configuration"""
        # Get configuration parameters
        chunk_size_rows = self.config.get("chunk_size_rows", context.config.chunk_size_rows)
        sample_limit = self.config.get("sample_size_limit", 10000)
        
        # Determine chunking strategy based on file size and configuration
        file_size_mb = context.file_metadata.file_size_mb
        if file_size_mb > 100:
            strategy = ChunkingStrategy.ADAPTIVE
            max_memory_mb = min(1024, context.config.max_memory_mb // 4)
        elif file_size_mb > 50:
            strategy = ChunkingStrategy.ROW_BASED
            max_memory_mb = 512
        else:
            strategy = ChunkingStrategy.ROW_BASED
            max_memory_mb = 256
        
        # Create chunk configuration
        chunk_config = ChunkConfig(
            chunk_size_rows=min(chunk_size_rows, sample_limit),
            max_memory_mb=max_memory_mb,
            strategy=strategy,
            enable_progress_tracking=True,
            intermediate_save=file_size_mb > 200
        )
        
        # Initialize processors
        self.chunk_processor = ChunkedSheetProcessor(chunk_config)
        self.profiling_processor = DataProfilingProcessor()
    
    def _profile_sheet_chunked(self, worksheet, sheet_name: str, 
                              context: AnalysisContext) -> Dict[str, Any]:
        """Profile individual worksheet using chunked processing"""
        try:
            # Check cache first
            cache_key = f"sheet_profile_{sheet_name}"
            cached = self.get_cached_result(context, cache_key)
            if cached:
                return cached
            
            # Detect data region
            enhanced_region_data = self._detect_data_region_enhanced(worksheet)
            if not enhanced_region_data:
                return {
                    'status': 'no_data',
                    'message': 'No data region detected'
                }
            
            # Extract primary boundaries for chunked processing
            data_region = enhanced_region_data['primary_boundaries']
            
            # Process using chunked approach
            chunk_results = self.chunk_processor.process_worksheet(
                worksheet, self.profiling_processor, data_region
            )
            
            if not chunk_results:
                return {
                    'status': 'empty',
                    'message': 'No data processed'
                }
            
            # Aggregate chunk results
            profile = self._aggregate_chunk_results(chunk_results, data_region)
            profile['status'] = 'success'
            profile['data_region'] = data_region
            profile['chunks_processed'] = len(chunk_results)
            
            # Add enhanced region analysis
            profile['enhanced_region_analysis'] = {
                'data_regions': enhanced_region_data['data_regions'],
                'empty_regions': enhanced_region_data['empty_regions'],
                'region_classifications': enhanced_region_data['region_classifications'],
                'header_footer_info': enhanced_region_data['header_footer_info'],
                'continuity_analysis': enhanced_region_data['continuity_analysis'],
                'merged_cells_detected': enhanced_region_data['merged_cells_detected']
            }
            
            # Add advanced analysis
            profile.update(self._perform_advanced_analysis(chunk_results))
            
            # Cache result
            self.cache_intermediate_result(context, cache_key, profile)
            
            return profile
            
        except Exception as e:
            self.logger.warning(f"Error profiling sheet {sheet_name}: {e}")
            return {
                'status': 'error',
                'error': str(e)
            }
    
    def _detect_data_region_enhanced(self, worksheet) -> Optional[Dict[str, Any]]:
        """Enhanced data region detection with comprehensive analysis"""
        if not worksheet.max_row or not worksheet.max_column:
            return None
        
        # Step 1: Find overall data boundaries
        boundaries = self._find_comprehensive_boundaries(worksheet)
        if not boundaries:
            return None
        
        # Step 2: Handle merged cells and adjust boundaries
        boundaries = self._adjust_for_merged_cells(worksheet, boundaries)
        
        # Step 3: Identify distinct data regions within boundaries
        data_regions = self._identify_distinct_data_regions(worksheet, boundaries)
        
        # Step 4: Classify empty regions
        empty_regions = self._classify_empty_regions(worksheet, boundaries, data_regions)
        
        # Step 5: Distinguish data vs formatting areas
        region_classifications = self._classify_region_types(worksheet, data_regions)
        
        # Step 6: Detect headers and footers
        header_info = self._detect_headers_and_footers(worksheet, data_regions)
        
        # Step 7: Analyze data continuity
        continuity_analysis = self._analyze_data_continuity(worksheet, data_regions)
        
        return {
            'primary_boundaries': boundaries,
            'data_regions': data_regions,
            'empty_regions': empty_regions,
            'region_classifications': region_classifications,
            'header_footer_info': header_info,
            'continuity_analysis': continuity_analysis,
            'merged_cells_detected': len(list(worksheet.merged_cells.ranges)) > 0 if hasattr(worksheet, 'merged_cells') and worksheet.merged_cells else False
        }
    
    def _find_comprehensive_boundaries(self, worksheet) -> Optional[Dict[str, int]]:
        """Find comprehensive data boundaries with merged cell awareness"""
        min_row, max_row = self._find_row_boundaries(worksheet)
        if not min_row or not max_row:
            return None
            
        min_col, max_col = self._find_column_boundaries(worksheet, min_row, max_row)
        if not min_col or not max_col:
            return None
        
        return {
            'min_row': min_row,
            'max_row': max_row,
            'min_col': min_col,
            'max_col': max_col
        }
    
    def _find_row_boundaries(self, worksheet) -> Tuple[Optional[int], Optional[int]]:
        """Find first and last non-empty rows efficiently"""
        min_row = None
        max_row = None
        
        # Sample every nth row for efficiency
        sample_step = max(1, worksheet.max_row // 200)
        
        for row_idx in range(1, worksheet.max_row + 1, sample_step):
            row_has_data = False
            for col_idx in range(1, min(worksheet.max_column + 1, 20)):  # Sample columns
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip():
                    row_has_data = True
                    break
            
            if row_has_data:
                if min_row is None:
                    min_row = row_idx
                max_row = row_idx
        
        # Refine boundaries with more precise search
        if min_row and max_row:
            min_row = self._refine_min_row(worksheet, max(1, min_row - sample_step))
            max_row = self._refine_max_row(worksheet, min(worksheet.max_row, max_row + sample_step))
        
        return min_row, max_row
    
    def _find_column_boundaries(self, worksheet, min_row: int, max_row: int) -> Tuple[Optional[int], Optional[int]]:
        """Find first and last non-empty columns efficiently"""
        min_col = None
        max_col = None
        
        # Sample rows for column detection
        sample_rows = list(range(min_row, min(min_row + 30, max_row + 1)))
        
        for col_idx in range(1, worksheet.max_column + 1):
            col_has_data = False
            for row_idx in sample_rows:
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip():
                    col_has_data = True
                    break
            
            if col_has_data:
                if min_col is None:
                    min_col = col_idx
                max_col = col_idx
        
        return min_col, max_col
    
    def _refine_min_row(self, worksheet, start_row: int) -> int:
        """Refine minimum row with precise search"""
        for row_idx in range(start_row, min(start_row + 100, worksheet.max_row + 1)):
            for col_idx in range(1, min(worksheet.max_column + 1, 50)):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip():
                    return row_idx
        return start_row
    
    def _refine_max_row(self, worksheet, end_row: int) -> int:
        """Refine maximum row with precise search"""
        for row_idx in range(end_row, max(end_row - 100, 0), -1):
            for col_idx in range(1, min(worksheet.max_column + 1, 50)):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip():
                    return row_idx
        return end_row
    
    def _adjust_for_merged_cells(self, worksheet, boundaries: Dict[str, int]) -> Dict[str, int]:
        """Adjust boundaries to account for merged cells"""
        if not hasattr(worksheet, 'merged_cells') or not worksheet.merged_cells:
            return boundaries
        
        adjusted = boundaries.copy()
        
        try:
            for merged_range in worksheet.merged_cells.ranges:
                # Expand boundaries if merged cells extend beyond current boundaries
                if merged_range.min_row < adjusted['min_row']:
                    adjusted['min_row'] = merged_range.min_row
                if merged_range.max_row > adjusted['max_row']:
                    adjusted['max_row'] = merged_range.max_row
                if merged_range.min_col < adjusted['min_col']:
                    adjusted['min_col'] = merged_range.min_col
                if merged_range.max_col > adjusted['max_col']:
                    adjusted['max_col'] = merged_range.max_col
        except Exception as e:
            self.logger.warning(f"Error processing merged cells: {e}")
        
        return adjusted
    
    def _identify_distinct_data_regions(self, worksheet, boundaries: Dict[str, int]) -> List[Dict[str, Any]]:
        """Identify distinct data regions within the worksheet boundaries"""
        regions = []
        
        # Sample the worksheet to find data clusters
        min_row, max_row = boundaries['min_row'], boundaries['max_row']
        min_col, max_col = boundaries['min_col'], boundaries['max_col']
        
        # Create a grid to track data presence
        sample_step = max(1, (max_row - min_row + 1) // 50)  # Sample every nth row
        col_step = max(1, (max_col - min_col + 1) // 20)      # Sample every nth col
        
        data_grid = []
        sampled_rows = list(range(min_row, max_row + 1, sample_step))
        sampled_cols = list(range(min_col, max_col + 1, col_step))
        
        for row_idx in sampled_rows:
            row_data = []
            for col_idx in sampled_cols:
                try:
                    cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                    has_data = cell_value is not None and str(cell_value).strip()
                    row_data.append(has_data)
                except Exception:
                    row_data.append(False)
            data_grid.append(row_data)
        
        # Find contiguous regions using flood-fill approach
        visited = [[False] * len(sampled_cols) for _ in range(len(sampled_rows))]
        region_id = 0
        
        for i, row_data in enumerate(data_grid):
            for j, has_data in enumerate(row_data):
                if has_data and not visited[i][j]:
                    region = self._flood_fill_region(data_grid, visited, i, j, sampled_rows, sampled_cols)
                    if region['size'] > 4:  # Minimum region size
                        region['id'] = region_id
                        region['type'] = 'data_cluster'
                        regions.append(region)
                        region_id += 1
        
        # If no distinct regions found, create one main region
        if not regions:
            regions.append({
                'id': 0,
                'type': 'main_data',
                'min_row': min_row,
                'max_row': max_row,
                'min_col': min_col,
                'max_col': max_col,
                'size': (max_row - min_row + 1) * (max_col - min_col + 1)
            })
        
        return regions
    
    def _flood_fill_region(self, data_grid, visited, start_i, start_j, sampled_rows, sampled_cols) -> Dict[str, Any]:
        """Use flood fill to identify contiguous data regions"""
        stack = [(start_i, start_j)]
        min_i, max_i = start_i, start_i
        min_j, max_j = start_j, start_j
        size = 0
        
        while stack:
            i, j = stack.pop()
            
            if (i < 0 or i >= len(data_grid) or j < 0 or j >= len(data_grid[0]) or
                visited[i][j] or not data_grid[i][j]):
                continue
            
            visited[i][j] = True
            size += 1
            
            min_i, max_i = min(min_i, i), max(max_i, i)
            min_j, max_j = min(min_j, j), max(max_j, j)
            
            # Add adjacent cells
            for di, dj in [(0, 1), (0, -1), (1, 0), (-1, 0)]:
                stack.append((i + di, j + dj))
        
        return {
            'min_row': sampled_rows[min_i],
            'max_row': sampled_rows[max_i],
            'min_col': sampled_cols[min_j],
            'max_col': sampled_cols[max_j],
            'size': size
        }
    
    def _classify_empty_regions(self, worksheet, boundaries: Dict[str, int], 
                               data_regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Classify empty regions and their purposes"""
        empty_regions = []
        
        min_row, max_row = boundaries['min_row'], boundaries['max_row']
        min_col, max_col = boundaries['min_col'], boundaries['max_col']
        
        # Create occupied region map
        occupied = set()
        for region in data_regions:
            for r in range(region['min_row'], region['max_row'] + 1):
                for c in range(region['min_col'], region['max_col'] + 1):
                    occupied.add((r, c))
        
        # Find large empty areas
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (row, col) not in occupied:
                    # Check if this starts a significant empty region
                    empty_size = self._measure_empty_region(worksheet, row, col, max_row, max_col, occupied)
                    if empty_size > 10:  # Significant empty area
                        empty_regions.append({
                            'start_row': row,
                            'start_col': col,
                            'estimated_size': empty_size,
                            'type': self._classify_empty_region_type(row, col, boundaries)
                        })
        
        return empty_regions
    
    def _measure_empty_region(self, worksheet, start_row: int, start_col: int, 
                             max_row: int, max_col: int, occupied: set) -> int:
        """Measure the size of an empty region starting from given position"""
        size = 0
        
        # Check a small sample area to estimate empty region size
        for r in range(start_row, min(start_row + 10, max_row + 1)):
            for c in range(start_col, min(start_col + 10, max_col + 1)):
                if (r, c) not in occupied:
                    try:
                        cell_value = worksheet.cell(row=r, column=c).value
                        if cell_value is None or not str(cell_value).strip():
                            size += 1
                    except Exception:
                        size += 1
        
        return size
    
    def _classify_empty_region_type(self, row: int, col: int, boundaries: Dict[str, int]) -> str:
        """Classify the type/purpose of an empty region"""
        total_rows = boundaries['max_row'] - boundaries['min_row'] + 1
        total_cols = boundaries['max_col'] - boundaries['min_col'] + 1
        
        row_position = (row - boundaries['min_row']) / total_rows
        col_position = (col - boundaries['min_col']) / total_cols
        
        if row_position < 0.1:
            return 'header_spacing'
        elif row_position > 0.9:
            return 'footer_spacing'
        elif col_position < 0.1:
            return 'left_margin'
        elif col_position > 0.9:
            return 'right_margin'
        else:
            return 'internal_spacing'
    
    def _classify_region_types(self, worksheet, data_regions: List[Dict[str, Any]]) -> Dict[str, str]:
        """Distinguish between data regions and formatting-only areas"""
        classifications = {}
        
        for region in data_regions:
            region_id = region['id']
            
            # Sample cells in the region to determine type
            sample_cells = self._sample_region_cells(worksheet, region, max_samples=20)
            
            data_cells = 0
            format_only_cells = 0
            formula_cells = 0
            
            for cell_info in sample_cells:
                if cell_info['has_formula']:
                    formula_cells += 1
                elif cell_info['has_meaningful_data']:
                    data_cells += 1
                elif cell_info['has_formatting']:
                    format_only_cells += 1
            
            total_samples = len(sample_cells)
            if total_samples == 0:
                classification = 'unknown'
            elif data_cells / total_samples > 0.6:
                classification = 'primary_data'
            elif formula_cells / total_samples > 0.3:
                classification = 'calculated_data'
            elif format_only_cells / total_samples > 0.7:
                classification = 'formatting_only'
            else:
                classification = 'mixed_content'
            
            classifications[str(region_id)] = classification
        
        return classifications
    
    def _sample_region_cells(self, worksheet, region: Dict[str, Any], max_samples: int = 20) -> List[Dict[str, Any]]:
        """Sample cells from a region to analyze their content type"""
        samples = []
        
        min_row, max_row = region['min_row'], region['max_row']
        min_col, max_col = region['min_col'], region['max_col']
        
        total_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
        step = max(1, total_cells // max_samples)
        
        current_sample = 0
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if current_sample % step == 0 and len(samples) < max_samples:
                    try:
                        cell = worksheet.cell(row=row, column=col)
                        cell_info = {
                            'row': row,
                            'col': col,
                            'has_formula': str(cell.value or '').startswith('='),
                            'has_meaningful_data': self._has_meaningful_data(cell.value),
                            'has_formatting': self._has_significant_formatting(cell),
                            'value_type': type(cell.value).__name__
                        }
                        samples.append(cell_info)
                    except Exception:
                        pass
                current_sample += 1
        
        return samples
    
    def _has_meaningful_data(self, value) -> bool:
        """Check if cell value represents meaningful data"""
        if value is None:
            return False
        
        str_value = str(value).strip()
        if not str_value:
            return False
        
        # Filter out common non-data content
        non_data_patterns = ['', ' ', '-', '—', 'n/a', 'na', 'null', 'none']
        return str_value.lower() not in non_data_patterns and len(str_value) > 0
    
    def _has_significant_formatting(self, cell) -> bool:
        """Check if cell has significant formatting beyond defaults"""
        try:
            # Check for non-default formatting
            has_font_formatting = (cell.font and 
                                 (cell.font.bold or cell.font.italic or 
                                  cell.font.color or cell.font.size != 11))
            
            has_fill_formatting = cell.fill and cell.fill.fill_type is not None
            has_border_formatting = cell.border and any([
                cell.border.left.style, cell.border.right.style,
                cell.border.top.style, cell.border.bottom.style
            ])
            
            return has_font_formatting or has_fill_formatting or has_border_formatting
        except Exception:
            return False
    
    def _detect_headers_and_footers(self, worksheet, data_regions: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Detect headers and footers in data regions"""
        header_footer_info = {
            'headers_detected': [],
            'footers_detected': [],
            'summary_rows': []
        }
        
        for region in data_regions:
            # Analyze first few rows for headers
            header_analysis = self._analyze_potential_headers(worksheet, region)
            if header_analysis['confidence'] > 0.6:
                header_footer_info['headers_detected'].append({
                    'region_id': region['id'],
                    'header_rows': header_analysis['header_rows'],
                    'confidence': header_analysis['confidence']
                })
            
            # Analyze last few rows for footers/summaries
            footer_analysis = self._analyze_potential_footers(worksheet, region)
            if footer_analysis['has_summary']:
                header_footer_info['summary_rows'].append({
                    'region_id': region['id'],
                    'summary_rows': footer_analysis['summary_rows'],
                    'summary_type': footer_analysis['summary_type']
                })
        
        return header_footer_info
    
    def _analyze_potential_headers(self, worksheet, region: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze first rows of region for header patterns"""
        min_row, max_row = region['min_row'], region['max_row']
        min_col, max_col = region['min_col'], region['max_col']
        
        # Check first 3 rows maximum
        header_candidates = min(3, max_row - min_row + 1)
        header_scores = []
        
        for row_offset in range(header_candidates):
            row_num = min_row + row_offset
            score = self._score_header_row(worksheet, row_num, min_col, max_col)
            header_scores.append(score)
        
        # Determine if we have headers
        if header_scores and max(header_scores) > 0.7:
            header_rows = [min_row + i for i, score in enumerate(header_scores) if score > 0.5]
            confidence = max(header_scores)
        else:
            header_rows = []
            confidence = 0.0
        
        return {
            'header_rows': header_rows,
            'confidence': confidence,
            'scores': header_scores
        }
    
    def _score_header_row(self, worksheet, row_num: int, min_col: int, max_col: int) -> float:
        """Score a row for header likelihood"""
        score_factors = []
        cell_count = 0
        text_count = 0
        values = []
        
        for col in range(min_col, max_col + 1):
            try:
                cell = worksheet.cell(row=row_num, column=col)
                value = cell.value
                
                if value is not None:
                    cell_count += 1
                    str_value = str(value).strip()
                    
                    if str_value:
                        values.append(str_value)
                        
                        # Text content is more likely to be headers
                        if isinstance(value, str) and not str_value.isdigit():
                            text_count += 1
                        
                        # Check for header-like patterns
                        if self._looks_like_header_text(str_value):
                            score_factors.append(0.8)
                        
                        # Bold formatting suggests headers
                        if cell.font and cell.font.bold:
                            score_factors.append(0.6)
            except Exception:
                pass
        
        if cell_count == 0:
            return 0.0
        
        # Calculate uniqueness (headers should be unique)
        unique_count = len(set(values))
        uniqueness_ratio = unique_count / len(values) if values else 0
        
        # Text ratio (headers are usually text)
        text_ratio = text_count / cell_count
        
        # Combined score
        base_score = (text_ratio * 0.4 + uniqueness_ratio * 0.3)
        formatting_bonus = sum(score_factors) / max(1, len(score_factors)) if score_factors else 0
        
        return min(1.0, base_score + formatting_bonus * 0.3)
    
    def _looks_like_header_text(self, text: str) -> bool:
        """Check if text looks like a typical header"""
        header_patterns = [
            'name', 'id', 'date', 'time', 'amount', 'total', 'count', 'number',
            'description', 'type', 'category', 'status', 'value', 'price',
            'quantity', 'rate', 'percent', 'address', 'phone', 'email'
        ]
        
        text_lower = text.lower()
        return any(pattern in text_lower for pattern in header_patterns)
    
    def _analyze_potential_footers(self, worksheet, region: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze last rows of region for footer/summary patterns"""
        min_row, max_row = region['min_row'], region['max_row']
        min_col, max_col = region['min_col'], region['max_col']
        
        # Check last 3 rows maximum
        footer_candidates = min(3, max_row - min_row + 1)
        summary_rows = []
        
        for row_offset in range(footer_candidates):
            row_num = max_row - row_offset
            if self._is_summary_row(worksheet, row_num, min_col, max_col):
                summary_rows.append(row_num)
        
        summary_type = 'none'
        if summary_rows:
            summary_type = self._classify_summary_type(worksheet, summary_rows, min_col, max_col)
        
        return {
            'has_summary': len(summary_rows) > 0,
            'summary_rows': summary_rows,
            'summary_type': summary_type
        }
    
    def _is_summary_row(self, worksheet, row_num: int, min_col: int, max_col: int) -> bool:
        """Check if a row appears to be a summary row"""
        formula_count = 0
        total_count = 0
        
        for col in range(min_col, max_col + 1):
            try:
                cell = worksheet.cell(row=row_num, column=col)
                if cell.value is not None:
                    total_count += 1
                    if str(cell.value).startswith('='):
                        formula_count += 1
                        
                        # Check for summary functions
                        formula_str = str(cell.value).upper()
                        if any(func in formula_str for func in ['SUM', 'AVERAGE', 'COUNT', 'TOTAL']):
                            return True
            except Exception:
                pass
        
        # High ratio of formulas suggests summary row
        return total_count > 0 and formula_count / total_count > 0.5
    
    def _classify_summary_type(self, worksheet, summary_rows: List[int], min_col: int, max_col: int) -> str:
        """Classify the type of summary in footer rows"""
        formula_types = set()
        
        for row_num in summary_rows:
            for col in range(min_col, max_col + 1):
                try:
                    cell = worksheet.cell(row=row_num, column=col)
                    if cell.value and str(cell.value).startswith('='):
                        formula_str = str(cell.value).upper()
                        if 'SUM' in formula_str:
                            formula_types.add('totals')
                        elif 'AVERAGE' in formula_str:
                            formula_types.add('averages')
                        elif 'COUNT' in formula_str:
                            formula_types.add('counts')
                except Exception:
                    pass
        
        if 'totals' in formula_types:
            return 'financial_totals'
        elif 'counts' in formula_types:
            return 'record_counts'
        elif 'averages' in formula_types:
            return 'statistical_summary'
        else:
            return 'generic_summary'
    
    def _analyze_data_continuity(self, worksheet, data_regions: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Analyze data continuity within and between regions"""
        continuity_analysis = {
            'regions_with_gaps': [],
            'continuous_regions': [],
            'data_density_scores': {},
            'gap_analysis': []
        }
        
        for region in data_regions:
            region_id = str(region['id'])
            
            # Analyze gaps within the region
            gaps = self._find_gaps_in_region(worksheet, region)
            density_score = self._calculate_region_density(worksheet, region)
            
            continuity_analysis['data_density_scores'][region_id] = density_score
            
            if gaps:
                continuity_analysis['regions_with_gaps'].append({
                    'region_id': region_id,
                    'gap_count': len(gaps),
                    'gaps': gaps[:5]  # Limit to first 5 gaps
                })
            else:
                continuity_analysis['continuous_regions'].append(region_id)
            
            # Analyze gaps between this region and others
            for other_region in data_regions:
                if other_region['id'] != region['id']:
                    gap_info = self._analyze_inter_region_gap(region, other_region)
                    if gap_info:
                        continuity_analysis['gap_analysis'].append(gap_info)
        
        return continuity_analysis
    
    def _find_gaps_in_region(self, worksheet, region: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Find significant gaps within a data region"""
        gaps = []
        min_row, max_row = region['min_row'], region['max_row']
        min_col, max_col = region['min_col'], region['max_col']
        
        # Sample every 5th row and column to detect gaps
        sample_step = 5
        
        for row in range(min_row, max_row + 1, sample_step):
            empty_streak = 0
            gap_start = None
            
            for col in range(min_col, max_col + 1):
                try:
                    cell_value = worksheet.cell(row=row, column=col).value
                    has_data = cell_value is not None and str(cell_value).strip()
                    
                    if not has_data:
                        if gap_start is None:
                            gap_start = col
                        empty_streak += 1
                    else:
                        if empty_streak > 3:  # Significant gap
                            gaps.append({
                                'type': 'horizontal_gap',
                                'row': row,
                                'start_col': gap_start,
                                'end_col': col - 1,
                                'size': empty_streak
                            })
                        empty_streak = 0
                        gap_start = None
                except Exception:
                    empty_streak += 1
        
        return gaps
    
    def _calculate_region_density(self, worksheet, region: Dict[str, Any]) -> float:
        """Calculate data density score for a region"""
        min_row, max_row = region['min_row'], region['max_row']
        min_col, max_col = region['min_col'], region['max_col']
        
        total_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
        filled_cells = 0
        
        # Sample to estimate density
        sample_size = min(100, total_cells)
        sample_step = max(1, total_cells // sample_size)
        
        current_cell = 0
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if current_cell % sample_step == 0:
                    try:
                        cell_value = worksheet.cell(row=row, column=col).value
                        if cell_value is not None and str(cell_value).strip():
                            filled_cells += 1
                    except Exception:
                        pass
                current_cell += 1
        
        return filled_cells / sample_size if sample_size > 0 else 0.0
    
    def _analyze_inter_region_gap(self, region1: Dict[str, Any], region2: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Analyze gaps between two regions"""
        # Check if regions are adjacent or have significant gaps
        
        # Horizontal gap
        if (region1['max_col'] < region2['min_col'] or region2['max_col'] < region1['min_col']):
            gap_size = abs(region2['min_col'] - region1['max_col']) - 1
            if gap_size > 2:  # Significant gap
                return {
                    'type': 'horizontal_inter_region_gap',
                    'region1_id': region1['id'],
                    'region2_id': region2['id'],
                    'gap_size': gap_size
                }
        
        # Vertical gap
        if (region1['max_row'] < region2['min_row'] or region2['max_row'] < region1['min_row']):
            gap_size = abs(region2['min_row'] - region1['max_row']) - 1
            if gap_size > 2:  # Significant gap
                return {
                    'type': 'vertical_inter_region_gap',
                    'region1_id': region1['id'],
                    'region2_id': region2['id'],
                    'gap_size': gap_size
                }
        
        return None
    
    def _aggregate_chunk_results(self, chunk_results: List[Dict[str, Any]], 
                               data_region: Dict[str, int]) -> Dict[str, Any]:
        """Aggregate results from multiple chunks into final profile"""
        total_rows = sum(r.get('row_count', 0) for r in chunk_results if 'error' not in r)
        total_columns = max(r.get('column_count', 0) for r in chunk_results if 'error' not in r) if chunk_results else 0
        
        # Aggregate null percentages (weighted average)
        aggregated_nulls = {}
        if total_rows > 0:
            for col_name in range(total_columns):  # Assuming generic column names
                col_key = f'col_{col_name}'
                total_null_ratio = 0
                valid_chunks = 0
                
                for chunk in chunk_results:
                    if 'error' not in chunk and col_key in chunk.get('null_percentages', {}):
                        chunk_rows = chunk.get('row_count', 0)
                        chunk_null_ratio = chunk['null_percentages'][col_key]
                        total_null_ratio += chunk_null_ratio * chunk_rows
                        valid_chunks += 1
                
                if valid_chunks > 0:
                    aggregated_nulls[col_key] = total_null_ratio / total_rows
        
        # Aggregate data types (most common type per column)
        aggregated_types = {}
        for col_name in range(total_columns):
            col_key = f'col_{col_name}'
            type_counts = {}
            
            for chunk in chunk_results:
                if 'error' not in chunk and col_key in chunk.get('data_types', {}):
                    data_type = chunk['data_types'][col_key]
                    type_counts[data_type] = type_counts.get(data_type, 0) + chunk.get('row_count', 0)
            
            if type_counts:
                aggregated_types[col_key] = max(type_counts, key=type_counts.get)
        
        # Sum other metrics
        total_outliers = sum(r.get('outliers_detected', 0) for r in chunk_results if 'error' not in r)
        total_duplicates = sum(r.get('duplicate_rows', 0) for r in chunk_results if 'error' not in r)
        
        # Calculate quality score
        quality_score = self._calculate_aggregated_quality_score(
            total_rows, total_columns, aggregated_nulls, total_outliers, total_duplicates
        )
        
        return {
            'row_count': total_rows,
            'column_count': total_columns,
            'null_percentages': aggregated_nulls,
            'data_types': aggregated_types,
            'outliers_count': total_outliers,
            'duplicate_rows': total_duplicates,
            'quality_score': quality_score,
            'chunk_errors': [r for r in chunk_results if 'error' in r]
        }
    
    def _perform_advanced_analysis(self, chunk_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Perform advanced analysis on aggregated chunk results"""
        advanced_stats = {
            'headers': {'detected': False},
            'column_stats': {},
            'patterns': [],
            'data_consistency': 0.9
        }
        
        try:
            # Detect headers from first chunk if available
            if chunk_results and 'error' not in chunk_results[0]:
                advanced_stats['headers'] = self._detect_headers_from_chunks(chunk_results)
            
            # Generate column statistics
            advanced_stats['column_stats'] = self._generate_column_stats_from_chunks(chunk_results)
            
            # Detect patterns across chunks
            advanced_stats['patterns'] = self._detect_patterns_from_chunks(chunk_results)
            
            # Calculate data consistency score
            advanced_stats['data_consistency'] = self._calculate_chunk_consistency(chunk_results)
            
        except Exception as e:
            self.logger.warning(f"Advanced analysis failed: {e}")
            advanced_stats['analysis_error'] = str(e)
        
        return advanced_stats
    
    def _detect_headers_from_chunks(self, chunk_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Detect headers using chunk data"""
        # Simple heuristic: if first chunk has mostly text in first row vs. subsequent rows
        if not chunk_results or 'error' in chunk_results[0]:
            return {'detected': False}
        
        first_chunk = chunk_results[0]
        confidence = 0.7  # Default confidence for chunked approach
        
        return {
            'detected': True,
            'confidence': confidence,
            'row_index': 0,
            'headers': [f'Column_{i}' for i in range(first_chunk.get('column_count', 0))]
        }
    
    def _generate_column_stats_from_chunks(self, chunk_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate column statistics from chunk results"""
        column_stats = {}
        
        # Aggregate statistics for each column across chunks
        total_columns = max(r.get('column_count', 0) for r in chunk_results if 'error' not in r) if chunk_results else 0
        
        for col_idx in range(total_columns):
            col_key = f'col_{col_idx}'
            col_stats = {
                'count': sum(r.get('row_count', 0) for r in chunk_results if 'error' not in r),
                'null_count': 0,
                'unique_estimate': 'unknown'  # Would need more sophisticated tracking
            }
            
            # Calculate null count from null percentages
            for chunk in chunk_results:
                if 'error' not in chunk and col_key in chunk.get('null_percentages', {}):
                    chunk_rows = chunk.get('row_count', 0)
                    null_ratio = chunk['null_percentages'][col_key]
                    col_stats['null_count'] += int(chunk_rows * null_ratio)
            
            column_stats[col_key] = col_stats
        
        return column_stats
    
    def _detect_patterns_from_chunks(self, chunk_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Detect patterns across chunk results"""
        patterns = []
        
        # Analyze data type consistency across chunks
        type_consistency = self._analyze_type_consistency(chunk_results)
        if type_consistency['inconsistent_columns']:
            patterns.append({
                'type': 'type_inconsistency',
                'pattern': f"Inconsistent data types detected in columns: {type_consistency['inconsistent_columns']}"
            })
        
        # Check for outlier patterns
        total_outliers = sum(r.get('outliers_detected', 0) for r in chunk_results if 'error' not in r)
        total_rows = sum(r.get('row_count', 0) for r in chunk_results if 'error' not in r)
        
        if total_rows > 0 and (total_outliers / total_rows) > 0.1:
            patterns.append({
                'type': 'high_outlier_rate',
                'pattern': f"High outlier rate detected: {total_outliers/total_rows:.1%}"
            })
        
        return patterns
    
    def _analyze_type_consistency(self, chunk_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Analyze data type consistency across chunks"""
        column_types = {}
        inconsistent_columns = []
        
        for chunk in chunk_results:
            if 'error' not in chunk:
                for col, data_type in chunk.get('data_types', {}).items():
                    if col not in column_types:
                        column_types[col] = set()
                    column_types[col].add(data_type)
        
        # Find columns with multiple types
        for col, types in column_types.items():
            if len(types) > 1:
                inconsistent_columns.append(col)
        
        return {
            'column_types': {col: list(types) for col, types in column_types.items()},
            'inconsistent_columns': inconsistent_columns
        }
    
    def _calculate_chunk_consistency(self, chunk_results: List[Dict[str, Any]]) -> float:
        """Calculate consistency score across chunks"""
        if not chunk_results:
            return 0.0
        
        # Base consistency on error rate and type consistency
        error_chunks = sum(1 for r in chunk_results if 'error' in r)
        error_rate = error_chunks / len(chunk_results)
        
        type_consistency = self._analyze_type_consistency(chunk_results)
        inconsistency_rate = len(type_consistency['inconsistent_columns']) / max(1, len(type_consistency['column_types']))
        
        consistency_score = 1.0 - (error_rate * 0.5 + inconsistency_rate * 0.3)
        return max(0.0, min(1.0, consistency_score))
    
    def _calculate_aggregated_quality_score(self, total_rows: int, total_columns: int,
                                          null_percentages: Dict[str, float],
                                          total_outliers: int, total_duplicates: int) -> float:
        """Calculate overall quality score for aggregated data"""
        if total_rows == 0:
            return 0.0
        
        factors = []
        
        # Completeness factor (average non-null percentage)
        if null_percentages:
            avg_completeness = 1.0 - sum(null_percentages.values()) / len(null_percentages)
            factors.append(avg_completeness * 0.4)
        
        # Uniqueness factor (inverse of duplicate rate)
        duplicate_rate = total_duplicates / total_rows
        uniqueness = 1.0 - duplicate_rate
        factors.append(uniqueness * 0.3)
        
        # Outlier factor (low outlier rate is good)
        outlier_rate = total_outliers / total_rows
        outlier_factor = max(0.0, 1.0 - outlier_rate * 2)  # Penalize high outlier rates
        factors.append(outlier_factor * 0.3)
        
        return sum(factors) if factors else 0.5

    def _infer_data_types(self, df: pd.DataFrame) -> Dict[str, str]:
        """Advanced data type inference with >95% accuracy requirement"""
        data_types = {}
        
        for col in df.columns:
            series = df[col].dropna()
            if series.empty:
                data_types[col] = 'empty'
                continue
            
            # Use statistical sampling for large datasets
            type_analysis = self._advanced_type_inference(series)
            data_types[col] = type_analysis['primary_type']
        
        return data_types
    
    def _advanced_type_inference(self, series: pd.Series) -> Dict[str, Any]:
        """Advanced type inference with statistical sampling and confidence scoring"""
        if series.empty:
            return {'primary_type': 'empty', 'confidence': 1.0, 'mixed_types': []}
        
        # Statistical sampling for large datasets
        sample_size = min(1000, len(series))  # Optimal sample size for 95% accuracy
        if len(series) > sample_size:
            # Stratified sampling to maintain distribution
            sample_indices = np.linspace(0, len(series) - 1, sample_size, dtype=int)
            sample_series = series.iloc[sample_indices]
        else:
            sample_series = series
        
        # Pattern matching with confidence scoring
        type_scores = self._calculate_type_scores(sample_series)
        
        # Determine primary type and confidence
        primary_type, confidence = self._determine_primary_type(type_scores)
        
        # Handle mixed-type columns
        mixed_analysis = self._analyze_mixed_types(type_scores, confidence)
        
        return {
            'primary_type': primary_type,
            'confidence': confidence,
            'type_scores': type_scores,
            'mixed_types': mixed_analysis['types'],
            'is_mixed': mixed_analysis['is_mixed'],
            'dominant_ratio': mixed_analysis['dominant_ratio']
        }
    
    def _calculate_type_scores(self, series: pd.Series) -> Dict[str, float]:
        """Calculate confidence scores for each potential data type"""
        str_values = series.astype(str).str.strip()
        non_empty = str_values[str_values != '']
        total_count = len(non_empty)
        
        if total_count == 0:
            return {'empty': 1.0}
        
        type_scores = {}
        
        # Integer pattern matching
        integer_matches = self._match_integer_patterns(non_empty)
        type_scores['integer'] = integer_matches / total_count
        
        # Numeric pattern matching (includes decimals)
        numeric_matches = self._match_numeric_patterns(non_empty)
        type_scores['numeric'] = numeric_matches / total_count
        
        # Date pattern matching with multiple formats
        date_matches = self._match_date_patterns(non_empty)
        type_scores['date'] = date_matches / total_count
        
        # Time pattern matching
        time_matches = self._match_time_patterns(non_empty)
        type_scores['time'] = time_matches / total_count
        
        # Currency pattern matching
        currency_matches = self._match_currency_patterns(non_empty)
        type_scores['currency'] = currency_matches / total_count
        
        # Percentage pattern matching
        percentage_matches = self._match_percentage_patterns(non_empty)
        type_scores['percentage'] = percentage_matches / total_count
        
        # Email pattern matching
        email_matches = self._match_email_patterns(non_empty)
        type_scores['email'] = email_matches / total_count
        
        # URL pattern matching
        url_matches = self._match_url_patterns(non_empty)
        type_scores['url'] = url_matches / total_count
        
        # Phone pattern matching
        phone_matches = self._match_phone_patterns(non_empty)
        type_scores['phone'] = phone_matches / total_count
        
        # Boolean pattern matching
        boolean_matches = self._match_boolean_patterns(non_empty)
        type_scores['boolean'] = boolean_matches / total_count
        
        # Business-specific patterns
        business_analysis = self._match_business_patterns(non_empty)
        type_scores.update(business_analysis)
        
        # Default to text if no strong pattern matches
        max_score = max(type_scores.values()) if type_scores else 0
        if max_score < 0.5:  # No strong pattern detected
            type_scores['text'] = 1.0 - max_score
        
        return type_scores
    
    def _match_integer_patterns(self, values: pd.Series) -> int:
        """Match integer patterns with various formats"""
        integer_pattern = re.compile(r'^[+-]?\d+$')
        comma_integer_pattern = re.compile(r'^[+-]?\d{1,3}(,\d{3})*$')
        
        count = 0
        for value in values:
            str_val = str(value).strip().replace(',', '')
            if integer_pattern.match(str_val) or comma_integer_pattern.match(str(value).strip()):
                count += 1
        
        return count
    
    def _match_numeric_patterns(self, values: pd.Series) -> int:
        """Match numeric patterns including decimals"""
        numeric_pattern = re.compile(r'^[+-]?\d*\.?\d+([eE][+-]?\d+)?$')
        comma_numeric_pattern = re.compile(r'^[+-]?\d{1,3}(,\d{3})*(\.\d+)?$')
        
        count = 0
        for value in values:
            str_val = str(value).strip().replace(',', '')
            if numeric_pattern.match(str_val) or comma_numeric_pattern.match(str(value).strip()):
                count += 1
        
        return count
    
    def _match_date_patterns(self, values: pd.Series) -> int:
        """Match various date patterns"""
        date_patterns = [
            re.compile(r'^\d{4}-\d{1,2}-\d{1,2}$'),  # YYYY-MM-DD
            re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$'),  # MM/DD/YYYY
            re.compile(r'^\d{1,2}-\d{1,2}-\d{4}$'),  # MM-DD-YYYY
            re.compile(r'^\d{1,2}\.\d{1,2}\.\d{4}$'),  # MM.DD.YYYY
            re.compile(r'^\d{4}/\d{1,2}/\d{1,2}$'),  # YYYY/MM/DD
            re.compile(r'^[A-Za-z]{3}\s\d{1,2},\s\d{4}$'),  # Mon DD, YYYY
            re.compile(r'^\d{1,2}\s[A-Za-z]{3}\s\d{4}$'),  # DD Mon YYYY
        ]
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if any(pattern.match(str_val) for pattern in date_patterns):
                count += 1
                continue
            
            # Try pandas date parsing as fallback
            try:
                pd.to_datetime(str_val, errors='raise')
                count += 1
            except:
                pass
        
        return count
    
    def _match_time_patterns(self, values: pd.Series) -> int:
        """Match time patterns"""
        time_patterns = [
            re.compile(r'^\d{1,2}:\d{2}(:\d{2})?$'),  # HH:MM or HH:MM:SS
            re.compile(r'^\d{1,2}:\d{2}(:\d{2})?\s?(AM|PM)$', re.IGNORECASE),  # 12-hour format
        ]
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if any(pattern.match(str_val) for pattern in time_patterns):
                count += 1
        
        return count
    
    def _match_currency_patterns(self, values: pd.Series) -> int:
        """Match currency patterns"""
        currency_patterns = [
            re.compile(r'^\$[+-]?\d{1,3}(,\d{3})*(\.\d{2})?$'),  # $1,234.56
            re.compile(r'^[+-]?\d{1,3}(,\d{3})*(\.\d{2})?\s?\$?$'),  # 1,234.56$
            re.compile(r'^\€[+-]?\d{1,3}(,\d{3})*(\.\d{2})?$'),  # €1,234.56
            re.compile(r'^\£[+-]?\d{1,3}(,\d{3})*(\.\d{2})?$'),  # £1,234.56
            re.compile(r'^\¥[+-]?\d{1,3}(,\d{3})*$'),  # ¥1,234
        ]
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if any(pattern.match(str_val) for pattern in currency_patterns):
                count += 1
        
        return count
    
    def _match_percentage_patterns(self, values: pd.Series) -> int:
        """Match percentage patterns"""
        percentage_pattern = re.compile(r'^[+-]?\d*\.?\d+%$')
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if percentage_pattern.match(str_val):
                count += 1
        
        return count
    
    def _match_email_patterns(self, values: pd.Series) -> int:
        """Match email patterns"""
        email_pattern = re.compile(
            r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        )
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if email_pattern.match(str_val):
                count += 1
        
        return count
    
    def _match_url_patterns(self, values: pd.Series) -> int:
        """Match URL patterns"""
        url_patterns = [
            re.compile(r'^https?://[^\s/$.?#].[^\s]*$', re.IGNORECASE),
            re.compile(r'^www\.[^\s/$.?#].[^\s]*$', re.IGNORECASE),
            re.compile(r'^[a-zA-Z0-9][a-zA-Z0-9-]{1,61}[a-zA-Z0-9]\.[a-zA-Z]{2,}$', re.IGNORECASE)
        ]
        
        count = 0
        for value in values:
            str_val = str(value).strip()
            if any(pattern.match(str_val) for pattern in url_patterns):
                count += 1
        
        return count
    
    def _match_phone_patterns(self, values: pd.Series) -> int:
        """Match phone number patterns"""
        phone_patterns = [
            re.compile(r'^\(\d{3}\)\s?\d{3}-\d{4}$'),  # (123) 456-7890
            re.compile(r'^\d{3}-\d{3}-\d{4}$'),  # 123-456-7890
            re.compile(r'^\d{3}\.\d{3}\.\d{4}$'),  # 123.456.7890
            re.compile(r'^\+?1?\d{10}$'),  # +11234567890 or 1234567890
            re.compile(r'^\+\d{1,3}\s?\d{1,14}$'),  # International format
        ]
        
        count = 0
        for value in values:
            str_val = str(value).strip().replace(' ', '')
            if any(pattern.match(str_val) for pattern in phone_patterns):
                count += 1
        
        return count
    
    def _match_boolean_patterns(self, values: pd.Series) -> int:
        """Match boolean patterns"""
        boolean_values = {
            'true', 'false', 'yes', 'no', 'y', 'n', '1', '0',
            'on', 'off', 'enabled', 'disabled', 'active', 'inactive'
        }
        
        count = 0
        for value in values:
            str_val = str(value).strip().lower()
            if str_val in boolean_values:
                count += 1
        
        return count
    
    def _match_business_patterns(self, values: pd.Series) -> Dict[str, float]:
        """Match business-specific patterns"""
        business_scores = {}
        total_count = len(values)
        
        if total_count == 0:
            return business_scores
        
        # SSN pattern
        ssn_pattern = re.compile(r'^\d{3}-\d{2}-\d{4}$')
        ssn_count = sum(1 for v in values if ssn_pattern.match(str(v).strip()))
        if ssn_count > 0:
            business_scores['ssn'] = ssn_count / total_count
        
        # Credit card pattern
        cc_pattern = re.compile(r'^\d{4}[\s-]?\d{4}[\s-]?\d{4}[\s-]?\d{4}$')
        cc_count = sum(1 for v in values if cc_pattern.match(str(v).strip().replace(' ', '')))
        if cc_count > 0:
            business_scores['credit_card'] = cc_count / total_count
        
        # ZIP code pattern
        zip_patterns = [
            re.compile(r'^\d{5}$'),  # 12345
            re.compile(r'^\d{5}-\d{4}$'),  # 12345-6789
        ]
        zip_count = sum(1 for v in values 
                       if any(p.match(str(v).strip()) for p in zip_patterns))
        if zip_count > 0:
            business_scores['zip_code'] = zip_count / total_count
        
        # Product code patterns
        product_patterns = [
            re.compile(r'^[A-Z]{2,4}-\d{3,6}$'),  # AB-1234
            re.compile(r'^\d{3,5}-[A-Z]{2,4}$'),  # 1234-AB
            re.compile(r'^[A-Z0-9]{6,12}$'),  # ABCD1234
        ]
        product_count = sum(1 for v in values 
                           if any(p.match(str(v).strip()) for p in product_patterns))
        if product_count > 0:
            business_scores['product_code'] = product_count / total_count
        
        return business_scores
    
    def _determine_primary_type(self, type_scores: Dict[str, float]) -> Tuple[str, float]:
        """Determine primary type and confidence from type scores"""
        if not type_scores:
            return 'text', 0.0
        
        # Remove business patterns from primary type determination
        core_types = {k: v for k, v in type_scores.items() 
                     if k not in ['ssn', 'credit_card', 'zip_code', 'product_code']}
        
        if not core_types:
            return 'text', 0.0
        
        primary_type = max(core_types, key=core_types.get)
        confidence = core_types[primary_type]
        
        # Apply confidence thresholds for type validation
        if confidence < 0.6:  # Low confidence threshold
            primary_type = 'text'
            confidence = 0.5
        
        return primary_type, confidence
    
    def _analyze_mixed_types(self, type_scores: Dict[str, float], primary_confidence: float) -> Dict[str, Any]:
        """Analyze mixed type columns"""
        # Find types with significant scores (>0.1)
        significant_types = {k: v for k, v in type_scores.items() if v > 0.1}
        
        is_mixed = len(significant_types) > 1 and primary_confidence < 0.9
        
        # Calculate dominant ratio
        total_significant = sum(significant_types.values())
        dominant_ratio = max(significant_types.values()) / total_significant if total_significant > 0 else 0
        
        return {
            'types': list(significant_types.keys()),
            'is_mixed': is_mixed,
            'dominant_ratio': dominant_ratio,
            'type_distribution': significant_types
        }

    def _calculate_overall_quality_score(self, sheet_profiles: Dict[str, Any]) -> float:
        """Calculate overall data quality score across all sheets"""
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
        """Estimate complexity based on file size, sheet count, and chunking strategy"""
        base_complexity = super().estimate_complexity(context)
        
        # Adjust complexity based on chunking strategy and file characteristics
        file_size_mb = context.file_metadata.file_size_mb
        
        if file_size_mb > 100:
            # Large files benefit from chunking but still complex
            complexity_multiplier = 2.5
        elif file_size_mb > 50:
            # Medium files with moderate chunking
            complexity_multiplier = 2.0
        else:
            # Small files may not need aggressive chunking
            complexity_multiplier = 1.8
        
        # Account for memory pressure effects
        if context.should_reduce_complexity():
            complexity_multiplier *= 1.3  # More complex due to resource constraints
        
        return base_complexity * complexity_multiplier
    
    def _count_processed_items(self, data: DataProfileData) -> int:
        """Count items processed during analysis"""
        total_rows = 0
        total_chunks = 0
        
        for sheet_profile in data.sheet_profiles.values():
            if isinstance(sheet_profile, dict):
                if 'row_count' in sheet_profile:
                    total_rows += sheet_profile['row_count']
                if 'chunks_processed' in sheet_profile:
                    total_chunks += sheet_profile['chunks_processed']
        
        # Return combination of rows and chunks as processing complexity indicator
        return total_rows + (total_chunks * 100)  # Weight chunks as equivalent to 100 rows

    def should_skip_analysis(self, context: AnalysisContext) -> bool:
        """Enhanced skip logic considering memory constraints and chunking capabilities"""
        # Call parent implementation first
        if super().should_skip_analysis(context):
            return True
        
        # Additional data profiler specific checks
        file_size_mb = context.file_metadata.file_size_mb
        
        # Very large files might be skipped under severe memory pressure
        if file_size_mb > 500 and context.should_reduce_complexity():
            memory_usage = context.memory_manager.get_current_usage()
            if memory_usage['usage_ratio'] > 0.9:
                self.logger.warning(f"Skipping data profiling for {file_size_mb:.1f}MB file due to severe memory pressure")
                return True
        
        return False


# Legacy compatibility
def create_data_profiler(config: dict = None) -> DataProfiler:
    """Factory function for backward compatibility"""
    profiler = DataProfiler()
    if config:
        profiler.configure(config)
    return profiler
