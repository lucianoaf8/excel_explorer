"""
Pivot table analysis module
Framework-compatible placeholder with basic pivot detection.
"""

from typing import List, Optional, Dict, Any

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import PivotAnalysisData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class PivotIntelligence(BaseAnalyzer):
    """Framework-compatible pivot intelligence with basic implementation"""
    
    def __init__(self, name: str = "pivot_intelligence", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["structure_mapper", "data_profiler"])
    
    def _perform_analysis(self, context: AnalysisContext) -> PivotAnalysisData:
        """Perform basic pivot table analysis
        
        Args:
            context: AnalysisContext with workbook access
            
        Returns:
            PivotAnalysisData with pivot table inventory
        """
        try:
            # Get structure information
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
            
            # Pivot analysis inventories
            pivot_tables = []
            pivot_charts = []
            data_sources = []
            calculated_fields = []
            slicers = []
            refresh_metadata = {}
            
            with context.get_workbook_access().get_workbook() as wb:
                for sheet_name in sheet_names:
                    try:
                        ws = wb[sheet_name]
                        
                        # Analyze pivot tables
                        sheet_pivots = self._analyze_pivot_tables(ws, sheet_name)
                        pivot_tables.extend(sheet_pivots['tables'])
                        data_sources.extend(sheet_pivots['sources'])
                        calculated_fields.extend(sheet_pivots['calculated_fields'])
                        
                        # Analyze pivot charts
                        sheet_pivot_charts = self._analyze_pivot_charts(ws, sheet_name)
                        pivot_charts.extend(sheet_pivot_charts)
                        
                        # Analyze slicers
                        sheet_slicers = self._analyze_slicers(ws, sheet_name)
                        slicers.extend(sheet_slicers)
                        
                    except Exception as e:
                        self.logger.warning(f"Error analyzing pivots in {sheet_name}: {e}")
                
                # Analyze refresh metadata
                refresh_metadata = self._analyze_refresh_metadata(wb)
            
            return PivotAnalysisData(
                pivot_tables=pivot_tables,
                pivot_charts=pivot_charts,
                data_sources=list(set(data_sources)),  # Remove duplicates
                calculated_fields=calculated_fields,
                slicers=slicers,
                refresh_metadata=refresh_metadata
            )
            
        except Exception as e:
            # Return minimal data rather than failing
            self.logger.error(f"Pivot analysis failed: {e}")
            return PivotAnalysisData(
                pivot_tables=[],
                pivot_charts=[],
                data_sources=[],
                calculated_fields=[],
                slicers=[],
                refresh_metadata={'error': str(e)}
            )
    
    def _validate_result(self, data: PivotAnalysisData, context: AnalysisContext) -> ValidationResult:
        """Validate pivot analysis results
        
        Args:
            data: PivotAnalysisData to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness based on pivot element detection
        total_pivot_elements = (
            len(data.pivot_tables) + len(data.pivot_charts) + len(data.slicers)
        )
        
        # Many workbooks don't have pivot tables, so finding none is often correct
        if total_pivot_elements > 10:
            completeness = 0.9  # High pivot usage
        elif total_pivot_elements > 2:
            completeness = 0.7  # Moderate pivot usage
        elif total_pivot_elements > 0:
            completeness = 0.6  # Some pivots found
        else:
            completeness = 0.5  # No pivots (often valid)
        
        # Accuracy - assume good for basic detection
        accuracy = 0.8
        
        # Consistency checks
        consistency = 0.9
        if 'error' in data.refresh_metadata:
            consistency -= 0.3
            validation_notes.append("Refresh metadata analysis failed")
        
        # Check logical relationships
        if len(data.pivot_charts) > len(data.pivot_tables):
            # This could be valid if charts reference external pivots
            validation_notes.append("More pivot charts than tables detected")
        
        # Confidence based on pivot complexity
        if total_pivot_elements > 5:
            confidence = ConfidenceLevel.MEDIUM
        elif total_pivot_elements > 1:
            confidence = ConfidenceLevel.LOW
        else:
            confidence = ConfidenceLevel.UNCERTAIN
        
        if len(data.calculated_fields) > 0:
            validation_notes.append("Calculated fields detected")
        if len(data.slicers) > 0:
            validation_notes.append("Slicers detected")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _analyze_pivot_tables(self, worksheet, sheet_name: str) -> Dict[str, List[Dict[str, Any]]]:
        """Analyze pivot tables in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            Dict with pivot table information
        """
        result = {
            'tables': [],
            'sources': [],
            'calculated_fields': []
        }
        
        try:
            # Check for pivot-related content in cell values
            pivot_indicators = ['pivot', 'table', 'sum of', 'count of', 'average of']
          
            for row_idx, row in enumerate(worksheet.iter_rows(max_row=50, values_only=True), 1):
                for col_idx, cell in enumerate(row, 1):
                    if cell and isinstance(cell, str):
                        cell_lower = cell.lower()
                        if any(indicator in cell_lower for indicator in pivot_indicators):
                            result['tables'].append({
                                'sheet': sheet_name,
                                'id': f"potential_pivot_{len(result['tables'])}",
                                'location': f"{chr(64+col_idx)}{row_idx}",
                                'indicator': cell[:50],
                                'type': 'detected_content'
                            })
                          
            # Check for common pivot table patterns
            if len(result['tables']) > 3:
                result['sources'].append(f"Multiple pivot indicators in {sheet_name}")
        
        except Exception as e:
            self.logger.warning(f"Error analyzing pivot tables in {sheet_name}: {e}")
        
        return result
    
    def _analyze_pivot_charts(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Analyze pivot charts in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of pivot chart information
        """
        pivot_charts = []
        
        try:
            if hasattr(worksheet, '_charts') and worksheet._charts:
                for i, chart in enumerate(worksheet._charts):
                    chart_info = {
                        'sheet': sheet_name,
                        'id': f"chart_{i}",
                        'type': 'chart',
                        'chart_type': getattr(chart, 'tagname', 'unknown'),
                        'is_pivot_chart': False  # Default assumption
                    }
                  
                    # Basic heuristic for pivot charts
                    if hasattr(chart, 'title') and chart.title:
                        title_text = str(chart.title).lower()
                        if any(keyword in title_text for keyword in ['sum', 'count', 'average', 'total']):
                            chart_info['is_pivot_chart'] = True
                  
                    pivot_charts.append(chart_info)
        
        except Exception as e:
            self.logger.warning(f"Error analyzing pivot charts in {sheet_name}: {e}")
        
        return pivot_charts
    
    def _analyze_slicers(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Analyze slicers in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of slicer information
        """
        slicers = []
        
        try:
            # Check for slicers
            # Note: openpyxl has limited support for slicers
            # This is a placeholder implementation
            
            # Look for slicer-like content in cells
            for row_idx, row in enumerate(worksheet.iter_rows(max_row=30, values_only=True), 1):
                for col_idx, cell in enumerate(row, 1):
                    if cell and isinstance(cell, str):
                        if any(keyword in cell.lower() for keyword in ['filter', 'slicer', 'select']):
                            slicers.append({
                                'sheet': sheet_name,
                                'id': f"potential_slicer_{len(slicers)}",
                                'location': f"{chr(64+col_idx)}{row_idx}",
                                'indicator': cell[:50],
                                'type': 'content_based_detection'
                            })
        
        except Exception as e:
            self.logger.warning(f"Error analyzing slicers in {sheet_name}: {e}")
        
        return slicers
    
    def _analyze_refresh_metadata(self, workbook) -> Dict[str, Any]:
        """Analyze pivot refresh metadata
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            Dict with refresh metadata
        """
        refresh_metadata = {
            'auto_refresh': False,
            'refresh_on_open': False,
            'last_refresh': None,
            'refresh_interval': None
        }
        
        try:
            # Check workbook properties for refresh settings
            if hasattr(workbook, 'properties'):
                props = workbook.properties
                refresh_metadata['last_refresh'] = str(props.modified) if props.modified else None
            
            # Check for pivot cache refresh settings
            # This would require deeper inspection of pivot cache properties
            # which is not well supported in openpyxl
            
        except Exception as e:
            self.logger.warning(f"Error analyzing refresh metadata: {e}")
            refresh_metadata['error'] = str(e)
        
        return refresh_metadata
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity for pivot analysis
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier
        """
        base_complexity = super().estimate_complexity(context)
        
        # Pivot analysis is moderately complex
        return base_complexity * 1.3
    
    def _count_processed_items(self, data: PivotAnalysisData) -> int:
        """Count pivot elements processed
        
        Args:
            data: PivotAnalysisData result
            
        Returns:
            int: Number of pivot elements processed
        """
        return (
            len(data.pivot_tables) + len(data.pivot_charts) + 
            len(data.slicers) + len(data.calculated_fields)
        )


# Legacy compatibility
def create_pivot_intelligence(config: dict = None) -> PivotIntelligence:
    """Factory function for backward compatibility"""
    intelligence = PivotIntelligence()
    if config:
        intelligence.configure(config)
    return intelligence
