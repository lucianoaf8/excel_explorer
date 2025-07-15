"""
Charts, images, shapes inventory module
Framework-compatible placeholder with basic visual element detection.
"""

from typing import List, Optional, Dict, Any

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import VisualCatalogData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class VisualCataloger(BaseAnalyzer):
    """Framework-compatible visual cataloger with basic implementation"""
    
    def __init__(self, name: str = "visual_cataloger", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["structure_mapper"])
    
    def _perform_analysis(self, context: AnalysisContext) -> VisualCatalogData:
        """Perform basic visual element cataloging
        
        Args:
            context: AnalysisContext with workbook access
            
        Returns:
            VisualCatalogData with visual element inventory
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
            
            # Visual element inventories
            charts = []
            images = []
            shapes = []
            conditional_formatting = []
            data_validation = []
            
            with context.get_workbook_access().get_workbook() as wb:
                for sheet_name in sheet_names:
                    try:
                        ws = wb[sheet_name]
                        
                        # Catalog charts
                        sheet_charts = self._catalog_charts(ws, sheet_name)
                        charts.extend(sheet_charts)
                        
                        # Catalog images/drawings
                        sheet_images = self._catalog_images(ws, sheet_name)
                        images.extend(sheet_images)
                        
                        # Catalog shapes (basic detection)
                        sheet_shapes = self._catalog_shapes(ws, sheet_name)
                        shapes.extend(sheet_shapes)
                        
                        # Catalog conditional formatting
                        sheet_cf = self._catalog_conditional_formatting(ws, sheet_name)
                        conditional_formatting.extend(sheet_cf)
                        
                        # Catalog data validation
                        sheet_dv = self._catalog_data_validation(ws, sheet_name)
                        data_validation.extend(sheet_dv)
                        
                    except Exception as e:
                        self.logger.warning(f"Error cataloging visuals in {sheet_name}: {e}")
            
            # Calculate visual complexity score
            visual_complexity_score = self._calculate_visual_complexity(
                charts, images, shapes, conditional_formatting, data_validation
            )
            
            return VisualCatalogData(
                charts=charts,
                images=images,
                shapes=shapes,
                conditional_formatting=conditional_formatting,
                data_validation=data_validation,
                visual_complexity_score=visual_complexity_score
            )
            
        except Exception as e:
            # Return minimal data rather than failing
            self.logger.error(f"Visual cataloging failed: {e}")
            return VisualCatalogData(
                charts=[],
                images=[],
                shapes=[],
                conditional_formatting=[],
                data_validation=[],
                visual_complexity_score=0.0
            )
    
    def _validate_result(self, data: VisualCatalogData, context: AnalysisContext) -> ValidationResult:
        """Validate visual cataloging results
        
        Args:
            data: VisualCatalogData to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness based on visual element detection
        total_elements = (
            len(data.charts) + len(data.images) + len(data.shapes) +
            len(data.conditional_formatting) + len(data.data_validation)
        )
        
        if total_elements > 50:
            completeness = 0.9  # High visual content
        elif total_elements > 10:
            completeness = 0.7  # Moderate visual content
        elif total_elements > 0:
            completeness = 0.5  # Some visual content
        else:
            completeness = 0.3  # No visual elements (could be valid)
        
        # Accuracy - assume high for basic cataloging
        accuracy = 0.8
        
        # Consistency checks
        consistency = 0.9
        if not (0.0 <= data.visual_complexity_score <= 1.0):
            consistency -= 0.3
            validation_notes.append("Complexity score out of range")
        
        # Confidence based on element count
        if total_elements > 20:
            confidence = ConfidenceLevel.MEDIUM
        elif total_elements > 5:
            confidence = ConfidenceLevel.LOW
        else:
            confidence = ConfidenceLevel.UNCERTAIN
        
        if len(data.charts) > 10:
            validation_notes.append("High chart count detected")
        if len(data.images) > 20:
            validation_notes.append("High image count detected")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _catalog_charts(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Catalog charts in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of chart information dictionaries
        """
        charts = []
        
        try:
            if hasattr(worksheet, '_charts'):
                for i, chart in enumerate(worksheet._charts):
                    chart_info = {
                        'sheet': sheet_name,
                        'id': f"chart_{i}",
                        'type': getattr(chart, 'tagname', 'unknown'),
                        'title': getattr(chart, 'title', None),
                        'position': {
                            'anchor': str(getattr(chart, 'anchor', 'unknown'))
                        }
                    }
                    
                    # Try to get data source information
                    try:
                        if hasattr(chart, 'series') and chart.series:
                            chart_info['series_count'] = len(chart.series)
                            # Get first series as example
                            if chart.series[0]:
                                series = chart.series[0]
                                if hasattr(series, 'val') and series.val:
                                    chart_info['data_source'] = str(series.val)
                    except Exception:
                        pass
                    
                    charts.append(chart_info)
        
        except Exception as e:
            self.logger.warning(f"Error cataloging charts in {sheet_name}: {e}")
        
        return charts
    
    def _catalog_images(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Catalog images in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of image information dictionaries
        """
        images = []
        
        try:
            if hasattr(worksheet, '_images'):
                for i, image in enumerate(worksheet._images):
                    image_info = {
                        'sheet': sheet_name,
                        'id': f"image_{i}",
                        'type': 'image',
                        'position': {
                            'anchor': str(getattr(image, 'anchor', 'unknown'))
                        }
                    }
                    
                    # Try to get image metadata
                    try:
                        if hasattr(image, 'ref'):
                            image_info['reference'] = image.ref
                        if hasattr(image, 'format'):
                            image_info['format'] = image.format
                    except Exception:
                        pass
                    
                    images.append(image_info)
        
        except Exception as e:
            self.logger.warning(f"Error cataloging images in {sheet_name}: {e}")
        
        return images
    
    def _catalog_shapes(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Catalog shapes in a worksheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of shape information dictionaries
        """
        shapes = []
        
        try:
            # Check for drawing objects
            if hasattr(worksheet, '_drawings') and worksheet._drawings:
                for i, drawing in enumerate(worksheet._drawings):
                    shape_info = {
                        'sheet': sheet_name,
                        'id': f"drawing_{i}",
                        'type': 'drawing_object',
                        'description': str(type(drawing).__name__)
                    }
                  
                    # Try to get more specific information
                    if hasattr(drawing, 'anchor'):
                        shape_info['anchor'] = str(drawing.anchor)[:50]
                      
                    shapes.append(shape_info)
                  
            # Check for text boxes and other shapes by examining cell formatting
            shape_indicators = 0
            for row in worksheet.iter_rows(max_row=50):
                for cell in row:
                    if cell.value and hasattr(cell, 'fill'):
                        if (cell.fill and cell.fill.fill_type and 
                            cell.fill.fill_type != 'none'):
                            shape_indicators += 1
                          
            if shape_indicators > 10:  # Threshold for shape-heavy worksheet
                shapes.append({
                    'sheet': sheet_name,
                    'id': 'formatting_shapes',
                    'type': 'formatted_cells',
                    'count': shape_indicators,
                    'description': f'Detected {shape_indicators} formatted cells indicating shapes'
                })
        
        except Exception as e:
            self.logger.warning(f"Error cataloging shapes in {sheet_name}: {e}")
        
        return shapes
    
    def _catalog_conditional_formatting(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Catalog conditional formatting rules
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of conditional formatting rules
        """
        cf_rules = []
        
        try:
            if hasattr(worksheet, 'conditional_formatting'):
                for i, cf in enumerate(worksheet.conditional_formatting):
                    cf_info = {
                        'sheet': sheet_name,
                        'id': f"cf_{i}",
                        'type': 'conditional_formatting',
                        'ranges': [str(range_) for range_ in cf.cells] if hasattr(cf, 'cells') else [],
                        'rule_count': len(cf.cfRule) if hasattr(cf, 'cfRule') else 0
                    }
                    cf_rules.append(cf_info)
        
        except Exception as e:
            self.logger.warning(f"Error cataloging conditional formatting in {sheet_name}: {e}")
        
        return cf_rules
    
    def _catalog_data_validation(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """Catalog data validation rules
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            
        Returns:
            List of data validation rules
        """
        dv_rules = []
        
        try:
            if hasattr(worksheet, 'data_validations'):
                for i, dv in enumerate(worksheet.data_validations.dataValidation):
                    dv_info = {
                        'sheet': sheet_name,
                        'id': f"dv_{i}",
                        'type': 'data_validation',
                        'ranges': dv.ranges if hasattr(dv, 'ranges') else [],
                        'validation_type': getattr(dv, 'type', 'unknown'),
                        'operator': getattr(dv, 'operator', None)
                    }
                    dv_rules.append(dv_info)
        
        except Exception as e:
            self.logger.warning(f"Error cataloging data validation in {sheet_name}: {e}")
        
        return dv_rules
    
    def _calculate_visual_complexity(self, charts: List[Dict], images: List[Dict], 
                                   shapes: List[Dict], cf_rules: List[Dict], 
                                   dv_rules: List[Dict]) -> float:
        """Calculate visual complexity score
        
        Args:
            charts: List of chart elements
            images: List of image elements
            shapes: List of shape elements
            cf_rules: List of conditional formatting rules
            dv_rules: List of data validation rules
            
        Returns:
            Complexity score between 0.0 and 1.0
        """
        # Weight different visual elements
        weighted_score = (
            len(charts) * 0.3 +      # Charts are complex
            len(images) * 0.2 +      # Images add complexity
            len(shapes) * 0.1 +      # Shapes are simpler
            len(cf_rules) * 0.15 +   # Conditional formatting is moderately complex
            len(dv_rules) * 0.1      # Data validation is simpler
        )
        
        # Normalize to 0-1 scale (assuming 20 weighted elements = complexity 1.0)
        return min(1.0, weighted_score / 20.0)
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity for visual analysis
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier
        """
        base_complexity = super().estimate_complexity(context)
        
        # Visual analysis is moderately intensive
        return base_complexity * 1.5
    
    def _count_processed_items(self, data: VisualCatalogData) -> int:
        """Count visual elements processed
        
        Args:
            data: VisualCatalogData result
            
        Returns:
            int: Number of visual elements processed
        """
        return (
            len(data.charts) + len(data.images) + len(data.shapes) +
            len(data.conditional_formatting) + len(data.data_validation)
        )


# Legacy compatibility
def create_visual_cataloger(config: dict = None) -> VisualCataloger:
    """Factory function for backward compatibility"""
    cataloger = VisualCataloger()
    if config:
        cataloger.configure(config)
    return cataloger
