"""
Formula dependencies and logic analysis module
Framework-compatible placeholder with basic formula detection.
"""

from typing import List, Optional, Dict, Any
import re

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import FormulaAnalysisData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class FormulaAnalyzer(BaseAnalyzer):
    """Framework-compatible formula analyzer with basic implementation"""
    
    def __init__(self, name: str = "formula_analyzer", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["structure_mapper"])
    
    def _perform_analysis(self, context: AnalysisContext) -> FormulaAnalysisData:
        """Perform basic formula analysis
        
        Args:
            context: AnalysisContext with workbook access
            
        Returns:
            FormulaAnalysisData with formula statistics
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
            
            # Basic formula analysis
            total_formulas = 0
            formula_errors = []
            external_references = []
            circular_references = []
            dependency_chains = []
            volatile_formulas = 0
            array_formulas = 0
            
            max_formulas = self.config.get("max_formulas_analyze", 50000)
            analyzed_formulas = 0
            
            with context.get_workbook_access().get_workbook() as wb:
                for sheet_name in sheet_names:
                    if analyzed_formulas >= max_formulas:
                        self.logger.warning(f"Formula analysis limit reached: {max_formulas}")
                        break
                    
                    try:
                        ws = wb[sheet_name]
                        sheet_formulas = self._analyze_sheet_formulas(
                            ws, sheet_name, max_formulas - analyzed_formulas
                        )
                        
                        total_formulas += sheet_formulas['count']
                        formula_errors.extend(sheet_formulas['errors'])
                        external_references.extend(sheet_formulas['external_refs'])
                        volatile_formulas += sheet_formulas['volatile']
                        array_formulas += sheet_formulas['array']
                        analyzed_formulas += sheet_formulas['count']
                        
                    except Exception as e:
                        self.logger.warning(f"Error analyzing formulas in {sheet_name}: {e}")
                        formula_errors.append({
                            'sheet': sheet_name,
                            'error': f"Sheet analysis failed: {e}"
                        })
            
            # Calculate complexity score
            complexity_score = self._calculate_complexity_score(
                total_formulas, len(formula_errors), volatile_formulas, array_formulas
            )
            
            return FormulaAnalysisData(
                total_formulas=total_formulas,
                formula_complexity_score=complexity_score,
                circular_references=circular_references,
                external_references=list(set(external_references)),
                formula_errors=formula_errors,
                dependency_chains=dependency_chains,
                volatile_formulas=volatile_formulas,
                array_formulas=array_formulas
            )
            
        except Exception as e:
            # For now, return minimal data rather than failing
            self.logger.error(f"Formula analysis failed: {e}")
            return FormulaAnalysisData(
                total_formulas=0,
                formula_complexity_score=0.0,
                circular_references=[],
                external_references=[],
                formula_errors=[{'error': str(e)}],
                dependency_chains=[],
                volatile_formulas=0,
                array_formulas=0
            )
    
    def _validate_result(self, data: FormulaAnalysisData, context: AnalysisContext) -> ValidationResult:
        """Validate formula analysis results
        
        Args:
            data: FormulaAnalysisData to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness based on whether we found formulas
        if data.total_formulas > 0:
            completeness = 0.8  # Found and analyzed formulas
        else:
            completeness = 0.5  # No formulas found (could be valid)
        
        # Accuracy based on error rate
        if data.total_formulas > 0:
            error_rate = len(data.formula_errors) / data.total_formulas
            accuracy = max(0.0, 1.0 - error_rate)
        else:
            accuracy = 1.0  # No formulas to analyze incorrectly
        
        # Consistency checks
        consistency = 0.9
        if data.total_formulas < 0:
            consistency -= 0.5
            validation_notes.append("Negative formula count")
        
        if not (0.0 <= data.formula_complexity_score <= 1.0):
            consistency -= 0.3
            validation_notes.append("Complexity score out of range")
        
        # Confidence assessment
        if data.total_formulas > 100:
            confidence = ConfidenceLevel.MEDIUM  # Substantial analysis
        elif data.total_formulas > 10:
            confidence = ConfidenceLevel.LOW  # Limited analysis
        else:
            confidence = ConfidenceLevel.UNCERTAIN  # Very limited
        
        if data.total_formulas > 1000:
            validation_notes.append("High formula count detected")
        if len(data.circular_references) > 0:
            validation_notes.append("Circular references detected")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _analyze_sheet_formulas(self, worksheet, sheet_name: str, max_check: int) -> Dict[str, Any]:
        """Basic formula analysis for a single sheet
        
        Args:
            worksheet: openpyxl worksheet object
            sheet_name: Name of the sheet
            max_check: Maximum formulas to analyze
            
        Returns:
            Dict with sheet formula statistics
        """
        results = {
            'count': 0,
            'errors': [],
            'external_refs': [],
            'volatile': 0,
            'array': 0
        }
        
        checked = 0
        volatile_functions = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'}
        
        try:
            for row in worksheet.iter_rows():
                if checked >= max_check:
                    break
                
                for cell in row:
                    if checked >= max_check:
                        break
                    
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                        results['count'] += 1
                        checked += 1
                        formula = cell.value
                        
                        try:
                            # Check for external references
                            if '[' in formula and ']' in formula:
                                # Pattern like [workbook.xlsx]Sheet1!A1
                                ext_refs = re.findall(r'\[([^\]]+)\]', formula)
                                results['external_refs'].extend(ext_refs)
                            
                            # Check for volatile functions
                            formula_upper = formula.upper()
                            for func in volatile_functions:
                                if func in formula_upper:
                                    results['volatile'] += 1
                                    break
                            
                            # Check for array formulas (basic detection)
                            if '{' in formula and '}' in formula:
                                results['array'] += 1
                        
                        except Exception as e:
                            results['errors'].append({
                                'sheet': sheet_name,
                                'cell': cell.coordinate,
                                'formula': formula[:100],  # Truncate long formulas
                                'error': str(e)
                            })
        
        except Exception as e:
            results['errors'].append({
                'sheet': sheet_name,
                'error': f"Sheet iteration failed: {e}"
            })
        
        return results
    
    def _calculate_complexity_score(self, total_formulas: int, error_count: int, 
                                  volatile_count: int, array_count: int) -> float:
        """Calculate formula complexity score
        
        Args:
            total_formulas: Total number of formulas
            error_count: Number of formula errors
            volatile_count: Number of volatile formulas
            array_count: Number of array formulas
            
        Returns:
            Complexity score between 0.0 and 1.0
        """
        if total_formulas == 0:
            return 0.0
        
        # Base complexity from formula count
        base_score = min(1.0, total_formulas / 1000.0)  # Normalize to 1000 formulas
        
        # Adjust for special formula types
        volatile_ratio = volatile_count / total_formulas
        array_ratio = array_count / total_formulas
        error_ratio = error_count / total_formulas
        
        # Higher ratios increase complexity
        complexity_score = base_score + (volatile_ratio * 0.2) + (array_ratio * 0.3)
        
        # Errors actually reduce the score (indicate problems)
        complexity_score *= (1.0 - error_ratio * 0.5)
        
        return min(1.0, max(0.0, complexity_score))
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity for formula analysis
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier
        """
        base_complexity = super().estimate_complexity(context)
        
        # Formula analysis can be very intensive
        return base_complexity * 3.0
    
    def _count_processed_items(self, data: FormulaAnalysisData) -> int:
        """Count formulas processed
        
        Args:
            data: FormulaAnalysisData result
            
        Returns:
            int: Number of formulas processed
        """
        return data.total_formulas


# Legacy compatibility
def create_formula_analyzer(config: dict = None) -> FormulaAnalyzer:
    """Factory function for backward compatibility"""
    analyzer = FormulaAnalyzer()
    if config:
        analyzer.configure(config)
    return analyzer
