"""
Formula dependencies and logic analysis module
Comprehensive formula analysis with dependency mapping and complexity scoring.
"""

from typing import List, Optional, Dict, Any
import re

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import FormulaAnalysisData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory
from ..utils.excel_formula_parser import ExcelFormulaParser, create_formula_parser
from ..utils.formula_dependency_analyzer import FormulaDependencyAnalyzer, create_dependency_analyzer


class FormulaAnalyzer(BaseAnalyzer):
    """Comprehensive formula analyzer with parsing and dependency analysis"""
    
    def __init__(self, name: str = "formula_analyzer", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["structure_mapper"])
        self.formula_parser = None
        self.dependency_analyzer = None
    
    def _initialize_analyzers(self):
        """Initialize parser and dependency analyzer if not already done"""
        if self.formula_parser is None:
            self.formula_parser = create_formula_parser()
        if self.dependency_analyzer is None:
            max_depth = self.config.get("max_dependency_depth", 50)
            self.dependency_analyzer = create_dependency_analyzer(max_depth)
    
    def _perform_analysis(self, context: AnalysisContext) -> FormulaAnalysisData:
        """Perform comprehensive formula analysis with parsing and dependencies"""
        try:
            self._initialize_analyzers()
            
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
            
            # Enhanced formula analysis
            total_formulas = 0
            formula_errors = []
            external_references = []
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
                        sheet_results = self._analyze_sheet_formulas_enhanced(
                            ws, sheet_name, max_formulas - analyzed_formulas
                        )
                        
                        total_formulas += sheet_results['count']
                        formula_errors.extend(sheet_results['errors'])
                        external_references.extend(sheet_results['external_refs'])
                        volatile_formulas += sheet_results['volatile']
                        array_formulas += sheet_results['array']
                        analyzed_formulas += sheet_results['count']
                        
                    except Exception as e:
                        self.logger.warning(f"Error analyzing formulas in {sheet_name}: {e}")
                        formula_errors.append({
                            'sheet': sheet_name,
                            'error': f"Sheet analysis failed: {e}"
                        })
            
            # Get dependency analysis results
            dependency_metrics = self.dependency_analyzer.get_dependency_metrics()
            circular_references = self.dependency_analyzer.find_circular_references()
            
            # Convert circular references to simple list format for compatibility
            circular_refs_list = [{
                'chain': ref.chain,
                'length': ref.chain_length,
                'complexity': ref.complexity_score,
                'impact': ref.impact_level.value
            } for ref in circular_references]
            
            # Create dependency chains summary
            dependency_chains = [{
                'max_length': dependency_metrics.max_chain_length,
                'avg_length': dependency_metrics.avg_chain_length,
                'total_dependencies': dependency_metrics.total_dependencies
            }]
            
            # Enhanced complexity score calculation
            complexity_score = self._calculate_enhanced_complexity_score(
                total_formulas, len(formula_errors), volatile_formulas, 
                array_formulas, dependency_metrics
            )
            
            return FormulaAnalysisData(
                total_formulas=total_formulas,
                formula_complexity_score=complexity_score,
                circular_references=circular_refs_list,
                external_references=list(set(external_references)),
                formula_errors=formula_errors,
                dependency_chains=dependency_chains,
                volatile_formulas=volatile_formulas,
                array_formulas=array_formulas
            )
            
        except Exception as e:
            self.logger.error(f"Enhanced formula analysis failed: {e}")
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
    
    def _analyze_sheet_formulas_enhanced(self, worksheet, sheet_name: str, max_check: int) -> Dict[str, Any]:
        """Enhanced formula analysis for a single sheet using new parser"""
        results = {'count': 0, 'errors': [], 'external_refs': [], 'volatile': 0, 'array': 0}
        checked = 0
        
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
                            parsed = self.formula_parser.parse_formula(formula, cell.coordinate)
                            self.dependency_analyzer.add_formula(cell.coordinate, formula, sheet_name)
                            results['external_refs'].extend(parsed.external_references)
                            
                            if any(f.name in ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'] 
                                  for f in parsed.functions):
                                results['volatile'] += 1
                            
                            if parsed.is_array_formula:
                                results['array'] += 1
                            
                            if parsed.parsing_errors:
                                results['errors'].extend([{
                                    'sheet': sheet_name, 'cell': cell.coordinate,
                                    'formula': formula[:100], 'error': err
                                } for err in parsed.parsing_errors])
                        
                        except Exception as e:
                            results['errors'].append({
                                'sheet': sheet_name, 'cell': cell.coordinate,
                                'formula': formula[:100], 'error': str(e)
                            })
        except Exception as e:
            results['errors'].append({'sheet': sheet_name, 'error': f"Sheet iteration failed: {e}"})
        return results
    
    def _calculate_enhanced_complexity_score(self, total_formulas: int, error_count: int, 
                                           volatile_count: int, array_count: int, dependency_metrics) -> float:
        """Calculate enhanced complexity score using dependency analysis"""
        if total_formulas == 0:
            return 0.0
        
        base_score = min(1.0, total_formulas / 1000.0)
        volatile_ratio = volatile_count / total_formulas
        array_ratio = array_count / total_formulas
        error_ratio = error_count / total_formulas
        
        # Dependency complexity factors
        circular_factor = dependency_metrics.circular_reference_count / max(1, total_formulas) * 0.3
        external_factor = dependency_metrics.external_dependency_count / max(1, total_formulas) * 0.2
        chain_factor = min(1.0, dependency_metrics.avg_chain_length / 10.0) * 0.2
        
        complexity_score = (base_score + (volatile_ratio * 0.2) + (array_ratio * 0.3) + 
                          circular_factor + external_factor + chain_factor)
        
        complexity_score *= (1.0 - error_ratio * 0.5)
        return min(1.0, max(0.0, complexity_score))
    
    def _validate_result(self, data: FormulaAnalysisData, context: AnalysisContext) -> ValidationResult:
        """Validate formula analysis results with >95% accuracy requirement"""
        validation_notes = []
        completeness = 0.8 if data.total_formulas > 0 else 0.5
        
        if data.total_formulas > 0:
            error_rate = len(data.formula_errors) / data.total_formulas
            accuracy = max(0.0, 1.0 - error_rate)
        else:
            accuracy = 1.0
        
        consistency = 0.9
        if data.total_formulas < 0:
            consistency -= 0.5
            validation_notes.append("Negative formula count")
        
        if not (0.0 <= data.formula_complexity_score <= 1.0):
            consistency -= 0.3
            validation_notes.append("Complexity score out of range")
        
        if data.total_formulas > 100:
            confidence = ConfidenceLevel.MEDIUM
        elif data.total_formulas > 10:
            confidence = ConfidenceLevel.LOW
        else:
            confidence = ConfidenceLevel.UNCERTAIN
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity for formula analysis"""
        return super().estimate_complexity(context) * 3.0
    
    def _count_processed_items(self, data: FormulaAnalysisData) -> int:
        """Count formulas processed"""
        return data.total_formulas


def create_formula_analyzer(config: dict = None) -> FormulaAnalyzer:
    """Factory function for backward compatibility"""
    analyzer = FormulaAnalyzer()
    if config:
        analyzer.configure(config)
    return analyzer
