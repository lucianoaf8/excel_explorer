"""
Complexity Scoring Algorithm - Multi-dimensional formula complexity evaluation
"""

from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum
import logging

from .excel_formula_parser import ParsedFormula, FormulaFunction, FormulaComplexity


class ComplexityCategory(Enum):
    """Formula complexity categories"""
    SIMPLE = "simple"        # 0-25
    MODERATE = "moderate"    # 26-50  
    COMPLEX = "complex"      # 51-75
    CRITICAL = "critical"    # 76-100


class FunctionCategory(Enum):
    """Excel function categories with complexity weights"""
    MATHEMATICAL = "mathematical"
    LOGICAL = "logical"
    LOOKUP = "lookup"
    FINANCIAL = "financial"
    STATISTICAL = "statistical"
    ADVANCED = "advanced"
    TEXT = "text"
    DATE_TIME = "date_time"
    DATABASE = "database"
    INFORMATION = "information"


@dataclass
class ComplexityFactors:
    """Complexity factors for scoring"""
    formula_length: float
    nesting_depth: float
    function_complexity: float
    dependency_complexity: float
    reference_complexity: float
    special_features: float
    performance_impact: float


@dataclass
class ComplexityDistribution:
    """Workbook complexity distribution"""
    simple_count: int
    moderate_count: int
    complex_count: int
    critical_count: int
    total_formulas: int
    average_score: float
    max_score: float
    min_score: float
    variance: float


@dataclass
class ComplexityAnalysis:
    """Complete complexity analysis results"""
    complexity_score: float
    complexity_category: ComplexityCategory
    complexity_factors: ComplexityFactors
    performance_prediction: str
    maintenance_risk: str
    optimization_suggestions: List[str]


class ComplexityScorer:
    """Sophisticated formula complexity scoring system"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._init_function_categories()
        self._init_complexity_weights()
    
    def _init_function_categories(self):
        """Initialize function category mappings"""
        self.function_categories = {
            # Mathematical functions (weight: 0.8)
            FunctionCategory.MATHEMATICAL: {
                'SUM': 0.1, 'AVERAGE': 0.1, 'COUNT': 0.1, 'MAX': 0.1, 'MIN': 0.1,
                'ROUND': 0.2, 'ABS': 0.1, 'SQRT': 0.2, 'POWER': 0.3, 'MOD': 0.2,
                'PRODUCT': 0.2, 'SUBTOTAL': 0.3, 'AGGREGATE': 0.4, 'SUMPRODUCT': 0.6
            },
            
            # Logical functions (weight: 1.0)
            FunctionCategory.LOGICAL: {
                'IF': 0.3, 'AND': 0.2, 'OR': 0.2, 'NOT': 0.1, 'IFERROR': 0.3,
                'IFS': 0.4, 'SWITCH': 0.4, 'CHOOSE': 0.4, 'XOR': 0.3
            },
            
            # Lookup functions (weight: 1.4)
            FunctionCategory.LOOKUP: {
                'VLOOKUP': 0.6, 'HLOOKUP': 0.6, 'INDEX': 0.5, 'MATCH': 0.5,
                'XLOOKUP': 0.7, 'LOOKUP': 0.5, 'FILTER': 0.8, 'UNIQUE': 0.7,
                'SORT': 0.7, 'SORTBY': 0.8
            },
            
            # Financial functions (weight: 1.3)
            FunctionCategory.FINANCIAL: {
                'PMT': 0.5, 'PV': 0.5, 'FV': 0.5, 'RATE': 0.6, 'NPV': 0.6,
                'IRR': 0.7, 'XIRR': 0.8, 'XNPV': 0.7, 'MIRR': 0.7, 'SLN': 0.4
            },
            
            # Statistical functions (weight: 1.1)
            FunctionCategory.STATISTICAL: {
                'STDEV': 0.3, 'VAR': 0.3, 'MEDIAN': 0.3, 'MODE': 0.3, 'PERCENTILE': 0.4,
                'RANK': 0.4, 'CORREL': 0.5, 'SLOPE': 0.5, 'INTERCEPT': 0.5,
                'FORECAST': 0.6, 'TREND': 0.7, 'LINEST': 0.8
            },
            
            # Advanced functions (weight: 1.8)
            FunctionCategory.ADVANCED: {
                'INDIRECT': 0.9, 'OFFSET': 0.7, 'ARRAY': 0.8, 'EVALUATE': 1.0,
                'FORMULA': 0.9, 'SEQUENCE': 0.6, 'RANDARRAY': 0.7
            },
            
            # Text functions (weight: 0.9)
            FunctionCategory.TEXT: {
                'LEFT': 0.2, 'RIGHT': 0.2, 'MID': 0.3, 'LEN': 0.1, 'FIND': 0.3,
                'SEARCH': 0.3, 'SUBSTITUTE': 0.4, 'REPLACE': 0.4, 'CONCATENATE': 0.3,
                'TEXTJOIN': 0.5, 'REGEX': 0.7
            },
            
            # Date/Time functions (weight: 1.2)
            FunctionCategory.DATE_TIME: {
                'NOW': 0.8, 'TODAY': 0.8, 'DATE': 0.2, 'TIME': 0.2, 'YEAR': 0.1,
                'MONTH': 0.1, 'DAY': 0.1, 'WEEKDAY': 0.2, 'NETWORKDAYS': 0.4,
                'WORKDAY': 0.4, 'EDATE': 0.3, 'EOMONTH': 0.3
            },
            
            # Database functions (weight: 1.3)
            FunctionCategory.DATABASE: {
                'DSUM': 0.6, 'DCOUNT': 0.6, 'DAVERAGE': 0.6, 'DMAX': 0.6, 'DMIN': 0.6,
                'DGET': 0.7, 'DVAR': 0.6, 'DSTDEV': 0.6
            },
            
            # Information functions (weight: 0.7)
            FunctionCategory.INFORMATION: {
                'ISBLANK': 0.1, 'ISERROR': 0.1, 'ISNUMBER': 0.1, 'ISTEXT': 0.1,
                'CELL': 0.4, 'INFO': 0.3, 'TYPE': 0.2, 'FORMULA': 0.5
            }
        }
    
    def _init_complexity_weights(self):
        """Initialize complexity scoring weights"""
        self.weights = {
            'length': 0.15,           # Formula character length
            'nesting': 0.25,          # Function nesting depth
            'functions': 0.30,        # Function complexity
            'dependencies': 0.15,     # Dependency chain complexity
            'references': 0.10,       # Reference complexity
            'special_features': 0.05  # Array formulas, tables, etc.
        }
        
        self.category_weights = {
            FunctionCategory.MATHEMATICAL: 0.8,
            FunctionCategory.LOGICAL: 1.0,
            FunctionCategory.LOOKUP: 1.4,
            FunctionCategory.FINANCIAL: 1.3,
            FunctionCategory.STATISTICAL: 1.1,
            FunctionCategory.ADVANCED: 1.8,
            FunctionCategory.TEXT: 0.9,
            FunctionCategory.DATE_TIME: 1.2,
            FunctionCategory.DATABASE: 1.3,
            FunctionCategory.INFORMATION: 0.7
        }
    
    def calculate_complexity_score(self, parsed_formula: ParsedFormula, 
                                 dependency_chain_length: int = 0,
                                 execution_frequency: float = 1.0) -> ComplexityAnalysis:
        """Calculate comprehensive complexity score"""
        factors = self._calculate_complexity_factors(parsed_formula, dependency_chain_length)
        
        # Weighted total score
        total_score = (
            factors.formula_length * self.weights['length'] +
            factors.nesting_depth * self.weights['nesting'] +
            factors.function_complexity * self.weights['functions'] +
            factors.dependency_complexity * self.weights['dependencies'] +
            factors.reference_complexity * self.weights['references'] +
            factors.special_features * self.weights['special_features']
        )
        
        # Performance impact adjustment
        performance_multiplier = self._calculate_performance_impact(
            parsed_formula, execution_frequency
        )
        final_score = min(100.0, total_score * performance_multiplier)
        
        category = self._get_complexity_category(final_score)
        performance_prediction = self._predict_performance_impact(final_score, parsed_formula)
        maintenance_risk = self._assess_maintenance_risk(final_score, factors)
        optimization_suggestions = self._generate_optimization_suggestions(parsed_formula, factors)
        
        return ComplexityAnalysis(
            complexity_score=final_score,
            complexity_category=category,
            complexity_factors=factors,
            performance_prediction=performance_prediction,
            maintenance_risk=maintenance_risk,
            optimization_suggestions=optimization_suggestions
        )
    
    def calculate_workbook_distribution(self, formula_scores: List[float]) -> ComplexityDistribution:
        """Calculate complexity distribution across workbook"""
        if not formula_scores:
            return ComplexityDistribution(0, 0, 0, 0, 0, 0.0, 0.0, 0.0, 0.0)
        
        simple_count = sum(1 for score in formula_scores if score <= 25)
        moderate_count = sum(1 for score in formula_scores if 25 < score <= 50)
        complex_count = sum(1 for score in formula_scores if 50 < score <= 75)
        critical_count = sum(1 for score in formula_scores if score > 75)
        
        total_formulas = len(formula_scores)
        average_score = sum(formula_scores) / total_formulas
        max_score = max(formula_scores)
        min_score = min(formula_scores)
        
        # Calculate variance
        variance = sum((score - average_score) ** 2 for score in formula_scores) / total_formulas
        
        return ComplexityDistribution(
            simple_count=simple_count,
            moderate_count=moderate_count,
            complex_count=complex_count,
            critical_count=critical_count,
            total_formulas=total_formulas,
            average_score=average_score,
            max_score=max_score,
            min_score=min_score,
            variance=variance
        )
    
    def _calculate_complexity_factors(self, parsed_formula: ParsedFormula, 
                                    dependency_chain_length: int) -> ComplexityFactors:
        """Calculate individual complexity factors"""
        # Formula length score (0-100)
        length_score = min(100.0, (len(parsed_formula.original_formula) / 500.0) * 100)
        
        # Nesting depth score (0-100)
        max_nesting = max([f.nesting_level for f in parsed_formula.functions], default=0)
        nesting_score = min(100.0, (max_nesting / 10.0) * 100)
        
        # Function complexity score (0-100)
        function_score = self._calculate_function_complexity(parsed_formula.functions)
        
        # Dependency complexity (0-100)
        dependency_score = min(100.0, (dependency_chain_length / 20.0) * 100)
        
        # Reference complexity (0-100)
        reference_score = self._calculate_reference_complexity(parsed_formula)
        
        # Special features score (0-100)
        special_score = self._calculate_special_features_score(parsed_formula)
        
        return ComplexityFactors(
            formula_length=length_score,
            nesting_depth=nesting_score,
            function_complexity=function_score,
            dependency_complexity=dependency_score,
            reference_complexity=reference_score,
            special_features=special_score,
            performance_impact=0.0  # Calculated separately
        )
    
    def _calculate_function_complexity(self, functions: List[FormulaFunction]) -> float:
        """Calculate function complexity score"""
        if not functions:
            return 0.0
        
        total_complexity = 0.0
        for func in functions:
            # Base function complexity
            base_complexity = func.complexity_weight * 20
            
            # Category weight multiplier
            category = self._get_function_category(func.name)
            category_multiplier = self.category_weights.get(category, 1.0)
            
            # Nesting penalty
            nesting_penalty = func.nesting_level * 5
            
            # Parameter complexity
            param_complexity = min(func.parameter_count * 2, 20)
            
            func_complexity = (base_complexity * category_multiplier + 
                             nesting_penalty + param_complexity)
            total_complexity += func_complexity
        
        return min(100.0, total_complexity / len(functions))
    
    def _calculate_reference_complexity(self, parsed_formula: ParsedFormula) -> float:
        """Calculate reference complexity score"""
        base_score = len(parsed_formula.cell_references) * 2
        range_score = len(parsed_formula.ranges) * 5
        external_score = len(parsed_formula.external_references) * 15
        
        # Reference type complexity
        type_complexity = 0
        for ref in parsed_formula.cell_references:
            if ref.is_external:
                type_complexity += 10
            elif ref.sheet:
                type_complexity += 3
            else:
                type_complexity += 1
        
        total_score = base_score + range_score + external_score + type_complexity
        return min(100.0, total_score)
    
    def _calculate_special_features_score(self, parsed_formula: ParsedFormula) -> float:
        """Calculate special features complexity score"""
        score = 0.0
        
        if parsed_formula.is_array_formula:
            score += 25.0
        
        if parsed_formula.is_table_formula:
            score += 15.0
        
        # Volatile function penalty
        volatile_functions = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT', 'OFFSET'}
        has_volatile = any(f.name in volatile_functions for f in parsed_formula.functions)
        if has_volatile:
            score += 20.0
        
        return min(100.0, score)
    
    def _get_function_category(self, function_name: str) -> FunctionCategory:
        """Determine function category"""
        for category, functions in self.function_categories.items():
            if function_name in functions:
                return category
        return FunctionCategory.MATHEMATICAL  # Default
    
    def _calculate_performance_impact(self, parsed_formula: ParsedFormula, 
                                    execution_frequency: float) -> float:
        """Calculate performance impact multiplier"""
        base_multiplier = 1.0
        
        # Volatile functions impact
        volatile_functions = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'}
        has_volatile = any(f.name in volatile_functions for f in parsed_formula.functions)
        if has_volatile:
            base_multiplier += 0.3
        
        # Array formula impact
        if parsed_formula.is_array_formula:
            base_multiplier += 0.2
        
        # External reference impact
        if parsed_formula.external_references:
            base_multiplier += 0.4
        
        # Frequency multiplier
        frequency_multiplier = min(2.0, 1.0 + (execution_frequency / 100.0))
        
        return base_multiplier * frequency_multiplier
    
    def _get_complexity_category(self, score: float) -> ComplexityCategory:
        """Convert score to complexity category"""
        if score <= 25:
            return ComplexityCategory.SIMPLE
        elif score <= 50:
            return ComplexityCategory.MODERATE
        elif score <= 75:
            return ComplexityCategory.COMPLEX
        else:
            return ComplexityCategory.CRITICAL
    
    def _predict_performance_impact(self, score: float, parsed_formula: ParsedFormula) -> str:
        """Predict performance impact based on complexity"""
        if score <= 25:
            return "Minimal performance impact. Fast execution expected."
        elif score <= 50:
            return "Low performance impact. Acceptable execution time."
        elif score <= 75:
            return "Moderate performance impact. May slow calculations."
        else:
            volatile_count = sum(1 for f in parsed_formula.functions 
                               if f.name in ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'])
            if volatile_count > 0:
                return "High performance impact. Volatile functions cause frequent recalculation."
            else:
                return "High performance impact. Complex calculations may cause delays."
    
    def _assess_maintenance_risk(self, score: float, factors: ComplexityFactors) -> str:
        """Assess maintenance risk level"""
        risk_factors = []
        
        if factors.nesting_depth > 60:
            risk_factors.append("deep nesting")
        if factors.function_complexity > 70:
            risk_factors.append("complex functions")
        if factors.dependency_complexity > 50:
            risk_factors.append("long dependency chains")
        
        if score <= 25:
            return "Low maintenance risk. Easy to understand and modify."
        elif score <= 50:
            base_risk = "Moderate maintenance risk."
        elif score <= 75:
            base_risk = "High maintenance risk."
        else:
            base_risk = "Critical maintenance risk."
        
        if risk_factors:
            return f"{base_risk} Key concerns: {', '.join(risk_factors)}."
        return base_risk
    
    def _generate_optimization_suggestions(self, parsed_formula: ParsedFormula, 
                                         factors: ComplexityFactors) -> List[str]:
        """Generate optimization recommendations"""
        suggestions = []
        
        if factors.nesting_depth > 60:
            suggestions.append("Break down nested functions into intermediate cells")
        
        if factors.function_complexity > 70:
            suggestions.append("Consider using helper columns for complex calculations")
        
        if len(parsed_formula.external_references) > 0:
            suggestions.append("Minimize external references for better performance")
        
        # Check for volatile functions
        volatile_functions = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'}
        has_volatile = any(f.name in volatile_functions for f in parsed_formula.functions)
        if has_volatile:
            suggestions.append("Reduce volatile functions usage or calculate values manually")
        
        if parsed_formula.is_array_formula and factors.formula_length > 60:
            suggestions.append("Consider breaking array formulas into smaller components")
        
        if len(parsed_formula.ranges) > 3:
            suggestions.append("Consolidate multiple ranges into named ranges for clarity")
        
        return suggestions


def create_complexity_scorer() -> ComplexityScorer:
    """Create ComplexityScorer instance"""
    return ComplexityScorer()
