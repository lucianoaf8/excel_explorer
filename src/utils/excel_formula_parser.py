"""
Excel Formula Parser - Comprehensive formula parsing with AST generation
Handles all Excel formula formats including cell references, functions, and external workbook links.
"""

import re
import logging
from typing import List, Dict, Any, Optional, Tuple, Set, Union
from enum import Enum
from dataclasses import dataclass
from abc import ABC, abstractmethod


class ReferenceType(Enum):
    """Types of cell references"""
    RELATIVE = "relative"  # A1
    ABSOLUTE = "absolute"  # $A$1
    MIXED_COLUMN = "mixed_col"  # $A1
    MIXED_ROW = "mixed_row"  # A$1


class FormulaComplexity(Enum):
    """Formula complexity levels"""
    SIMPLE = "simple"  # 0-25
    MODERATE = "moderate"  # 26-50
    COMPLEX = "complex"  # 51-75
    CRITICAL = "critical"  # 76-100


@dataclass
class CellReference:
    """Represents a cell reference with metadata"""
    sheet: Optional[str]
    column: str
    row: int
    reference_type: ReferenceType
    workbook: Optional[str] = None
    is_external: bool = False
    original_text: str = ""


@dataclass
class FormulaFunction:
    """Represents a function call in a formula"""
    name: str
    parameters: List[str]
    parameter_count: int
    nesting_level: int
    complexity_weight: float
    start_pos: int
    end_pos: int


@dataclass
class ParsedFormula:
    """Complete parsed formula with all components"""
    original_formula: str
    cell_references: List[CellReference]
    functions: List[FormulaFunction]
    ranges: List[str]
    external_references: List[str]
    is_array_formula: bool
    is_table_formula: bool
    complexity_score: float
    complexity_level: FormulaComplexity
    parsing_errors: List[str] = None


class ExcelFormulaParser:
    """Comprehensive Excel formula parser with AST generation"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._init_patterns()
        self._init_function_weights()
    
    def _init_patterns(self):
        """Initialize regex patterns for formula parsing"""
        self.patterns = {
            'cell_ref': re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,7})'),
            'sheet_ref': re.compile(r"(?:'([^']+)'|([A-Za-z0-9_]+))!(\$?)([A-Z]{1,3})(\$?)(\d{1,7})"),
            'external_ref': re.compile(r"\[([^\]]+)\](?:'([^']+)'|([A-Za-z0-9_]+))!(\$?)([A-Z]{1,3})(\$?)(\d{1,7})"),
            'range_ref': re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Z]{1,3})(\$?)(\d{1,7})'),
            'function_call': re.compile(r'([A-Z][A-Z0-9_]*)\s*\('),
            'array_formula': re.compile(r'^\{=.*\}$'),
            'table_ref': re.compile(r'([A-Za-z0-9_]+)\[(@?)([A-Za-z0-9_\s]+)\]'),
            'string_literal': re.compile(r'"([^"]*)"'),
            'number_literal': re.compile(r'-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?'),
            'operators': re.compile(r'(\+|\-|\*|\/|\^|&|=|<>|<=|>=|<|>)')
        }
    
    def _init_function_weights(self):
        """Initialize complexity weights for Excel functions"""
        self.function_weights = {
            # Mathematical functions
            'SUM': 0.1, 'AVERAGE': 0.1, 'COUNT': 0.1, 'MAX': 0.1, 'MIN': 0.1,
            'ROUND': 0.2, 'ABS': 0.1, 'SQRT': 0.2, 'POWER': 0.3, 'MOD': 0.2,
            
            # Logical functions
            'IF': 0.3, 'AND': 0.2, 'OR': 0.2, 'NOT': 0.1, 'IFERROR': 0.3,
            'IFS': 0.4, 'SWITCH': 0.4,
            
            # Lookup functions
            'VLOOKUP': 0.6, 'HLOOKUP': 0.6, 'INDEX': 0.5, 'MATCH': 0.5,
            'XLOOKUP': 0.7, 'LOOKUP': 0.5, 'CHOOSE': 0.4,
            
            # Text functions
            'LEFT': 0.2, 'RIGHT': 0.2, 'MID': 0.3, 'LEN': 0.1, 'FIND': 0.3,
            'SEARCH': 0.3, 'SUBSTITUTE': 0.4, 'REPLACE': 0.4, 'CONCATENATE': 0.3,
            'TEXTJOIN': 0.5,
            
            # Date/Time functions
            'NOW': 0.8, 'TODAY': 0.8, 'DATE': 0.2, 'TIME': 0.2, 'YEAR': 0.1,
            'MONTH': 0.1, 'DAY': 0.1, 'WEEKDAY': 0.2, 'NETWORKDAYS': 0.4,
            
            # Statistical functions
            'STDEV': 0.3, 'VAR': 0.3, 'MEDIAN': 0.3, 'MODE': 0.3, 'PERCENTILE': 0.4,
            'RANK': 0.4, 'CORREL': 0.5, 'REGRESSION': 0.7,
            
            # Financial functions
            'PMT': 0.5, 'PV': 0.5, 'FV': 0.5, 'RATE': 0.6, 'NPV': 0.6,
            'IRR': 0.7, 'XIRR': 0.8, 'XNPV': 0.7,
            
            # Advanced functions
            'SUMPRODUCT': 0.6, 'SUMIFS': 0.5, 'COUNTIFS': 0.5, 'AVERAGEIFS': 0.5,
            'INDIRECT': 0.9, 'OFFSET': 0.7, 'ARRAY': 0.8,
            
            # Database functions
            'DSUM': 0.6, 'DCOUNT': 0.6, 'DAVERAGE': 0.6, 'DMAX': 0.6, 'DMIN': 0.6,
            
            # Information functions
            'ISBLANK': 0.1, 'ISERROR': 0.1, 'ISNUMBER': 0.1, 'ISTEXT': 0.1,
            'CELL': 0.4, 'INFO': 0.3
        }
    
    def parse_formula(self, formula: str, cell_address: str = None) -> ParsedFormula:
        """Parse an Excel formula into its components"""
        try:
            if not formula.startswith('='):
                formula = '=' + formula
            
            clean_formula = formula.strip()
            
            result = ParsedFormula(
                original_formula=clean_formula,
                cell_references=[],
                functions=[],
                ranges=[],
                external_references=[],
                is_array_formula=False,
                is_table_formula=False,
                complexity_score=0.0,
                complexity_level=FormulaComplexity.SIMPLE,
                parsing_errors=[]
            )
            
            # Check for array formula
            result.is_array_formula = bool(self.patterns['array_formula'].match(clean_formula))
            
            # Parse components
            result.cell_references = self._parse_cell_references(clean_formula)
            result.functions = self._parse_functions(clean_formula)
            result.ranges = self._parse_ranges(clean_formula)
            result.external_references = self._parse_external_references(clean_formula)
            result.is_table_formula = bool(self.patterns['table_ref'].search(clean_formula))
            
            # Calculate complexity
            result.complexity_score = self._calculate_complexity_score(result)
            result.complexity_level = self._get_complexity_level(result.complexity_score)
            
            return result
            
        except Exception as e:
            self.logger.error(f"Formula parsing failed for '{formula}': {str(e)}")
            return ParsedFormula(
                original_formula=formula,
                cell_references=[], functions=[], ranges=[], external_references=[],
                is_array_formula=False, is_table_formula=False,
                complexity_score=0.0, complexity_level=FormulaComplexity.SIMPLE,
                parsing_errors=[str(e)]
            )
    
    def _parse_cell_references(self, formula: str) -> List[CellReference]:
        """Parse all cell references in the formula"""
        references = []
        
        # Parse external references first
        for match in self.patterns['external_ref'].finditer(formula):
            workbook = match.group(1)
            sheet = match.group(2) or match.group(3)
            col_abs = match.group(4)
            column = match.group(5)
            row_abs = match.group(6)
            row = int(match.group(7))
            
            ref_type = self._determine_reference_type(col_abs, row_abs)
            
            references.append(CellReference(
                sheet=sheet, column=column, row=row, reference_type=ref_type,
                workbook=workbook, is_external=True, original_text=match.group(0)
            ))
        
        # Parse sheet references
        for match in self.patterns['sheet_ref'].finditer(formula):
            if any(ext_ref.original_text in match.group(0) for ext_ref in references):
                continue
                
            sheet = match.group(1) or match.group(2)
            col_abs = match.group(3)
            column = match.group(4)
            row_abs = match.group(5)
            row = int(match.group(6))
            
            ref_type = self._determine_reference_type(col_abs, row_abs)
            
            references.append(CellReference(
                sheet=sheet, column=column, row=row, reference_type=ref_type,
                is_external=False, original_text=match.group(0)
            ))
        
        # Parse basic cell references
        existing_refs = {ref.original_text for ref in references}
        for match in self.patterns['cell_ref'].finditer(formula):
            if match.group(0) not in existing_refs:
                col_abs = match.group(1)
                column = match.group(2)
                row_abs = match.group(3)
                row = int(match.group(4))
                
                ref_type = self._determine_reference_type(col_abs, row_abs)
                
                references.append(CellReference(
                    sheet=None, column=column, row=row, reference_type=ref_type,
                    is_external=False, original_text=match.group(0)
                ))
        
        return references
    
    def _parse_functions(self, formula: str) -> List[FormulaFunction]:
        """Parse all function calls in the formula"""
        functions = []
        
        for match in self.patterns['function_call'].finditer(formula):
            func_name = match.group(1).upper()
            start_pos = match.start()
            
            # Find matching closing parenthesis
            paren_count = 1
            pos = match.end()
            while pos < len(formula) and paren_count > 0:
                if formula[pos] == '(':
                    paren_count += 1
                elif formula[pos] == ')':
                    paren_count -= 1
                pos += 1
            
            if paren_count == 0:
                end_pos = pos
                func_content = formula[match.end():end_pos-1]
                parameters = self._parse_function_parameters(func_content)
                nesting_level = formula[:start_pos].count('(') - formula[:start_pos].count(')')
                complexity_weight = self.function_weights.get(func_name, 0.5)
                
                functions.append(FormulaFunction(
                    name=func_name, parameters=parameters, parameter_count=len(parameters),
                    nesting_level=max(0, nesting_level), complexity_weight=complexity_weight,
                    start_pos=start_pos, end_pos=end_pos
                ))
        
        return functions
    
    def _parse_function_parameters(self, param_string: str) -> List[str]:
        """Parse function parameters from parameter string"""
        if not param_string.strip():
            return []
        
        parameters = []
        current_param = ""
        paren_depth = 0
        in_quotes = False
        
        for char in param_string:
            if char == '"' and (not current_param or current_param[-1] != '\\'):
                in_quotes = not in_quotes
            elif not in_quotes:
                if char == '(':
                    paren_depth += 1
                elif char == ')':
                    paren_depth -= 1
                elif char == ',' and paren_depth == 0:
                    parameters.append(current_param.strip())
                    current_param = ""
                    continue
            
            current_param += char
        
        if current_param.strip():
            parameters.append(current_param.strip())
        
        return parameters
    
    def _parse_ranges(self, formula: str) -> List[str]:
        """Parse cell ranges in the formula"""
        ranges = []
        for match in self.patterns['range_ref'].finditer(formula):
            ranges.append(match.group(0))
        return list(set(ranges))
    
    def _parse_external_references(self, formula: str) -> List[str]:
        """Parse external workbook references"""
        external_refs = []
        for match in self.patterns['external_ref'].finditer(formula):
            workbook = match.group(1)
            if workbook not in external_refs:
                external_refs.append(workbook)
        return external_refs
    
    def _determine_reference_type(self, col_abs: str, row_abs: str) -> ReferenceType:
        """Determine the type of cell reference"""
        if col_abs and row_abs:
            return ReferenceType.ABSOLUTE
        elif col_abs and not row_abs:
            return ReferenceType.MIXED_COLUMN
        elif not col_abs and row_abs:
            return ReferenceType.MIXED_ROW
        else:
            return ReferenceType.RELATIVE
    
    def _calculate_complexity_score(self, parsed: ParsedFormula) -> float:
        """Calculate complexity score for the parsed formula"""
        # Base complexity from formula length
        length_score = min(len(parsed.original_formula) / 200.0, 1.0) * 20
        
        # Function complexity
        function_score = sum(f.complexity_weight * (1 + f.nesting_level * 0.2) 
                           for f in parsed.functions) * 15
        
        # Reference complexity
        ref_score = len(parsed.cell_references) * 2
        external_score = len(parsed.external_references) * 10
        
        # Special formula types
        array_score = 15 if parsed.is_array_formula else 0
        table_score = 5 if parsed.is_table_formula else 0
        
        # Range complexity
        range_score = len(parsed.ranges) * 3
        
        total_score = (length_score + function_score + ref_score + 
                      external_score + array_score + table_score + range_score)
        
        return min(100.0, max(0.0, total_score))
    
    def _get_complexity_level(self, score: float) -> FormulaComplexity:
        """Convert complexity score to complexity level"""
        if score <= 25:
            return FormulaComplexity.SIMPLE
        elif score <= 50:
            return FormulaComplexity.MODERATE
        elif score <= 75:
            return FormulaComplexity.COMPLEX
        else:
            return FormulaComplexity.CRITICAL
    
    def convert_reference_format(self, reference: str, target_type: ReferenceType) -> str:
        """Convert cell reference between different formats"""
        match = self.patterns['cell_ref'].match(reference)
        if not match:
            return reference
        
        col_abs, column, row_abs, row = match.groups()
        
        if target_type == ReferenceType.ABSOLUTE:
            return f"${column}${row}"
        elif target_type == ReferenceType.RELATIVE:
            return f"{column}{row}"
        elif target_type == ReferenceType.MIXED_COLUMN:
            return f"${column}{row}"
        elif target_type == ReferenceType.MIXED_ROW:
            return f"{column}${row}"
        
        return reference
    
    def extract_workbook_references(self, formula: str) -> List[str]:
        """Extract all external workbook references"""
        return self._parse_external_references(formula)
    
    def is_volatile_formula(self, formula: str) -> bool:
        """Check if formula contains volatile functions"""
        volatile_functions = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT', 'OFFSET'}
        formula_upper = formula.upper()
        return any(func in formula_upper for func in volatile_functions)
    
    def validate_formula_syntax(self, formula: str) -> Tuple[bool, List[str]]:
        """Validate basic formula syntax"""
        errors = []
        
        if not formula.startswith('='):
            errors.append("Formula must start with '='")
        
        # Check parentheses balance
        open_count = formula.count('(')
        close_count = formula.count(')')
        if open_count != close_count:
            errors.append(f"Unbalanced parentheses: {open_count} open, {close_count} close")
        
        # Check quotes balance
        quote_count = formula.count('"')
        if quote_count % 2 != 0:
            errors.append("Unbalanced quotes")
        
        # Check for common syntax errors
        if ',,,' in formula:
            errors.append("Multiple consecutive commas")
        
        if formula.endswith(','):
            errors.append("Formula ends with comma")
        
        return len(errors) == 0, errors


def create_formula_parser() -> ExcelFormulaParser:
    """Create a new ExcelFormulaParser instance"""
    return ExcelFormulaParser()
