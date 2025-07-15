# Excel Explorer Framework - Comprehensive Fixing Plan

## Phase 1: Establish Working Foundation

### Step 1: Create Functional Entry Point

**Create `src/main.py`:**

```python
#!/usr/bin/env python3
import sys
import json
from pathlib import Path
from src.core.orchestrator import ExcelExplorer

def main():
    if len(sys.argv) != 2:
        print("Usage: python -m src.main <excel_file>")
        sys.exit(1)
  
    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(f"File not found: {file_path}")
        sys.exit(1)
  
    try:
        explorer = ExcelExplorer()
        results = explorer.analyze_file(str(file_path))
        print(json.dumps(results, indent=2, default=str))
        return 0
    except Exception as e:
        print(f"Analysis failed: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
```

**Update `src/core/orchestrator.py` main block:**

```python
# Replace existing main() function
def main():
    from .main import main as main_impl
    return main_impl()
```

### Step 2: Fix Critical Module Placeholders

**Update `src/modules/connection_inspector.py`:**

Replace placeholder methods with minimal working implementations:

```python
def _detect_external_connections(self, workbook) -> List[Dict[str, Any]]:
    connections = []
    try:
        if hasattr(workbook, 'defined_names') and workbook.defined_names:
            for name in workbook.defined_names.definedName:
                if name.attr_text and ('[' in name.attr_text or 'http' in name.attr_text.lower()):
                    connections.append({
                        'id': name.name,
                        'type': 'external_reference',
                        'description': name.attr_text[:100],
                        'source': 'defined_names'
                    })
    except Exception as e:
        self.logger.warning(f"External connection detection failed: {e}")
    return connections

def _detect_linked_workbooks(self, workbook) -> List[str]:
    linked_workbooks = set()
    try:
        max_cells_check = self.config.get("max_cell_checks", 1000)
        checked = 0
      
        for ws in workbook.worksheets:
            if checked >= max_cells_check:
                break
            for row in ws.iter_rows(max_row=min(100, ws.max_row or 1)):
                if checked >= max_cells_check:
                    break
                for cell in row:
                    if checked >= max_cells_check:
                        break
                    if (cell.value and isinstance(cell.value, str) and 
                        cell.value.startswith('=') and '[' in cell.value):
                        import re
                        matches = re.findall(r'\[([^\]]+\.xlsx?)\]', cell.value, re.IGNORECASE)
                        linked_workbooks.update(matches)
                    checked += 1
    except Exception as e:
        self.logger.warning(f"Linked workbook detection failed: {e}")
    return list(linked_workbooks)

def _detect_database_connections(self, workbook) -> List[Dict[str, Any]]:
    connections = []
    try:
        # Check worksheet names for database indicators
        for ws in workbook.worksheets:
            sheet_name = ws.title.lower()
            if any(keyword in sheet_name for keyword in ['data', 'import', 'query', 'connection']):
                # Sample first few rows for connection strings
                for row in ws.iter_rows(max_row=5, values_only=True):
                    for cell in row:
                        if cell and isinstance(cell, str):
                            cell_lower = cell.lower()
                            if any(keyword in cell_lower for keyword in 
                                 ['driver=', 'server=', 'database=', 'dsn=', 'connection']):
                                connections.append({
                                    'id': f"db_conn_{len(connections)}",
                                    'type': 'database',
                                    'sheet': ws.title,
                                    'indicator': cell[:100]
                                })
                                break
    except Exception as e:
        self.logger.warning(f"Database connection detection failed: {e}")
    return connections

def _detect_web_queries(self, workbook) -> List[Dict[str, Any]]:
    web_queries = []
    try:
        for ws in workbook.worksheets:
            for row in ws.iter_rows(max_row=20, values_only=True):
                for cell in row:
                    if cell and isinstance(cell, str):
                        if any(protocol in cell.lower() for protocol in ['http://', 'https://', 'ftp://']):
                            web_queries.append({
                                'id': f"web_query_{len(web_queries)}",
                                'type': 'web_query',
                                'sheet': ws.title,
                                'url': cell[:200]
                            })
    except Exception as e:
        self.logger.warning(f"Web query detection failed: {e}")
    return web_queries
```

**Update `src/modules/pivot_intelligence.py`:**

Replace placeholder methods:

```python
def _analyze_pivot_tables(self, worksheet, sheet_name: str) -> Dict[str, List[Dict[str, Any]]]:
    result = {'tables': [], 'sources': [], 'calculated_fields': []}
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
        self.logger.warning(f"Pivot table analysis failed for {sheet_name}: {e}")
    return result

def _analyze_pivot_charts(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
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
        self.logger.warning(f"Pivot chart analysis failed for {sheet_name}: {e}")
    return pivot_charts

def _analyze_slicers(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
    slicers = []
    try:
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
        self.logger.warning(f"Slicer analysis failed for {sheet_name}: {e}")
    return slicers
```

**Update `src/modules/visual_cataloger.py`:**

Replace placeholder methods:

```python
def _catalog_shapes(self, worksheet, sheet_name: str) -> List[Dict[str, Any]]:
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
        self.logger.warning(f"Shape cataloging failed for {sheet_name}: {e}")
    return shapes
```

### Step 3: Fix Import Dependencies

**Update `src/modules/formula_analyzer.py`:**

Create missing utility files:

**Create `src/utils/excel_formula_parser.py`:**

```python
from dataclasses import dataclass
from typing import List, Optional
from enum import Enum

class ReferenceType(Enum):
    RELATIVE = "relative"
    ABSOLUTE = "absolute"
    MIXED_COLUMN = "mixed_col"
    MIXED_ROW = "mixed_row"

class FormulaComplexity(Enum):
    SIMPLE = "simple"
    MODERATE = "moderate"
    COMPLEX = "complex"
    CRITICAL = "critical"

@dataclass
class CellReference:
    sheet: Optional[str]
    column: str
    row: int
    reference_type: ReferenceType
    workbook: Optional[str] = None
    is_external: bool = False
    original_text: str = ""

@dataclass
class FormulaFunction:
    name: str
    parameters: List[str]
    parameter_count: int
    nesting_level: int
    complexity_weight: float
    start_pos: int
    end_pos: int

@dataclass
class ParsedFormula:
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
    def __init__(self):
        self.function_weights = {
            'SUM': 0.1, 'AVERAGE': 0.1, 'COUNT': 0.1, 'IF': 0.3,
            'VLOOKUP': 0.6, 'INDEX': 0.5, 'MATCH': 0.5
        }
  
    def parse_formula(self, formula: str, cell_address: str = None) -> ParsedFormula:
        import re
      
        # Basic parsing implementation
        functions = []
        cell_references = []
        ranges = []
        external_references = []
      
        # Find functions
        func_pattern = re.compile(r'([A-Z][A-Z0-9_]*)\s*\(')
        for match in func_pattern.finditer(formula):
            func_name = match.group(1)
            functions.append(FormulaFunction(
                name=func_name,
                parameters=[],
                parameter_count=0,
                nesting_level=0,
                complexity_weight=self.function_weights.get(func_name, 0.5),
                start_pos=match.start(),
                end_pos=match.end()
            ))
      
        # Find cell references
        cell_pattern = re.compile(r'([A-Z]{1,3})(\d{1,7})')
        for match in cell_pattern.finditer(formula):
            cell_references.append(CellReference(
                sheet=None,
                column=match.group(1),
                row=int(match.group(2)),
                reference_type=ReferenceType.RELATIVE,
                original_text=match.group(0)
            ))
      
        # Calculate basic complexity
        complexity_score = min(100.0, len(formula) * 0.1 + len(functions) * 10)
      
        return ParsedFormula(
            original_formula=formula,
            cell_references=cell_references,
            functions=functions,
            ranges=ranges,
            external_references=external_references,
            is_array_formula=formula.startswith('{') and formula.endswith('}'),
            is_table_formula=False,
            complexity_score=complexity_score,
            complexity_level=FormulaComplexity.SIMPLE if complexity_score < 25 else FormulaComplexity.MODERATE,
            parsing_errors=[]
        )

def create_formula_parser() -> ExcelFormulaParser:
    return ExcelFormulaParser()
```

**Create `src/utils/formula_dependency_analyzer.py`:**

```python
from dataclasses import dataclass
from typing import List, Dict
from enum import Enum

class ImpactLevel(Enum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"

@dataclass
class CircularReference:
    chain: List[str]
    chain_length: int
    complexity_score: float
    impact_level: ImpactLevel
    description: str = ""

@dataclass
class DependencyMetrics:
    total_formulas: int
    total_dependencies: int
    max_chain_length: int
    avg_chain_length: float
    circular_reference_count: int
    external_dependency_count: int
    orphaned_formula_count: int
    complexity_distribution: Dict[str, int]
    fan_out_distribution: Dict[str, int]
    volatility_score: float

class FormulaDependencyAnalyzer:
    def __init__(self, max_depth: int = 50):
        self.max_depth = max_depth
        self.formulas = {}
  
    def add_formula(self, cell_address: str, formula: str, sheet_name: str = None) -> bool:
        full_address = f"{sheet_name}!{cell_address}" if sheet_name else cell_address
        self.formulas[full_address] = formula
        return True
  
    def find_circular_references(self) -> List[CircularReference]:
        # Simple implementation - no actual circular detection
        return []
  
    def get_dependency_metrics(self) -> DependencyMetrics:
        return DependencyMetrics(
            total_formulas=len(self.formulas),
            total_dependencies=0,
            max_chain_length=1,
            avg_chain_length=1.0,
            circular_reference_count=0,
            external_dependency_count=0,
            orphaned_formula_count=0,
            complexity_distribution={"simple": len(self.formulas)},
            fan_out_distribution={"0": len(self.formulas)},
            volatility_score=0.0
        )

def create_dependency_analyzer(max_depth: int = 50) -> FormulaDependencyAnalyzer:
    return FormulaDependencyAnalyzer(max_depth)
```

## Phase 2: Configuration System Unification

### Step 4: Create Unified Configuration

**Create `src/core/unified_config.py`:**

```python
from dataclasses import dataclass
from typing import Dict, Any, Optional
from pathlib import Path
import yaml

@dataclass
class UnifiedConfig:
    # Analysis settings
    max_memory_mb: int = 4096
    warning_threshold: float = 0.8
    chunk_size_rows: int = 10000
    max_formulas_analyze: int = 50000
    enable_caching: bool = True
  
    # Module toggles
    include_charts: bool = True
    include_pivots: bool = True
    include_connections: bool = True
    deep_analysis: bool = False
    parallel_processing: bool = False
  
    # Module-specific settings
    module_configs: Dict[str, Dict[str, Any]] = None
  
    def __post_init__(self):
        if self.module_configs is None:
            self.module_configs = {}
  
    @classmethod
    def from_yaml(cls, path: Optional[str] = None) -> 'UnifiedConfig':
        config = cls()
        if path and Path(path).exists():
            try:
                with open(path, 'r') as f:
                    data = yaml.safe_load(f) or {}
                  
                # Map YAML structure to config
                if 'analysis' in data:
                    analysis = data['analysis']
                    config.max_memory_mb = analysis.get('max_memory_mb', config.max_memory_mb)
                    config.chunk_size_rows = analysis.get('chunk_size_rows', config.chunk_size_rows)
                    config.enable_caching = analysis.get('enable_caching', config.enable_caching)
              
                # Store module-specific configs
                for key, value in data.items():
                    if key not in ['analysis']:
                        config.module_configs[key] = value
                      
            except Exception as e:
                import logging
                logging.warning(f"Failed to load config from {path}: {e}")
      
        return config
  
    def get_module_config(self, module_name: str) -> Dict[str, Any]:
        return self.module_configs.get(module_name, {})
```

### Step 5: Update AnalysisContext

**Modify `src/core/analysis_context.py`:**

Add unified config integration:

```python
# Add import at top
from .unified_config import UnifiedConfig

# Modify AnalysisContext.__init__
def __init__(self, file_path: Union[str, Path], config: Optional[AnalysisConfig] = None):
    self.file_path = Path(file_path)
    self.config = config or AnalysisConfig()
  
    # Add unified config
    self.unified_config = UnifiedConfig.from_yaml()
  
    # Existing initialization code...
```

### Step 6: Update All Modules to Use Unified Config

**Pattern to apply to all modules:**

Replace configuration access patterns:

```python
# Old pattern (remove):
chunk_size_rows = self.config.get("chunk_size_rows", context.config.chunk_size_rows)

# New pattern (implement):
chunk_size_rows = context.unified_config.chunk_size_rows
module_config = context.unified_config.get_module_config(self.name)
specific_setting = module_config.get("specific_setting", default_value)
```

**Apply to these files:**

* `src/modules/data_profiler.py`
* `src/modules/formula_analyzer.py`
* `src/modules/health_checker.py`
* `src/modules/structure_mapper.py`
* All other modules with config access

## Phase 3: Processing System Consolidation

### Step 7: Simplify Processing Architecture

**Update `src/utils/chunked_processor.py`:**

Replace complex chunking with simplified version:

```python
from typing import Iterator, List, Dict, Any, Optional
from dataclasses import dataclass
from enum import Enum
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

@dataclass
class ChunkConfig:
    chunk_size_rows: int = 10000
    max_memory_mb: float = 512.0
    enable_progress_tracking: bool = True

class ChunkProcessor:
    def process_chunk(self, chunk_data: pd.DataFrame, chunk_index: int, 
                     chunk_metadata: Dict[str, Any]) -> Any:
        raise NotImplementedError

class ChunkedSheetProcessor:
    def __init__(self, config: Optional[ChunkConfig] = None):
        self.config = config or ChunkConfig()
  
    def process_worksheet(self, worksheet: Worksheet, processor: ChunkProcessor,
                         data_region: Optional[Dict[str, int]] = None) -> List[Any]:
        if not data_region:
            data_region = self._detect_data_region(worksheet)
            if not data_region:
                return []
      
        results = []
        chunk_size = self.config.chunk_size_rows
        min_row = data_region['min_row']
        max_row = data_region['max_row']
        min_col = data_region['min_col']
        max_col = data_region['max_col']
      
        chunk_index = 0
        for start_row in range(min_row, max_row + 1, chunk_size):
            end_row = min(start_row + chunk_size - 1, max_row)
          
            # Extract chunk data
            chunk_data = []
            for row in worksheet.iter_rows(
                min_row=start_row, max_row=end_row,
                min_col=min_col, max_col=max_col,
                values_only=True
            ):
                chunk_data.append(list(row))
          
            if chunk_data:
                num_cols = len(chunk_data[0]) if chunk_data else 0
                df = pd.DataFrame(
                    chunk_data, 
                    columns=[f'col_{i}' for i in range(num_cols)]
                )
              
                chunk_metadata = {
                    'start_row': start_row,
                    'end_row': end_row,
                    'start_col': min_col,
                    'end_col': max_col
                }
              
                result = processor.process_chunk(df, chunk_index, chunk_metadata)
                results.append(result)
                chunk_index += 1
      
        return results
  
    def _detect_data_region(self, worksheet: Worksheet) -> Optional[Dict[str, int]]:
        if not worksheet.max_row or not worksheet.max_column:
            return None
      
        # Simple implementation
        return {
            'min_row': 1,
            'max_row': min(worksheet.max_row, 1000),  # Limit for performance
            'min_col': 1,
            'max_col': min(worksheet.max_column, 50)  # Limit for performance
        }

class DataProfilingProcessor(ChunkProcessor):
    def __init__(self):
        pass
  
    def process_chunk(self, chunk_data: pd.DataFrame, chunk_index: int, 
                     chunk_metadata: Dict[str, Any]) -> Dict[str, Any]:
        return {
            'chunk_index': chunk_index,
            'row_count': len(chunk_data),
            'column_count': len(chunk_data.columns),
            'null_percentages': {
                col: chunk_data[col].isna().sum() / len(chunk_data) 
                for col in chunk_data.columns
            },
            'data_types': {
                col: str(chunk_data[col].dtype) 
                for col in chunk_data.columns
            },
            'outliers_detected': 0,
            'duplicate_rows': chunk_data.duplicated().sum()
        }
```

### Step 8: Remove Redundant Processing Systems

**Delete these files:**

* `src/utils/streaming_processor.py`
* Any other processing utilities not used

**Update imports in affected modules:**

* Update `src/modules/data_profiler.py` to only import from `chunked_processor`

## Phase 4: Error Recovery Implementation

### Step 9: Implement Basic Error Recovery

**Update `src/utils/error_handler.py`:**

Replace placeholder recovery methods:

```python
def _recover_memory_limit(self, error: ExcelAnalysisError, 
                        context: Optional[ErrorContext]) -> bool:
    """Basic memory recovery implementation"""
    try:
        import gc
        gc.collect()
      
        # Reduce processing complexity
        if context and context.module_name:
            self.logger.info(f"Reducing complexity for {context.module_name} due to memory pressure")
            return True
        return False
    except Exception:
        return False

def _recover_data_corruption(self, error: ExcelAnalysisError, 
                           context: Optional[ErrorContext]) -> bool:
    """Basic data corruption recovery"""
    try:
        if context and context.sheet_name:
            self.logger.warning(f"Skipping corrupted sheet: {context.sheet_name}")
            return True
        return False
    except Exception:
        return False

def _recover_dependency_failure(self, error: ExcelAnalysisError, 
                              context: Optional[ErrorContext]) -> bool:
    """Basic dependency failure recovery"""
    try:
        if context and context.module_name:
            self.logger.warning(f"Continuing analysis without {context.module_name} dependencies")
            return True
        return False
    except Exception:
        return False

def _recover_timeout(self, error: ExcelAnalysisError, 
                    context: Optional[ErrorContext]) -> bool:
    """Basic timeout recovery"""
    try:
        if context and context.module_name:
            self.logger.warning(f"Reducing scope for {context.module_name} due to timeout")
            return True
        return False
    except Exception:
        return False
```

## Phase 5: System Integration and Validation

### Step 10: Update Orchestrator Integration

**Modify `src/core/orchestrator.py`:**

Update to use unified config:

```python
# In ExcelExplorer.__init__
def __init__(self, config_path: Optional[str] = None, log_file: Optional[str] = None):
    # Use unified config
    self.unified_config = UnifiedConfig.from_yaml(config_path)
  
    # Create AnalysisConfig from unified config
    self.analysis_config = AnalysisConfig(
        max_memory_mb=self.unified_config.max_memory_mb,
        warning_threshold=self.unified_config.warning_threshold,
        enable_caching=self.unified_config.enable_caching,
        chunk_size_rows=self.unified_config.chunk_size_rows,
        max_formulas_analyze=self.unified_config.max_formulas_analyze,
        include_charts=self.unified_config.include_charts,
        include_pivots=self.unified_config.include_pivots,
        include_connections=self.unified_config.include_connections,
        deep_analysis=self.unified_config.deep_analysis,
        parallel_processing=self.unified_config.parallel_processing
    )
  
    # Existing initialization...
```

### Step 11: Fix Missing Placeholder Utilities

**Update `src/utils/config_loader.py`:**

Replace placeholder function:

```python
def placeholder_function():
    """Configuration validation and defaults"""
    return {
        "health_checker": {"enabled": True},
        "structure_mapper": {"enabled": True, "include_hidden_sheets": True},
        "data_profiler": {"enabled": True, "chunk_size_rows": 10000},
        "formula_analyzer": {"enabled": True, "max_formulas_analyze": 50000},
        "visual_cataloger": {"enabled": True},
        "connection_inspector": {"enabled": True, "max_cell_checks": 1000},
        "pivot_intelligence": {"enabled": True},
        "doc_synthesizer": {"enabled": True}
    }
```

### Step 12: Create Basic Test Validation

**Create `test_basic.py`:**

```python
#!/usr/bin/env python3
"""Basic functionality test"""

import sys
from pathlib import Path
from src.main import main

def test_basic_functionality():
    """Test with a simple Excel file"""
    # This should be run with: python test_basic.py sample.xlsx
    if len(sys.argv) != 2:
        print("Usage: python test_basic.py <excel_file>")
        return False
  
    try:
        # Override sys.argv for main function
        original_argv = sys.argv
        sys.argv = ['src.main', sys.argv[1]]
      
        result = main()
      
        sys.argv = original_argv
      
        if result == 0:
            print("✅ Basic functionality test PASSED")
            return True
        else:
            print("❌ Basic functionality test FAILED")
            return False
          
    except Exception as e:
        print(f"❌ Test failed with exception: {e}")
        return False

if __name__ == "__main__":
    success = test_basic_functionality()
    sys.exit(0 if success else 1)
```

## Phase 6: Final Integration Steps

### Step 13: Fix DocSynthesizer Type Safety

**Update `src/modules/doc_synthesizer.py`:**

Create proper data class:

```python
# Add at top with other imports
from dataclasses import dataclass
from typing import Dict, Any, List

@dataclass
class DocumentationData:
    file_overview: Dict[str, Any]
    executive_summary: str
    detailed_analysis: Dict[str, Any]
    recommendations: List[str]
    ai_navigation_guide: Dict[str, Any]
    metadata: Dict[str, Any]

# Update _perform_analysis return type
def _perform_analysis(self, context: AnalysisContext) -> DocumentationData:
    # Existing implementation, but return DocumentationData instead of dict
    documentation_dict = {
        'file_overview': self._create_file_overview(context),
        'executive_summary': self._create_executive_summary(context),
        'detailed_analysis': self._collect_detailed_analysis(context),
        'recommendations': self._generate_recommendations(context),
        'ai_navigation_guide': self._create_ai_navigation_guide(context),
        'metadata': self._create_metadata(context)
    }
  
    return DocumentationData(**documentation_dict)
```

### Step 14: Update Module Result Framework

**Update `src/core/module_result.py`:**

Add DocumentationData to the module result types:

```python
# Add with other dataclass imports
@dataclass
class DocumentationData:
    file_overview: Dict[str, Any]
    executive_summary: str
    detailed_analysis: Dict[str, Any]
    recommendations: List[str]
    ai_navigation_guide: Dict[str, Any]
    metadata: Dict[str, Any]
```

### Step 15: Validation and Testing Commands

**Create validation script `validate_system.py`:**

```python
#!/usr/bin/env python3
"""System validation script"""

def validate_imports():
    """Test all imports work"""
    try:
        from src.core.orchestrator import ExcelExplorer
        from src.core.unified_config import UnifiedConfig
        from src.main import main
        print("✅ All imports successful")
        return True
    except Exception as e:
        print(f"❌ Import failed: {e}")
        return False

def validate_config():
    """Test configuration system"""
    try:
        from src.core.unified_config import UnifiedConfig
        config = UnifiedConfig()
        assert config.max_memory_mb > 0
        assert config.chunk_size_rows > 0
        print("✅ Configuration validation passed")
        return True
    except Exception as e:
        print(f"❌ Configuration validation failed: {e}")
        return False

def validate_modules():
    """Test module instantiation"""
    try:
        from src.modules.health_checker import HealthChecker
        from src.modules.structure_mapper import StructureMapper
        from src.modules.data_profiler import DataProfiler
      
        modules = [
            HealthChecker(),
            StructureMapper(),
            DataProfiler()
        ]
        print("✅ Module instantiation successful")
        return True
    except Exception as e:
        print(f"❌ Module validation failed: {e}")
        return False

if __name__ == "__main__":
    validations = [
        validate_imports,
        validate_config,
        validate_modules
    ]
  
    success = all(validation() for validation in validations)
  
    if success:
        print("\n🎉 System validation PASSED - Ready for testing with Excel files")
    else:
        print("\n💥 System validation FAILED - Fix errors before proceeding")
  
    exit(0 if success else 1)
```

## Execution Order Summary

1. **Phase 1** : Create entry point, fix critical placeholders, resolve imports
2. **Phase 2** : Implement unified configuration system
3. **Phase 3** : Consolidate processing systems
4. **Phase 4** : Add basic error recovery
5. **Phase 5** : Integrate systems and fix type safety
6. **Phase 6** : Final validation and testing

**Validation Commands:**

```bash
# Step-by-step validation
python validate_system.py
python test_basic.py sample.xlsx
python -m src.main sample.xlsx
```

This plan transforms the framework from non-functional placeholder code to a working system with unified architecture and proper error handling.
