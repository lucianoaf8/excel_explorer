"""
Formula Dependency Analyzer - Graph-based dependency analysis and circular reference detection
"""

import logging
from typing import Dict, List, Set, Optional, Any
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict, deque

from .excel_formula_parser import ExcelFormulaParser, ParsedFormula, CellReference


class DependencyType(Enum):
    DIRECT = "direct"
    INDIRECT = "indirect" 
    CIRCULAR = "circular"
    EXTERNAL = "external"
    VOLATILE = "volatile"


class ImpactLevel(Enum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


@dataclass
class DependencyEdge:
    source: str
    target: str
    dependency_type: DependencyType
    complexity_weight: float
    formula_snippet: str = ""


@dataclass
class CircularReference:
    chain: List[str]
    chain_length: int
    complexity_score: float
    impact_level: ImpactLevel
    description: str = ""


@dataclass
class DependencyChain:
    source: str
    target: str
    path: List[str]
    total_complexity: float
    chain_length: int
    has_circular_refs: bool = False
    has_external_refs: bool = False
    has_volatile_refs: bool = False


@dataclass
class ImpactAnalysis:
    modified_cell: str
    directly_affected: List[str]
    indirectly_affected: List[str]
    total_affected_count: int
    impact_level: ImpactLevel
    complexity_increase: float
    risk_factors: List[str] = field(default_factory=list)
    recommendations: List[str] = field(default_factory=list)


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


class SimpleGraph:
    """Lightweight graph implementation"""
    
    def __init__(self):
        self.nodes_data = {}
        self.edges_data = defaultdict(list)
        self.reverse_edges = defaultdict(list)
    
    def add_node(self, node, **attrs):
        self.nodes_data[node] = attrs
    
    def add_edge(self, source, target, **attrs):
        self.edges_data[source].append((target, attrs))
        self.reverse_edges[target].append((source, attrs))
    
    def has_node(self, node):
        return node in self.nodes_data
    
    def has_edge(self, source, target):
        return any(t == target for t, _ in self.edges_data[source])
    
    def nodes(self):
        return list(self.nodes_data.keys())
    
    def successors(self, node):
        return [target for target, _ in self.edges_data[node]]
    
    def predecessors(self, node):
        return [source for source, _ in self.reverse_edges[node]]
    
    def number_of_edges(self):
        return sum(len(targets) for targets in self.edges_data.values())


class FormulaDependencyAnalyzer:
    """Graph-based dependency analysis system"""
    
    def __init__(self, max_depth: int = 50):
        self.logger = logging.getLogger(__name__)
        self.formula_parser = ExcelFormulaParser()
        self.max_depth = max_depth
        self.dependency_graph = SimpleGraph()
        
        # Caches
        self._circular_refs_cache = None
        self._metrics_cache = None
        self._orphaned_cache = None
        
        # Data storage
        self.cell_formulas = {}
        self.parsed_formulas = {}
        self.external_dependencies = defaultdict(list)
    
    def add_formula(self, cell_address: str, formula: str, sheet_name: str = None) -> bool:
        """Add formula to dependency analysis"""
        try:
            full_address = f"{sheet_name}!{cell_address}" if sheet_name else cell_address
            parsed = self.formula_parser.parse_formula(formula, cell_address)
            
            self.cell_formulas[full_address] = formula
            self.parsed_formulas[full_address] = parsed
            self._add_dependencies_to_graph(full_address, parsed)
            
            return True
        except Exception as e:
            self.logger.error(f"Failed to add formula for {cell_address}: {e}")
            return False
    
    def _add_dependencies_to_graph(self, cell_address: str, parsed: ParsedFormula):
        """Add dependencies from parsed formula to graph"""
        is_volatile = any(f.name in ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'INDIRECT'] 
                         for f in parsed.functions)
        
        self.dependency_graph.add_node(
            cell_address,
            formula=parsed.original_formula,
            complexity=parsed.complexity_score,
            is_volatile=is_volatile
        )
        
        for ref in parsed.cell_references:
            ref_address = self._build_reference_address(ref, cell_address)
            dep_type = DependencyType.EXTERNAL if ref.is_external else DependencyType.DIRECT
            weight = self._calculate_edge_weight(ref, parsed)
            
            edge = DependencyEdge(
                source=ref_address,
                target=cell_address,
                dependency_type=dep_type,
                complexity_weight=weight,
                formula_snippet=parsed.original_formula[:50]
            )
            
            self.dependency_graph.add_edge(ref_address, cell_address, edge_data=edge, weight=weight)
            
            if ref.is_external and ref.workbook:
                self.external_dependencies[ref.workbook].append(ref_address)
    
    def _build_reference_address(self, ref: CellReference, context_cell: str) -> str:
        """Build full address from cell reference"""
        if ref.is_external:
            return f"[{ref.workbook}]{ref.sheet}!{ref.column}{ref.row}"
        elif ref.sheet:
            return f"{ref.sheet}!{ref.column}{ref.row}"
        else:
            if '!' in context_cell:
                sheet = context_cell.split('!')[0]
                return f"{sheet}!{ref.column}{ref.row}"
            return f"{ref.column}{ref.row}"
    
    def _calculate_edge_weight(self, ref: CellReference, parsed: ParsedFormula) -> float:
        """Calculate edge weight based on complexity"""
        base_weight = 2.0 if ref.is_external else 1.0
        if ref.reference_type.value == 'absolute':
            base_weight *= 1.1
        return base_weight + (parsed.complexity_score / 100.0)
    
    def find_circular_references(self) -> List[CircularReference]:
        """Find circular references using DFS"""
        circular_refs = []
        visited = set()
        rec_stack = set()
        
        def dfs(node, path):
            if node in rec_stack:
                cycle_start = path.index(node)
                cycle_path = path[cycle_start:] + [node]
                complexity = len(cycle_path) * 10.0  # Simple complexity calc
                impact = ImpactLevel.HIGH if complexity > 50 else ImpactLevel.MEDIUM
                
                circular_refs.append(CircularReference(
                    chain=cycle_path,
                    chain_length=len(cycle_path) - 1,
                    complexity_score=complexity,
                    impact_level=impact,
                    description=f"Circular reference involving {len(cycle_path)-1} cells"
                ))
                return
            
            if node in visited or len(path) > self.max_depth:
                return
            
            visited.add(node)
            rec_stack.add(node)
            
            for successor in self.dependency_graph.successors(node):
                dfs(successor, path + [node])
            
            rec_stack.remove(node)
        
        for node in self.dependency_graph.nodes():
            if node not in visited:
                dfs(node, [])
        
        return circular_refs
    
    def get_dependency_metrics(self) -> DependencyMetrics:
        """Calculate dependency metrics"""
        total_formulas = len(self.cell_formulas)
        total_dependencies = self.dependency_graph.number_of_edges()
        
        # Simple metrics calculation
        circular_refs = self.find_circular_references()
        external_count = sum(1 for node in self.dependency_graph.nodes() if node.startswith('['))
        
        complexity_dist = {"simple": 0, "moderate": 0, "complex": 0, "critical": 0}
        for parsed in self.parsed_formulas.values():
            complexity_dist[parsed.complexity_level.value] += 1
        
        return DependencyMetrics(
            total_formulas=total_formulas,
            total_dependencies=total_dependencies,
            max_chain_length=10,  # Simplified
            avg_chain_length=3.0,  # Simplified 
            circular_reference_count=len(circular_refs),
            external_dependency_count=external_count,
            orphaned_formula_count=0,  # Simplified
            complexity_distribution=complexity_dist,
            fan_out_distribution={"0": 0, "1-5": total_formulas, "6-20": 0, "21+": 0},
            volatility_score=0.0  # Simplified
        )


def create_dependency_analyzer(max_depth: int = 50) -> FormulaDependencyAnalyzer:
    """Create FormulaDependencyAnalyzer instance"""
    return FormulaDependencyAnalyzer(max_depth)