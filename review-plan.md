# Claude Code Prompt: Comprehensive Excel Analysis Framework Architectural Review

 **Documentation Standards Applied** : Following Claude Code Practical Prompting Guide 10-step methodology and Essentials best practices for systematic codebase analysis.

 **Validation Scores** : Clarity: 5, Specificity: 5, Completeness: 5, Compliance: 5, Effectiveness: 5

## Context and Objective

We're conducting a comprehensive architectural review of an Excel analysis framework with identified issues: overengineered architecture, circular dependencies, overlapping systems, and incomplete implementations. The goal is systematic analysis and concrete remediation recommendations.

 **Project Structure** : Python-based Excel analysis framework with src/ directory containing core/, modules/, and utils/ subdirectories.

## Comprehensive Analysis Instructions

Think step-by-step through each analysis task. Consider architectural patterns, dependency relationships, and performance implications. For each task, provide specific file references, severity classifications, and actionable recommendations.

### Phase 1: Foundation Analysis (Critical Path)

**Task 1.1: Dependency Mapping and Circular Import Detection**

1. **Analyze import structure across the entire codebase:**
   ```bash
   find src/ -name "*.py" -exec grep -H "^from \.\." {} \;
   find src/ -name "*.py" -exec grep -H "^import.*src\." {} \;
   find src/ -name "*.py" -exec grep -H "^from src\." {} \;
   ```
2. **Create dependency graph:**
   * Map all inter-module dependencies
   * Identify circular import chains (A imports B imports A)
   * Check for missing module implementations referenced in imports
   * Validate relative vs absolute import consistency
3. **Generate dependency report:**
   * List all circular dependencies with exact file paths and line numbers
   * Identify phantom imports (importing non-existent modules)
   * Classify dependency violations by severity (Critical/High/Medium)
   * Propose dependency restructuring plan

**Task 1.2: Module Existence Validation**

1. **Verify all imported modules exist:**
   ```bash
   find src/ -name "*.py" -exec python3 -m py_compile {} \; 2>&1 | grep -i "import\|module"
   ```
2. **Check for stub/placeholder modules:**
   ```bash
   find src/ -name "*.py" -exec grep -l "pass$" {} \;
   find src/ -name "*.py" -exec grep -l "NotImplementedError" {} \;
   ```

### Phase 2: Interface and Framework Analysis

**Task 2.1: BaseAnalyzer Framework Audit**

1. **Examine base class implementation:**
   * Open and analyze `src/core/base_analyzer.py` (or equivalent)
   * Document all abstract methods and their signatures
   * Check method naming conventions and parameter consistency
2. **Validate analyzer implementations:**
   ```bash
   grep -r "class.*BaseAnalyzer" src/
   grep -r "_perform_analysis" src/
   grep -r "_validate_result" src/
   grep -r "def analyze" src/
   ```
3. **Interface consistency check:**
   * Compare method signatures across all analyzer implementations
   * Identify missing abstract method implementations
   * Check data structure compatibility between analyzers
   * Document interface violations with specific fix recommendations

**Task 2.2: Data Structure Compatibility Analysis**

1. **Map data flow patterns:**
   ```bash
   grep -r "ModuleResult" src/
   grep -r "AnalysisContext" src/
   grep -r "get_module_result" src/
   ```
2. **Validate data structure consistency:**
   * Check ModuleResult class definition and usage patterns
   * Verify AnalysisContext state management
   * Identify data transformation bottlenecks
   * Document type mismatches and incompatibilities

### Phase 3: System Configuration and Memory Management

**Task 3.1: Configuration System Consolidation Analysis**

1. **Map configuration access patterns:**
   ```bash
   find src/ -name "*.py" -exec grep -H "config\." {} \;
   find src/ -name "*config*" -type f
   grep -r "self\.config\." src/
   ```
2. **Analyze configuration conflicts:**
   * Identify multiple configuration systems
   * Check default value inconsistencies
   * Map configuration validation logic
   * Document configuration access pattern violations

**Task 3.2: Memory Management System Review**

1. **Analyze memory management implementations:**
   ```bash
   grep -r "memory_manager" src/
   grep -r "MemoryManager" src/
   grep -r "get_memory_manager" src/
   find src/ -name "*.py" -exec grep -l "memory" {} \;
   ```
2. **Check memory management patterns:**
   * Identify overlapping memory management systems
   * Review chunk processing memory usage
   * Analyze cache management implementations
   * Check resource cleanup patterns in try/finally blocks

### Phase 4: Error Handling and Robustness

**Task 4.1: Error Handling Pattern Analysis**

1. **Map error handling across modules:**
   ```bash
   grep -r "ExcelAnalysisError" src/
   grep -r "try:" src/ | wc -l
   grep -r "except Exception" src/
   grep -r "raise.*Error" src/
   ```
2. **Analyze error propagation:**
   * Check exception hierarchy consistency
   * Identify bare except clauses
   * Map error logging patterns
   * Document error handling anti-patterns

**Task 4.2: Logging Configuration Analysis**

1. **Check logging setup:**
   ```bash
   find src/ -name "*.py" -exec grep -l "logging" {} \;grep -r "logger" src/ | head -20grep -r "log\." src/ | head -20
   ```

### Phase 5: Performance and Implementation Completeness

**Task 5.1: Processing Pattern Analysis**

1. **Review processing implementations:**
   ```bash
   grep -r "chunk" src/
   grep -r "stream" src/
   grep -r "process_worksheet" src/
   grep -r "batch" src/
   ```
2. **Identify performance bottlenecks:**
   * Map chunked vs streaming processing conflicts
   * Check for redundant data processing
   * Analyze memory usage in large file processing
   * Document processing inefficiencies

**Task 5.2: Implementation Completeness Audit**

1. **Find incomplete implementations:**
   ```bash
   grep -r "placeholder" src/ -i
   grep -r "TODO" src/ -i
   grep -r "FIXME" src/ -i
   grep -r "NotImplementedError" src/
   ```
2. **Check core functionality coverage:**
   * Identify stub methods and empty implementations
   * Map missing test coverage
   * Document incomplete feature implementations

## Analysis Output Requirements

For each task, generate a structured report with:

1. **Issue Classification:**
   * Critical: Blocking system functionality
   * High: Significant architectural problems
   * Medium: Code quality/maintainability issues
   * Low: Minor optimization opportunities
2. **Specific References:**
   * Exact file paths and line numbers
   * Code snippets demonstrating issues
   * Current vs expected behavior
3. **Remediation Plan:**
   * Specific fix recommendations
   * Implementation effort estimates
   * Dependency order for fixes
   * Risk assessment for changes
4. **Priority Matrix:**
   * Order fixes by impact vs effort
   * Identify quick wins vs architectural changes
   * Dependencies between fixes

## Final Deliverable

Compile all analysis results into a comprehensive architectural review document with:

1. **Executive Summary** : Top 5 critical issues requiring immediate attention
2. **Detailed Findings** : Complete analysis results by task
3. **Remediation Roadmap** : Prioritized fix sequence with effort estimates
4. **Architectural Recommendations** : Simplified design proposals
5. **Implementation Plan** : Step-by-step execution strategy

 **Safety Considerations** :

* Run analysis in read-only mode initially
* Create backup branch before any modifications
* Validate all grep/find commands in small scope first
* Do not modify any files during analysis phase

 **Testing Phase** : After analysis completion, run existing test suite to establish baseline before any remediation efforts.

Execute this analysis systematically, providing detailed findings for each task before proceeding to generate the final architectural review report.
