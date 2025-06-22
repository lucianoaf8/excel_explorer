---
trigger: always_on
---

<excel_explorer_context>
- This is the Excel Explorer project: a modular Excel file analysis system with 8 specialized analyzer modules that must be built in dependency order.
- Project uses Python 3.8+ with openpyxl, pandas, xlwings libraries for Excel processing and analysis.
- Follow the exact folder structure: src/core/, src/modules/, src/utils/, output/, config/, tests/, docs/.
- Respect strict module dependency chain: Health Checker → Structure Mapper → Data Profiler → Formula Analyzer → Visual Cataloger → Connection Inspector → Pivot Intelligence → Documentation Synthesizer.
- Focus on core functionality implementation first - no tests or documentation until MVP modules are working.
</excel_explorer_context>

<development_priorities>
- Implement modules in dependency order: cannot start module N+1 until module N is functionally complete.
- All modules must inherit from base_analyzer.py and follow the established interface contracts.
- Prioritize memory management and error resilience in all implementations - these are critical for large Excel files.
- Use streaming processors and chunk-based analysis for data-heavy operations to prevent memory exhaustion.
- Implement graceful degradation and comprehensive exception handling in every module.
- Each module must produce both JSON structured output and human-readable summaries as specified.
</development_priorities>

<technical_requirements>
- Use read-only file access throughout - never modify source Excel files.
- Implement lazy loading for resource-intensive operations to optimize performance.
- Follow the success metrics defined for each module in the documentation.
- Handle Excel formats: .xlsx, .xlsm, .xls with appropriate library selection.
- Maintain security considerations: sandbox operations, disable macro execution, validate external connections.
- Use the utils/ modules (file_handler, memory_manager, error_handler, config_loader) for common operations.
</technical_requirements>

<module_specific_guidance>
- Health Checker: First module, no dependencies, must validate file integrity before any other analysis.
- Structure Mapper: Creates architectural blueprint, dependency for most other modules.
- Data Profiler: Analyzes data quality and patterns, feeds into Formula Analyzer and Pivot Intelligence.
- Formula Analyzer: Complex dependency mapping, requires Structure Mapper and Data Profiler context.
- Visual Cataloger: Charts and visual elements, depends on data source context from previous modules.
- Connection Inspector: External connections mapping, security-focused, minimal dependencies.
- Pivot Intelligence: Requires Structure Mapper and Data Profiler for source data context.
- Documentation Synthesizer: Final module, depends on all others, produces comprehensive output.
</module_specific_guidance>

<implementation_standards>
- Each module must handle files up to 500MB efficiently with memory usage <4x source file size.
- Implement comprehensive error handling with partial result preservation when components fail.
- Use openpyxl for structure analysis, pandas for data processing, xlwings for advanced Excel features.
- Follow the JSON schemas and interface contracts to ensure module integration works correctly.
- Maintain processing standards: Health check <30 seconds, complete analysis <5 minutes per 100MB.
</implementation_standards>