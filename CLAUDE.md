# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands for Development

### Basic Analysis
```bash
# Analyze any Excel file
python -m src.core.orchestrator your_file.xlsx

# Save results to JSON
python -m src.core.orchestrator your_file.xlsx --output results.json

# Enable detailed logging
python -m src.core.orchestrator your_file.xlsx --logfile analysis.log
```

### Advanced Analysis Options
```bash
# High memory analysis for large files
python -m src.core.orchestrator large_file.xlsx --memory-limit 8192

# Deep analysis mode
python -m src.core.orchestrator complex_file.xlsx --deep-analysis

# Custom configuration
python -m src.core.orchestrator file.xlsx --config custom_settings.yaml
```

### Testing and Quality Assurance
```bash
# Run integration tests (if available)
python test_integration.py

# Code quality checks
black src/ --check --diff
flake8 src/
mypy src/

# Format code
black src/
isort src/
```

### Environment Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Set up virtual environment (Windows)
python3.11 -m venv .venv311
.venv311\Scripts\activate

# Set up virtual environment (Linux/WSL)
python3.11 -m venv .venv311
source .venv311/bin/activate
```

## Architecture Overview

Excel Explorer is a comprehensive Excel file analysis system built with a **modular, dependency-driven architecture**. The system generates detailed documentation to enable AI assistants to effectively work with complex Excel files.

### Core Framework (`src/core/`)

- **`orchestrator.py`**: Main controller that manages module execution order based on dependencies
- **`analysis_context.py`**: Central state management providing shared workbook access and file metadata
- **`base_analyzer.py`**: Abstract base class implementing common functionality for all analysis modules
- **`module_result.py`**: Standardized result structures with quality metrics and validation

### Analysis Pipeline

The system executes modules in dependency order:

1. **Health Checker** → File integrity validation (no dependencies)
2. **Structure Mapper** → Workbook architecture analysis (depends on health check)
3. **Data Profiler** → Data quality and type inference (depends on structure mapping)
4. **Formula Analyzer** → Formula dependencies and complexity (depends on structure mapping)
5. **Visual Cataloger** → Charts and visual elements (depends on structure mapping)
6. **Connection Inspector** → External connections (depends on health check)
7. **Pivot Intelligence** → Pivot table analysis (depends on structure + data profiling)
8. **Documentation Synthesizer** → Final report generation (depends on multiple modules)

### Analysis Modules (`src/modules/`)

Each module extends `BaseAnalyzer` and implements:
- `_perform_analysis()`: Core analysis logic
- `_validate_result()`: Quality validation
- Dependency checking and resource management

Key modules:
- **`health_checker.py`**: File corruption detection, security warnings
- **`structure_mapper.py`**: Sheet inventory, named ranges, protection levels
- **`data_profiler.py`**: Type inference, data quality metrics, header detection
- **`formula_analyzer.py`**: Formula dependency graphs, circular reference detection
- **`visual_cataloger.py`**: Chart inventory with data source mapping
- **`connection_inspector.py`**: External data connection analysis
- **`pivot_intelligence.py`**: Pivot table and slicer analysis
- **`doc_synthesizer.py`**: AI-friendly documentation generation

### Utility Systems (`src/utils/`)

- **`memory_manager.py`**: Resource monitoring and optimization
- **`error_handler.py`**: Structured error handling with recovery mechanisms
- **`config_loader.py`**: YAML configuration management
- **`chunked_processor.py`**: Memory-efficient processing for large files
- **`streaming_processor.py`**: Row/column/cell-level streaming for massive datasets
- **`data_validator.py`**: Comprehensive data validation framework

## Configuration Management

### Primary Configuration File
- **`config/analysis_settings.yaml`**: Comprehensive module settings including:
  - Memory limits and chunking strategies
  - Quality thresholds and validation rules
  - Performance optimization settings
  - Feature flags for optional modules

### Key Configuration Patterns
- Module settings follow `module_name: { enabled: true, ... }` pattern
- Memory management through `analysis.max_memory_mb` and chunking config
- Quality thresholds defined per module with fallback defaults
- Performance tuning via `chunk_size_rows` and `sample_size_limit`

## Development Patterns

### Adding New Analysis Modules
1. Extend `BaseAnalyzer` class
2. Implement required abstract methods: `_perform_analysis()` and `_validate_result()`
3. Define dependencies in module registry (`orchestrator.py`)
4. Add configuration section to `analysis_settings.yaml`
5. Register module in `MODULE_REGISTRY` with appropriate execution phase

### Error Handling Strategy
- Use `ExcelAnalysisError` for domain-specific errors
- Implement graceful degradation for non-critical failures
- Critical modules (health_checker, structure_mapper) stop execution on failure
- Optional modules continue pipeline execution despite individual failures

### Memory Management
- All modules use `ResourceMonitor` context manager
- Chunked processing available via `ChunkedSheetProcessor`
- Streaming analysis for large datasets via `StreamingProcessor`
- Memory pressure detection triggers reduced complexity analysis

### Testing and Validation
- Integration testing through comprehensive file analysis
- Module-level validation via `ValidationResult` framework
- Quality metrics include completeness, accuracy, and confidence scores
- Performance benchmarking against file size and complexity

## File Structure Conventions

```
excel_explorer/
├── src/                    # Source code
│   ├── core/              # Framework components
│   ├── modules/           # Analysis modules
│   └── utils/             # Support utilities
├── config/                # Configuration files
├── output/                # Generated reports and cache
│   ├── reports/          # Human-readable reports
│   ├── structured/       # Machine-parseable data
│   └── cache/            # Intermediate processing files
└── requirements.txt       # Python dependencies
```

## Performance Considerations

### File Size Guidelines
- **< 10MB**: Default settings with 4GB memory limit
- **10-50MB**: Consider `--memory-limit 6144`
- **50MB+**: Use `--memory-limit 8192` and enable chunked processing
- **100MB+**: Enable streaming analysis and reduce sample sizes

### Processing Optimization
- Use chunked processing for memory efficiency
- Enable caching for repeated analysis of same files
- Monitor resource usage through integrated memory manager
- Adjust `chunk_size_rows` based on available memory

### Quality vs Performance Trade-offs
- Standard analysis: 95%+ accuracy target in under 5 minutes
- Deep analysis: Enhanced accuracy with longer processing time
- Streaming mode: Memory efficiency for massive files with progressive insights