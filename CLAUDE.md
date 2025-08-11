# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Installation and Setup
```bash
# Install core dependencies
pip install -r requirements.txt

# Install screenshot dependencies (Windows only)
pip install xlwings pillow pywin32

# Install as editable package (optional)
pip install -e .
```

### Running the Application
```bash
# GUI mode (default)
python main.py

# CLI mode with file analysis
python main.py --mode cli --file data.xlsx

# CLI with specific output format and directory
python main.py --mode cli --file data.xlsx --format json --output ./reports

# CLI with custom configuration
python main.py --mode cli --file data.xlsx --config config/config.yaml --verbose

# CLI with screenshot capture (Windows only)
python main.py --mode cli --file data.xlsx --screenshots

# CLI with screenshots and verbose output
python main.py --mode cli --file data.xlsx --screenshots --verbose
```

### Testing
```bash
# Run integration tests
python tests/test_architecture.py

# Run individual test modules
python tests/test_json_structure.py
```

### Report Validation
```bash
# Validate report consistency across formats
python -m src.utils.validate_reports data.xlsx

# Test screenshot functionality
python test_screenshot.py
```

## Architecture Overview

Excel Explorer v2.0 is a comprehensive Excel analysis tool with dual interfaces (GUI/CLI), modular package-based architecture, and advanced features including Excel sheet screenshot capture.

### Core Structure
- **src/**: Main package directory with modular components
- **main.py**: Unified entry point that routes to GUI or CLI modes
- **config/**: YAML-based configuration with environment variable overrides
- **output/reports/**: Generated analysis reports with timestamps

### Key Modules

#### Entry Points
- **main.py** (root): Launch script that adds src/ to Python path
- **src/main.py**: Unified entry point with argument parsing for GUI/CLI modes
- **src/gui/excel_explorer_gui.py**: Tkinter-based GUI application
- **src/cli/cli_runner.py**: Command-line interface implementation

#### Core Analysis Engine (Refactored v2.0)
- **src/core/analysis_service.py**: High-level service providing unified interface for CLI and GUI
- **src/core/analyzers/orchestrator.py**: Coordinates execution of all analyzer modules
- **src/core/analyzers/structure.py**: Structure analysis module (sheets, ranges, tables)
- **src/core/analyzers/data.py**: Data profiling and quality analysis module
- **src/core/analyzers/formula.py**: Formula analysis and complexity scoring module
- **src/core/analyzers/screenshot.py**: Excel sheet screenshot capture module (Windows only)
- **src/core/config.py**: Simplified configuration loading (81 lines vs 325 lines)

#### Report Generation
- **src/reports/report_generator.py**: Multi-format report generation (HTML, JSON, text, markdown)
- **src/reports/report_base.py**: Unified data model for consistent reporting
- **src/reports/report_adapter.py**: Adapter bridging new AnalysisService with existing ReportDataModel
- **src/reports/structured_text_report.py**: Text and markdown output formatters

#### Utilities
- **src/utils/validate_reports.py**: Cross-format report consistency validation
- **src/utils/markdown_utils.py**: Markdown formatting utilities

### Configuration System
- **config/config.yaml**: Central configuration file with analysis parameters
- Environment variables: `EXCEL_EXPLORER_*` prefix overrides config values
- Key settings: `sample_rows`, `max_formula_check`, `memory_limit_mb`, `detail_level`, `screenshot` options

### Analysis Modules Architecture (Refactored v2.0)
The system now uses a modular, service-oriented architecture with clear separation of concerns:

#### Modular Analyzers
- **StructureAnalyzer**: Analyzes workbook structure, sheets, ranges, tables
- **DataAnalyzer**: Performs data profiling, quality analysis, type distribution
- **FormulaAnalyzer**: Analyzes formulas, complexity scoring, dependencies
- **ScreenshotAnalyzer**: Captures exact visual screenshots of Excel sheets (Windows only)

#### Service Layer
- **AnalysisService**: High-level service providing unified API for CLI and GUI
- **AnalyzerOrchestrator**: Coordinates module execution with error handling and progress tracking
- **ReportService**: Handles report generation with adapter pattern for backward compatibility

#### Key Features
- Plugin-based module registration and execution tracking
- Error isolation per module with graceful degradation
- Configurable timeout and memory limits per module
- Standardized data model output across all modules
- Progress callbacks with module-level granularity

### Report Data Model
- **ReportDataModel**: Standardizes analysis results across all output formats
- Sections: `file_summary`, `quality_metrics`, `module_execution`, `sheet_details`
- Ensures identical metrics in HTML, JSON, text, and markdown outputs
- Timestamp and version tracking for report provenance

### GUI Features
- Custom progress indicators with module-level tracking
- Embedded HTML preview within application
- Asynchronous analysis with responsive UI
- Automatic browser integration for full report viewing

### Performance Design
- Memory-aware processing with `read_only=True` for openpyxl
- Configurable sampling limits for large files
- Streaming analysis with chunked processing
- Module-level timeouts with fallback mechanisms

## Development Guidelines

### Package Structure
All code resides under `src/` following standard Python package conventions. The root `main.py` is a launcher that adds `src/` to the Python path.

### Configuration Management
Use the simplified `config.py` module for all settings. Environment variables automatically override YAML config using the pattern `EXCEL_EXPLORER_<SECTION>_<KEY>`. The old 325-line ConfigManager class has been replaced with a streamlined 81-line implementation focusing on DRY principles.

### Adding New Analysis Modules
1. Create module class extending `BaseAnalyzer` in `src/core/analyzers/`
2. Implement the `analyze(workbook)` method returning standardized results
3. Register module in `AnalyzerOrchestrator._initialize_analyzers()`
4. Follow error handling patterns with graceful degradation
5. Update `ReportAdapter` if new data sections are needed for report compatibility

### Report Format Consistency
All new report formats must use the `ReportDataModel` to ensure consistency. Test with `validate_reports.py` to verify identical metrics across formats.

### Screenshot Feature (New in v2.0)
The screenshot analyzer captures Excel sheets exactly as they appear in Excel using xlwings and Windows COM automation:

#### Features
- **Pixel-perfect capture**: Uses Excel's CopyPicture API to capture sheets exactly as displayed
- **No content modification**: Captures visual appearance without altering sheet data
- **Organized output**: Screenshots saved in timestamped folders with sheet names
- **Multiple capture modes**: `used_range`, `full_sheet`, or `print_area`
- **Windows only**: Requires xlwings, PIL, and pywin32 dependencies

#### Usage
```bash
# Enable screenshots in CLI
python main.py --mode cli --file data.xlsx --screenshots

# Configuration in config.yaml
screenshot:
  enabled: false
  show_excel: false          # Show Excel during capture (debugging)
  capture_mode: used_range   # used_range | full_sheet | print_area
  output_dir: output/screenshots
  format: PNG
  quality: 95
```

#### Dependencies
```bash
# Windows only - install screenshot dependencies
pip install xlwings pillow pywin32
```

### Error Handling
- Module-level isolation prevents single failures from breaking analysis
- User-friendly error messages in both GUI and CLI modes
- Comprehensive logging with configurable levels