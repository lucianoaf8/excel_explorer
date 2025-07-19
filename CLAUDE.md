# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Installation and Setup
```bash
# Install dependencies
pip install -r requirements.txt

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
```

## Architecture Overview

Excel Explorer v2.0 is a comprehensive Excel analysis tool with dual interfaces (GUI/CLI) and a modular, package-based architecture.

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

#### Core Analysis Engine
- **src/core/analyzer.py**: Modular analysis engine with pluggable components
- **src/core/config_manager.py**: Configuration loading with environment variable support

#### Report Generation
- **src/reports/report_generator.py**: Multi-format report generation (HTML, JSON, text, markdown)
- **src/reports/report_base.py**: Unified data model for consistent reporting
- **src/reports/structured_text_report.py**: Text and markdown output formatters

#### Utilities
- **src/utils/validate_reports.py**: Cross-format report consistency validation
- **src/utils/markdown_utils.py**: Markdown formatting utilities

### Configuration System
- **config/config.yaml**: Central configuration file with analysis parameters
- Environment variables: `EXCEL_EXPLORER_*` prefix overrides config values
- Key settings: `sample_rows`, `max_formula_check`, `memory_limit_mb`, `detail_level`

### Analysis Modules Architecture
The analyzer uses a plugin-based system where each module handles specific aspects:
- Module registration and execution tracking
- Error isolation per module with graceful degradation
- Configurable timeout and memory limits per module
- Standardized data model output across all modules

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
Use the ConfigManager class for all settings. Environment variables automatically override YAML config using the pattern `EXCEL_EXPLORER_<SECTION>_<KEY>`.

### Adding New Analysis Modules
1. Implement module in `src/core/analyzer.py`
2. Register module in the analyzer's module list
3. Follow the standard module interface with error handling
4. Update `ReportDataModel` if new data sections are needed

### Report Format Consistency
All new report formats must use the `ReportDataModel` to ensure consistency. Test with `validate_reports.py` to verify identical metrics across formats.

### Error Handling
- Module-level isolation prevents single failures from breaking analysis
- User-friendly error messages in both GUI and CLI modes
- Comprehensive logging with configurable levels