# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Setup
```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
# or
venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

### Running the Application
```bash
python main.py
# or
python excel_explorer_gui.py
```

### Dependencies
- Python ≥ 3.9
- openpyxl ≥ 3.1.2 (Excel file processing)
- PyYAML ≥ 6.0 (configuration)
- tkinter (GUI - bundled with Python)

## Architecture Overview

This is a desktop Excel analysis application with a modular architecture:

### Core Components
- **main.py**: Entry point with error handling
- **excel_explorer_gui.py**: Tkinter GUI with custom progress indicators and embedded HTML preview
- **analyzer.py**: Core analysis engine with pluggable modules
- **report_generator.py**: HTML and JSON report generation

### Analysis Modules (in analyzer.py)
The analyzer uses a module-based architecture where each module handles a specific aspect:
- `health_checker`: File validation and loading
- `structure_mapper`: Sheet structure and metadata
- `data_profiler`: Data type analysis with configurable sampling
- `formula_analyzer`: Formula detection and complexity analysis
- `visual_cataloger`: Charts, images, and formatting detection
- `connection_inspector`: External data connections (stub)
- `pivot_intelligence`: Pivot table analysis (stub)
- `doc_synthesizer`: Documentation generation (stub)

### Configuration System
- **config.yaml**: Runtime configuration with performance tuning
  - Analysis limits: `max_formula_check`, `sample_rows`, `memory_limit_mb`
  - Output formats: `json_enabled`, `html_enabled`
  - Performance: `timeout_per_sheet_seconds`, `chunk_size`

### GUI Architecture
- Custom `CircularProgress` widget for visual feedback
- `ProgressTracker` for coordinating UI updates during analysis
- Asynchronous analysis with threading to keep UI responsive
- Embedded HTML preview using tkinter.html (if available)

### Report Generation
- HTML reports with collapsible sections and timestamps
- JSON export capability
- Reports saved to `reports/` directory with timestamped filenames
- Cross-platform browser integration for viewing reports

## Key Features

### Analysis Capabilities
- Read-only Excel file processing (no modifications)
- Configurable sampling for large files (default: 100 rows per sheet)
- Memory-aware processing with fallback mechanisms
- Cross-sheet dependency analysis
- Data quality profiling with type detection

### GUI Features
- File selection dialog with Excel file filtering
- Real-time progress tracking with circular indicator
- Automatic tab switching to report view on completion
- "Open Last Report" button for quick access
- Embedded HTML preview within the application

### Performance Considerations
- Uses `read_only=True` for openpyxl to minimize memory usage
- Configurable timeouts and sampling limits
- Streaming analysis for large datasets
- Memory limit checks with graceful degradation

## Development Notes

### Code Style
- Uses type hints throughout
- Modular design with clear separation of concerns
- Exception handling with user-friendly error messages
- Configuration-driven behavior via YAML

### Testing
- No formal test framework currently implemented
- Manual testing workflow via GUI interaction

### Output
- Reports directory: `reports/` (auto-created)
- HTML reports: Interactive with collapsible sections
- JSON exports: Structured data for programmatic access

### Error Handling
- Graceful degradation when modules fail
- Module-level error tracking in reports
- User-friendly error messages in GUI
- Automatic cleanup of resources (workbook closing)