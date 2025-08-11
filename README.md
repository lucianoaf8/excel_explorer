# Excel Explorer v2.0

A comprehensive Excel file analysis tool with both GUI and CLI interfaces. Features modular architecture, pixel-perfect Excel sheet screenshots, data anonymization, and detailed reporting with security analysis, data profiling, and structure mapping for Excel workbooks.

## Features

### Core Analysis
- **Unified Entry Point** – Single `main.py` script supports both GUI and CLI modes
- **Cross-platform GUI** – Tkinter-based interface for interactive analysis
- **Command-Line Interface** – Perfect for automation and batch processing
- **Multiple Report Formats** – HTML, JSON, Text, and Markdown outputs
- **Modular Architecture** – Service-oriented design with pluggable analyzers

### Advanced Features
- **📸 Excel Screenshots** – Pixel-perfect capture of sheets exactly as they appear in Excel (Windows only)
- **🔒 Data Anonymization** – Secure data masking with reversible mapping for sensitive information
- **🔍 Security Analysis** – Detect macros, external references, and potential threats
- **📊 Data Profiling** – Column-wise analysis with type detection and quality metrics
- **🗂️ Structure Mapping** – Sheet relationships, named ranges, and workbook features

### Configuration & Reporting
- **Centralized Configuration** – YAML-based config with environment variable overrides
- **Report Consistency** – Unified data model ensures identical metrics across formats
- **Progress Tracking** – Module-level progress callbacks and error isolation
- **Flexible Output** – Timestamped reports with organized folder structures

## Installation

### From Source
```bash
git clone <repository-url>
cd excel_explorer
pip install -e .
```

### Local Development (no virtual-environment)
You can run Excel Explorer directly from the cloned repository—no virtual-environment required:

```bash
git clone <repository-url>
cd excel_explorer
pip install -r requirements.txt   # install core dependencies globally or in your preferred environment

# Optional: Install screenshot dependencies (Windows only)
pip install xlwings pillow pywin32

# Optional: Install anonymization dependencies
pip install faker

python main.py                    # launches GUI (default)
# or
python main.py --mode cli --file data.xlsx  # CLI
```

### After Installation
If you installed the package (e.g. `pip install -e .` or from PyPI), use the console script or module entry point:

```bash
excel-explorer --help             # console script
python -m excel_explorer.main --mode gui  # module invocation
```

## Usage

### GUI Mode (Default)
```bash
python main.py                    # Launch GUI
excel-explorer                    # If installed via pip
```

### CLI Mode
```bash
# Basic analysis with HTML report
python main.py --mode cli --file data.xlsx

# Specify output format and directory
python main.py --mode cli --file data.xlsx --format json --output ./reports

# Use custom configuration
python main.py --mode cli --file data.xlsx --config config/config.yaml --verbose

# Enable screenshot capture (Windows only)
python main.py --mode cli --file data.xlsx --screenshots

# Enable data anonymization
python main.py --mode cli --file data.xlsx --anonymize

# Anonymize specific columns
python main.py --mode cli --file data.xlsx --anonymize --anonymize-columns "Sheet1:Name" "Sheet1:Email"

# Reverse anonymization
python main.py --mode cli --file anonymized.xlsx --reverse mappings.json

# All available formats
python main.py --mode cli --file data.xlsx --format html     # Default
python main.py --mode cli --file data.xlsx --format json    # Structured data
python main.py --mode cli --file data.xlsx --format text    # Plain text
python main.py --mode cli --file data.xlsx --format markdown # Markdown
```

## Configuration

Excel Explorer uses a centralized configuration system:

```yaml
# config/config.yaml
analysis:
  sample_rows: 100
  max_formula_check: 1000
  memory_limit_mb: 512
  detail_level: 'comprehensive'

output:
  json_enabled: true
  html_enabled: true
  include_raw_data: false
  
performance:
  timeout_seconds: 300
  parallel_processing: false

screenshot:
  enabled: false                # Enable/disable screenshot capture
  show_excel: false            # Show Excel window during capture
  capture_mode: used_range     # used_range | full_sheet | print_area
  output_dir: output/screenshots
  format: PNG
  quality: 95

logging:
  level: INFO
  include_timestamps: true
```

Environment variables override config file settings:
```bash
export EXCEL_EXPLORER_ANALYSIS_SAMPLE_ROWS=200
export EXCEL_EXPLORER_ANALYSIS_DETAIL_LEVEL=basic
export EXCEL_EXPLORER_SCREENSHOT_ENABLED=true
```

## Special Features

### Excel Sheet Screenshots (Windows Only)
Capture pixel-perfect screenshots of Excel sheets exactly as they appear:

```bash
# Enable screenshots via CLI flag
python main.py --mode cli --file data.xlsx --screenshots

# Or configure in config.yaml
screenshot:
  enabled: true
  capture_mode: used_range  # or 'full_sheet' or 'print_area'
  show_excel: false        # set true for debugging
```

**Requirements**: `xlwings`, `pillow`, `pywin32` (Windows only)

### Data Anonymization
Protect sensitive data while maintaining analysis accuracy:

```bash
# Anonymize all detected sensitive data
python main.py --mode cli --file data.xlsx --anonymize

# Anonymize specific columns
python main.py --mode cli --file data.xlsx --anonymize --anonymize-columns "Sheet1:Name" "Sheet1:Email"

# Reverse anonymization when needed
python main.py --mode cli --file anonymized.xlsx --reverse mapping.json
```

**Requirements**: `faker` library

## Validation

Verify report consistency across formats:
```bash
python -m src.utils.validate_reports data.xlsx

# Test screenshot functionality
python test_screenshot.py
```

## Project Structure

```
excel_explorer/
├── src/                          # Main source package
│   ├── core/                     # Analysis engine & config
│   │   ├── analyzers/            # Modular analyzers (v2.0 refactor)
│   │   │   ├── structure.py      # Sheet structure analysis
│   │   │   ├── data.py          # Data profiling & quality
│   │   │   ├── formula.py       # Formula analysis
│   │   │   ├── screenshot.py    # Excel screenshots (Windows)
│   │   │   └── orchestrator.py  # Module coordination
│   │   ├── analysis_service.py  # Unified service interface
│   │   └── config.py            # Simplified configuration (81 lines)
│   ├── gui/                      # GUI components  
│   ├── cli/                      # CLI functionality
│   │   ├── cli_runner.py        # Main CLI interface
│   │   └── anonymizer_command.py # Anonymization CLI
│   ├── reports/                  # Report generation
│   │   ├── report_adapter.py    # New/old system bridge
│   │   └── report_generator.py  # Multi-format reports
│   └── utils/                    # Utilities & validation
├── tests/                        # Test suite
├── config/                       # Configuration files
├── output/                       # Generated reports & screenshots
│   ├── reports/                  # Analysis reports
│   └── screenshots/              # Excel sheet screenshots
├── main.py                       # Unified entry point
├── test_screenshot.py            # Screenshot testing utility
├── setup.py                      # Package installation
└── requirements.txt              # Dependencies
```

## Dependencies

### Core Dependencies
- Python ≥ 3.9
- `openpyxl==3.1.2` - Excel file processing
- `PyYAML==6.0.1` - Configuration management  
- `mdutils==1.6.0` - Markdown report generation
- `Pillow==10.1.0` - Image processing
- `tkinter` - GUI framework (included with Python)

### Optional Dependencies
- `xlwings>=0.30.0` - Excel automation for screenshots (Windows only)
- `pywin32>=306` - Windows COM automation (Windows only)  
- `faker>=20.0.0` - Data anonymization and masking

## Development

### Running Tests
```bash
python tests/test_architecture.py
```

### Building Package
```bash
python setup.py sdist bdist_wheel
```

### Code Organization (v2.0 Refactored Architecture)
- **Modular Analyzers**: `src/core/analyzers/` - Plugin-based analysis modules with service orchestration
- **Service Layer**: `src/core/analysis_service.py` - Unified interface providing consistent API for CLI and GUI
- **User Interfaces**: `src/gui/` and `src/cli/` - Separate interface implementations using common service
- **Report Generation**: `src/reports/` - Multi-format reporting with adapter pattern for backward compatibility
- **Configuration**: `src/core/config.py` - Simplified 81-line configuration system (reduced from 325 lines)
- **Testing**: `tests/` - Integration tests and feature demos

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run tests
5. Submit a pull request

---

*Excel Explorer v2.0 - Professional Excel Analysis Made Simple*
