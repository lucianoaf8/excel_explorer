# Excel Explorer v2.0

A comprehensive Excel file analysis tool with both GUI and CLI interfaces. Provides detailed reporting, security analysis, data profiling, and structure mapping for Excel workbooks.

## Features

- **Unified Entry Point** – Single `main.py` script supports both GUI and CLI modes
- **Cross-platform GUI** – Tkinter-based interface for interactive analysis
- **Command-Line Interface** – Perfect for automation and batch processing
- **Multiple Report Formats** – HTML, JSON, Text, and Markdown outputs
- **Centralized Configuration** – YAML-based config with environment variable overrides
- **Report Consistency** – Unified data model ensures identical metrics across formats
- **Security Analysis** – Detect macros, external references, and potential threats
- **Data Profiling** – Column-wise analysis with type detection and quality metrics
- **Structure Mapping** – Sheet relationships, named ranges, and workbook features

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
pip install -r requirements.txt   # install dependencies globally or in your preferred environment
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
  auto_export: true
  timestamp_reports: true
  
performance:
  timeout_seconds: 300
  parallel_processing: false
```

Environment variables override config file settings:
```bash
export EXCEL_EXPLORER_SAMPLE_ROWS=200
export EXCEL_EXPLORER_DETAIL_LEVEL=basic
```

## Validation

Verify report consistency across formats:
```bash
python -m excel_explorer.utils.validate_reports data.xlsx
```

## Project Structure

```
excel_explorer/
├── src/
│   └── excel_explorer/           # Main package
│       ├── core/                 # Analysis engine & config
│       ├── gui/                  # GUI components  
│       ├── cli/                  # CLI functionality
│       ├── reports/              # Report generation
│       └── utils/                # Utilities & validation
├── tests/                        # Test suite
├── docs/                         # Documentation
├── config/                       # Configuration files
├── output/                       # Generated reports
├── main.py                       # Entry point
├── setup.py                      # Package installation
└── requirements.txt              # Dependencies
```

## Dependencies

- Python ≥ 3.9
- `openpyxl==3.1.2` - Excel file processing
- `PyYAML==6.0.1` - Configuration management
- `tkinter` - GUI framework (included with Python)

## Development

### Running Tests
```bash
python tests/test_architecture.py
```

### Building Package
```bash
python setup.py sdist bdist_wheel
```

### Code Organization
- **Core Logic**: `src/excel_explorer/core/` - Analysis engine and configuration
- **User Interfaces**: `src/excel_explorer/gui/` and `src/excel_explorer/cli/`
- **Report Generation**: `src/excel_explorer/reports/` - Unified reporting system
- **Testing**: `tests/` - Integration tests and demos

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
