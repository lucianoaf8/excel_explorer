# Excel Explorer v2.0 - Comprehensive Project Documentation

## Table of Contents

1. [Project Overview](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#project-overview)
2. [Architecture &amp; Design](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#architecture--design)
3. [Installation &amp; Setup](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#installation--setup)
4. [Usage Guide](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#usage-guide)
5. [Configuration](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#configuration)
6. [Development Guide](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#development-guide)
7. [API Reference](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#api-reference)
8. [Testing &amp; Validation](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#testing--validation)
9. [Deployment](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#deployment)
10. [Troubleshooting](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#troubleshooting)
11. [Contributing](https://claude.ai/chat/1f81d95d-7550-47ed-84ce-3bf318a8b96e#contributing)

---

## Project Overview

### Purpose

Excel Explorer v2.0 is a comprehensive Excel file analysis tool designed for data analysts, auditors, and business users who need rapid, detailed insights into Excel workbooks without opening them in Excel. It provides security analysis, data profiling, structure mapping, and generates interactive reports in multiple formats.

### Key Capabilities

* **Rapid File Inspection** : Analyze Excel files without Excel installation
* **Security Assessment** : Detect macros, external references, sensitive data patterns
* **Data Profiling** : Column-wise analysis with type detection and quality metrics
* **Structure Mapping** : Sheet relationships, hidden content, workbook features
* **Multi-format Reporting** : HTML (interactive), JSON, Text, and Markdown
* **Real-time Progress** : GUI with circular progress indicators and CLI verbose mode
* **Cross-platform** : Windows, macOS, Linux support via Python and Tkinter

### Supported File Formats

* `.xlsx` (Excel 2007+)
* `.xlsm` (Excel 2007+ Macro-enabled)
* `.xlsb` (Excel 2007+ Binary)
* `.xls` (Excel 97-2003)

---

## Architecture & Design

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        main.py                              │
│                  (Unified Entry Point)                      │
└─────────────────┬───────────────────────┬───────────────────┘
                  │                       │
        ┌─────────▼─────────┐   ┌─────────▼─────────┐
        │   GUI Mode        │   │   CLI Mode        │
        │ (Tkinter-based)   │   │ (Command-line)    │
        └─────────┬─────────┘   └─────────┬─────────┘
                  │                       │
                  └───────────┬───────────┘
                              │
                  ┌───────────▼───────────┐
                  │     Core Analysis     │
                  │      Engine           │
                  └───────────┬───────────┘
                              │
              ┌───────────────┼───────────────┐
              │               │               │
    ┌─────────▼─────────┐ ┌───▼───┐ ┌─────────▼─────────┐
    │  Report Generation│ │Config │ │    Utilities      │
    │  (Multi-format)   │ │ Mgmt  │ │  & Validation     │
    └───────────────────┘ └───────┘ └───────────────────┘
```

### Project Structure

```
excel_explorer/
├── main.py                          # Unified entry point
├── config/
│   └── config.yaml                  # Default configuration
├── src/
│   ├── __init__.py
│   ├── main.py                      # Module entry point
│   ├── core/                        # Analysis engine
│   │   ├── __init__.py
│   │   ├── analyzer.py              # Main analysis logic
│   │   └── config_manager.py        # Configuration management
│   ├── gui/                         # GUI components
│   │   ├── __init__.py
│   │   └── excel_explorer_gui.py    # Tkinter interface
│   ├── cli/                         # CLI components
│   │   ├── __init__.py
│   │   └── cli_runner.py            # Command-line execution
│   ├── reports/                     # Report generation
│   │   ├── __init__.py
│   │   ├── report_base.py           # Unified data model
│   │   ├── report_generator.py      # HTML/JSON reports
│   │   └── structured_text_report.py # Text/Markdown reports
│   └── utils/                       # Utilities
│       ├── __init__.py
│       ├── markdown_utils.py        # Markdown helpers
│       └── validate_reports.py      # Report validation
├── tests/                           # Test suite
│   ├── __init__.py
│   ├── test_architecture.py         # Integration tests
│   └── test_json_structure.py       # JSON validation
├── output/                          # Generated reports
│   ├── reports/                     # Auto-exported reports
│   └── cache/                       # Temporary files
├── assets/                          # Assets and resources
│   └── logos/                       # Application logos
├── requirements.txt                 # Python dependencies
├── setup.py                         # Package installation
├── README.md                        # Quick start guide
└── .gitignore                       # Version control exclusions
```

### Core Analysis Modules

#### 1. Health Checker

* **Purpose** : Initial file validation and integrity check
* **Implementation** : `_get_file_info()` in `analyzer.py`
* **Checks** : File accessibility, signature validation, corruption detection

#### 2. Structure Mapper

* **Purpose** : Workbook structure analysis
* **Implementation** : `_analyze_structure()` in `analyzer.py`
* **Features** : Sheet enumeration, hidden content, named ranges, tables, protection

#### 3. Data Profiler

* **Purpose** : Cell-level data analysis and quality assessment
* **Implementation** : `_analyze_data()` in `analyzer.py`
* **Features** : Type detection, column profiling, data density, quality metrics

#### 4. Formula Analyzer

* **Purpose** : Formula complexity and dependency analysis
* **Implementation** : `_analyze_formulas()` in `analyzer.py`
* **Features** : Formula counting, complexity scoring, external reference detection

#### 5. Visual Cataloger

* **Purpose** : Charts, images, and visual elements analysis
* **Implementation** : `_analyze_visuals()` in `analyzer.py`
* **Features** : Chart counting, image detection, conditional formatting

#### 6. Security Inspector

* **Purpose** : Security risk assessment and threat detection
* **Implementation** : `_analyze_security()` in `analyzer.py`
* **Features** : Macro detection, sensitive data patterns, security scoring

---

## Installation & Setup

### Prerequisites

* Python ≥ 3.9
* Operating System: Windows, macOS, or Linux
* Memory: 4GB RAM recommended (8GB for large files)
* Storage: 100MB for installation + space for reports

### Installation Methods

#### Method 1: Direct Repository Usage (Recommended for Development)

```bash
# Clone repository
git clone <repository-url>
cd excel_explorer

# Install dependencies
pip install -r requirements.txt

# Run directly
python main.py                    # GUI mode
python main.py --mode cli --help  # CLI help
```

#### Method 2: Package Installation

```bash
# Install as package
pip install -e .

# Use console script
excel-explorer --help
# OR module invocation
python -m excel_explorer.main --mode gui
```

#### Method 3: Virtual Environment (Recommended for Production)

```bash
# Create virtual environment
python -m venv excel_explorer_env

# Activate environment
# Windows:
excel_explorer_env\Scripts\activate
# macOS/Linux:
source excel_explorer_env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Dependency Details

```
openpyxl==3.1.2      # Excel file processing (read-only streaming)
PyYAML==6.0.1        # Configuration management
mdutils==1.6.0       # Markdown report generation
Pillow==10.1.0       # Image processing for GUI logos
```

### Platform-Specific Setup

#### Windows

* Tkinter included with Python
* No additional setup required

#### macOS

* Tkinter included with Python
* May require: `brew install python-tk` if using Homebrew Python

#### Linux (Ubuntu/Debian)

```bash
# Install tkinter
sudo apt-get install python3-tk

# Install dependencies
pip install -r requirements.txt
```

---

## Usage Guide

### GUI Mode

#### Starting the GUI

```bash
python main.py                # Default GUI mode
python main.py --mode gui     # Explicit GUI mode
```

#### GUI Workflow

1. **File Selection** : Click "Browse Excel Files" to select target file
2. **Analysis Execution** : Click "Start Analysis" to begin processing
3. **Progress Monitoring** : Real-time circular progress indicator with module status
4. **Report Review** : Automatic tab switch to "Analysis Report" upon completion
5. **Export Options** : Multiple export formats available

#### GUI Features

* **Circular Progress Indicator** : Visual progress with percentage and elapsed time
* **Module Progress Tracking** : Real-time updates for each analysis module
* **Tabbed Interface** : Separate tabs for logs, results, and reports
* **Search Functionality** : Find text within generated reports
* **Multiple Export Formats** : HTML, JSON, Text, Markdown
* **Auto-Export** : Automatic HTML report generation upon completion

### CLI Mode

#### Basic Usage

```bash
# Basic analysis with HTML report
python main.py --mode cli --file data.xlsx

# Specify output format and directory
python main.py --mode cli --file data.xlsx --format json --output ./reports

# Verbose output with custom configuration
python main.py --mode cli --file data.xlsx --config custom.yaml --verbose
```

#### CLI Options

```bash
Options:
  --mode {gui,cli}           Execution mode (default: gui)
  --file FILE               Excel file to analyze (required for CLI)
  --output OUTPUT           Output directory (default: ./reports)
  --format {html,json,text,markdown}  Report format (default: html)
  --config CONFIG           Configuration file path (default: config.yaml)
  --verbose, -v             Enable verbose output
```

#### CLI Examples

```bash
# Generate all report formats
python main.py --mode cli --file data.xlsx --format html
python main.py --mode cli --file data.xlsx --format json
python main.py --mode cli --file data.xlsx --format text
python main.py --mode cli --file data.xlsx --format markdown

# Custom configuration with verbose output
python main.py --mode cli --file data.xlsx --config production.yaml --verbose

# Batch processing with output directory
python main.py --mode cli --file large_file.xlsx --output /path/to/reports --format json
```

### Report Formats

#### HTML Report (Default)

* **Features** : Interactive, tabbed interface with expandable sections
* **Tabs** : Overview, Structure, Data Quality, Sheet Analysis, Security, Recommendations
* **Interactive Elements** : Collapsible sections, data tables, progress bars
* **Styling** : Professional CSS with responsive design
* **Use Case** : Detailed analysis review, stakeholder presentations

#### JSON Report

* **Features** : Structured data format for programmatic access
* **Content** : Complete analysis results with standardized schema
* **Use Case** : API integration, further data processing, automation

#### Text Report

* **Features** : Plain text format with structured sections
* **Content** : Key metrics, findings, and recommendations
* **Use Case** : Documentation, email reports, system logs

#### Markdown Report

* **Features** : GitHub-flavored markdown with tables and formatting
* **Content** : Structured analysis with tables and lists
* **Use Case** : Documentation websites, wikis, version control

---

## Configuration

### Configuration File Structure

```yaml
# config/config.yaml
analysis:
  max_cells_check: 1000              # Maximum cells to check for formulas
  max_formula_check: 1000            # Maximum formulas to analyze
  sample_rows: 100                   # Rows sampled per sheet
  max_sample_rows: 1000              # Upper bound for sampling
  memory_limit_mb: 512               # Soft memory limit
  timeout_per_sheet_seconds: 30      # Timeout per sheet analysis
  enable_cross_sheet_analysis: true  # Enable relationship analysis
  enable_data_quality_checks: true   # Enable data quality assessment
  detail_level: comprehensive        # basic | standard | comprehensive

output:
  json_enabled: true                 # Enable JSON report generation
  html_enabled: true                 # Enable HTML report generation
  include_raw_data: false            # Include raw data in reports
  auto_export: true                  # Auto-export reports after analysis
  timestamp_reports: true            # Add timestamps to report names

performance:
  parallel_processing: false         # Enable parallel processing
  chunk_size: 1000                   # Processing chunk size
  timeout_seconds: 300               # Overall analysis timeout
  memory_warning_threshold_mb: 1024  # Memory warning threshold

logging:
  level: INFO                        # DEBUG | INFO | WARNING | ERROR
  include_timestamps: true           # Include timestamps in logs
  log_to_file: false                 # Enable file logging
  log_file_path: excel_explorer.log  # Log file path

security:
  enable_pattern_detection: true     # Enable sensitive data detection
  scan_for_pii: true                 # Scan for personally identifiable information
  security_threshold: 8.0            # Security score threshold
```

### Environment Variable Overrides

```bash
# Analysis settings
export EXCEL_EXPLORER_SAMPLE_ROWS=200
export EXCEL_EXPLORER_MAX_FORMULA_CHECK=2000
export EXCEL_EXPLORER_MEMORY_LIMIT_MB=1024
export EXCEL_EXPLORER_TIMEOUT_SECONDS=600
export EXCEL_EXPLORER_DETAIL_LEVEL=comprehensive

# Performance settings
export EXCEL_EXPLORER_CHUNK_SIZE=2000
export EXCEL_EXPLORER_PARALLEL_PROCESSING=true

# Output settings
export EXCEL_EXPLORER_AUTO_EXPORT=true
export EXCEL_EXPLORER_INCLUDE_RAW_DATA=false

# Logging settings
export EXCEL_EXPLORER_LOG_LEVEL=DEBUG
export EXCEL_EXPLORER_LOG_TO_FILE=true
```

### Configuration Management

* **Priority** : Environment variables > Configuration file > Defaults
* **Validation** : Automatic constraint validation and bounds checking
* **Reload** : Dynamic configuration reloading without restart
* **Export** : Export current configuration to YAML file

---

## Development Guide

### Development Setup

```bash
# Clone and setup development environment
git clone <repository-url>
cd excel_explorer

# Create development virtual environment
python -m venv dev_env
source dev_env/bin/activate  # or dev_env\Scripts\activate on Windows

# Install dependencies
pip install -r requirements.txt

# Install development tools (optional)
pip install pytest black isort flake8
```

### Code Organization Principles

#### Modular Design

* **Core Logic** : Isolated in `src/core/` for reusability
* **Interface Separation** : GUI and CLI in separate modules
* **Report Generation** : Unified data model ensures consistency
* **Configuration** : Centralized management with environment overrides

#### Key Classes and Functions

##### `SimpleExcelAnalyzer` (src/core/analyzer.py)

* **Purpose** : Main analysis engine
* **Key Methods** :
* `analyze(file_path, progress_callback)`: Primary analysis method
* `_analyze_structure(wb)`: Structure analysis
* `_analyze_data(wb)`: Data profiling
* `_analyze_security(wb, data)`: Security assessment

##### `ConfigManager` (src/core/config_manager.py)

* **Purpose** : Configuration management with singleton pattern
* **Key Methods** :
* `load_config(config_path)`: Load configuration with overrides
* `get(key_path, default)`: Get configuration value using dot notation
* `reload_config()`: Force configuration reload

##### `ReportDataModel` (src/reports/report_base.py)

* **Purpose** : Unified data model for consistent reporting
* **Key Methods** :
* `get_standardized_data()`: Return standardized data structure
* `_validate_completeness()`: Ensure all required sections exist

##### `ExcelExplorerApp` (src/gui/excel_explorer_gui.py)

* **Purpose** : Tkinter-based GUI application
* **Key Features** :
* Circular progress tracking
* Real-time module updates
* Embedded report display
* Multi-format export

### Adding New Analysis Modules

#### Step 1: Implement Analysis Logic

```python
# Add to SimpleExcelAnalyzer class in analyzer.py
def _analyze_new_feature(self, wb) -> Dict[str, Any]:
    """Analyze new feature in workbook"""
    results = {
        'feature_count': 0,
        'feature_details': [],
        'quality_score': 0.0
    }
  
    # Implementation logic
    for ws in wb.worksheets:
        # Analysis code here
        pass
  
    return results
```

#### Step 2: Integrate into Main Analysis

```python
# Add to analyze() method in SimpleExcelAnalyzer
new_feature = _safe_run("new_feature_analyzer", "Analyzing new feature", 
                       lambda: self._analyze_new_feature(wb))
```

#### Step 3: Update Report Generation

```python
# Add to _compile_results() method
'module_results': {
    # existing modules...
    'new_feature_analyzer': new_feature
}
```

#### Step 4: Update Progress Tracking

```python
# Add to ProgressTracker modules list in excel_explorer_gui.py
self.modules = [
    'health_checker', 'structure_mapper', 'data_profiler', 
    'formula_analyzer', 'visual_cataloger', 'security_inspector',
    'new_feature_analyzer'  # Add new module
]
```

### Adding New Report Formats

#### Step 1: Create Generator Class

```python
# Create new file: src/reports/new_format_report.py
from .report_base import BaseReportGenerator

class NewFormatReportGenerator(BaseReportGenerator):
    def _generate_content(self) -> str:
        """Generate new format content"""
        file_summary = self._get_file_summary()
        quality_metrics = self._get_quality_metrics()
      
        # Format-specific generation logic
        content = f"New Format Report for {file_summary['name']}\n"
        content += f"Quality Score: {quality_metrics['overall_quality_score']:.1%}\n"
      
        return content
```

#### Step 2: Integrate into CLI

```python
# Update cli_runner.py
elif format_type == 'new_format':
    output_file = output_dir / f"{base_name}.newext"
    generator = NewFormatReportGenerator()
    generator.generate_report(results, str(output_file))
```

#### Step 3: Update Argument Parser

```python
# Update main.py
parser.add_argument('--format', 
                   choices=['html', 'json', 'text', 'markdown', 'new_format'], 
                   default='html', help='Report format (default: html)')
```

### Error Handling Patterns

#### Safe Module Execution

```python
def _safe_run(module_name: str, description: str, func: Callable) -> Any:
    """Execute module with error handling"""
    try:
        self._update_progress(module_name, "starting", description)
        result = func()
        self._update_progress(module_name, "complete")
        return result
    except Exception as e:
        self._update_progress(module_name, "error", str(e))
        return self._get_fallback_result(module_name)
```

#### Configuration Validation

```python
def _validate_config(self):
    """Validate configuration values"""
    constraints = {
        ('analysis', 'sample_rows'): (1, 10000),
        ('analysis', 'memory_limit_mb'): (64, 8192)
    }
  
    for path, (min_val, max_val) in constraints.items():
        value = self.get('.'.join(path))
        if value < min_val or value > max_val:
            # Apply constraint and log warning
            pass
```

### Testing Strategy

#### Unit Tests

* Test individual analysis modules
* Validate configuration management
* Test report generation components

#### Integration Tests

* End-to-end CLI analysis
* GUI functionality testing
* Report consistency validation

#### Performance Tests

* Large file handling
* Memory usage monitoring
* Processing time benchmarks

---

## API Reference

### Core Analysis API

#### `SimpleExcelAnalyzer`

```python
from src.core import SimpleExcelAnalyzer

analyzer = SimpleExcelAnalyzer(config_path="config.yaml")
results = analyzer.analyze(file_path, progress_callback=None)
```

**Parameters:**

* `file_path` (str): Path to Excel file
* `progress_callback` (Callable, optional): Progress update function

**Returns:**

* `Dict[str, Any]`: Complete analysis results

**Progress Callback Signature:**

```python
def progress_callback(module_name: str, status: str, detail: str = ""):
    """
    Args:
        module_name: Name of current analysis module
        status: 'starting', 'step', 'complete', or 'error'
        detail: Additional status information
    """
    pass
```

#### `ConfigManager`

```python
from src.core import ConfigManager

config = ConfigManager()
config.load_config("config.yaml")
value = config.get("analysis.sample_rows", default=100)
```

**Key Methods:**

* `load_config(config_path)`: Load configuration file
* `get(key_path, default)`: Get configuration value
* `reload_config()`: Reload configuration
* `export_current_config(output_path)`: Export current config

### Report Generation API

#### `ReportGenerator`

```python
from src.reports import ReportGenerator

generator = ReportGenerator()
html_path = generator.generate_html_report(results, "report.html")
json_path = generator.generate_json_report(results, "report.json")
```

#### `StructuredTextReportGenerator`

```python
from src.reports.structured_text_report import StructuredTextReportGenerator

generator = StructuredTextReportGenerator()
text_content = generator.generate_report(results)
markdown_content = generator.generate_markdown_report(results, "Title")
generator.export_to_file(content, "report.txt", "text")
```

### CLI API

#### `run_cli_analysis`

```python
from src.cli import run_cli_analysis

exit_code = run_cli_analysis(
    file_path="data.xlsx",
    output_dir="./reports",
    format_type="html",
    config_path="config.yaml",
    verbose=True
)
```

**Parameters:**

* `file_path` (str): Excel file to analyze
* `output_dir` (str, optional): Output directory
* `format_type` (str): Report format ('html', 'json', 'text', 'markdown')
* `config_path` (str): Configuration file path
* `verbose` (bool): Enable verbose output

**Returns:**

* `int`: Exit code (0 = success, 1 = error)

### Data Structures

#### Analysis Results Schema

```python
{
    'file_info': {
        'name': str,
        'size_mb': float,
        'path': str,
        'created': str,
        'modified': str,
        'sheet_count': int,
        'sheets': List[str]
    },
    'analysis_metadata': {
        'timestamp': float,
        'total_duration_seconds': float,
        'success_rate': float,
        'quality_score': float,
        'security_score': float
    },
    'module_results': {
        'structure_mapper': {...},
        'data_profiler': {...},
        'formula_analyzer': {...},
        'visual_cataloger': {...},
        'security_inspector': {...}
    },
    'execution_summary': {
        'total_modules': int,
        'successful_modules': int,
        'failed_modules': int,
        'module_statuses': Dict[str, str]
    },
    'recommendations': List[str]
}
```

#### Standardized Data Model

```python
{
    'file_summary': {...},
    'quality_metrics': {...},
    'security_analysis': {...},
    'structure_analysis': {...},
    'data_analysis': {...},
    'sheet_details': [...],
    'formula_analysis': {...},
    'visual_analysis': {...},
    'performance_metrics': {...},
    'recommendations': [...],
    'module_execution': {...}
}
```

---

## Testing & Validation

### Test Suite Organization

```
tests/
├── __init__.py
├── test_architecture.py        # Integration tests
├── test_json_structure.py      # JSON validation
├── demo_architecture.py        # Demo script
└── sample_files/               # Test Excel files
```

### Running Tests

#### Integration Tests

```bash
# Run full integration test suite
python tests/test_architecture.py

# Run specific test categories
python -c "from tests.test_architecture import test_cli_functionality; test_cli_functionality()"
```

#### Report Consistency Validation

```bash
# Validate report consistency across formats
python -m src.utils.validate_reports test_file.xlsx

# Validate with custom configuration
python -m src.utils.validate_reports test_file.xlsx custom_config.yaml
```

#### Demo Architecture

```bash
# Run architecture demonstration
python tests/demo_architecture.py
```

### Manual Testing Checklist

#### GUI Testing

* [ ] File selection dialog works
* [ ] Analysis starts and completes successfully
* [ ] Progress indicators update correctly
* [ ] All tabs display content properly
* [ ] Search functionality works in reports
* [ ] Export buttons generate correct files
* [ ] Error handling displays appropriate messages

#### CLI Testing

* [ ] Help command displays all options
* [ ] All report formats generate successfully
* [ ] Verbose output provides detailed information
* [ ] Custom configuration files are respected
* [ ] Error conditions are handled gracefully

#### Report Testing

* [ ] HTML reports display correctly in browsers
* [ ] JSON reports contain valid structure
* [ ] Text reports are readable and well-formatted
* [ ] Markdown reports render properly
* [ ] All formats contain consistent core metrics

### Performance Testing

#### Test File Categories

* **Small Files** : < 1MB, < 10 sheets
* **Medium Files** : 1-10MB, 10-50 sheets
* **Large Files** : 10-100MB, 50+ sheets
* **Very Large Files** : > 100MB, complex structure

#### Performance Benchmarks

```python
# Memory usage monitoring
def test_memory_usage():
    import psutil
    process = psutil.Process()
  
    initial_memory = process.memory_info().rss / 1024 / 1024  # MB
  
    # Run analysis
    analyzer = SimpleExcelAnalyzer()
    results = analyzer.analyze("large_file.xlsx")
  
    peak_memory = process.memory_info().rss / 1024 / 1024  # MB
    memory_used = peak_memory - initial_memory
  
    assert memory_used < 1024, f"Memory usage too high: {memory_used}MB"
```

---

## Deployment

### Production Deployment

#### Docker Deployment

```dockerfile
# Dockerfile
FROM python:3.9-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    python3-tk \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Set environment variables
ENV EXCEL_EXPLORER_LOG_LEVEL=INFO
ENV EXCEL_EXPLORER_AUTO_EXPORT=true

# Create volume for reports
VOLUME ["/app/output"]

# Default to CLI mode for containers
ENTRYPOINT ["python", "main.py", "--mode", "cli"]
```

```bash
# Build and run Docker container
docker build -t excel-explorer .
docker run -v $(pwd)/reports:/app/output excel-explorer --file /app/sample.xlsx
```

#### Systemd Service (Linux)

```ini
# /etc/systemd/system/excel-explorer.service
[Unit]
Description=Excel Explorer Analysis Service
After=network.target

[Service]
Type=simple
User=excel-explorer
WorkingDirectory=/opt/excel-explorer
ExecStart=/opt/excel-explorer/venv/bin/python main.py --mode cli --file %i
Environment=EXCEL_EXPLORER_LOG_TO_FILE=true
Environment=EXCEL_EXPLORER_AUTO_EXPORT=true
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

#### Windows Service

```python
# windows_service.py
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import subprocess
import os

class ExcelExplorerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ExcelExplorerService"
    _svc_display_name_ = "Excel Explorer Analysis Service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                            servicemanager.PYS_SERVICE_STARTED,
                            (self._svc_name_, ''))
        self.main()

    def main(self):
        # Service main loop
        pass

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ExcelExplorerService)
```

### Packaging for Distribution

#### Python Package

```bash
# Build distribution packages
python setup.py sdist bdist_wheel

# Upload to PyPI (requires API token)
pip install twine
twine upload dist/*
```

#### Executable Generation

```bash
# Install PyInstaller
pip install pyinstaller

# Generate executable (GUI mode)
pyinstaller --onefile --windowed --name ExcelExplorer main.py

# Generate executable (CLI mode)
pyinstaller --onefile --console --name ExcelExplorerCLI main.py
```

#### MSI Installer (Windows)

```python
# installer.py using cx_Freeze
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["openpyxl", "yaml", "tkinter"],
    "include_files": ["config/", "assets/"]
}

setup(
    name="ExcelExplorer",
    version="2.0.0",
    description="Excel Analysis Tool",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base="Win32GUI")]
)
```

### Environment Configuration

#### Production Settings

```yaml
# production.yaml
analysis:
  sample_rows: 200
  memory_limit_mb: 1024
  timeout_per_sheet_seconds: 60
  detail_level: comprehensive

performance:
  parallel_processing: true
  timeout_seconds: 600
  memory_warning_threshold_mb: 2048

logging:
  level: INFO
  log_to_file: true
  log_file_path: /var/log/excel_explorer.log

output:
  auto_export: true
  timestamp_reports: true
```

#### Monitoring and Logging

```python
# logging_config.py
import logging
import logging.handlers

def setup_production_logging():
    logger = logging.getLogger('excel_explorer')
    logger.setLevel(logging.INFO)
  
    # File handler with rotation
    file_handler = logging.handlers.RotatingFileHandler(
        '/var/log/excel_explorer.log',
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
  
    # Console handler for containers
    console_handler = logging.StreamHandler()
  
    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
  
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
  
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
  
    return logger
```

---

## Troubleshooting

### Common Issues and Solutions

#### Installation Issues

 **Issue** : `ModuleNotFoundError: No module named 'tkinter'`

```bash
# Linux (Ubuntu/Debian)
sudo apt-get install python3-tk

# CentOS/RHEL
sudo yum install tkinter
# or
sudo dnf install python3-tkinter

# macOS (Homebrew)
brew install python-tk
```

 **Issue** : `ImportError: No module named 'openpyxl'`

```bash
# Install missing dependencies
pip install -r requirements.txt

# Or install individual package
pip install openpyxl==3.1.2
```

#### Memory Issues

 **Issue** : `MemoryError` or `Out of memory` during analysis

```yaml
# Reduce memory usage in config.yaml
analysis:
  sample_rows: 50          # Reduce from 100
  max_sample_rows: 500     # Reduce from 1000
  memory_limit_mb: 256     # Reduce from 512
```

 **Issue** : Analysis hangs on large files

```yaml
# Set timeouts in config.yaml
analysis:
  timeout_per_sheet_seconds: 30
performance:
  timeout_seconds: 300
```

#### GUI Issues

 **Issue** : GUI doesn't start on headless systems

```bash
# Use CLI mode instead
python main.py --mode cli --file data.xlsx

# Or enable X11 forwarding (SSH)
ssh -X user@server
python main.py --mode gui
```

 **Issue** : Progress indicator not updating

* Ensure file is accessible and not locked by another application
* Check file permissions
* Verify file is valid Excel format

#### Report Generation Issues

 **Issue** : HTML report doesn't display correctly

* Ensure modern browser (Chrome, Firefox, Edge, Safari)
* Check for JavaScript errors in browser console
* Verify output file was completely written

 **Issue** : JSON report validation errors

```python
# Validate JSON structure
import json
with open('report.json', 'r') as f:
    data = json.load(f)
    print("JSON is valid")
```

### Performance Optimization

#### For Large Files (>100MB)

```yaml
# config.yaml optimizations
analysis:
  sample_rows: 25              # Reduce sampling
  max_formula_check: 500       # Reduce formula analysis
  enable_cross_sheet_analysis: false  # Disable expensive analysis
  detail_level: standard       # Reduce detail level

performance:
  chunk_size: 500              # Smaller chunks
  timeout_seconds: 1800        # Longer timeout (30 min)
```

#### Memory Management

```python
# Custom memory monitoring
import psutil

def monitor_memory_usage():
    process = psutil.Process()
    memory_mb = process.memory_info().rss / 1024 / 1024
  
    if memory_mb > 1024:  # 1GB threshold
        print(f"Warning: High memory usage: {memory_mb:.1f}MB")
        # Trigger garbage collection
        import gc
        gc.collect()
```

### Debugging Techniques

#### Enable Debug Logging

```bash
# Environment variable
export EXCEL_EXPLORER_LOG_LEVEL=DEBUG

# Or in config.yaml
logging:
  level: DEBUG
  include_timestamps: true
```

#### Verbose CLI Output

```bash
# Enable verbose mode for detailed progress
python main.py --mode cli --file data.xlsx --verbose
```

#### Progress Callback Debugging

```python
def debug_progress_callback(module_name, status, detail):
    timestamp = datetime.now().strftime('%H:%M:%S')
    print(f"[{timestamp}] {module_name}: {status} - {detail}")

analyzer = SimpleExcelAnalyzer()
results = analyzer.analyze("file.xlsx", progress_callback=debug_progress_callback)
```

### Error Recovery Strategies

#### Partial Analysis Recovery

* Individual module failures don't stop entire analysis
* Fallback results provided for failed modules
* Success rate calculation includes partial failures

#### Configuration Reset

```python
# Reset to default configuration
from src.core import ConfigManager

config = ConfigManager()
config._config = None  # Clear cached config
default_config = config._get_default_config()
config.export_current_config("reset_config.yaml")
```

#### Safe Mode Analysis

```python
# Minimal analysis for problematic files
safe_config = {
    'analysis': {
        'sample_rows': 10,
        'max_formula_check': 100,
        'enable_cross_sheet_analysis': False,
        'detail_level': 'basic'
    }
}

analyzer = SimpleExcelAnalyzer()
analyzer.config = safe_config
results = analyzer.analyze("problematic_file.xlsx")
```

---

## Contributing

### Development Workflow

#### Setting Up Development Environment

```bash
# Fork and clone repository
git clone https://github.com/your-username/excel-explorer.git
cd excel-explorer

# Create development branch
git checkout -b feature/your-feature-name

# Set up virtual environment
python -m venv dev_env
source dev_env/bin/activate  # Linux/macOS
# or dev_env\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt

# Install development tools
pip install pytest black isort flake8 mypy
```

#### Code Quality Standards

##### Code Formatting

```bash
# Format code with Black
black src/ tests/

# Sort imports with isort
isort src/ tests/

# Lint with flake8
flake8 src/ tests/ --max-line-length=100

# Type checking with mypy
mypy src/ --ignore-missing-imports
```

##### Pre-commit Setup

```yaml
# .pre-commit-config.yaml
repos:
  - repo: https://github.com/psf/black
    rev: 23.7.0
    hooks:
      - id: black
        language_version: python3.9

  - repo: https://github.com/pycqa/isort
    rev: 5.12.0
    hooks:
      - id: isort

  - repo: https://github.com/pycqa/flake8
    rev: 6.0.0
    hooks:
      - id: flake8
        args: [--max-line-length=100]
```

#### Contribution Guidelines

##### Branching Strategy

* `main`: Production-ready code
* `develop`: Integration branch for features
* `feature/*`: Individual feature development
* `bugfix/*`: Bug fixes
* `hotfix/*`: Critical production fixes

##### Commit Message Format

```
type(scope): short description

Longer description if needed.

- Bullet points for changes
- Reference issues: Fixes #123
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

##### Pull Request Process

1. **Create Feature Branch** : `git checkout -b feature/new-analysis-module`
2. **Implement Changes** : Follow coding standards and add tests
3. **Run Tests** : Ensure all tests pass
4. **Update Documentation** : Update relevant documentation
5. **Submit PR** : Create pull request with detailed description
6. **Code Review** : Address reviewer feedback
7. **Merge** : Squash and merge after approval

### Adding New Features

#### Feature Request Template

```markdown
## Feature Description
Brief description of the proposed feature.

## Use Case
Why is this feature needed? What problem does it solve?

## Proposed Implementation
High-level approach to implementing the feature.

## Dependencies
Any new dependencies or requirements.

## Testing Strategy
How will the feature be tested?
```

#### Implementation Checklist

* [ ] Feature implementation with error handling
* [ ] Unit tests with >80% coverage
* [ ] Integration tests if applicable
* [ ] Configuration options if needed
* [ ] Documentation updates
* [ ] CLI/GUI integration if applicable
* [ ] Performance impact assessment
* [ ] Backward compatibility verification

### Bug Reports

#### Bug Report Template

```markdown
## Bug Description
Clear description of the bug.

## Steps to Reproduce
1. Step one
2. Step two
3. Step three

## Expected Behavior
What should happen.

## Actual Behavior
What actually happens.

## Environment
- OS: [Windows/macOS/Linux]
- Python Version: [3.9/3.10/3.11]
- Excel Explorer Version: [2.0.0]
- File Size: [if relevant]

## Error Messages
```

Any error messages or stack traces.

```

## Additional Context
Screenshots, logs, or other relevant information.
```

### Documentation Contributions

#### Documentation Standards

* **Clarity** : Use clear, concise language
* **Examples** : Provide practical examples
* **Completeness** : Cover all aspects of features
* **Accuracy** : Ensure technical accuracy
* **Formatting** : Follow Markdown standards

#### Documentation Types

* **README** : Quick start and overview
* **API Docs** : Detailed API reference
* **User Guide** : Step-by-step instructions
* **Developer Guide** : Development information
* **Configuration** : Configuration options

### Release Process

#### Version Management

* **Semantic Versioning** : MAJOR.MINOR.PATCH
* **Release Notes** : Detailed changelog
* **Migration Guide** : Breaking changes documentation

#### Release Checklist

* [ ] All tests passing
* [ ] Documentation updated
* [ ] Version numbers updated
* [ ] Changelog updated
* [ ] Security review completed
* [ ] Performance benchmarks passed
* [ ] Backward compatibility verified
* [ ] Release notes prepared

---

This comprehensive documentation provides everything needed to understand, use, develop, and contribute to the Excel Explorer v2.0 project. The modular architecture, extensive configuration options, and multiple interfaces make it a powerful tool for Excel file analysis while maintaining ease of use and extensibility.
