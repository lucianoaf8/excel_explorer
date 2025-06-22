# Excel Explorer

A comprehensive Excel file analysis system that generates detailed documentation to enable AI assistants to effectively work with complex Excel files.

## Overview

Excel Explorer analyzes Excel files through multiple specialized modules to create comprehensive documentation that makes complex spreadsheets AI-navigable.

## Quick Start

1. Install dependencies: `pip install -r requirements.txt`
2. Configure settings in `config/analysis_settings.yaml`
3. Run analysis: `python -m src.core.orchestrator <excel_file>`

## Project Structure

- `src/core/` - Main orchestration logic
- `src/modules/` - Specialized analysis modules  
- `src/utils/` - Supporting utilities
- `config/` - Configuration files
- `output/` - Generated documentation
- `tests/` - Test files and cases
- `docs/` - Project documentation

## Modules

- **Health Checker**: File integrity validation
- **Structure Mapper**: Workbook architecture analysis
- **Data Profiler**: Data quality and patterns
- **Formula Analyzer**: Formula dependencies and logic
- **Visual Cataloger**: Charts, images, shapes inventory
- **Connection Inspector**: External connections mapping
- **Pivot Intelligence**: Pivot table analysis
- **Documentation Synthesizer**: Final documentation generator

## Requirements

- Python 3.8+
- See `requirements.txt` for package dependencies

## Configuration

Edit `config/analysis_settings.yaml` to customize analysis parameters for your use case.

## Output

Analysis generates multiple output formats:
- HTML reports (human-readable)
- JSON data (machine-parseable)
- Markdown summaries (AI-friendly)

## Contributing

1. Follow the established module pattern in `src/core/base_analyzer.py`
2. Add comprehensive tests for new functionality
3. Update documentation for any new features

## License

[Add your license information here]
