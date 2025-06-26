# Excel Explorer

A comprehensive Excel file analysis system that generates detailed documentation to enable AI assistants to effectively work with complex Excel files.

## 🚀 Quick Start

### Basic Usage
```bash
# Analyze any Excel file
python analyze.py your_file.xlsx

# Save results to JSON
python analyze.py your_file.xlsx --output results.json

# Enable detailed logging
python analyze.py your_file.xlsx --logfile analysis.log
```

### Advanced Options
```bash
# High memory analysis
python analyze.py large_file.xlsx --memory-limit 8192

# Deep analysis mode
python analyze.py complex_file.xlsx --deep-analysis

# Custom configuration
python analyze.py file.xlsx --config custom_settings.yaml
```

## 📋 Installation

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd excel_explorer
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify installation**
   ```bash
   python test_integration.py
   ```

## 🏗️ Architecture

Excel Explorer uses a modular architecture with dependency-driven execution:

### Core Framework
- **AnalysisContext**: Central state management and workbook access
- **ModuleResult**: Standardized result structures with quality metrics
- **BaseAnalyzer**: Enhanced base class for all analysis modules
- **Orchestrator**: Dependency-aware module execution engine

### Analysis Modules
1. **Health Checker**: File integrity and accessibility validation
2. **Structure Mapper**: Workbook architecture analysis (sheets, tables, ranges)
3. **Data Profiler**: Data quality, patterns, and type inference
4. **Formula Analyzer**: Formula dependencies and complexity analysis
5. **Visual Cataloger**: Charts, images, and visual elements inventory
6. **Connection Inspector**: External connections and security assessment
7. **Pivot Intelligence**: Pivot table and slicer analysis
8. **Documentation Synthesizer**: AI-friendly documentation generation

### Support Systems
- **Memory Manager**: Resource monitoring and optimization
- **Error Handler**: Structured error handling with recovery mechanisms
- **Configuration Loader**: YAML-based configuration management

## 📊 Output Formats

Excel Explorer generates comprehensive documentation in multiple formats:

### JSON Output Structure
```json
{
  "file_info": {
    "name": "example.xlsx",
    "size_mb": 15.2,
    "path": "/path/to/file"
  },
  "analysis_metadata": {
    "success_rate": 0.95,
    "total_duration_seconds": 45.3,
    "quality_score": 0.87
  },
  "module_results": {
    "health_checker": { /* health analysis */ },
    "structure_mapper": { /* structure analysis */ },
    "data_profiler": { /* data quality metrics */ }
  },
  "recommendations": [
    "Data quality issues detected - review null values",
    "High formula count - consider optimization"
  ]
}
```

### Key Features
- **Executive Summary**: Human-readable overview
- **AI Navigation Guide**: Structured data for AI consumption
- **Quality Metrics**: Confidence scores and completeness indicators
- **Actionable Recommendations**: Performance and optimization suggestions

## ⚙️ Configuration

Customize analysis behavior through `config/analysis_settings.yaml`:

```yaml
# Global settings
analysis:
  max_memory_mb: 4096
  enable_caching: true
  deep_analysis: false

# Module-specific settings
health_checker:
  max_file_size_mb: 500
  
data_profiler:
  sample_size_limit: 10000
  
formula_analyzer:
  max_formulas_analyze: 50000
```

## 🧪 Testing

Run the integration test to verify all components:

```bash
python test_integration.py
```

Expected output:
```
🧪 Testing Phase 1 Core Components
========================================
1. Testing imports...
   ✅ All imports successful
2. Testing Memory Manager...
   ✅ Memory manager working
...
🎉 ALL TESTS PASSED - Phase 1 Ready for Use!
```

## 📈 Performance Guidelines

### Memory Management
- **Files < 10MB**: Default settings (4GB memory limit)
- **Files 10-50MB**: Consider `--memory-limit 6144`
- **Files 50MB+**: Use `--memory-limit 8192` or higher

### Processing Time Estimates
- **Simple files**: 10-30 seconds
- **Complex files**: 1-5 minutes
- **Large files (100MB+)**: 5-15 minutes

### Optimization Tips
- Enable caching for repeated analysis
- Use `--deep-analysis` only when needed
- Monitor resource usage in logs

## 🔧 Troubleshooting

### Common Issues

**"File not found" error**
```bash
# Ensure file path is correct
python analyze.py "/full/path/to/file.xlsx"
```

**Memory limit exceeded**
```bash
# Increase memory limit
python analyze.py large_file.xlsx --memory-limit 8192
```

**Module dependency failures**
```bash
# Check file integrity first
python analyze.py file.xlsx --logfile debug.log
# Review debug.log for detailed error information
```

### Performance Issues
- Large files: Enable streaming with reduced sample sizes
- Complex formulas: Set `max_formulas_analyze` limit
- Memory pressure: Reduce `chunk_size_rows` in configuration

## 🏆 MVP Achievement Status

✅ **Phase 1 Complete** - Core infrastructure and basic analysis
- Framework architecture implemented
- All core modules functional
- Dependency-driven execution
- Resource management and error handling
- Basic data profiling and structure analysis

🔄 **Next Phases** - Advanced analysis and optimization
- Enhanced formula parsing
- Parallel processing
- Advanced caching strategies
- Production optimization

## 📚 API Reference

### Command Line Interface

```
python analyze.py <excel_file> [options]

Required Arguments:
  excel_file              Path to Excel file (.xlsx, .xlsm, .xls)

Optional Arguments:
  --output PATH           Save results to JSON file
  --config PATH           Custom YAML configuration file
  --logfile PATH          Detailed log file path
  --memory-limit MB       Memory limit in megabytes
  --deep-analysis         Enable comprehensive analysis
  --parallel              Enable parallel processing (when available)
```

### Programmatic Usage

```python
from src.core.orchestrator import ExcelExplorer

# Initialize with custom config
explorer = ExcelExplorer(config_path="custom.yaml")

# Run analysis
results = explorer.analyze_file("data.xlsx")

# Access specific module results
health_data = results['module_results']['health_checker']
structure_data = results['module_results']['structure_mapper']
```

## 🤝 Contributing

1. Follow the established module pattern in `src/core/base_analyzer.py`
2. Add comprehensive tests for new functionality
3. Update documentation for any new features
4. Ensure all modules use the ModuleResult framework

## 📄 License

[Add your license information here]

---

**Excel Explorer** - Making complex spreadsheets AI-navigable since 2025.
