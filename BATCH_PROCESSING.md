# Batch Processing Guide

This guide explains how to analyze multiple Excel files at once using Excel Explorer.

## Quick Start - ComplyWorks Reports

To analyze all ComplyWorks Client Value Reports, simply run:

### Windows PowerShell
```powershell
.\analyze_cw_reports.ps1
```

### Windows Command Prompt
```cmd
analyze_cw_reports.bat
```

This will:
- Find all Excel files in `C:\Projects\contractor_analysis\data\complyworks\cw_monthly_raw_files`
- Analyze each file individually
- Generate HTML reports in `C:\Projects\contractor_analysis\reports\cw_analysis`
- Show a summary of successes and failures

## General Batch Analysis

### Analyze All Files in a Directory

```bash
# Basic directory analysis
python batch_analyze.py --directory C:\Projects\data

# With custom output location
python batch_analyze.py --directory C:\Projects\data --output C:\Reports

# Generate JSON reports instead of HTML
python batch_analyze.py --directory C:\Projects\data --format json

# Enable verbose output
python batch_analyze.py --directory C:\Projects\data --verbose
```

### Analyze Specific Files

```bash
# List specific files
python batch_analyze.py --files file1.xlsx file2.xlsx file3.xlsx

# Using full paths
python batch_analyze.py --files "C:\Data\Report1.xlsx" "C:\Data\Report2.xlsx"
```

### Using Glob Patterns

```bash
# All Excel files in a directory
python batch_analyze.py --files "C:\Data\*.xlsx"

# Files matching a pattern
python batch_analyze.py --files "C:\Data\Report_*.xlsx"

# Multiple patterns
python batch_analyze.py --files "C:\Data\Jan_*.xlsx" "C:\Data\Feb_*.xlsx"
```

### Combine Directory and Specific Files

```bash
# Analyze directory plus additional files
python batch_analyze.py --directory C:\Data --files C:\Extra\special.xlsx
```

## Advanced Options

### Report Formats

```bash
# HTML (default)
python batch_analyze.py --directory C:\Data --format html

# JSON
python batch_analyze.py --directory C:\Data --format json

# Text
python batch_analyze.py --directory C:\Data --format text

# Markdown
python batch_analyze.py --directory C:\Data --format markdown
```

### Screenshots (Windows Only)

```bash
# Enable screenshot capture
python batch_analyze.py --directory C:\Data --screenshots

# Screenshots with verbose output
python batch_analyze.py --directory C:\Data --screenshots --verbose
```

### Error Handling

```bash
# Stop on first error (default: continue)
python batch_analyze.py --directory C:\Data --stop-on-error

# Continue processing all files (default)
python batch_analyze.py --directory C:\Data
```

### Custom Configuration

```bash
# Use custom config file
python batch_analyze.py --directory C:\Data --config custom_config.yaml
```

## Output Structure

Reports are saved with timestamps to avoid overwriting:

```
reports/batch/20250105_143022/
├── Report1_20250105_143025.html
├── Report2_20250105_143028.html
├── Report3_20250105_143031.html
└── ...
```

Or with custom output directory:

```
C:\Reports\
├── Report1_20250105_143025.html
├── Report2_20250105_143028.html
└── ...
```

## Exit Codes

- `0` - All files processed successfully
- `1` - All files failed
- `2` - Some files succeeded, some failed (partial success)

## Examples

### Example 1: Monthly Reports
```bash
# Analyze all monthly reports for 2025
python batch_analyze.py --files "C:\Reports\*2025.xlsx" --output C:\Analysis\2025
```

### Example 2: Multiple Formats
```bash
# Generate both HTML and JSON reports
python batch_analyze.py --directory C:\Data --format html
python batch_analyze.py --directory C:\Data --format json --output C:\Data\json_reports
```

### Example 3: Department Reports
```bash
# Analyze all HR department reports
python batch_analyze.py --files "C:\Reports\HR\*.xlsx" --output C:\Analysis\HR --verbose
```

### Example 4: With Screenshots
```bash
# Full analysis with screenshots (Windows)
python batch_analyze.py --directory C:\Data --screenshots --verbose
```

## Customizing ComplyWorks Script

Edit `analyze_cw_reports.ps1` or `analyze_cw_reports.bat` to customize:

```powershell
$DATA_DIR = "C:\Your\Custom\Path"          # Source directory
$OUTPUT_DIR = "C:\Your\Output\Path"        # Output directory
$FORMAT = "json"                            # Change format: html, json, text, markdown
```

```batch
SET DATA_DIR=C:\Your\Custom\Path
SET OUTPUT_DIR=C:\Your\Output\Path
SET FORMAT=json
```

## Troubleshooting

### No Files Found
- Verify the directory path exists
- Check file extensions (.xlsx, .xls, .xlsm)
- Use `--verbose` to see which files were found

### All Files Failing
- Test with a single file first: `python main.py --mode cli --file test.xlsx`
- Check if dependencies are installed: `pip install -r requirements.txt`
- Use `--verbose` to see detailed error messages

### Some Files Failing
- Review the summary to see which files failed and why
- Use `--verbose` for detailed error information
- Failed files might be corrupted or password-protected

### Screenshots Not Working
- Screenshots only work on Windows
- Install dependencies: `pip install xlwings pillow pywin32`
- Excel must be installed on the system

## Tips

1. **Test First**: Always test with a few files before running on large batches
2. **Use Verbose Mode**: Add `--verbose` to see detailed progress
3. **Check Disk Space**: Ensure adequate space for reports (screenshots can be large)
4. **Custom Configs**: Use different configs for different file types
5. **Incremental Processing**: Process files in smaller batches if you have many files

## See Also

- Main documentation: `README.md`
- Configuration guide: `CLAUDE.md`
- Single file analysis: `python main.py --help`
