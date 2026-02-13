# Batch analyze ComplyWorks Client Value Reports
# Analyzes all CW monthly summary reports from August 2024 - September 2025

$DATA_DIR = "C:\Projects\contractor_analysis\data\complyworks\cw_monthly_raw_files"
$OUTPUT_DIR = "C:\Projects\contractor_analysis\reports\cw_analysis"
$FORMAT = "markdown"  # Options: html, json, text, markdown

Write-Host "="*70 -ForegroundColor Cyan
Write-Host "ComplyWorks Report Batch Analyzer" -ForegroundColor Cyan
Write-Host "="*70 -ForegroundColor Cyan
Write-Host ""
Write-Host "Data directory: $DATA_DIR" -ForegroundColor Yellow
Write-Host "Output directory: $OUTPUT_DIR" -ForegroundColor Yellow
Write-Host "Report format: $FORMAT" -ForegroundColor Yellow
Write-Host ""

# Check if data directory exists
if (-not (Test-Path $DATA_DIR)) {
    Write-Host "ERROR: Data directory not found: $DATA_DIR" -ForegroundColor Red
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $OUTPUT_DIR)) {
    New-Item -ItemType Directory -Path $OUTPUT_DIR -Force | Out-Null
    Write-Host "Created output directory: $OUTPUT_DIR" -ForegroundColor Green
}

# Run batch analysis
Write-Host "Starting batch analysis..." -ForegroundColor Green
Write-Host ""

python batch_analyze.py `
    --directory "$DATA_DIR" `
    --output "$OUTPUT_DIR" `
    --format $FORMAT `
    --verbose

# Check exit code
if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "="*70 -ForegroundColor Green
    Write-Host "All files processed successfully!" -ForegroundColor Green
    Write-Host "="*70 -ForegroundColor Green
} elseif ($LASTEXITCODE -eq 2) {
    Write-Host ""
    Write-Host "="*70 -ForegroundColor Yellow
    Write-Host "Some files failed - check summary above" -ForegroundColor Yellow
    Write-Host "="*70 -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "="*70 -ForegroundColor Red
    Write-Host "Batch processing failed" -ForegroundColor Red
    Write-Host "="*70 -ForegroundColor Red
}

exit $LASTEXITCODE
