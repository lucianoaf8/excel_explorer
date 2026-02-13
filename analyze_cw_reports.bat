@echo off
REM Batch analyze ComplyWorks Client Value Reports
REM Analyzes all CW monthly summary reports

SET DATA_DIR=C:\Projects\contractor_analysis\data\complyworks\cw_monthly_raw_files
SET OUTPUT_DIR=C:\Projects\contractor_analysis\reports\cw_analysis
SET FORMAT=markdown

echo ======================================================================
echo ComplyWorks Report Batch Analyzer
echo ======================================================================
echo.
echo Data directory: %DATA_DIR%
echo Output directory: %OUTPUT_DIR%
echo Report format: %FORMAT%
echo.

REM Check if data directory exists
if not exist "%DATA_DIR%" (
    echo ERROR: Data directory not found: %DATA_DIR%
    exit /b 1
)

REM Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
    echo Created output directory: %OUTPUT_DIR%
)

echo Starting batch analysis...
echo.

REM Run batch analysis
python batch_analyze.py --directory "%DATA_DIR%" --output "%OUTPUT_DIR%" --format %FORMAT% --verbose

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ======================================================================
    echo All files processed successfully!
    echo ======================================================================
) else if %ERRORLEVEL% EQU 2 (
    echo.
    echo ======================================================================
    echo Some files failed - check summary above
    echo ======================================================================
) else (
    echo.
    echo ======================================================================
    echo Batch processing failed
    echo ======================================================================
)

pause
