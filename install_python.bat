@echo off
setlocal EnableDelayedExpansion

:: Config
set "PYTHON_VERSION=3.11.9"
set "PYTHON_SHORT=311"
set "INSTALLER_NAME=python-%PYTHON_VERSION%-amd64.exe"
set "INSTALLER_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%INSTALLER_NAME%"
set "INSTALLER_PATH=%TEMP%\%INSTALLER_NAME%"
set "TARGET_DIR=%USERPROFILE%\AppData\Local\Programs\Python\Python%PYTHON_SHORT%"
set "LOG_FILE=%TEMP%\python_install_log.txt"

:: Clean up log
if exist "%LOG_FILE%" del /f /q "%LOG_FILE%"

:: Check if already installed
if exist "%TARGET_DIR%\python.exe" (
    echo Python %PYTHON_VERSION% already installed at: %TARGET_DIR%
    "%TARGET_DIR%\python.exe" --version
    goto :EOF
)

echo Downloading Python %PYTHON_VERSION% installer...

:: Download fallback logic
:download
if exist "%INSTALLER_PATH%" del /f /q "%INSTALLER_PATH%"

:: Try curl
curl --version >nul 2>&1
if !ERRORLEVEL! == 0 (
    curl -L -o "%INSTALLER_PATH%" "%INSTALLER_URL%" >>"%LOG_FILE%" 2>&1
) else (
    :: Try powershell
    powershell -Command "try { Invoke-WebRequest -Uri '%INSTALLER_URL%' -OutFile '%INSTALLER_PATH%' -UseBasicParsing } catch { exit 1 }" >>"%LOG_FILE%" 2>&1
    if not exist "%INSTALLER_PATH%" (
        :: Try bitsadmin
        bitsadmin /transfer pythonInstaller /priority high "%INSTALLER_URL%" "%INSTALLER_PATH%" >>"%LOG_FILE%" 2>&1
    )
)

if not exist "%INSTALLER_PATH%" (
    echo [ERROR] Failed to download installer. Check your internet connection.
    echo See log: %LOG_FILE%
    goto :EOF
)

echo Installing Python silently to: %TARGET_DIR%

:: Run installer
"%INSTALLER_PATH%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 TargetDir="%TARGET_DIR%" >>"%LOG_FILE%" 2>&1

:: Check install
if exist "%TARGET_DIR%\python.exe" (
    echo Installation successful.
    "%TARGET_DIR%\python.exe" --version
) else (
    echo [ERROR] Python installation failed. See log: %LOG_FILE%
)

:: Cleanup installer
if exist "%INSTALLER_PATH%" del /f /q "%INSTALLER_PATH%"

endlocal
