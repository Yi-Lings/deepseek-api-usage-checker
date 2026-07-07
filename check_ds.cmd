@echo off
REM DeepSeek API Usage Checker -- Windows Launcher
REM Double-click or run from PowerShell/CMD.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Auto-detect Git Bash and set PATH for coreutils (dirname, grep, etc.)
set "BASH="
where bash >nul 2>&1 && set "BASH=bash"
if not defined BASH (
    if exist "D:\Git\usr\bin\bash.exe" set "BASH=D:\Git\usr\bin\bash.exe"
)
if not defined BASH (
    if exist "%ProgramFiles%\Git\usr\bin\bash.exe" set "BASH=%ProgramFiles%\Git\usr\bin\bash.exe"
)
if not defined BASH (
    if exist "%ProgramFiles(x86)%\Git\usr\bin\bash.exe" set "BASH=%ProgramFiles(x86)%\Git\usr\bin\bash.exe"
)

if not defined BASH (
    echo Error: Git Bash not found.
    echo Install from: https://git-scm.com/downloads
    echo.
    pause
    exit /b 1
)

REM Add Git usr/bin to PATH so bash finds dirname, grep, etc.
for %%i in ("%BASH%") do set "BASH_DIR=%%~dpi"
set "PATH=%BASH_DIR%;%PATH%"

"%BASH%" "%SCRIPT_DIR%check_ds.sh" %*
set EXIT_CODE=%ERRORLEVEL%
echo.
if %EXIT_CODE% NEQ 0 (
    echo Error: script exited with code %EXIT_CODE%
)
pause
