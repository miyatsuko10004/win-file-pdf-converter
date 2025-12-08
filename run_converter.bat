@echo off
chcp 65001 > nul
cd /d %~dp0

echo ==========================================
echo PDF Converter Execution
echo ==========================================

uv run converter.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] An error occurred during execution.
) else (
    echo.
    echo [SUCCESS] Execution completed successfully.
)

echo.
pause
