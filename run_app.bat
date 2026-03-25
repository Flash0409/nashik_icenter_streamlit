@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe"

if exist "%VENV_PYTHON%" (
    set "PYTHON_EXE=%VENV_PYTHON%"
) else (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" -m streamlit run "%SCRIPT_DIR%app.py"

if errorlevel 1 (
    echo.
    echo Failed to start Streamlit. Make sure dependencies are installed.
    pause
)

endlocal
