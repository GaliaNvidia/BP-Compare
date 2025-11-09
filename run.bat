@echo off
REM PO Line Comparison Tool - Start Script for Windows

echo.
echo Starting PO Line Comparison Tool...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python 3.8 or higher.
    pause
    exit /b 1
)

REM Check if requirements are installed
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
    echo.
)

REM Run Streamlit app
echo Starting application...
echo The app will open in your default browser
echo.
echo Press Ctrl+C to stop the server
echo.

streamlit run app.py

pause

