@echo off
REM Batch file to set up virtual environment and run the Excel to PDF converter

REM Create virtual environment if it doesn't exist
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate

REM Install dependencies
echo Installing dependencies...
pip install -e .

REM Run the GUI application
echo Starting Excel to PDF Converter...
python excel_to_pdf_gui.py

REM Deactivate virtual environment
call deactivate