@echo off

:: Check if PyInstaller is installed, and install it if not
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
) else (
    echo PyInstaller is already installed.
)

:: Create the executable using PyInstaller
echo Creating executable...
pyinstaller --onefile --windowed excel_to_pdf_gui.py

:: Check if the executable was created successfully
if exist "dist\\excel_to_pdf_gui.exe" (
    echo Executable created successfully: dist\\excel_to_pdf_gui.exe
) else (
    echo Failed to create executable.
)