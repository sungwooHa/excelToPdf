#!/bin/bash
# Shell script to set up virtual environment and run the Excel to PDF converter

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
pip install -e .

# Run the GUI application
echo "Starting Excel to PDF Converter..."
python excel_to_pdf_gui.py

# Deactivate virtual environment
deactivate