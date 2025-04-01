# Excel to PDF Converter

A Python utility to convert Excel files to PDF format with both command-line and GUI interfaces.

## Requirements

- Python 3.6+
- pywin32 package (for Windows COM automation)
- tkinter (for GUI version - included in standard Python distribution)
- Microsoft Excel (must be installed on the system)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/excelToPdf.git
   cd excelToPdf
   ```

2. Using Virtual Environment (Recommended):

   **Windows:**
   ```
   # Run the setup batch file
   run.bat
   ```

   **Linux/macOS:**
   ```
   # Make the script executable
   chmod +x run.sh
   
   # Run the setup script
   ./run.sh
   ```
   
   The scripts will:
   - Create a virtual environment
   - Install required dependencies
   - Launch the application

3. Manual Installation:
   ```
   # Create virtual environment
   python -m venv venv
   
   # Activate it (Windows)
   venv\Scripts\activate
   
   # Activate it (Linux/macOS)
   source venv/bin/activate
   
   # Install in development mode
   pip install -e .
   
   # Run the application
   python excel_to_pdf_gui.py
   ```

## Usage

### GUI Version

Launch the graphical user interface:

```
python excel_to_pdf_gui.py
```

The GUI allows you to:
- Select multiple Excel files or a folder containing Excel files
- Set output directory for the PDFs
- Include subfolders when selecting a directory
- Track conversion progress with a progress bar
- View conversion logs in real-time

### Command-Line Version

#### Basic Usage

Convert a single Excel file to PDF:

```
python excel_to_pdf.py path/to/excel/file.xlsx
```

This will create a PDF file with the same name in the same directory.

#### Specify Output Path

```
python excel_to_pdf.py path/to/excel/file.xlsx -o path/to/output/file.pdf
```

#### Convert All Excel Files in a Directory

```
python excel_to_pdf.py path/to/directory/
```

#### Convert All Excel Files in a Directory and Subdirectories

```
python excel_to_pdf.py path/to/directory/ -r
```

#### Specify Output Directory for Batch Conversion

```
python excel_to_pdf.py path/to/directory/ -o path/to/output/directory/
```

#### Enable Verbose Mode for Detailed Error Messages

```
python excel_to_pdf.py path/to/excel/file.xlsx -v
```

#### Specify Number of Retries for Failed Conversions

```
python excel_to_pdf.py path/to/directory/ --retry 3
```

## Troubleshooting Common Issues

If conversion fails, here are some common issues and solutions:

1. **Excel Process Issues**
   - Close all running Excel instances before conversion
   - Ensure Excel is not running in protected view mode
   - Try running the script with administrator privileges

2. **File/Permission Problems**
   - Check if Excel files are password-protected
   - Ensure you have write permissions to the output directory
   - Verify file paths don't contain special characters

3. **Excel Configuration**
   - Make sure Excel is properly installed and licensed
   - Check if Excel is configured to allow automation
   - Verify Excel can manually export to PDF

4. **System Resources**
   - Ensure you have enough disk space
   - Close other applications to free up memory
   - Try converting fewer files at once

## Features

- Convert Excel files (.xlsx, .xls, .xlsm) to PDF format
- Process multiple files at once through batch conversion
- Automatically handle Korean/non-ASCII filenames
- Create unique filenames to avoid overwriting existing PDFs
- Visual progress tracking with color-coded logs
- Detailed error messages and troubleshooting information

## Notes

- This script requires Microsoft Excel to be installed on the system as it uses Excel's COM interface.
- The script works on Windows only due to the dependency on the Windows COM interface.
- Make sure Excel is not running in protected view mode for the files being converted.
- The GUI interface provides real-time conversion status with detailed logging.
- Files with Korean or non-ASCII characters in their filenames are automatically handled through temporary file creation.
- If a PDF with the same name already exists, a new unique filename (with _1, _2, etc. suffix) will be created instead of overwriting.

## License

See the [LICENSE](LICENSE) file for details.