#!/usr/bin/env python3
import os
import sys
import argparse
from pathlib import Path
import importlib.util

# Check if pywin32 is installed
try:
    import win32com.client
    from pywintypes import com_error
except ImportError:
    print("ERROR: The pywin32 package is required but not installed.")
    print("Please install it using pip: pip install pywin32>=228")
    sys.exit(1)

def convert_excel_to_pdf(excel_path, output_path=None, verbose=False):
    """
    Convert an Excel file to PDF using the Excel COM object.
    
    Args:
        excel_path (str): Path to the Excel file
        output_path (str, optional): Path to save the PDF file. If not provided,
                                     it will use the same path as the Excel file with .pdf extension.
        verbose (bool, optional): Whether to print detailed error messages
    
    Returns:
        tuple: (success, result_or_error_message)
              success is a boolean indicating whether the conversion succeeded
              result_or_error_message is either the path to the created PDF file or an error message
    """
    # Get absolute paths
    excel_path = os.path.abspath(excel_path)
    
    # If output path not provided, use the same name with .pdf extension (with duplicate handling)
    if not output_path:
        base_filename = os.path.splitext(excel_path)[0]
        base_output_path = base_filename + '.pdf'
        
        # Check if file already exists and create a unique name if needed
        output_path = base_output_path
        counter = 1
        
        while os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            # File exists, create a new name with counter
            output_path = f"{base_filename}_{counter}.pdf"
            counter += 1
    else:
        output_path = os.path.abspath(output_path)
        
        # If an explicit output path was provided, still check for duplicates
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            base_path = os.path.splitext(output_path)[0]
            ext = os.path.splitext(output_path)[1]
            counter = 1
            
            while os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                output_path = f"{base_path}_{counter}{ext}"
                counter += 1
    
    # Check if the Excel file exists
    if not os.path.exists(excel_path):
        error_msg = f"Excel file not found: {excel_path}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    
    # Check if the file is actually an Excel file
    file_ext = os.path.splitext(excel_path)[1].lower()
    if file_ext not in ['.xlsx', '.xls', '.xlsm']:
        error_msg = f"Not an Excel file: {excel_path}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    
    # Try to create the output directory
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
    except PermissionError:
        error_msg = f"Permission denied when creating output directory: {os.path.dirname(output_path)}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    except Exception as e:
        error_msg = f"Failed to create output directory: {str(e)}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    
    # Check if the output PDF path is writable
    try:
        # Try to create an empty file to check permissions
        with open(output_path, 'a'):
            pass
        os.remove(output_path)  # Remove the test file
    except PermissionError:
        error_msg = f"Permission denied when writing to output file: {output_path}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    except Exception as e:
        error_msg = f"Cannot write to output file: {str(e)}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    
    # Special handling for Korean filenames or other non-ASCII characters
    temp_excel_path = None
    original_has_korean = "통합" in excel_path or any(ord(c) > 127 for c in excel_path)
    
    if original_has_korean:
        # Create a temporary copy with an English filename
        import tempfile
        import shutil
        
        temp_dir = tempfile.gettempdir()
        
        # Fix encoding issues with Korean filenames (remove URL encoding like %20)
        import urllib.parse
        basename = os.path.basename(excel_path)
        try:
            # Try to decode URL encoded characters if present
            decoded_basename = urllib.parse.unquote(basename)
            # Replace spaces with underscores
            safe_basename = decoded_basename.replace(' ', '_')
        except:
            # If decoding fails, just use the original with spaces replaced
            safe_basename = basename.replace(' ', '_')
            
        temp_excel_path = os.path.join(temp_dir, f"temp_excel_file_{safe_basename}")
        
        # If the temporary filename still has non-ASCII characters, use a generic name
        if any(ord(c) > 127 for c in temp_excel_path):
            import uuid
            temp_excel_path = os.path.join(temp_dir, f"excel_file_{uuid.uuid4().hex[:8]}{file_ext}")
        
        try:
            # Create a copy of the original file with an English name
            shutil.copy2(excel_path, temp_excel_path)
            # Use this copy for conversion
            excel_path_to_use = temp_excel_path
            if verbose:
                print(f"Created temporary copy for Korean filename: {os.path.basename(excel_path)}")
        except Exception as e:
            # If copying fails, use the original path
            excel_path_to_use = excel_path
            if verbose:
                print(f"Warning: Could not create temporary file: {str(e)}")
    else:
        # Use the original path if no special characters
        excel_path_to_use = excel_path
    
    excel = None
    wb = None
    
    try:
        # Create Excel COM object
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            # Set properties safely
            try:
                excel.DisplayAlerts = False
            except:
                pass
            try:
                excel.Visible = False
            except:
                pass
            # Add additional settings to increase reliability
            try:
                excel.AskToUpdateLinks = False
            except:
                pass
            try:
                excel.EnableEvents = False
            except:
                pass
        except com_error as e:
            error_msg = f"Failed to create Excel application: {str(e)}"
            if verbose:
                print(f"Error: {error_msg}")
            return False, error_msg
        
        # Open the workbook with additional parameters for better reliability
        try:
            # Use additional parameters to handle problematic files
            wb = excel.Workbooks.Open(
                excel_path_to_use,  # Use temporary English filename if created
                UpdateLinks=0,        # Don't update links
                ReadOnly=True,        # Open in read-only mode
                IgnoreReadOnlyRecommended=True,  # Ignore read-only recommendation
                Notify=False,         # Don't notify about updates
                CorruptLoad=1         # Try to open even if file might be corrupt
            )
        except com_error as e:
            error_msg = f"Failed to open Excel file: {str(e)}"
            if verbose:
                print(f"Error: {error_msg}")
                print("This might be due to a password-protected file or file corruption.")
            return False, error_msg
        
        # Save as PDF with a more robust approach
        try:
            # Handle files with Korean characters or special symbols better
            # Try different methods to export as PDF
            
            # Method 1: Using standard export with parameters for better compatibility
            try:
                wb.ExportAsFixedFormat(
                    Type=0,                 # PDF format
                    Filename=output_path,
                    Quality=0,              # Standard quality
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
            except Exception as e1:
                # Method 2: Try exporting active sheet if the whole workbook fails
                try:
                    activeSheet = wb.ActiveSheet
                    activeSheet.ExportAsFixedFormat(
                        Type=0,
                        Filename=output_path,
                        Quality=0,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False
                    )
                except Exception as e2:
                    # Method 3: Try with minimal parameters
                    try:
                        wb.ExportAsFixedFormat(0, output_path)
                    except Exception as e3:
                        # If all methods fail, raise the original error
                        raise e1
                        
        except com_error as e:
            error_msg = f"Failed to export to PDF: {str(e)}"
            if verbose:
                print(f"Error: {error_msg}")
                print("This might be due to insufficient permissions or issues with Excel's PDF export functionality.")
                
                # Don't show Korean character warnings if we're on verbose mode since these typically work fine
                if verbose and "통합" in excel_path or any(ord(c) > 127 for c in excel_path):
                    print("\nSpecial note for files with non-English characters:")
                    print("- These files often work despite showing COM errors")
                    print("- If PDF was created despite the error, you can ignore this message")
            
            return False, error_msg
        
        # Close the workbook
        wb.Close(False)
        excel.Quit()
        
        # Clean up the temporary file if we created one
        if temp_excel_path and os.path.exists(temp_excel_path):
            try:
                os.remove(temp_excel_path)
            except:
                pass
        
        # Always explicitly check if the PDF exists and is valid regardless of previous errors
        # This handles cases where files with Korean names might show COM errors but still create valid PDFs
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            # PDF was successfully created despite possible COM errors
            return True, output_path
        else:
            if original_has_korean:
                error_msg = "Korean filename caused conversion issue. Try renaming the file to use English characters."
            else:
                error_msg = "PDF file was not created successfully"
                
            if verbose:
                print(f"Error: {error_msg}")
            return False, error_msg
    
    except com_error as e:
        error_code = e.excepinfo[5] if hasattr(e, 'excepinfo') and len(e.excepinfo) > 5 else None
        error_msg = f"COM Error: {str(e)}"
        if error_code:
            error_msg += f" (Error code: {error_code})"
        
        if verbose:
            print(f"Error: {error_msg}")
            print("\nTroubleshooting suggestions:")
            print("1. Make sure Excel is installed and properly configured")
            print("2. Check if Excel is not running in protected view")
            print("3. Close any open Excel instances and try again")
            print("4. Run the script with administrative privileges")
        
        return False, error_msg
    
    except Exception as e:
        error_msg = f"Unexpected error: {str(e)}"
        if verbose:
            print(f"Error: {error_msg}")
        return False, error_msg
    
    finally:
        # Make sure Excel gets closed even if there's an error
        try:
            if wb:
                wb.Close(False)
            if excel:
                excel.Quit()
        except:
            pass

def main():
    parser = argparse.ArgumentParser(description='Convert Excel files to PDF')
    parser.add_argument('input', help='Path to Excel file or directory containing Excel files')
    parser.add_argument('-o', '--output', help='Output PDF file or directory')
    parser.add_argument('-r', '--recursive', action='store_true', help='Process directories recursively')
    parser.add_argument('-v', '--verbose', action='store_true', help='Print detailed error messages')
    parser.add_argument('--retry', type=int, default=1, help='Number of retries for failed conversions (default: 1)')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if input_path.is_file():
        # Convert a single file
        if not input_path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
            print(f"Error: {input_path} is not an Excel file")
            sys.exit(1)
        
        success, result = convert_excel_to_pdf(str(input_path), args.output, args.verbose)
        
        if success:
            print(f"Converted {input_path} to {result}")
        else:
            print(f"Failed to convert {input_path}: {result}")
            sys.exit(1)
    
    elif input_path.is_dir():
        # Convert all Excel files in directory
        if args.output and not os.path.isdir(args.output):
            try:
                os.makedirs(args.output, exist_ok=True)
            except Exception as e:
                print(f"Error creating output directory: {e}")
                sys.exit(1)
        
        excel_files = []
        if args.recursive:
            for ext in ['.xlsx', '.xls', '.xlsm']:
                excel_files.extend(input_path.glob(f'**/*{ext}'))
        else:
            for ext in ['.xlsx', '.xls', '.xlsm']:
                excel_files.extend(input_path.glob(f'*{ext}'))
        
        if not excel_files:
            print(f"No Excel files found in {input_path}")
            sys.exit(1)
        
        # Initialize counters
        success_count = 0
        failure_count = 0
        retry_failures = []
        
        for excel_file in excel_files:
            output_file = None
            if args.output:
                rel_path = excel_file.relative_to(input_path)
                output_file = Path(args.output) / rel_path.with_suffix('.pdf')
            
            success, result = convert_excel_to_pdf(
                str(excel_file), 
                str(output_file) if output_file else None,
                args.verbose
            )
            
            if success:
                print(f"✓ Converted {excel_file} to {result}")
                success_count += 1
            else:
                print(f"✗ Failed to convert {excel_file}: {result}")
                failure_count += 1
                retry_failures.append((excel_file, output_file))
        
        # Retry failures if requested
        retried_success = 0
        for retry in range(args.retry):
            if not retry_failures:
                break
                
            if args.verbose:
                print(f"\nRetrying {len(retry_failures)} failed conversions (attempt {retry+1}/{args.retry})...")
            
            still_failing = []
            for excel_file, output_file in retry_failures:
                # Wait before retrying
                import time
                time.sleep(1)
                
                # Try to kill any lingering Excel processes (Windows only)
                try:
                    import subprocess
                    subprocess.call("taskkill /f /im excel.exe", shell=True)
                except:
                    pass
                
                success, result = convert_excel_to_pdf(
                    str(excel_file), 
                    str(output_file) if output_file else None,
                    args.verbose
                )
                
                if success:
                    print(f"✓ Retry succeeded: {excel_file} to {result}")
                    retried_success += 1
                    success_count += 1
                    failure_count -= 1
                else:
                    if args.verbose:
                        print(f"✗ Retry failed: {excel_file}")
                    still_failing.append((excel_file, output_file))
            
            retry_failures = still_failing
        
        # Print summary
        total = success_count + failure_count
        print(f"\nSummary: {success_count}/{total} files converted successfully ({retried_success} after retry)")
        if failure_count > 0:
            print(f"Failed to convert {failure_count} files")
            
            if args.verbose:
                print("\nTroubleshooting suggestions:")
                print("1. Make sure Excel is properly installed and licensed")
                print("2. Close all running Excel instances")
                print("3. Run the script with administrative privileges")
                print("4. Check if files are password-protected or corrupted")
                print("5. Check if there are permission issues with the output directory")
                print("6. Ensure Excel is not set to open files in Protected View")
    
    else:
        print(f"Error: {input_path} does not exist")
        sys.exit(1)

if __name__ == '__main__':
    main()