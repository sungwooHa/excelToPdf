#!/usr/bin/env python3
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import importlib.util
from packaging import version

# Check for required packages
def check_package(package_name, min_version=None):
    try:
        package_spec = importlib.util.find_spec(package_name)
        if package_spec is None:
            return False
        
        if min_version:
            pkg = importlib.import_module(package_name)
            pkg_version = getattr(pkg, '__version__', '0.0.0')
            return version.parse(pkg_version) >= version.parse(min_version)
        return True
    except ImportError:
        return False

# Check for pywin32
try:
    import win32com.client
    from pywintypes import com_error
except ImportError:
    # Show error message with tkinter if possible
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing Dependencies", 
            "The pywin32 package is required but not installed.\n\n"
            "Please install it using pip:\n"
            "pip install pywin32>=228"
        )
        root.destroy()
    except:
        print("ERROR: The pywin32 package is required but not installed.")
        print("Please install it using pip: pip install pywin32>=228")
    sys.exit(1)

class ExcelToPdfConverter:
    def __init__(self, master):
        self.master = master
        master.title("Excel to PDF Converter")
        master.geometry("800x700")  # Increased window size to provide more space for log panel
        master.resizable(True, True)
        
        # Set styles
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#4CAF50")
        self.style.configure("TLabel", padding=6)
        self.style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        
        # Try to set a theme if available (makes the UI look better)
        try:
            available_themes = self.style.theme_names()
            if 'vista' in available_themes:
                self.style.theme_use('vista')
            elif 'clam' in available_themes:
                self.style.theme_use('clam')
        except:
            pass
        
        # Create main frame with more padding
        main_frame = ttk.Frame(master, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input section with a frame for better visual grouping
        input_section = ttk.LabelFrame(main_frame, text="Input Files", padding="10")
        input_section.grid(column=0, row=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_section.columnconfigure(0, weight=1)
        
        # File selection frame
        file_frame = ttk.Frame(input_section)
        file_frame.grid(column=0, row=0, sticky=(tk.W, tk.E), pady=(0, 5))
        file_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(file_frame)
        self.input_entry.grid(column=0, row=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        browse_button = ttk.Button(file_frame, text="Browse Files", command=self.browse_files)
        browse_button.grid(column=1, row=0)
        
        browse_dir_button = ttk.Button(file_frame, text="Browse Folder", command=self.browse_directory)
        browse_dir_button.grid(column=2, row=0, padx=(5, 0))
        
        # Options frame for input-related options
        input_options = ttk.Frame(input_section)
        input_options.grid(column=0, row=1, sticky=(tk.W, tk.E))
        
        self.recursive_var = tk.BooleanVar(value=False)
        self.recursive_check = ttk.Checkbutton(input_options, text="Retrieve Subfolders", variable=self.recursive_var, command=self.toggle_output_settings)
        self.recursive_check.grid(column=0, row=0, sticky=tk.W)
        
        # Output section with a frame for better visual grouping
        output_section = ttk.LabelFrame(main_frame, text="Output Settings", padding="10")
        output_section.grid(column=0, row=1, sticky=(tk.W, tk.E), pady=(0, 10))
        output_section.columnconfigure(0, weight=1)
        
        # Output folder selection frame
        output_frame = ttk.Frame(output_section)
        output_frame.grid(column=0, row=0, sticky=(tk.W, tk.E), pady=(0, 5))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame)
        self.output_entry.grid(column=0, row=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        output_button = ttk.Button(output_frame, text="Browse", command=self.browse_output)
        output_button.grid(column=1, row=0)
        self.output_button = output_button  # Store reference to the output button
        
        # Output options frame
        output_options = ttk.Frame(output_section)
        output_options.grid(column=0, row=1, sticky=(tk.W, tk.E))
        
        self.overwrite_var = tk.BooleanVar(value=False)
        overwrite_check = ttk.Checkbutton(output_options, text="Overwrite existing PDF files", variable=self.overwrite_var)
        overwrite_check.grid(column=0, row=0, sticky=tk.W)
        
        # Convert button with more emphasis
        convert_button = ttk.Button(main_frame, text="Convert to PDF", command=self.start_conversion)
        convert_button.grid(column=0, row=2, pady=(0, 10))
        
        # Progress section with a frame for better visual grouping
        progress_section = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_section.grid(column=0, row=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_section.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(progress_section, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.grid(column=0, row=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(progress_section, textvariable=self.status_var)
        status_label.grid(column=0, row=1, sticky=tk.W)
        
        # Log section with a frame for better visual grouping
        log_section = ttk.LabelFrame(main_frame, text="Conversion Log", padding="10")
        log_section.grid(column=0, row=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_section.columnconfigure(0, weight=1)
        log_section.rowconfigure(0, weight=1)
        
        # Create log text widget with scrollbar
        log_frame = ttk.Frame(log_section)
        log_frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Configure the log text widget with enhanced styling
        self.log_text = tk.Text(
            log_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=20,
            font=('Consolas', 10),  # Monospaced font for better readability
            bg='#f5f5f5',           # Light background
            padx=5,                  # Padding for text
            pady=5
        )
        self.log_text.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Define tags for different message types
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("info", foreground="blue")
        self.log_text.tag_configure("timestamp", foreground="gray")
        self.log_text.tag_configure("warning", foreground="orange")
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(column=1, row=0, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=3)  # Log section takes more space
        
        # Store data
        self.files = []
        self.is_converting = False
        
    def browse_files(self):
        """Open file dialog to select multiple Excel files"""
        filetypes = (
            ('Excel files', '*.xlsx *.xls *.xlsm'),
            ('All files', '*.*')
        )
        
        files = filedialog.askopenfilenames(
            title='Select Excel Files',
            initialdir='/',
            filetypes=filetypes
        )
        
        if files:
            self.files = list(files)
            if len(self.files) <= 5:
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, "; ".join(self.files))
            else:
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, f"Selected {len(self.files)} Excel files")
    
    def browse_directory(self):
        """Open directory dialog to select a folder with Excel files"""
        directory = filedialog.askdirectory(
            title='Select Folder with Excel Files',
            initialdir='/'
        )
        
        if directory:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, directory)
            
            # Store directory path for later processing
            self.files = [directory]
    
    def browse_output(self):
        """Open directory dialog to select output folder"""
        directory = filedialog.askdirectory(
            title='Select Output Folder',
            initialdir='/'
        )
        
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, directory)
    
    def log(self, message, message_type="info"):
        """
        Add message to log widget and scroll to end
        
        Args:
            message (str): The message to log
            message_type (str): Type of message - "info", "success", "error"
        """
        # Insert text with timestamp
        import datetime
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S] ")
        
        # Insert timestamp with gray color
        self.log_text.insert(tk.END, timestamp, "timestamp")
        
        # Detect message type if not specified
        if message_type == "info" and message.startswith("✓"):
            message_type = "success"
        elif message_type == "info" and message.startswith("✗"):
            message_type = "error"
        
        # Insert message with appropriate tag
        self.log_text.insert(tk.END, message + "\n", message_type)
        
        # Auto-scroll to end
        self.log_text.see(tk.END)
        
        # Update the GUI immediately to show the new log entry
        self.master.update_idletasks()
        
    def update_status(self, message, progress_value=None):
        """Update status text and progress bar"""
        self.status_var.set(message)
        if progress_value is not None:
            self.progress['value'] = progress_value
        self.master.update_idletasks()
    
    def convert_excel_to_pdf(self, excel_path, output_dir=None):
        """
        Convert an Excel file to PDF using the Excel COM object
        
        Args:
            excel_path (str): Path to the Excel file
            output_dir (str, optional): Directory to save the PDF file
        
        Returns:
            tuple: (success, result_or_error_message)
        """
        # Get absolute paths
        excel_path = os.path.abspath(excel_path)
        
        # Check if the Excel file exists
        if not os.path.exists(excel_path):
            return False, f"Excel file not found: {excel_path}"
        
        # Check if the file is actually an Excel file
        file_ext = os.path.splitext(excel_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls', '.xlsm']:
            return False, f"Not an Excel file: {excel_path}"
        
        # Generate output path with duplicate handling
        base_filename = os.path.basename(os.path.splitext(excel_path)[0])
        output_filename = base_filename + '.pdf'
        
        if output_dir:
            base_output_path = os.path.join(output_dir, output_filename)
        else:
            base_output_path = os.path.join(os.path.dirname(excel_path), output_filename)
            
        # Check if file already exists and create a unique name if needed
        output_path = base_output_path
        counter = 1
        
        while os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            if self.overwrite_var.get():
                # If overwrite is enabled, just use the existing path
                break
            # File exists, create a new name with counter
            new_filename = f"{base_filename}_{counter}.pdf"
            
            if output_dir:
                output_path = os.path.join(output_dir, new_filename)
            else:
                output_path = os.path.join(os.path.dirname(excel_path), new_filename)
                
            counter += 1
            
        # Now output_path contains a unique filename that doesn't exist or is empty
        
        # Create output directory if it doesn't exist
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
        except PermissionError:
            return False, f"Permission denied when creating output directory"
        except Exception as e:
            return False, f"Failed to create output directory: {str(e)}"
        
        # Check if the output PDF path is writable
        try:
            # Try to create an empty file to check permissions
            with open(output_path, 'a'):
                pass
            os.remove(output_path)  # Remove the test file
        except PermissionError:
            return False, f"Permission denied when writing to output file"
        except Exception as e:
            return False, f"Cannot write to output file: {str(e)}"
            
        # Special handling for Korean filenames or other non-ASCII characters
        temp_excel_path = None
        original_has_korean = "통합" in excel_path or any(ord(c) > 127 for c in excel_path)
        
        if original_has_korean:
            # Create a temporary copy with a unique English filename
            import tempfile
            import shutil
            import uuid
            
            temp_dir = tempfile.gettempdir()
            temp_excel_path = os.path.join(temp_dir, f"excel_file_{uuid.uuid4().hex[:8]}{file_ext}")
            
            try:
                # Create a copy of the original file with a unique English name
                shutil.copy2(excel_path, temp_excel_path)
                # Use this copy for conversion
                excel_path_to_use = temp_excel_path
                self.log(f"Created temporary copy for Korean filename: {os.path.basename(excel_path)}", "info")
            except Exception as e:
                # If copying fails, use the original path
                excel_path_to_use = excel_path
                self.log(f"Warning: Could not create temporary file: {str(e)}", "error")
        else:
            # Use the original path if no special characters
            excel_path_to_use = excel_path
        
        excel = None
        wb = None
        
        try:
            # Create Excel COM object with enhanced error handling
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
                return False, f"Failed to create Excel application: {str(e)}"
            
            # Open the workbook with additional parameters for reliability
            try:
                wb = excel.Workbooks.Open(
                    excel_path_to_use,  # Use the encoded filename if created
                    UpdateLinks=0,        # Don't update links
                    ReadOnly=True,        # Open in read-only mode
                    IgnoreReadOnlyRecommended=True,  # Ignore read-only recommendation
                    Notify=False,         # Don't notify about updates
                    CorruptLoad=1         # Try to open even if file might be corrupt
                )
            except com_error as e:
                return False, f"Failed to open Excel file: {str(e)}"
            
            # Save as PDF using SaveAs method
            try:
                # Use SaveAs method to save as PDF
                wb.SaveAs(
                    output_path,
                    FileFormat=57  # 57 is the file format code for PDF
                )
            except com_error as e:
                # Capture the error but don't immediately return
                # We'll check if the PDF was actually created despite the error
                com_error_msg = f"COM Error during SaveAs: {str(e)}"
                pass
            
            # Close the workbook
            try:
                wb.Close(False)
                excel.Quit()
            except:
                pass
            
            # Clean up the temporary file if we created one
            if temp_excel_path and os.path.exists(temp_excel_path):
                try:
                    os.remove(temp_excel_path)
                except:
                    pass
            
            # Always check if the PDF exists and is valid, even if there were COM errors
            # This handles cases where files with Korean names show COM errors but still create valid PDFs
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                # PDF was successfully created, with or without COM errors
                if original_has_korean:
                    # Properly decode Korean filename for display
                    import urllib.parse
                    try:
                        decoded_name = urllib.parse.unquote(os.path.basename(excel_path))
                        # Check if the output filename is URL encoded
                        output_basename = os.path.basename(output_path)
                        if '%' in output_basename:
                            self.log(f"Warning: PDF created but filename may be URL encoded: {output_basename}", "warning")
                            self.log(f"✓ Successfully converted Korean filename: {decoded_name}", "success")
                        else:
                            self.log(f"✓ Successfully converted Korean filename: {decoded_name}", "success")
                    except:
                        # If decoding fails, show warning about filename display
                        self.log(f"Warning: PDF created but filename may not display correctly: {os.path.basename(output_path)}", "warning")
                        self.log(f"✓ Successfully converted Korean filename: {os.path.basename(excel_path)}", "success")
                    
                # Check if output_path is different from the original expected path (duplicate handling)
                original_filename = os.path.splitext(excel_path)[0]
                # Decode URL encoding if present for better display
                import urllib.parse
                try:
                    decoded_original = urllib.parse.unquote(original_filename)
                except:
                    decoded_original = original_filename
                
                expected_pdf_name = os.path.basename(decoded_original) + '.pdf'
                expected_path = os.path.join(os.path.dirname(output_path), expected_pdf_name)
                
                if output_path != expected_path:
                    # Show proper decoded filename in the log
                    try:
                        decoded_output = urllib.parse.unquote(os.path.basename(output_path))
                    except:
                        decoded_output = os.path.basename(output_path)
                        
                    self.log(f"Created with unique name to avoid overwriting: {decoded_output}", "info")
                    
                return True, output_path
            else:
                # If we captured a COM error earlier but no PDF was created, return that error
                if 'com_error_msg' in locals():
                    return False, com_error_msg
                else:
                    if original_has_korean:
                        # Get properly decoded filename for error message
                        import urllib.parse
                        try:
                            decoded_name = urllib.parse.unquote(os.path.basename(excel_path))
                        except:
                            decoded_name = os.path.basename(excel_path)
                            
                        return False, f"Failed to convert Korean filename: {decoded_name}. Please try renaming the file to use English characters."
                    else:
                        return False, "PDF file was not created successfully"
        
        except com_error as e:
            error_code = e.excepinfo[5] if hasattr(e, 'excepinfo') and len(e.excepinfo) > 5 else None
            error_msg = f"COM Error: {str(e)}"
            if error_code:
                error_msg += f" (Error code: {error_code})"
            return False, error_msg
        
        except Exception as e:
            return False, f"Unexpected error: {str(e)}"
        
        finally:
            # Make sure Excel gets closed even if there's an error
            try:
                if wb:
                    wb.Close(False)
                if excel:
                    excel.Quit()
            except:
                pass
    
    def process_files(self, files, output_dir, recursive):
        """Process the list of Excel files or directories"""
        excel_files = []
        
        # Modify the process_files method to only decide whether to include subfolders
        if len(files) == 1 and os.path.isdir(files[0]):
            directory = files[0]
            path = Path(directory)
            
            # Collect all Excel files
            if recursive:
                for ext in ['.xlsx', '.xls', '.xlsm']:
                    excel_files.extend(path.glob(f'**/*{ext}'))
            else:
                for ext in ['.xlsx', '.xls', '.xlsm']:
                    excel_files.extend(path.glob(f'*{ext}'))
                    
            excel_files = [str(f) for f in excel_files]
        else:
            # We already have a list of files
            excel_files = files
        
        # Now process all Excel files
        total_files = len(excel_files)
        successful = 0
        failed = 0
        failed_files = []
        
        for i, file_path in enumerate(excel_files):
            progress = int((i / total_files) * 100)
            self.update_status(f"Converting {i+1}/{total_files}: {os.path.basename(file_path)}", progress)
            
            success, result = self.convert_excel_to_pdf(file_path, output_dir)
            
            if success:
                self.log(f"✓ Converted: {os.path.basename(file_path)} → {os.path.basename(result)}", "success")
                successful += 1
            else:
                self.log(f"✗ Failed to convert: {os.path.basename(file_path)} - {result}", "error")
                failed += 1
                failed_files.append(file_path)
        
        # Retry failed files
        if failed_files and successful > 0:  # Only retry if at least one file worked
            self.log("\nRetrying failed conversions...", "info")
            
            # Try to kill any lingering Excel processes (Windows only)
            try:
                import subprocess
                subprocess.call("taskkill /f /im excel.exe", shell=True)
            except:
                pass
            
            # Wait a moment before retrying
            import time
            time.sleep(2)
            
            retried_success = 0
            for file_path in failed_files:
                self.update_status(f"Retrying: {os.path.basename(file_path)}")
                
                success, result = self.convert_excel_to_pdf(file_path, output_dir)
                
                if success:
                    self.log(f"✓ Retry successful: {os.path.basename(file_path)} → {os.path.basename(result)}", "success")
                    successful += 1
                    failed -= 1
                    retried_success += 1
                else:
                    self.log(f"✗ Retry failed: {os.path.basename(file_path)} - {result}", "error")
            
            if retried_success > 0:
                self.log(f"\nRetry recovered {retried_success} file(s)", "success")
        
        # Update final status
        self.update_status(f"Completed! {successful} successful, {failed} failed", 100)
        
        # Show troubleshooting tips if there were failures
        if failed > 0:
            self.log("\nTroubleshooting tips for failed conversions:", "info")
            self.log("1. Close all running Excel instances", "info")
            self.log("2. Make sure Excel is not in Protected View mode", "info")
            self.log("3. Check if the Excel files are password-protected", "info")
            self.log("4. Run the application as administrator", "info")
            self.log("5. Check if Excel is properly installed and licensed", "info")
            self.log("6. Check if there's enough disk space for the PDF files", "info")
            self.log("7. For files with Korean names (통합 문서), try renaming to English", "info")
            self.log("\nSee TROUBLESHOOTING.md for more detailed solutions", "info")
        
        self.is_converting = False
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        if self.is_converting:
            messagebox.showwarning("Already Running", "Conversion already in progress!")
            return
            
        if not self.files:
            messagebox.showwarning("No Files", "Please select Excel files or a directory first!")
            return
            
        # Get output directory
        output_dir = self.output_entry.get()
        if output_dir and not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create output directory: {e}")
                return
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Reset progress
        self.progress['value'] = 0
        
        # Update status
        self.is_converting = True
        self.update_status("Starting conversion...")
        
        # Start conversion in a separate thread
        conversion_thread = threading.Thread(
            target=self.process_files,
            args=(self.files, output_dir, self.recursive_var.get())
        )
        conversion_thread.daemon = True
        conversion_thread.start()

    # Add a method to enable/disable output settings based on 'Retrieve Subfolders'
    def toggle_output_settings(self):
        if self.recursive_var.get():
            self.output_entry.config(state='normal')
            self.output_button.config(state='normal')
        else:
            self.output_entry.config(state='disabled')
            self.output_button.config(state='disabled')

def main():
    root = tk.Tk()
    app = ExcelToPdfConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()