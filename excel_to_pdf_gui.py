#!/usr/bin/env python3
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import importlib.util
from packaging import version
import re

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

        # 한글 또는 공백이 있는 경우 안전한 파일명으로 처리
        if "통합" in output_filename or ' ' in output_filename or any(ord(c) > 127 for c in output_filename):
            # 공백을 언더스코어로 대체하고 안전한 파일명 생성
            safe_name = base_filename.replace(' ', '_')
            # 특수문자 처리
            safe_name = ''.join(c if c.isalnum() or c in '_-.' else '_' for c in safe_name)
            output_filename = safe_name + '.pdf'

        if output_dir:
            base_output_path = os.path.join(output_dir, output_filename)
        else:
            base_output_path = os.path.join(os.path.dirname(excel_path), output_filename)
            
        # Check if file already exists and create a unique name if needed
        output_path = base_output_path
        counter = 1
        
        while os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            if self.overwrite_var.get():
                # 덮어쓰기가 켜져 있으면 기존 경로 사용
                break
            
            # 파일이 존재하면 새 이름 생성 (카운터로)
            # 한글 또는 공백이 있는 경우 안전한 파일명으로 처리
            if "통합" in base_filename or ' ' in base_filename or any(ord(c) > 127 for c in base_filename):
                safe_base = base_filename.replace(' ', '_')
                # 특수문자 처리
                safe_base = ''.join(c if c.isalnum() or c in '_-.' else '_' for c in safe_base)
                new_filename = f"{safe_base}_{counter}.pdf"
            else:
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
            # 임시 파일 생성 - 한글 파일명에서 띄어쓰기를 언더스코어로 대체
            import tempfile
            import shutil
            import uuid
            
            temp_dir = tempfile.gettempdir()
            
            # 파일명에서 띄어쓰기를 언더스코어로 대체하고 안전한 임시 파일명 생성
            filename = os.path.basename(excel_path)
            
            # 특수문자 및 공백 처리
            safe_filename = filename.replace(' ', '_')
            safe_filename = ''.join(c if c.isalnum() or c in '_-.' else '_' for c in safe_filename)
            
            # 여전히 한글이나 비ASCII 문자가 있으면 UUID 사용
            if any(ord(c) > 127 for c in safe_filename):
                temp_excel_path = os.path.join(temp_dir, f"excel_file_{uuid.uuid4().hex[:8]}{file_ext}")
            else:
                temp_excel_path = os.path.join(temp_dir, safe_filename)
            
            try:
                # 원본 파일의 복사본 생성
                shutil.copy2(excel_path, temp_excel_path)
                # 이 복사본 사용
                excel_path_to_use = temp_excel_path
                self.log(f"한글 파일명용 임시 파일 생성: {os.path.basename(excel_path)}", "info")
            except Exception as e:
                # 복사 실패 시 원본 경로 사용
                excel_path_to_use = excel_path
                self.log(f"경고: 임시 파일을 생성할 수 없습니다: {str(e)}", "warning")
        else:
            # 특수 문자가 없으면 원본 경로 사용
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
                # excel_path_to_use 경로가 공백을 포함하거나 특수문자를 포함하는 경우 
                # 큰따옴표로 감싸서 안전하게 처리
                if ' ' in excel_path_to_use or any(c in excel_path_to_use for c in '()[]{}!@#$%^&*'):
                    # 이미 따옴표가 있으면 제거 후 새로 추가
                    if excel_path_to_use.startswith('"') and excel_path_to_use.endswith('"'):
                        safe_excel_path = excel_path_to_use
                    else:
                        safe_excel_path = f'"{excel_path_to_use.replace("\"", "")}"'
                else:
                    safe_excel_path = excel_path_to_use
                
                wb = excel.Workbooks.Open(
                    safe_excel_path,  # 안전하게 처리된 파일 경로 사용
                    UpdateLinks=0,        # Don't update links
                    ReadOnly=True,        # Open in read-only mode
                    IgnoreReadOnlyRecommended=True,  # Ignore read-only recommendation
                    Notify=False,         # Don't notify about updates
                    CorruptLoad=1         # Try to open even if file might be corrupt
                )
            except com_error as e:
                # 첫 번째 시도가 실패하면 원래 경로로 다시 시도
                try:
                    wb = excel.Workbooks.Open(
                        excel_path_to_use,  # 원래 경로 사용
                        UpdateLinks=0,
                        ReadOnly=True,
                        IgnoreReadOnlyRecommended=True,
                        Notify=False,
                        CorruptLoad=1
                    )
                    self.log("대체 방법으로 Excel 파일 열기 성공", "info")
                except com_error as e:
                    return False, f"Failed to open Excel file: {str(e)}"
            
            # Save as PDF using SaveAs method - 큰따옴표로 경로 감싸기
            try:
                # Use SaveAs method to save as PDF
                # 경로에 공백이 있을 경우를 처리하기 위해 큰따옴표로 감싸기
                # 경로에서 이미 있을 수 있는 따옴표 제거 후 다시 추가
                safe_output_path = f'"{output_path.replace("\"", "")}"'
                
                # Excel에 경로 전달 시 따옴표로 감싼 경로 사용
                wb.SaveAs(
                    safe_output_path,
                    FileFormat=57  # 57 is the file format code for PDF
                )
            except com_error as e:
                # 명시적으로 com_error_msg 변수 정의
                com_error_msg = f"COM Error during SaveAs: {str(e)}"
                # 에러를 로그에 기록하지만 즉시 반환하지 않음
                self.log(f"경고: COM 오류가 발생했습니다: {com_error_msg}", "warning")
                
                # 추가 시도: 따옴표 없이 원래 경로로 시도
                try:
                    wb.SaveAs(
                        output_path,
                        FileFormat=57
                    )
                    self.log("대체 방법으로 PDF 저장 시도", "info")
                except:
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
                # PDF가 성공적으로 생성됨
                
                # URL 인코딩 문제 해결을 위한 추가 처리
                # %20과 같은 URL 인코딩된 문자가 있는지 확인
                pdf_filename = os.path.basename(output_path)
                if '%' in pdf_filename:
                    # URL 디코딩 시도
                    try:
                        import urllib.parse
                        decoded_filename = urllib.parse.unquote(pdf_filename)
                        decoded_path = os.path.join(os.path.dirname(output_path), decoded_filename)
                        
                        # 디코딩된 이름으로 파일 복사 시도
                        if os.path.exists(output_path) and not os.path.exists(decoded_path):
                            import shutil
                            try:
                                shutil.copy2(output_path, decoded_path)
                                # 원본 삭제
                                os.remove(output_path)
                                # 출력 경로 업데이트
                                output_path = decoded_path
                                self.log(f"URL 인코딩된 PDF 파일명 수정됨: {decoded_filename}", "info")
                            except:
                                self.log(f"URL 인코딩된 파일명 수정 실패, 원본 유지: {pdf_filename}", "warning")
                    except:
                        self.log(f"PDF 생성됨 (URL 인코딩됨): {pdf_filename}", "warning")
                
                # 한글 또는 특수 문자가 있는 파일명 처리
                if original_has_korean:
                    # 한글 파일명 표시
                    try:
                        import urllib.parse
                        decoded_name = urllib.parse.unquote(os.path.basename(excel_path))
                        self.log(f"✓ 한글 파일명 변환 성공: {decoded_name}", "success")
                    except:
                        self.log(f"✓ 한글 파일명 변환 성공: {os.path.basename(excel_path)}", "success")
                else:
                    # 일반 파일명 성공 메시지
                    self.log(f"✓ 변환 성공: {os.path.basename(excel_path)}", "success")
                
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