#!/usr/bin/env python3
"""
Quip Exporter GUI - Advanced interface for exporting Quip folders with comprehensive error handling

Features:
- 5-attempt retry system for maximum reliability
- Automatic error reporting with detailed failure analysis
- Real-time logging with timestamps
- Smart format selection (DOCX/XLSX/PDF)
- Recursive folder traversal with structure preservation
- Comprehensive error tracking and recovery information
- User-friendly GUI with progress monitoring
- Automatic log file generation
- Network resilience and API rate limit handling

Version: 2.0 - Enhanced Reliability
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import time
from pathlib import Path
import re
import webbrowser

# Import our native export functionality
from native_export import extract_folder_id_from_url, export_folder_recursive, QuipExporter

class QuipExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Quip Folder Exporter")
        self.root.geometry("700x720")
        self.root.resizable(True, True)
        
        # Variables
        self.quip_token = tk.StringVar()
        self.quip_url = tk.StringVar()
        self.target_folder = tk.StringVar(value=str(Path.home() / "Downloads" / "QuipExports"))
        self.is_exporting = False
        self.start_time = None
        self.timer_job = None
        self.log_file_path = None
        self.log_messages = []
        self.failed_files = []  # Track files that failed to export
        self.unknown_folders = []  # Track folders with unknown names
        self.export_stats = {'folders': 0, 'documents': 0, 'failed_files': 0, 'unknown_folders': 0}
        self.failed_folders = []  # Track folders that failed to get proper names
        self.export_stats = {'folders': 0, 'documents': 0, 'failed_folders': 0, 'failed_documents': 0}  # Track export statistics
        self.stop_requested = False  # Flag to track if user requested to stop export
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(3, weight=0)  # For scrollbar
        
        # Title
        title_label = ttk.Label(main_frame, text="Quip Folder Exporter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Token instructions with clickable link (single line)
        token_frame = ttk.Frame(main_frame)
        token_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(token_frame, text="Get your API token from ", 
                 foreground="black", font=("Arial", 9)).pack(side=tk.LEFT)
        
        token_link = ttk.Label(token_frame, text="https://quip-amazon.com/dev/token", 
                              foreground="blue", font=("Arial", 9), cursor="hand2")
        token_link.pack(side=tk.LEFT)
        token_link.bind("<Button-1>", lambda e: webbrowser.open("https://quip-amazon.com/dev/token"))
        
        ttk.Label(token_frame, text=" and paste it below:", 
                 foreground="black", font=("Arial", 9)).pack(side=tk.LEFT)
        
        # Quip Token input
        ttk.Label(main_frame, text="Quip API Token:").grid(row=2, column=0, sticky=tk.W, pady=5)
        token_entry = ttk.Entry(main_frame, textvariable=self.quip_token, width=50, show="*")
        token_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        # Quip URL input
        ttk.Label(main_frame, text="Quip Folder URL:").grid(row=3, column=0, sticky=tk.W, pady=(15, 5))
        url_entry = ttk.Entry(main_frame, textvariable=self.quip_url, width=50)
        url_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        # URL help
        url_help = ttk.Label(main_frame, text="Example: https://quip-amazon.com/ABC123/FolderName", 
                            foreground="gray", font=("Arial", 9))
        url_help.grid(row=4, column=1, sticky=tk.W, padx=(10, 0))
        
        # Target folder input
        ttk.Label(main_frame, text="Save to Folder:").grid(row=5, column=0, sticky=tk.W, pady=(5, 5))
        folder_entry = ttk.Entry(main_frame, textvariable=self.target_folder, width=50)
        folder_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=(5, 5))
        
        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_folder)
        browse_btn.grid(row=5, column=2, padx=(5, 0), pady=(5, 5))
        
        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Export button
        self.export_btn = ttk.Button(btn_frame, text="🚀 Start Export", 
                                    command=self.start_export)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Stop button (initially hidden)
        self.stop_btn = ttk.Button(btn_frame, text="⏹️ Stop Export", 
                                  command=self.stop_export)
        # Don't pack the stop button initially - it will be shown during export
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status and timer frame
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=8, column=0, columnspan=3, pady=5)
        
        # Status label
        self.status_label = ttk.Label(status_frame, text="Ready to export", foreground="green")
        self.status_label.grid(row=0, column=0, padx=(0, 20))
        
        # Timer label
        self.timer_label = ttk.Label(status_frame, text="⏱️ 00:00:00", foreground="blue", font=("Arial", 10, "bold"))
        self.timer_label.grid(row=0, column=1)
        
        # Log output - create a visible text area
        ttk.Label(main_frame, text="Export Log:").grid(row=9, column=0, sticky=tk.W, pady=(10, 5))
        
        # Create log text widget with explicit size and border
        self.log_text = tk.Text(
            main_frame, 
            height=10, 
            width=80,
            wrap=tk.WORD,
            bg="white",
            fg="black",
            relief=tk.SUNKEN,
            bd=2
        )
        self.log_text.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 10), padx=5)
        
        # Add scrollbar manually
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=10, column=3, sticky=(tk.N, tk.S), pady=5)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Test that logging works immediately
        self.log_text.insert(tk.END, "GUI initialized - logging ready\n")
        self.log_text.see(tk.END)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(10, weight=1)
        
        # Info section - placed after the log text widget
        info_frame = ttk.LabelFrame(main_frame, text="Export Info", padding="10")
        info_frame.grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 10))
        
        info_text = """• Documents will be exported as DOCX files
• Spreadsheets will be exported as XLSX files  
• Presentations will be exported as PDF files
• Folder structure will be preserved
• Subfolders will be exported recursively

⚠️ IMPORTANT: After export completion, check the log files
   for any failed documents or "Unknown Folder" issues"""
        
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W)
        
    def browse_folder(self):
        """Open folder browser dialog."""
        folder = filedialog.askdirectory(
            title="Select folder to save exports",
            initialdir=self.target_folder.get()
        )
        if folder:
            self.target_folder.set(folder)
    
    def log_message(self, message):
        """Add message to log output and save to file."""
        try:
            # Add current system timestamp to log messages
            from datetime import datetime
            current_time = datetime.now().strftime("%H:%M:%S")
            timestamp = f"[{current_time}] "
            timestamped_message = timestamp + message
            
            # Store message for log file
            self.log_messages.append(timestamped_message)
            
            # Check if log_text exists
            if not hasattr(self, 'log_text') or self.log_text is None:
                return
            
            # Insert message to GUI
            self.log_text.configure(state=tk.NORMAL)  # Ensure it's editable
            self.log_text.insert(tk.END, timestamped_message + "\n")
            self.log_text.see(tk.END)
            self.log_text.update_idletasks()
            
            # Save to log file if we have a path
            if self.log_file_path:
                try:
                    with open(self.log_file_path, 'a', encoding='utf-8') as f:
                        f.write(timestamped_message + "\n")
                except Exception as e:
                    pass  # Silently handle log file errors
            
        except Exception as e:
            pass  # Silently handle logging errors
    
    def format_time(self, seconds):
        """Format seconds into HH:MM:SS format."""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    def update_timer(self):
        """Update the timer display."""
        if self.is_exporting and self.start_time:
            elapsed = time.time() - self.start_time
            self.timer_label.config(text=f"⏱️ {self.format_time(elapsed)}")
            # Schedule next update
            self.timer_job = self.root.after(1000, self.update_timer)
    
    def start_timer(self):
        """Start the timer."""
        self.start_time = time.time()
        self.timer_label.config(text="⏱️ 00:00:00")
        self.update_timer()
    
    def stop_timer(self):
        """Stop the timer."""
        if self.timer_job:
            self.root.after_cancel(self.timer_job)
            self.timer_job = None
    
    def stop_export(self):
        """Stop the export process when user clicks stop button."""
        if self.is_exporting:
            # Confirm with user
            if messagebox.askyesno("Stop Export", 
                                 "Are you sure you want to stop the export?\n\n"
                                 "The export will be incomplete and you'll need to restart it."):
                self.stop_requested = True
                self.log_message("🛑 Export stop requested by user...")
                self.status_label.config(text="⏹️ Stopping export...", foreground="orange")
    
    def initialize_log_file(self):
        """Initialize the log file in the target directory."""
        try:
            target_dir = self.target_folder.get().strip()
            os.makedirs(target_dir, exist_ok=True)
            
            # Create log filename with timestamp
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"quip_export_log_{timestamp}.txt"
            self.log_file_path = os.path.join(target_dir, log_filename)
            
            # Create initial log file with header
            with open(self.log_file_path, 'w', encoding='utf-8') as f:
                f.write("Quip Export Log\n")
                f.write("=" * 50 + "\n")
                f.write(f"Export started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Target folder: {target_dir}\n")
                f.write(f"Quip URL: {self.quip_url.get().strip()}\n")
                f.write(f"API Token: {'*' * (len(self.quip_token.get().strip()) - 4) + self.quip_token.get().strip()[-4:] if len(self.quip_token.get().strip()) > 4 else '****'}\n")
                f.write("=" * 50 + "\n\n")
            
            return True
        except Exception as e:
            return False
    
    def create_error_file_if_needed(self):
        """Create an error file listing all failed exports if any exist."""
        if not self.failed_files and not self.failed_folders:
            return  # No failed files or folders, nothing to do
        
        try:
            target_dir = self.target_folder.get().strip()
            
            # Create error filename with timestamp
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_filename = f"quip_export_errors_{timestamp}.txt"
            error_file_path = os.path.join(target_dir, error_filename)
            
            # Create error file
            with open(error_file_path, 'w', encoding='utf-8') as f:
                f.write("Quip Export - Failed Items Report\n")
                f.write("=" * 50 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Total failed files: {len(self.failed_files)}\n")
                f.write(f"Total failed folders: {len(self.failed_folders)}\n")
                f.write("=" * 50 + "\n\n")
                
                # Write failed folders first
                item_count = 1
                if self.failed_folders:
                    f.write("FAILED FOLDERS:\n")
                    f.write("-" * 20 + "\n\n")
                    for failed_folder in self.failed_folders:
                        f.write(f"{item_count}. FAILED FOLDER\n")
                        f.write(f"   Folder ID: {failed_folder['folder_id']}\n")
                        f.write(f"   Created As: {failed_folder['expected_name']}\n")
                        f.write(f"   Parent Path: {failed_folder['path']}\n")
                        f.write(f"   Failure Reason: {failed_folder['reason']}\n")
                        f.write("-" * 30 + "\n\n")
                        item_count += 1
                
                # Write failed files
                if self.failed_files:
                    f.write("FAILED FILES:\n")
                    f.write("-" * 20 + "\n\n")
                    for failed_file in self.failed_files:
                        f.write(f"{item_count}. FAILED FILE\n")
                        f.write(f"   Title: {failed_file['title']}\n")
                        f.write(f"   Document ID: {failed_file['doc_id']}\n")
                        f.write(f"   Expected Path: {failed_file['folder_path']}\n")
                        f.write(f"   Expected Format: {failed_file['format']}\n")
                        f.write(f"   Failure Reason: {failed_file['reason']}\n")
                        # Create sanitized filename
                        sanitized_title = failed_file['title'].replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                        expected_filename = f"{sanitized_title}.{failed_file['format']}"
                        expected_full_path = os.path.join(failed_file['folder_path'], expected_filename)
                        f.write(f"   Expected Full Path: {expected_full_path}\n")
                        f.write("-" * 30 + "\n\n")
                        item_count += 1
                
                f.write("=" * 50 + "\n")
                f.write("END OF REPORT\n")
            
            # Log the error file creation
            total_failures = len(self.failed_files) + len(self.failed_folders)
            if self.failed_folders:
                self.log_message(f"⚠️ {len(self.failed_folders)} folders could not get proper names")
            if self.failed_files:
                self.log_message(f"⚠️ {len(self.failed_files)} files failed to export")
            self.log_message(f"📄 Error report saved: {error_filename}")
            
        except Exception as e:
            # If we can't create the error file, at least log it
            total_failures = len(self.failed_files) + len(self.failed_folders)
            self.log_message(f"⚠️ {total_failures} items failed (could not create error file)")
    
    def validate_inputs(self):
        """Validate user inputs."""
        token = self.quip_token.get().strip()
        url = self.quip_url.get().strip()
        folder = self.target_folder.get().strip()
        
        if not token:
            messagebox.showerror("Error", "Please enter your Quip API token")
            return False
            
        if not url:
            messagebox.showerror("Error", "Please enter a Quip folder URL")
            return False
            
        if not folder:
            messagebox.showerror("Error", "Please select a target folder")
            return False
            
        # Try to extract folder ID to validate URL
        try:
            folder_id = extract_folder_id_from_url(url)
            if not folder_id:
                raise ValueError("Invalid URL format")
        except Exception as e:
            messagebox.showerror("Error", f"Invalid Quip URL: {e}")
            return False
            
        # Check if target folder exists or can be created
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create target folder: {e}")
            return False
            
        return True
    
    def start_export(self):
        """Start the export process."""
        if not self.validate_inputs():
            return
            
        if self.is_exporting:
            messagebox.showwarning("Warning", "Export already in progress")
            return
        
        # Clear log and initialize log file
        self.log_text.delete(1.0, tk.END)
        self.log_messages = []  # Clear stored messages
        self.failed_files = []  # Clear failed files list
        self.failed_folders = []  # Clear failed folders list
        self.export_stats = {'folders': 0, 'documents': 0, 'failed_folders': 0, 'failed_documents': 0}  # Reset statistics
        self.stop_requested = False  # Reset stop flag
        
        # Initialize log file
        if not self.initialize_log_file():
            messagebox.showerror("Error", "Could not create log file in target directory")
            return
        
        self.log_message("🚀 Initializing export process...")
        self.log_message("📋 Validating inputs...")
        self.log_message("🔗 Connecting to Quip API...")
        self.log_message(f"📝 Log file created: {os.path.basename(self.log_file_path)}")
        
        # Start export in separate thread
        self.is_exporting = True
        self.export_btn.config(text="⏳ Exporting...", state="disabled")
        
        # Show stop button and hide export button
        self.export_btn.pack_forget()
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.progress.start()
        self.status_label.config(text="Export in progress...", foreground="orange")
        
        # Start timer
        self.start_timer()
        
        # Run export in background thread
        export_thread = threading.Thread(target=self.run_export, daemon=True)
        export_thread.start()
    
    def run_export(self):
        """Run the actual export process."""
        try:
            url = self.quip_url.get().strip()
            target_folder = self.target_folder.get().strip()
            
            # Extract folder ID
            folder_id = extract_folder_id_from_url(url)
            
            # Log initial messages using root.after for thread safety
            self.root.after(0, lambda: self.log_message("🚀 Starting export..."))
            self.root.after(0, lambda: self.log_message(f"📂 Quip URL: {url}"))
            self.root.after(0, lambda: self.log_message(f"💾 Target folder: {target_folder}"))
            self.root.after(0, lambda: self.log_message(f"🔑 Folder ID: {folder_id}"))
            self.root.after(0, lambda: self.log_message("-" * 50))
            
            # Create exporter with user-provided token (silent mode for GUI)
            user_token = self.quip_token.get().strip()
            exporter = QuipExporter(user_token, silent=True)
            
            # Custom export function that logs to GUI
            def export_with_logging(exporter, folder_id, base_output_dir, depth=0):
                try:
                    indent = "  " * depth
                    
                    # Get folder contents (using the function from native_export)
                    from native_export import get_folder_contents, sanitize_folder_name
                    
                    # Log folder processing start
                    msg = f"{indent}📁 Processing folder: {folder_id}"
                    self.root.after(0, lambda m=msg: self.log_message(m))
                    
                    doc_ids, subfolder_ids, folder_title = get_folder_contents(exporter, folder_id)
                    safe_folder_name = sanitize_folder_name(folder_title)
                    
                    # Track folder statistics and failures
                    self.export_stats['folders'] += 1
                    if folder_title.startswith("Unknown Folder"):
                        self.export_stats['failed_folders'] += 1
                        # Track failed folder
                        failed_folder_info = {
                            'folder_id': folder_id,
                            'expected_name': folder_title,
                            'path': base_output_dir,
                            'reason': 'Could not retrieve folder name from Quip API after 5 attempts'
                        }
                        self.failed_folders.append(failed_folder_info)
                        
                        # Log warning
                        warning_msg = f"{indent}   ⚠️ Warning: Could not get proper folder name for {folder_id}"
                        self.root.after(0, lambda m=warning_msg: self.log_message(m))
                    
                    # Create folder path
                    if depth == 0:
                        folder_path = os.path.join(base_output_dir, safe_folder_name)
                    else:
                        folder_path = os.path.join(base_output_dir, safe_folder_name)
                    
                    # Log folder details
                    msg1 = f"{indent}   📋 Folder: {folder_title}"
                    msg2 = f"{indent}   📁 Path: {folder_path}"
                    msg3 = f"{indent}   📊 Found: {len(doc_ids)} documents, {len(subfolder_ids)} subfolders"
                    
                    self.root.after(0, lambda m=msg1: self.log_message(m))
                    self.root.after(0, lambda m=msg2: self.log_message(m))
                    self.root.after(0, lambda m=msg3: self.log_message(m))
                    
                    # Create the folder
                    os.makedirs(folder_path, exist_ok=True)
                    
                    # Export all documents in this folder
                    exported_docs = 0
                    for i, doc_id in enumerate(doc_ids):
                        # Check if user requested to stop
                        if self.stop_requested:
                            stop_msg = f"{indent}   🛑 Export stopped by user"
                            self.root.after(0, lambda m=stop_msg: self.log_message(m))
                            return False
                        
                        try:
                            # Get document info for logging with retry
                            doc_info = None
                            for info_attempt in range(3):
                                doc_info = exporter._make_request(f"threads/{doc_id}")
                                if doc_info is not None:
                                    break
                                elif info_attempt < 2:  # Not the last attempt
                                    time.sleep(1)  # Wait 1 second before retry
                            
                            if doc_info is None:
                                error_msg = f"{indent}     ❌ Failed to get document info after 3 attempts: {doc_id}"
                                self.root.after(0, lambda m=error_msg: self.log_message(m))
                                # Track failed file (no document info)
                                failed_file_info = {
                                    'title': f'Unknown Document ({doc_id})',
                                    'doc_id': doc_id,
                                    'folder_path': folder_path,
                                    'format': 'unknown',
                                    'reason': 'Failed to get document info after 3 attempts'
                                }
                                self.failed_files.append(failed_file_info)
                                continue
                            
                            title = doc_info.get("thread", {}).get("title", f"Document_{doc_id}")
                            doc_type = exporter.get_document_type(doc_id)
                            
                            # Choose format
                            if doc_type == "document":
                                best_format = "docx"
                            elif doc_type == "spreadsheet":
                                best_format = "xlsx"
                            elif doc_type == "presentation":
                                best_format = "pdf"
                            else:
                                best_format = "docx"
                            
                            # Sanitize filename
                            safe_title = title.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                            
                            # Log export start
                            start_msg = f"{indent}     📄 [{i+1}/{len(doc_ids)}] Exporting: {title} → {best_format.upper()}"
                            self.root.after(0, lambda m=start_msg: self.log_message(m))
                            
                            # Retry logic: attempt up to 5 times (initial + 4 retries)
                            result = False
                            max_attempts = 5
                            
                            for attempt in range(max_attempts):
                                try:
                                    result = exporter.download_document(doc_id, folder_path, title, best_format)
                                    
                                    if result:
                                        # Success - break out of retry loop
                                        break
                                    else:
                                        # Failed attempt
                                        if attempt < max_attempts - 1:  # Not the last attempt
                                            retry_msg = f"{indent}       🔄 Attempt {attempt + 1} failed, retrying..."
                                            self.root.after(0, lambda m=retry_msg: self.log_message(m))
                                            time.sleep(2)  # Wait 2 seconds before retry
                                        else:
                                            # Final attempt failed
                                            final_fail_msg = f"{indent}     ❌ Failed after {max_attempts} attempts: {title}"
                                            self.root.after(0, lambda m=final_fail_msg: self.log_message(m))
                                            # Track failed file
                                            failed_file_info = {
                                                'title': title,
                                                'doc_id': doc_id,
                                                'folder_path': folder_path,
                                                'format': best_format,
                                                'reason': 'Download failed after 5 attempts'
                                            }
                                            self.failed_files.append(failed_file_info)
                                            
                                except Exception as retry_error:
                                    if attempt < max_attempts - 1:  # Not the last attempt
                                        retry_error_msg = f"{indent}       🔄 Attempt {attempt + 1} error, retrying: {str(retry_error)[:50]}..."
                                        self.root.after(0, lambda m=retry_error_msg: self.log_message(m))
                                        time.sleep(2)  # Wait 2 seconds before retry
                                    else:
                                        # Final attempt failed with error
                                        final_error_msg = f"{indent}     ❌ Failed after {max_attempts} attempts with error: {str(retry_error)[:50]}..."
                                        self.root.after(0, lambda m=final_error_msg: self.log_message(m))
                                        # Track failed file with error
                                        failed_file_info = {
                                            'title': title,
                                            'doc_id': doc_id,
                                            'folder_path': folder_path,
                                            'format': best_format,
                                            'reason': f'Error after 5 attempts: {str(retry_error)[:100]}'
                                        }
                                        self.failed_files.append(failed_file_info)
                            
                            # Log final result and update statistics
                            self.export_stats['documents'] += 1
                            if result:
                                exported_docs += 1
                                # Create full file path for logging
                                file_path = os.path.join(folder_path, f"{safe_title}.{best_format}")
                                # Log success with full path
                                success_msg = f"{indent}     ✅ Saved: {title}.{best_format}"
                                path_msg = f"{indent}        📁 {file_path}"
                                self.root.after(0, lambda m=success_msg: self.log_message(m))
                                self.root.after(0, lambda m=path_msg: self.log_message(m))
                            else:
                                self.export_stats['failed_documents'] += 1
                                
                        except Exception as e:
                            error_msg = f"{indent}     ❌ Error exporting {title}: {str(e)[:50]}..."
                            self.root.after(0, lambda m=error_msg: self.log_message(m))
                    
                    # Recursively export subfolders
                    exported_subfolders = 0
                    for subfolder_id in subfolder_ids:
                        # Check if user requested to stop
                        if self.stop_requested:
                            stop_msg = f"{indent}   🛑 Export stopped by user"
                            self.root.after(0, lambda m=stop_msg: self.log_message(m))
                            return False
                        
                        try:
                            subfolder_result = export_with_logging(exporter, subfolder_id, folder_path, depth + 1)
                            if subfolder_result:
                                exported_subfolders += 1
                        except Exception as e:
                            error_msg = f"{indent}   ❌ Subfolder {subfolder_id} error: {str(e)[:50]}..."
                            self.root.after(0, lambda m=error_msg: self.log_message(m))
                    
                    # Log completion summary
                    summary_msg = f"{indent}   ✅ Completed: {exported_docs}/{len(doc_ids)} documents, {exported_subfolders}/{len(subfolder_ids)} subfolders"
                    self.root.after(0, lambda m=summary_msg: self.log_message(m))
                    return True
                    
                except Exception as e:
                    error_msg = f"❌ Error processing folder {folder_id}: {str(e)}"
                    self.root.after(0, lambda m=error_msg: self.log_message(m))
                    return False
            
            # Run the export
            success = export_with_logging(exporter, folder_id, target_folder)
            
            # Update UI on completion
            self.root.after(0, self.export_completed, success)
            
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            self.root.after(0, self.export_completed, False, error_msg)
    
    def export_completed(self, success, error_msg=None):
        """Handle export completion."""
        # Stop timer and calculate total time
        self.stop_timer()
        total_time = ""
        if self.start_time:
            elapsed = time.time() - self.start_time
            total_time = self.format_time(elapsed)
            self.timer_label.config(text=f"⏱️ {total_time} (Total)")
        
        self.is_exporting = False
        self.progress.stop()
        
        # Restore button visibility
        self.stop_btn.pack_forget()
        self.export_btn.pack(side=tk.LEFT, padx=5)
        self.export_btn.config(text="🚀 Start Export", state="normal")
        
        # Check if export was stopped by user
        if self.stop_requested:
            self.status_label.config(text="⏹️ Export stopped by user", foreground="orange")
            self.log_message(f"\n🛑 Export stopped by user!")
            self.log_message(f"⏱️ Time before stop: {total_time}")
            self.log_message(f"⚠️ Export was incomplete - some files may not have been exported")
            
            # Create error file and finalize logs for stopped export
            self.create_error_file_if_needed()
            
            # Finalize log file with stop information
            if self.log_file_path:
                try:
                    from datetime import datetime
                    with open(self.log_file_path, 'a', encoding='utf-8') as f:
                        f.write("\n" + "=" * 50 + "\n")
                        f.write(f"Export STOPPED by user: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write(f"Time before stop: {total_time}\n")
                        f.write("Export was INCOMPLETE - not all files were processed\n")
                        f.write("\nExport Statistics (Partial):\n")
                        f.write(f"Folders processed: {self.export_stats.get('folders', 0)}\n")
                        f.write(f"Documents processed: {self.export_stats.get('documents', 0)}\n")
                        f.write(f"Successful documents: {self.export_stats.get('documents', 0) - len(self.failed_files)}\n")
                        f.write(f"Failed documents: {len(self.failed_files)}\n")
                        if self.export_stats.get('failed_folders', 0) > 0:
                            f.write(f"Unknown folders created: {self.export_stats.get('failed_folders', 0)}\n")
                        f.write("=" * 50 + "\n")
                except Exception as e:
                    pass
            
            # Show stopped export dialog with folder opening option
            stop_dialog_msg = (f"Export was stopped by user after {total_time}.\n\n"
                              f"⚠️ Export is INCOMPLETE:\n"
                              f"   • Some files may not have been exported\n"
                              f"   • You may need to restart the export\n\n"
                              f"Partial results and logs saved to:\n{self.target_folder.get()}\n\n"
                              f"Would you like to open the export folder to review what was exported?")
            
            if messagebox.askyesno("Export Stopped - Open Folder?", stop_dialog_msg):
                try:
                    os.startfile(self.target_folder.get())  # Windows
                except:
                    try:
                        os.system(f'open "{self.target_folder.get()}"')  # macOS
                    except:
                        try:
                            os.system(f'xdg-open "{self.target_folder.get()}"')  # Linux
                        except:
                            pass
            
        elif success:
            self.status_label.config(text="✅ Export completed successfully!", foreground="green")
            self.log_message(f"\n🎉 Export completed successfully!")
            self.log_message(f"⏱️ Total time: {total_time}")
            self.log_message(f"📁 Files saved to: {self.target_folder.get()}")
            self.log_message(f"📝 Log saved to: {self.log_file_path}")
            
            # Log export statistics
            self.log_message(f"\n📊 Export Statistics:")
            self.log_message(f"   📁 Folders processed: {self.export_stats.get('folders', 0)}")
            self.log_message(f"   📄 Documents processed: {self.export_stats.get('documents', 0)}")
            self.log_message(f"   ✅ Successful documents: {self.export_stats.get('documents', 0) - len(self.failed_files)}")
            self.log_message(f"   ❌ Failed documents: {len(self.failed_files)}")
            if self.export_stats.get('failed_folders', 0) > 0:
                self.log_message(f"   ⚠️ Unknown folders created: {self.export_stats.get('failed_folders', 0)}")
            
            # Create error file if there were failed exports
            self.create_error_file_if_needed()
            
            # Finalize log file
            if self.log_file_path:
                try:
                    from datetime import datetime
                    with open(self.log_file_path, 'a', encoding='utf-8') as f:
                        f.write("\n" + "=" * 50 + "\n")
                        f.write(f"Export completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write(f"Total time: {total_time}\n")
                        f.write("\nExport Statistics:\n")
                        f.write(f"Folders processed: {self.export_stats.get('folders', 0)}\n")
                        f.write(f"Documents processed: {self.export_stats.get('documents', 0)}\n")
                        f.write(f"Successful documents: {self.export_stats.get('documents', 0) - len(self.failed_files)}\n")
                        f.write(f"Failed documents: {len(self.failed_files)}\n")
                        if self.export_stats.get('failed_folders', 0) > 0:
                            f.write(f"Unknown folders created: {self.export_stats.get('failed_folders', 0)}\n")
                        f.write("=" * 50 + "\n")
                except Exception as e:
                    pass
            
            # Check if there were any errors and show warning first
            has_errors = len(self.failed_files) > 0 or self.export_stats.get('failed_folders', 0) > 0
            
            if has_errors:
                # Show error warning first
                error_count = len(self.failed_files) + self.export_stats.get('failed_folders', 0)
                messagebox.showwarning("Export Completed with Errors", 
                                     f"Export completed in {total_time}, but {error_count} items failed!\n\n"
                                     f"⚠️ IMPORTANT: Please check the error log file:\n"
                                     f"   • Failed documents: {len(self.failed_files)}\n"
                                     f"   • Unknown folders: {self.export_stats.get('failed_folders', 0)}\n\n"
                                     f"Error details saved to: quip_export_errors_*.txt\n\n"
                                     f"Review the error file to identify which items need manual attention.")
            
            # Ask if user wants to open the folder
            folder_prompt_msg = f"Export completed in {total_time}!\n\nLog saved to: {os.path.basename(self.log_file_path)}"
            if has_errors:
                folder_prompt_msg += f"\n\n⚠️ Note: {error_count} items failed - check error log file"
            folder_prompt_msg += "\n\nWould you like to open the export folder?"
            
            if messagebox.askyesno("Open Export Folder", folder_prompt_msg):
                try:
                    os.startfile(self.target_folder.get())  # Windows
                except:
                    try:
                        os.system(f'open "{self.target_folder.get()}"')  # macOS
                    except:
                        try:
                            os.system(f'xdg-open "{self.target_folder.get()}"')  # Linux
                        except:
                            pass
        else:
            self.status_label.config(text="❌ Export failed", foreground="red")
            self.log_message(f"❌ Export failed: {error_msg if error_msg else 'Unknown error'}")
            
            # Create error file if there were failed exports
            self.create_error_file_if_needed()
            
            # Finalize log file even on failure
            if self.log_file_path:
                try:
                    from datetime import datetime
                    with open(self.log_file_path, 'a', encoding='utf-8') as f:
                        f.write("\n" + "=" * 50 + "\n")
                        f.write(f"Export failed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        if error_msg:
                            f.write(f"Error: {error_msg}\n")
                        f.write("=" * 50 + "\n")
                except Exception as e:
                    pass
            
            if error_msg:
                messagebox.showerror("Export Failed", error_msg)

def main():
    """Main function to run the GUI."""
    root = tk.Tk()
    
    # Set up styling
    style = ttk.Style()
    
    # Try to use a modern theme
    try:
        style.theme_use('clam')
    except:
        pass
    
    app = QuipExporterGUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()