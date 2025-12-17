#!/usr/bin/env python3
"""
Quip Native Export Module - Core functionality for Quip document and folder export

This module provides the QuipExporter class and supporting functions for:
- Quip API communication with retry logic
- Document type detection (documents, spreadsheets, presentations)
- Smart format selection and file downloads
- Folder traversal and structure preservation
- Error handling and network resilience

Security: No hardcoded tokens - all tokens must be provided by users
Reliability: Built-in retry mechanisms and error handling
Formats: DOCX for documents, XLSX for spreadsheets, PDF for presentations
"""

import sys
import os
import re
import requests
import time

# No hardcoded tokens - tokens must be provided by users

class QuipExporter:
    """Quip API client for exporting documents and folders."""
    
    def __init__(self, token, base_url="https://platform.quip-amazon.com/1", silent=False):
        self.token = token
        self.base_url = base_url
        self.silent = silent  # Flag to suppress console output
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {token}',
            'User-Agent': 'QuipExporter/1.0'
        })
    
    def _make_request(self, endpoint, params=None):
        """Make a request to the Quip API."""
        url = f"{self.base_url}/{endpoint}"
        try:
            response = self.session.get(url, params=params)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            if not self.silent:
                print(f"HTTP Error {response.status_code} for {endpoint}: {e}")
            return None
        except requests.exceptions.RequestException as e:
            if not self.silent:
                print(f"API request failed for {endpoint}: {e}")
            return None
    
    def get_document_type(self, doc_id):
        """Get the type of a document (document, spreadsheet, presentation)."""
        try:
            doc_info = self._make_request(f"threads/{doc_id}")
            if doc_info and "thread" in doc_info:
                thread_class = doc_info["thread"].get("type", "document")
                if thread_class == "spreadsheet":
                    return "spreadsheet"
                elif thread_class == "slides":
                    return "presentation"
                else:
                    return "document"
            return "document"
        except Exception as e:
            if not self.silent:
                print(f"Error getting document type: {e}")
            return "document"
    
    def download_document(self, doc_id, output_dir, title, format_type):
        """Download a document in the specified format."""
        try:
            # Sanitize filename
            safe_title = self._sanitize_filename(title)
            filename = f"{safe_title}.{format_type}"
            filepath = os.path.join(output_dir, filename)
            
            # Make export request
            export_url = f"{self.base_url}/threads/{doc_id}/export/{format_type}"
            response = self.session.get(export_url)
            
            if response.status_code == 200:
                os.makedirs(output_dir, exist_ok=True)
                with open(filepath, 'wb') as f:
                    f.write(response.content)
                return True
            else:
                if not self.silent:
                    print(f"Export failed with status {response.status_code}")
                return False
                
        except Exception as e:
            if not self.silent:
                print(f"Error downloading document: {e}")
            return False
    
    def _sanitize_filename(self, filename):
        """Sanitize filename for safe file system usage."""
        # Replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # Remove leading/trailing spaces and dots
        filename = filename.strip(' .')
        
        # Limit length
        if len(filename) > 200:
            filename = filename[:200]
        
        return filename or "untitled"

def extract_folder_id_from_url(url_or_id):
    """Extract folder ID from Quip URL or return the ID if it's already an ID."""
    # If it's already just an ID (no URL), return it
    if not url_or_id.startswith('http'):
        return url_or_id
    
    # Extract ID from URL patterns like:
    # https://quip-amazon.com/4LF8Ocl7mP4U/Artem
    # https://quip-amazon.com/4LF8Ocl7mP4U
    match = re.search(r'quip-amazon\.com/([A-Za-z0-9]+)', url_or_id)
    if match:
        return match.group(1)
    
    # If no match, try to extract anything that looks like an ID
    match = re.search(r'/([A-Za-z0-9]{10,})', url_or_id)
    if match:
        return match.group(1)
    
    raise ValueError(f"Could not extract folder ID from: {url_or_id}")

def sanitize_folder_name(name):
    """Sanitize folder name for file system."""
    # Remove or replace invalid characters for Windows/Mac/Linux
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name.strip()

def export_document_native(exporter, doc_id, output_dir):
    """Export a document using Quip's native export functionality in the best format."""
    try:
        print(f"  📄 Exporting document: {doc_id}")
        
        # Get document info
        doc_info = exporter._make_request(f"threads/{doc_id}")
        title = doc_info.get("thread", {}).get("title", f"Document_{doc_id}")
        print(f"     Title: {title}")
        
        # Get document type
        doc_type = exporter.get_document_type(doc_id)
        print(f"     Type: {doc_type}")
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        # Choose the best format based on document type
        if doc_type == "document":
            best_format = "docx"  # Word documents -> DOCX
        elif doc_type == "spreadsheet":
            best_format = "xlsx"  # Spreadsheets -> Excel
        elif doc_type == "presentation":
            best_format = "pdf"   # Presentations -> PDF
        else:
            best_format = "docx"  # Default to DOCX
        
        print(f"     Format: {best_format.upper()}")
        
        try:
            result = exporter.download_document(doc_id, output_dir, title, best_format)
            
            if result:
                print(f"     ✓ Success: {title}.{best_format}")
                return title
            else:
                print(f"     ✗ Failed: {title}")
                return None
                
        except KeyboardInterrupt:
            print(f"\n⚠ Export interrupted by user")
            raise
        except Exception as e:
            print(f"     ✗ Error: {str(e)[:50]}...")
            return None
        
    except Exception as e:
        print(f"Error exporting document {doc_id}: {e}")
        return None

def export_folder_recursive(exporter, folder_id, base_output_dir="native_export", depth=0):
    """Recursively export a folder and all its contents."""
    try:
        indent = "  " * depth
        print(f"{indent}📁 Processing folder: {folder_id}")
        
        # Get folder contents
        doc_ids, subfolder_ids, folder_title = get_folder_contents(exporter, folder_id)
        safe_folder_name = sanitize_folder_name(folder_title)
        
        # Create folder path
        if depth == 0:
            # Root folder
            folder_path = os.path.join(base_output_dir, safe_folder_name)
        else:
            # Subfolder - append to base path
            folder_path = os.path.join(base_output_dir, safe_folder_name)
        
        print(f"{indent}   Folder: {folder_title}")
        print(f"{indent}   Path: {folder_path}")
        print(f"{indent}   Documents: {len(doc_ids)}, Subfolders: {len(subfolder_ids)}")
        
        # Create the folder
        os.makedirs(folder_path, exist_ok=True)
        
        # Export all documents in this folder
        exported_docs = 0
        for doc_id in doc_ids:
            result = export_document_native(exporter, doc_id, folder_path)
            if result:
                exported_docs += 1
        
        # Recursively export subfolders
        exported_subfolders = 0
        for subfolder_id in subfolder_ids:
            try:
                subfolder_result = export_folder_recursive(exporter, subfolder_id, folder_path, depth + 1)
                if subfolder_result:
                    exported_subfolders += 1
            except Exception as e:
                print(f"{indent}   ✗ Error with subfolder {subfolder_id}: {e}")
        
        print(f"{indent}   ✓ Completed: {exported_docs}/{len(doc_ids)} docs, {exported_subfolders}/{len(subfolder_ids)} subfolders")
        return True
        
    except Exception as e:
        print(f"Error processing folder {folder_id}: {e}")
        return False

def get_folder_contents(exporter, folder_id):
    """Get all documents and subfolders from a folder."""
    try:
        print(f"Getting contents from folder: {folder_id}")
        
        # Get folder info with retry logic (5 attempts)
        folder_info = None
        max_attempts = 5
        
        for attempt in range(max_attempts):
            folder_info = exporter._make_request(f"folders/{folder_id}")
            if folder_info is not None:
                break
            elif attempt < max_attempts - 1:
                print(f"Folder info attempt {attempt + 1} failed, retrying...")
                time.sleep(2)  # Wait 2 seconds before retry
            else:
                print(f"Failed to get folder info after {max_attempts} attempts")
        
        if folder_info is None:
            print(f"ERROR: Could not retrieve folder information for {folder_id}")
            return [], [], f"Unknown Folder ({folder_id})"
        
        # Extract folder title safely
        folder_title = "Unknown Folder"
        if "folder" in folder_info and folder_info["folder"] is not None:
            folder_title = folder_info["folder"].get("title", f"Folder_{folder_id}")
        
        print(f"Folder title: {folder_title}")
        
        doc_ids = []
        subfolder_ids = []
        
        # Parse folder contents
        if "children" in folder_info:
            for child in folder_info["children"]:
                if child.get("thread_id"):
                    doc_ids.append(child["thread_id"])
                elif child.get("folder_id"):
                    subfolder_ids.append(child["folder_id"])
        elif "thread_ids" in folder_info:
            doc_ids = folder_info["thread_ids"]
        
        # Also check for folder_ids array
        if "folder_ids" in folder_info:
            subfolder_ids.extend(folder_info["folder_ids"])
        
        return doc_ids, subfolder_ids, folder_title
        
    except Exception as e:
        print(f"Error getting folder contents: {e}")
        return [], [], f"Unknown Folder ({folder_id})"

def native_export_folder(token, folder_url_or_id):
    """Export all documents from a folder and subfolders using Quip's native export."""
    try:
        exporter = QuipExporter(token)
        
        # Extract folder ID from URL or use as-is if it's already an ID
        folder_id = extract_folder_id_from_url(folder_url_or_id)
        print(f"Using folder ID: {folder_id}")
        
        print(f"\n🚀 Starting recursive export...")
        
        # Start recursive export
        success = export_folder_recursive(exporter, folder_id)
        
        if success:
            print(f"\n✅ Export complete! Check the 'native_export' folder.")
            print(f"📁 Folder structure preserved with proper file formats.")
        else:
            print(f"\n❌ Export failed or incomplete.")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Quip Folder Exporter - Export all documents from a Quip folder")
        print("")
        print("Usage: python native_export.py TOKEN FOLDER_URL_OR_ID [FORMATS]")
        print("")
        print("Arguments:")
        print("  TOKEN              Your Quip API token (get from https://quip-amazon.com/dev/token)")
        print("  FOLDER_URL_OR_ID   Quip folder URL or folder ID")
        print("  FORMATS            Comma-separated export formats (optional)")
        print("")
        print("Supported formats:")
        print("  docx     - Microsoft Word document (default)")
        print("  xlsx     - Excel format for spreadsheets (default)")
        print("  pdf      - PDF format for compatible documents (default)")
        print("  html     - HTML format")
        print("  markdown - Markdown format")
        print("")
        print("Examples:")
        print("  python native_export.py YOUR_TOKEN https://quip-amazon.com/4LF8Ocl7mP4U/Artem")
        print("  python native_export.py YOUR_TOKEN 4LF8Ocl7mP4U docx")
        print("  python native_export.py YOUR_TOKEN https://quip-amazon.com/ABC123/MyFolder docx,xlsx,pdf,html")
        print("")
        print("Security Note: Never share or commit your API token!")
        print("")
        sys.exit(1)
    
    user_token = sys.argv[1]
    folder_url_or_id = sys.argv[2]
    
    print("Quip Folder Exporter")
    print("=" * 50)
    print(f"Token: {'*' * (len(user_token) - 4) + user_token[-4:] if len(user_token) > 4 else '****'}")
    print(f"Folder: {folder_url_or_id}")
    print("Auto-selecting best format for each document type:")
    print("  • Documents → DOCX")
    print("  • Spreadsheets → XLSX") 
    print("  • Presentations → PDF")
    print("=" * 50)
    
    native_export_folder(user_token, folder_url_or_id)