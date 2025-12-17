# Quip Folder Exporter

**Repository:** https://github.com/j2bnyc/Quip-Copy-to-Sharepoint-Mac-20251217

A robust, cross-platform GUI application that recursively exports and downloads Quip folders and documents. Export entire folder hierarchies while maintaining the original structure, with intelligent format selection, comprehensive error handling, and detailed reporting.

**Supports:** Windows, macOS, and Linux with native launchers for each platform.

## Features

### **Core Functionality**
- **Easy-to-use GUI interface** - No command-line knowledge required
- **Recursive folder traversal** - Export entire folder hierarchies automatically
- **Smart format selection** - Automatically chooses the best format for each document type:
  - Documents → DOCX format
  - Spreadsheets → XLSX format  
  - Presentations → PDF format

### **Reliability & Error Handling**
- **5-attempt retry system** - Each file gets up to 5 download attempts for maximum reliability
- **Comprehensive error tracking** - Automatic error reports for any failed exports
- **Failed files documentation** - Detailed error files with exact paths and failure reasons
- **API rate limit handling** - Automatically manages Quip API limits with intelligent delays
- **Network resilience** - Handles temporary connectivity issues gracefully

### **Monitoring & Reporting**
- **Real-time logging** - See export progress with detailed status updates and timestamps
- **Automatic log file creation** - Complete export logs saved to your target directory
- **Error report generation** - Separate error files listing all failed exports with full details
- **Timer functionality** - Track export duration with built-in timer
- **Progress tracking** - Visual progress bar and status updates
- **Export statistics** - Summary of successful vs. failed exports

### **User Experience**
- **Export control** - Start/stop buttons with user-friendly export management
- **Clickable token link** - Direct access to Quip API token generation page
- **Filename sanitization** - Ensures safe file names across all operating systems
- **Windowless operation** - Clean launch without terminal windows (Windows)
- **Folder structure preservation** - Maintains original Quip folder hierarchy
- **Automatic folder opening** - Option to open export folder when complete or stopped
- **Graceful stop handling** - Safe export termination with complete partial results

## Installation

### **Quick Start (Windows)**
1. **Download/clone** this repository
2. **Install Python** from https://python.org (if not already installed)
   - ✅ Check "Add Python to PATH" during installation
3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
4. **Launch the application:**
   - **Double-click** `Quip Exporter.bat` (recommended)
   - Or run `python quip_exporter_gui.py` from command line

### **Quick Start (macOS)**
1. **Download/clone** this repository
2. **Install Python** (if not already installed):
   ```bash
   # Option 1: Download from https://python.org
   # Option 2: Use Homebrew
   brew install python
   ```
3. **Install dependencies:**
   ```bash
   pip3 install -r requirements.txt
   ```
4. **Launch the application:**
   ```bash
   # Method 1: Make launcher executable and double-click
   chmod +x launch_quip_exporter.sh
   # Then double-click launch_quip_exporter.sh in Finder
   
   # Method 2: Double-click the .command file
   # Double-click launch_quip_exporter_simple.command in Finder
   
   # Method 3: Terminal command
   python3 quip_exporter_gui.py
   ```

### **Quick Start (Linux)**
1. **Download/clone** this repository
2. **Install Python** (usually pre-installed):
   ```bash
   # Ubuntu/Debian
   sudo apt update && sudo apt install python3 python3-pip
   
   # CentOS/RHEL
   sudo yum install python3 python3-pip
   ```
3. **Install dependencies:**
   ```bash
   pip3 install -r requirements.txt
   ```
4. **Launch the application:**
   ```bash
   # Make launcher executable and run
   chmod +x launch_quip_exporter.sh
   ./launch_quip_exporter.sh
   
   # Or direct command
   python3 quip_exporter_gui.py
   ```

### **Alternative Launch Methods**
- **Windows Debug**: Double-click `Launch Quip Exporter - Debug.bat` for troubleshooting
- **macOS/Linux Debug**: Run `./launch_quip_exporter.sh` for detailed output
- **Command line**: `python quip_exporter_gui.py` or `python3 quip_exporter_gui.py`
- **Windows Windowless**: `pythonw quip_exporter_gui.py` (no terminal window)

## Requirements

- Python 3.6+
- `tkinter` (usually included with Python)
- `requests` library (installed via requirements.txt)
- Valid Quip API access token

### Required Python Packages

Install all dependencies with:
```bash
pip install -r requirements.txt
# or on macOS/Linux:
pip3 install -r requirements.txt
```

The core dependency is:
- `requests` - HTTP library for Quip API communication

If `tkinter` is not included with your Python installation:
```bash
# Ubuntu/Debian
sudo apt install python3-tk

# Fedora
sudo dnf install python3-tkinter

# macOS (usually included, but if missing)
brew install python-tk
```

## Getting Your Quip Personal Access Token (PAT)

To use this tool, you need a Quip API token:

1. Open your browser and go to: **https://quip-amazon.com/dev/token**
2. Log in with your Quip/Amazon credentials if prompted
3. Click **"Generate Token"** or similar button to create a new API token
4. **Copy the entire token** - it's a long string of characters
5. Paste the token into the application's "Quip API Token" field

**Important notes:**
- Keep your token secure - it provides full access to your Quip content
- Tokens do not expire automatically, but you can revoke them from the same page
- If you suspect your token is compromised, revoke it immediately and generate a new one
- Never commit your token to version control or share it publicly

## Finding Your CloudStorage/SharePoint Synced Folder Path

If you're syncing Quip exports to a SharePoint folder via OneDrive, here's how to find the correct local path:

### macOS

1. Open **Finder**
2. Look in the sidebar under **"Locations"** for your OneDrive folder (usually named "OneDrive - amazon.com" or similar)
3. Navigate to the SharePoint library you want to sync to
4. Right-click the folder and select **"Get Info"** or press `Cmd+I`
5. Copy the path from the "Where:" field

Alternatively, use Terminal:
```bash
# List CloudStorage folders
ls -la ~/Library/CloudStorage/

# Example paths:
# ~/Library/CloudStorage/OneDrive-SharedLibraries-amazon.com/
# ~/Library/CloudStorage/OneDrive-amazon.com/
```

The full path typically looks like:
```
/Users/YOUR_USERNAME/Library/CloudStorage/OneDrive-SharedLibraries-amazon.com/YOUR_TEAM_NAME - Documents/SUBFOLDER
```

### Windows

1. Open **File Explorer**
2. Look in the left sidebar for your OneDrive folder (blue cloud icon)
3. Navigate to the SharePoint library folder
4. Click in the address bar to see the full path
5. Copy the path

Typical Windows paths:
```
C:\Users\YOUR_USERNAME\OneDrive - amazon.com\YOUR_TEAM_NAME - Documents\SUBFOLDER
```

Or for shared libraries:
```
C:\Users\YOUR_USERNAME\amazon.com\YOUR_TEAM_NAME - Documents\SUBFOLDER
```

### Tips for SharePoint/OneDrive Sync

- Make sure the SharePoint library is synced to your local machine via OneDrive
- To sync a new SharePoint library: Open SharePoint in browser → Click "Sync" button → OneDrive will add it locally
- The `quip_sync_cli.py` script is designed for automated syncing to these locations
- Use incremental mode (`--mode incremental`) to only sync changed files

## Getting Your API Token

The application provides a direct link to get your token:

1. Click the blue link in the application: `https://quip-amazon.com/dev/token`
2. Log in to your Quip account if prompted
3. Generate an API access token
4. Copy and paste the token into the application
5. Keep this token secure - it provides access to your Quip content

## Usage

### GUI Application

1. **Launch the application:**
   
   **Windows (Recommended):**
   - Double-click `Quip Exporter.bat` (no terminal window)
   
   **Command Line (Alternative):**
   ```bash
   python quip_exporter_gui.py
   ```

2. **Get your API token:**
   - Click the blue link in the application to open the Quip token page
   - Generate and copy your API token
   - Paste it into the "Quip API Token" field

3. **Enter the folder URL:**
   - Copy the URL of the Quip folder you want to export
   - Paste it into the "Quip Folder URL" field
   - Example: `https://quip-amazon.com/ABC123/FolderName`

4. **Choose export location:**
   - Use the default location or click "Browse" to select a different folder
   - The application will create subfolders as needed

5. **Start the export:**
   - Click "🚀 Start Export" to begin
   - Watch the real-time progress in the log window
   - The timer shows elapsed time during export

### Application Interface

| Field | Description |
|-------|-------------|
| **Quip API Token** | Your personal API token (masked for security) |
| **Quip Folder URL** | Full URL of the Quip folder to export |
| **Save to Folder** | Local directory where files will be saved |
| **Export Log** | Real-time status updates and progress information |
| **Export Info** | Details about file formats and export behavior |

## Smart Export Formats

The application automatically selects the best format for each document type:

### Documents
- **DOCX format** - Microsoft Word compatible
- Preserves formatting, images, and document structure
- Best for text documents, reports, and collaborative writing

### Spreadsheets  
- **XLSX format** - Microsoft Excel compatible
- Maintains formulas, formatting, and data structure
- Perfect for data analysis and financial documents

### Presentations
- **PDF format** - Universal compatibility
- Preserves layout, images, and slide formatting
- Ideal for sharing and viewing presentations

### Automatic Format Selection
The application intelligently detects each document type and chooses the most appropriate format automatically. No manual format selection needed!

## Reliability Features

### **5-Attempt Retry System**
- Each document gets **up to 5 download attempts**
- 2-second delays between retry attempts
- Handles temporary network issues and API rate limiting
- Clear logging of each retry attempt

### **Comprehensive Error Handling**
- **Document info retrieval**: 3 attempts with 1-second delays
- **File downloads**: 5 attempts with 2-second delays
- **Graceful failure handling**: Export continues even if some files fail
- **Detailed error logging**: Specific reasons for each failure

### **Automatic Error Reports**
When exports complete, if any files failed to download, the application automatically creates:
- **Error report file**: `quip_export_errors_YYYYMMDD_HHMMSS.txt`
- **Complete failure details**: Document titles, IDs, expected paths, and failure reasons
- **Recovery information**: Exact file paths where documents should have been saved

## Example Export Process

1. **Launch the application**
   - Double-click `Quip Exporter.bat` (Windows)
   - Or run `python quip_exporter_gui.py` from command line
2. **Get your token** from https://quip-amazon.com/dev/token
3. **Enter folder URL** like: `https://quip-amazon.com/ABC123DEF456/MyProject`
4. **Choose save location** (default: `Downloads/QuipExports`)
5. **Click Start Export**

The application will:
- **Validate your inputs** and create timestamped log files
- **Show stop button** during export for user control
- **Recursively process** all folders and subfolders with 5-attempt retry logic
- **Export each document** in the optimal format (DOCX/XLSX/PDF)
- **Show real-time progress** with detailed logging and timestamps
- **Handle errors gracefully** with comprehensive retry mechanisms
- **Allow safe stopping** at any point with partial results preserved
- **Generate error reports** for any failed items (if needed)
- **Display completion summary** with detailed statistics
- **Show error warnings** if any items failed during export
- **Offer to open export folder** for easy access to results (complete or partial)

## Output Structure

The application maintains your Quip folder structure in the output directory:

```
QuipExports/
├── Project Folder/
│   ├── Meeting Notes.docx
│   ├── Budget Spreadsheet.xlsx
│   ├── Presentation.pdf
│   ├── Subfolder/
│   │   ├── Technical Specs.docx
│   │   └── Data Analysis.xlsx
│   ├── quip_export_log_20241022_143052.txt
│   └── quip_export_errors_20241022_143052.txt (if any failures)
└── Another Project/
    ├── Project Plan.docx
    └── Status Report.docx
```

### Log Files
Each export creates detailed log files:

#### **Main Log File** (`quip_export_log_YYYYMMDD_HHMMSS.txt`)
- Export timestamp and duration
- Real-time progress with timestamps
- Complete list of processed documents
- Success/failure status for each file
- Retry attempts and outcomes
- Final export statistics
- Masked API token info (last 4 characters only)

#### **Error Report File** (`quip_export_errors_YYYYMMDD_HHMMSS.txt`) - *Created only if failures occur*
- Complete list of failed documents
- Document titles and Quip IDs
- Expected file paths where documents should have been saved
- Detailed failure reasons for each document
- Recovery information for manual retry attempts

### Example Log Entries
```
[15:08:34] 📁 Processing folder: SCVentures
[15:08:34]    📊 Found: 6 documents, 0 subfolders
[15:08:35]       📄 [1/6] Exporting: Bloom → DOCX
[15:08:36]       🔄 Attempt 1 failed, retrying...
[15:08:38]       🔄 Attempt 2 failed, retrying...
[15:08:40]       ✅ Saved: Bloom.docx
[15:08:41]       📄 [2/6] Exporting: Budget Report → XLSX
[15:08:42]       ✅ Saved: Budget Report.xlsx
```

## Features in Detail

### **Advanced Retry System**
- **5-attempt downloads**: Each document gets up to 5 download attempts
- **3-attempt info retrieval**: Document information fetched with 3 retry attempts
- **Smart delays**: 1-2 second delays between attempts to respect API limits
- **Detailed retry logging**: See exactly which attempts succeeded or failed

### **Real-time Monitoring**
- **Live progress updates**: Real-time status in the application window
- **Timestamped logging**: Every action logged with precise timestamps
- **Detailed folder processing**: See document counts and subfolder structure
- **Export statistics**: Track successful vs. failed exports in real-time

### **Comprehensive Error Reporting**
- **Automatic error files**: Generated only when failures occur
- **Complete failure documentation**: Every failed file documented with full details
- **Recovery information**: Exact paths where files should have been saved
- **Failure categorization**: Different error types clearly identified

### **Robust Error Handling**
- **Network resilience**: Handles temporary connectivity issues
- **API rate limit management**: Automatic delays and retry logic
- **Graceful degradation**: Export continues even if some files fail
- **Clear error messages**: Specific reasons for each failure type

### **User-Friendly Interface**
- **No command-line knowledge required**: Pure GUI operation
- **Clickable token link**: Direct access to Quip API token generation
- **Progress visualization**: Progress bar and timer for long exports
- **Completion notifications**: Success/failure dialogs with folder opening option
- **Windowless operation**: Clean launch without terminal windows (Windows)

## Troubleshooting

### **Authentication Errors**
- **Check your API token**: Ensure you copied the complete token from the Quip developer page
- **Verify permissions**: Make sure your token has access to the folders you're trying to export
- **Test access**: Try accessing the folder in your browser first
- **Token expiration**: Generate a new token if the current one has expired

### **Invalid Folder URL**
- **Use complete URL**: Copy the full URL from your browser address bar
- **Check folder access**: Ensure you have permission to view the folder
- **Verify URL format**: Should look like `https://quip-amazon.com/ABC123/FolderName`
- **Private folders**: Ensure the folder is accessible with your API token

### **Export Issues**
- **Check internet connection**: Ensure stable connectivity during export
- **Review log files**: Check both main log and error report files for detailed information
- **Retry failed files**: Use the error report to identify and manually retry specific documents
- **Check disk space**: Ensure sufficient space in your target directory
- **Large folders**: For very large folders, consider exporting subfolders individually

### **Partial Export Issues**
- **Check error report**: Look for `quip_export_errors_*.txt` file in your export directory
- **Review retry attempts**: Failed files show all 5 retry attempts in the main log
- **Network stability**: Ensure stable internet connection for large exports
- **API rate limits**: The application handles these automatically, but very large exports may take time

### **Platform-Specific Issues**

#### **Windows**
- **Launcher not working**: Try `Launch Quip Exporter - Debug.bat` for detailed diagnostics
- **Python not found**: Install from https://python.org and check "Add Python to PATH"
- **Terminal window appears**: Use `Quip Exporter.bat` for windowless operation
- **Permission errors**: Run as administrator if needed

#### **macOS**
- **Script permission denied**: Run `chmod +x launch_quip_exporter.sh` in Terminal
- **Python not found**: Install via Homebrew: `brew install python`
- **GUI not appearing**: Ensure XQuartz is installed for X11 forwarding (if using SSH)
- **Gatekeeper blocking**: Right-click launcher and select "Open" to bypass security

#### **Linux**
- **Missing tkinter**: Install with `sudo apt install python3-tk` (Ubuntu/Debian)
- **Display issues**: Ensure `DISPLAY` environment variable is set correctly
- **Permission errors**: Make launcher executable: `chmod +x launch_quip_exporter.sh`
- **Python version**: Use `python3` command instead of `python`

#### **General Application Issues**
- **Restart the application**: Close and reopen if interface becomes unresponsive
- **Check Python version**: Ensure you're using Python 3.6 or newer
- **Verify dependencies**: Run `pip install -r requirements.txt` or `pip3 install -r requirements.txt`
- **Clear cache**: Delete `__pycache__` folders if experiencing import issues

### **Performance Optimization**
- **Close other applications**: Free up system resources for large exports
- **Stable network**: Use wired connection for very large exports
- **Sufficient disk space**: Ensure adequate free space before starting
- **Monitor progress**: Use the real-time logging to track export progress
- **Use stop function**: For very large exports, you can safely stop and resume later

### **Export Management**
- **Long exports**: Use the stop button if you need to interrupt long-running exports
- **Partial results**: Stopped exports create complete logs and preserve all exported files
- **Resume strategy**: Check error logs to identify what still needs to be exported
- **Safe stopping**: Export stops between documents/folders to prevent data corruption

## Error Handling & User Notifications

### **Automatic Error Detection**
The application automatically detects and reports:
- **Failed document downloads** (after 5 retry attempts)
- **Unknown folder names** (when Quip API can't provide folder titles)
- **Network connectivity issues** (with automatic retry logic)
- **API rate limiting** (with intelligent delay management)

### **User Notification System**
- **GUI warnings**: Prominent reminder to check error logs after export
- **Error dialogs**: Immediate notification when errors occur during export
- **Detailed breakdown**: Shows counts of failed documents vs. unknown folders
- **Guided troubleshooting**: Directs users to specific error log files

### **Error Recovery**
- **Complete documentation**: Every failure recorded with specific reasons
- **Recovery information**: Exact file paths where items should have been saved
- **Manual retry guidance**: Error reports include all necessary details for manual recovery
- **Folder identification**: Unknown folders include original Quip folder IDs for reference

## Security & Best Practices

### API Token Security
- **No hardcoded tokens**: The application never stores or uses hardcoded API tokens
- **User-provided tokens only**: All tokens must be entered by the user for each session
- **Never share your token**: Keep your API token private and secure
- **Use password managers**: Store tokens securely, not in plain text files
- **Revoke unused tokens**: Remove tokens you no longer need from your Quip account
- **Monitor token usage**: Check your Quip account for any unauthorized access

### Application Security
- **Token masking**: The application masks your token in the interface (shows only last 4 characters)
- **Log file security**: Log files show masked token information for security
- **Local processing**: All exports happen locally on your machine
- **No token persistence**: Tokens are not saved or stored anywhere in the application

## Project Structure

### **Core Files**
- `quip_exporter_gui.py` - Main GUI application with advanced error handling
- `native_export.py` - Core export engine with 5-attempt retry logic
- `requirements.txt` - Python dependencies (only requests library needed)
- `README.md` - Complete documentation

### **Platform-Specific Launchers**

#### **Windows**
- `Quip Exporter.bat` - Main launcher (recommended - no terminal window)
- `Launch Quip Exporter - Debug.bat` - Diagnostic launcher with detailed output

#### **macOS**
- `launch_quip_exporter_simple.command` - Simple double-click launcher
- `launch_quip_exporter.sh` - Full diagnostic launcher with error checking

#### **Linux**
- `launch_quip_exporter.sh` - Full diagnostic launcher with error checking

### **Generated Files (after export)**
- `quip_export_log_YYYYMMDD_HHMMSS.txt` - Complete export log with statistics
- `quip_export_errors_YYYYMMDD_HHMMSS.txt` - Error report (only if failures occur)

## System Requirements

### **Cross-Platform Support**
- **Windows**: 7, 8, 10, 11 (32-bit and 64-bit)
- **macOS**: 10.12 Sierra or newer (Intel and Apple Silicon)
- **Linux**: Ubuntu 16.04+, CentOS 7+, or equivalent distributions

### **Software Requirements**
- **Python**: Version 3.6 or newer (3.8+ recommended)
- **Dependencies**: Only `requests` library (automatically installed)
- **GUI Framework**: tkinter (included with most Python installations)

### **Hardware Requirements**
- **Memory**: At least 512MB RAM (1GB+ recommended for large exports)
- **Storage**: Sufficient disk space for exported documents
- **Network**: Stable internet connection for Quip API access
- **Display**: Any resolution supporting GUI applications

## Performance Tips

- **Large exports**: For folders with many documents, expect longer processing times
- **Network speed**: Faster internet connections will speed up the export process
- **Disk space**: Ensure adequate free space before starting large exports
- **Background apps**: Close unnecessary applications for better performance during large exports

## Creating Standalone Executables

For distribution to users without Python installed, you can create standalone executables:

### **Using PyInstaller (Recommended)**
```bash
# Install PyInstaller
pip install pyinstaller

# Create standalone executable
pyinstaller --onefile --windowed --name "Quip Exporter" quip_exporter_gui.py

# Find executable in dist/ folder
```

### **Platform-Specific Executables**
- **Windows**: Creates `.exe` file that runs without Python installation
- **macOS**: Creates `.app` bundle for easy distribution
- **Linux**: Creates binary executable for the target distribution

### **Distribution Notes**
- **File size**: Standalone executables are larger (50-100MB) but self-contained
- **No dependencies**: Users don't need Python or any libraries installed
- **Platform-specific**: Create executables on each target platform for best compatibility

## Quick Reference

### **Common Commands**
```bash
# Install dependencies
pip install -r requirements.txt          # Windows
pip3 install -r requirements.txt         # macOS/Linux

# Launch application
python quip_exporter_gui.py              # Windows
python3 quip_exporter_gui.py             # macOS/Linux

# Make launchers executable (macOS/Linux)
chmod +x launch_quip_exporter.sh
chmod +x launch_quip_exporter_simple.command

# Create standalone executable
pyinstaller --onefile --windowed quip_exporter_gui.py
```

### **File Locations After Export**
- **Main log**: `quip_export_log_YYYYMMDD_HHMMSS.txt`
- **Error report**: `quip_export_errors_YYYYMMDD_HHMMSS.txt` (if failures occur)
- **Exported documents**: Organized in original folder structure

### **API Endpoints Used**
- **Base URL**: `https://platform.quip-amazon.com/1`
- **Folder info**: `GET /folders/{folder_id}`
- **Document info**: `GET /threads/{doc_id}`
- **Document export**: `GET /threads/{doc_id}/export/{format}`

## License

This application is provided as-is for personal and organizational use.

## Support

For issues or questions:
1. **Check the log window**: The application provides detailed error information
2. **Review this documentation**: Most common issues are covered in the troubleshooting section
3. **Verify your setup**: Ensure your API token and folder URL are correct
4. **Test with small folders**: Try exporting a small folder first to verify everything works

## Version History

### **Current Version: v2.2 - User Control Enhanced**
- **5-attempt retry system**: Maximum reliability for both files AND folders
- **Export stop functionality**: Users can safely stop long-running exports at any time
- **Comprehensive error reporting**: Automatic error files with complete failure details
- **Advanced logging**: Real-time timestamps and detailed progress tracking
- **Robust error handling**: Graceful handling of network issues and API limits
- **Failed item documentation**: Complete recovery information for manual retry
- **Unknown folder detection**: Identifies and reports folders that couldn't be properly named
- **Export statistics**: Detailed counts of processed vs. failed items
- **User error notifications**: Immediate warnings when errors occur during export
- **Enhanced GUI warnings**: Clear reminders to check error logs after completion
- **Partial export support**: Complete logging and folder access for stopped exports

## Export Control Features

### **Start/Stop Functionality**
- **Dynamic button management**: Stop button appears only during active exports
- **User confirmation**: Confirmation dialog prevents accidental stops
- **Graceful termination**: Export stops safely between documents/folders
- **Immediate response**: Stop takes effect at the next safe checkpoint

### **Partial Export Handling**
- **Complete logging**: Full documentation of what was exported before stopping
- **Error file generation**: Includes any failures that occurred before stopping
- **Statistics tracking**: Accurate counts of processed items in partial exports
- **Folder access**: Option to open export folder to review partial results

### **Stop Process Flow**
1. **User clicks Stop Export** → Confirmation dialog appears
2. **User confirms stop** → Export begins graceful shutdown
3. **Current operation completes** → Export stops at safe checkpoint
4. **Logs finalized** → Complete documentation of partial export
5. **Folder opening offered** → Easy access to partial results and logs

### **Previous Versions:**
- **v2.1**: Production ready with comprehensive error handling
- **v2.0**: Enhanced reliability with 5-attempt retry system
- **v1.0**: Core GUI functionality with smart format selection

## Technical Specifications

### **Advanced Retry Logic**
- **Folder info retrieval**: 5 attempts with 2-second delays
- **Document info retrieval**: 3 attempts with 1-second delays  
- **File downloads**: 5 attempts with 2-second delays
- **Network resilience**: Handles temporary connectivity issues and API timeouts
- **API rate limiting**: Intelligent delay management and automatic retry

### **Smart Format Selection**
- **Documents**: DOCX (Microsoft Word format) - preserves formatting and structure
- **Spreadsheets**: XLSX (Microsoft Excel format) - maintains formulas and data
- **Presentations**: PDF (Universal format) - preserves layout and images
- **Automatic detection**: No manual format selection required
- **Fallback handling**: Graceful degradation for unsupported document types

### **Comprehensive Error Reporting**
- **Main log**: Complete export history with real-time timestamps
- **Error report**: Detailed failure analysis (auto-generated when needed)
- **Unknown folders**: Special tracking for folders that couldn't be properly named
- **Recovery data**: Exact file paths and specific failure reasons
- **Export statistics**: Complete counts of processed, successful, and failed items
- **Security**: Masked API tokens in all log files (shows only last 4 characters)

### **Production-Grade Features**
- **Export control**: Start/stop functionality with graceful termination
- **Windowless operation**: Clean launch without terminal windows (Windows)
- **Memory optimization**: Efficient handling of large folder structures
- **Progress monitoring**: Real-time status updates with precise timestamps
- **Resource management**: Automatic cleanup and comprehensive error recovery
- **User notifications**: Immediate error warnings and guided troubleshooting
- **Partial export support**: Complete documentation of stopped exports
- **Batch file launcher**: One-click execution with automatic Python detection
