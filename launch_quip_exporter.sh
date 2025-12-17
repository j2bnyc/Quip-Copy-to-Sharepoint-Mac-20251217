#!/bin/bash
# Quip Exporter GUI Launcher for macOS/Linux
# This script launches the Quip Exporter GUI

# Change to the directory where this script is located
cd "$(dirname "$0")"

echo "Quip Exporter Launcher"
echo "======================"
echo "Current directory: $(pwd)"
echo

# Check if the Python file exists
if [ ! -f "quip_exporter_gui.py" ]; then
    echo "ERROR: quip_exporter_gui.py not found"
    echo "Please make sure this script is in the same folder as quip_exporter_gui.py"
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Found: quip_exporter_gui.py ✓"
echo

# Try Method 1: python3
if command -v python3 &> /dev/null; then
    echo "SUCCESS: Found python3"
    echo "Launching Quip Exporter..."
    echo "If a GUI window appears, this launcher worked!"
    echo
    python3 quip_exporter_gui.py
    exit_code=$?
    
    if [ $exit_code -eq 0 ]; then
        echo "Application launched successfully"
        exit 0
    else
        echo "Failed to launch with python3 (exit code: $exit_code)"
        echo "Trying next method..."
    fi
else
    echo "python3 not found, trying python..."
fi

echo

# Try Method 2: python
if command -v python &> /dev/null; then
    echo "SUCCESS: Found python"
    echo "Launching Quip Exporter..."
    echo
    python quip_exporter_gui.py
    exit_code=$?
    
    if [ $exit_code -eq 0 ]; then
        echo "Application launched successfully"
        exit 0
    else
        echo "Failed to launch with python (exit code: $exit_code)"
    fi
else
    echo "python not found"
fi

echo
echo "========================================"
echo "ERROR: All launch methods failed"
echo "========================================"
echo
echo "Possible solutions:"
echo "1. Install Python from https://python.org"
echo "2. Install Python via Homebrew: brew install python"
echo "3. Try running this command manually in Terminal:"
echo "   python3 quip_exporter_gui.py"
echo
echo "4. Check if you have the required dependencies:"
echo "   pip3 install -r requirements.txt"
echo
read -p "Press Enter to exit..."