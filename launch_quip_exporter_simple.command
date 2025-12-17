#!/bin/bash
# Simple Quip Exporter Launcher for macOS
# Double-click this file to launch the application

cd "$(dirname "$0")"

# Try python3 first, then python
if command -v python3 &> /dev/null; then
    python3 quip_exporter_gui.py
elif command -v python &> /dev/null; then
    python quip_exporter_gui.py
else
    echo "Python not found. Please install Python from https://python.org"
    read -p "Press Enter to exit..."
fi