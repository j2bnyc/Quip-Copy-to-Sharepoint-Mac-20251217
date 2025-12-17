#!/bin/bash
# Quip Sync - Double-click to run on Mac
# This script syncs your Quip Citi Team Folder to SharePoint

cd "$(dirname "$0")"

echo "========================================"
echo "  Quip Sync for Citi Team Folder"
echo "========================================"
echo ""

# Check if token is set
if [ -z "$QUIP_TOKEN" ]; then
    echo "Enter your Quip API token (get from https://quip-amazon.com/dev/token):"
    read -s QUIP_TOKEN
    export QUIP_TOKEN
    echo ""
fi

# Ask for sync mode
echo "Select sync mode:"
echo "  1) Incremental (only sync modified files)"
echo "  2) Full (overwrite all files)"
echo ""
read -p "Enter choice [1/2]: " choice

case $choice in
    2)
        MODE="full"
        ;;
    *)
        MODE="incremental"
        ;;
esac

echo ""
echo "Starting $MODE sync..."
echo ""

python3 quip_sync_cli.py --mode "$MODE"

echo ""
echo "Press any key to close..."
read -n 1
