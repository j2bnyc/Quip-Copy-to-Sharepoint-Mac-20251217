#!/bin/bash
# Quip Sync Interactive - Configure and run Quip folder sync (Mac/Linux)
# 
# USAGE:
#   ./quip_sync_interactive.sh
#   
# Or double-click in Finder (Mac)
#
# For Windows, use: quip_sync_interactive.bat

cd "$(dirname "$0")"

echo "========================================"
echo "  Quip Folder Sync Tool (Mac/Linux)"
echo "========================================"
echo ""

# Get Quip folder URL
echo "Enter Quip folder URL (root folder you want to copy)"
echo "  Example: https://quip-mycompany.com/ABC123XYZ/My-Team-Folder"
read -p "> " SOURCE_URL
echo ""

if [ -z "$SOURCE_URL" ]; then
    echo "❌ Error: Quip folder URL is required"
    echo "Press any key to exit..."
    read -n 1
    exit 1
fi

# Get destination path
echo "Enter destination folder path (local folder synced to SharePoint)"
echo "  Example: /Users/username/Library/CloudStorage/OneDrive-SharedLibraries-mycompany.com/Team - Documents/Backup"
read -p "> " TARGET_PATH
echo ""

if [ -z "$TARGET_PATH" ]; then
    echo "❌ Error: Target folder path is required"
    echo "Press any key to exit..."
    read -n 1
    exit 1
fi

# Get PAT token
echo "Enter your Quip API token"
echo "  Get from: https://quip-mycompany.com/dev/token"
echo "  (input is hidden for security)"
read -s -p "> " QUIP_TOKEN
echo ""
echo ""

if [ -z "$QUIP_TOKEN" ]; then
    echo "❌ Error: Token is required"
    echo "Press any key to exit..."
    read -n 1
    exit 1
fi

# Select sync mode
echo "Select sync mode:"
echo "  1) Incremental - only sync modified files (default)"
echo "  2) Full - overwrite all files"
echo ""
read -p "Enter choice [1/2]: " MODE_CHOICE

case $MODE_CHOICE in
    2)
        SYNC_MODE="full"
        ;;
    *)
        SYNC_MODE="incremental"
        ;;
esac

echo ""
echo "========================================"
echo "  Configuration Summary"
echo "========================================"
echo "  Source: $SOURCE_URL"
echo "  Target: $TARGET_PATH"
echo "  Mode:   $SYNC_MODE"
echo "========================================"
echo ""
read -p "Press Enter to start sync (or Ctrl+C to cancel)..."
echo ""

# Export token and run sync
export QUIP_TOKEN
python3 quip_sync_cli.py --source "$SOURCE_URL" --target "$TARGET_PATH" --mode "$SYNC_MODE"

echo ""
echo "========================================"
echo "  Sync complete!"
echo "========================================"
echo "Press any key to close..."
read -n 1
