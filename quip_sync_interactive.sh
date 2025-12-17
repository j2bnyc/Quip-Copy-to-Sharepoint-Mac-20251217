#!/bin/bash
# Quip Sync Interactive - Configure and run Quip folder sync
# Double-click or run from terminal

cd "$(dirname "$0")"

echo "========================================"
echo "  Quip Folder Sync Tool"
echo "========================================"
echo ""

# Default values
DEFAULT_SOURCE="https://quip-amazon.com/jzKiObqykZYB/Citi-Team-Folder"
DEFAULT_TARGET="/Users/jimclrk/Library/CloudStorage/OneDrive-SharedLibraries-amazon.com/Citigroup Account Team - Documents/QS"

# Get Quip folder URL
echo "Enter Quip folder URL"
echo "  [Press Enter for default: $DEFAULT_SOURCE]"
read -p "> " SOURCE_URL
if [ -z "$SOURCE_URL" ]; then
    SOURCE_URL="$DEFAULT_SOURCE"
fi
echo ""

# Get destination path
echo "Enter destination folder path"
echo "  [Press Enter for default: $DEFAULT_TARGET]"
read -p "> " TARGET_PATH
if [ -z "$TARGET_PATH" ]; then
    TARGET_PATH="$DEFAULT_TARGET"
fi
echo ""

# Get PAT token
echo "Enter your Quip API token (get from https://quip-amazon.com/dev/token)"
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
