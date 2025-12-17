# Quip Folder Sync

Sync Quip folders to local directories (SharePoint/OneDrive) with full or incremental sync modes.

**Cross-platform:** Mac, Linux, and Windows

## Quick Start

### 1. Install requests (if needed)
```bash
pip install requests
```

### 2. Get Your API Token
Visit: https://quip-amazon.com/dev/token

### 3. Run the Sync

**Mac/Linux:**
```bash
./quip_sync_interactive.sh
```

**Windows:**
```
quip_sync_interactive.bat
```

## Sync Modes

| Mode | Description |
|------|-------------|
| **Incremental** (default) | Only syncs files modified since last sync |
| **Full** | Overwrites all files regardless of changes |

## Command Line Usage

```bash
# Mac/Linux - Set token as environment variable
export QUIP_TOKEN='your_token_here'
python3 quip_sync_cli.py --mode incremental

# Windows - Set token as environment variable  
set QUIP_TOKEN=your_token_here
python quip_sync_cli.py --mode incremental

# Or pass token directly (any platform)
python quip_sync_cli.py --token YOUR_TOKEN --mode full
```

### All Options

```bash
python quip_sync_cli.py \
  --token YOUR_TOKEN \
  --source "https://quip-amazon.com/ABC123/FolderName" \
  --target "/path/to/local/folder" \
  --mode incremental
```

| Option | Description |
|--------|-------------|
| `--token`, `-t` | Quip API token (or set `QUIP_TOKEN` env var) |
| `--source`, `-s` | Quip folder URL or ID |
| `--target`, `-d` | Local destination directory |
| `--mode`, `-m` | `full` or `incremental` (default) |
| `--dry-run` | Preview without downloading |

## File Format Selection

Documents are automatically exported in the best format:

| Quip Type | Export Format |
|-----------|---------------|
| Documents | DOCX |
| Spreadsheets | XLSX |
| Presentations | PDF |

## Files

| File | Platform | Description |
|------|----------|-------------|
| `quip_sync_cli.py` | All | Main sync tool |
| `quip_sync_interactive.sh` | Mac/Linux | Interactive launcher |
| `quip_sync_interactive.bat` | Windows | Interactive launcher |

## Sync State

Incremental sync tracks document modification times in `.quip_sync_state.json` in your target directory. Delete this file to force a full re-sync.

## Finding Your OneDrive/SharePoint Path

**Mac:**
```bash
ls ~/Library/CloudStorage/
# Example: ~/Library/CloudStorage/OneDrive-SharedLibraries-amazon.com/Team - Documents/
```

**Windows:**
```
C:\Users\USERNAME\OneDrive - amazon.com\Team - Documents\
```

## Troubleshooting

**Token has special characters (|, &, etc.):**
- Mac/Linux: Wrap in single quotes: `export QUIP_TOKEN='token|with|pipes'`
- Windows: No quotes needed with `set`

**Permission denied (Mac/Linux):**
```bash
chmod +x quip_sync_interactive.sh
```

**Python not found:**
- Mac: `brew install python` or download from python.org
- Windows: Download from python.org, check "Add to PATH"

## Legacy GUI Version

The original GUI exporter files (`quip_exporter_gui.py`, `native_export.py`, etc.) are also included for users who prefer a graphical interface.
