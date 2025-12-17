# Quip Folder Sync

Sync Quip folders to local directories (SharePoint/OneDrive) with full or incremental sync modes.

**Cross-platform:** Mac, Linux, and Windows

## Quick Start

### 1. Install requests (if needed)
```bash
pip install requests
```

### 2. Get Your API Token
Visit your Quip token page (e.g., `https://quip-mycompany.com/dev/token`)

### 3. Run the Sync

**Mac/Linux:**
```bash
# Option 1: Run with bash (recommended)
bash quip_sync_interactive.sh

# Option 2: Run directly (may require chmod first)
chmod +x quip_sync_interactive.sh
./quip_sync_interactive.sh
```

**Windows:**
```
quip_sync_interactive.bat
```

The interactive script will prompt you for:
- **Quip folder URL** - The root folder you want to copy
- **Target folder** - Local folder synced to SharePoint/OneDrive
- **API token** - Your personal access token
- **Sync mode** - Incremental (default) or Full

## Sync Modes

| Mode | Description |
|------|-------------|
| **Incremental** (default) | Only syncs files modified since last sync |
| **Full** | Overwrites all files regardless of changes |

## Command Line Usage

```bash
# Mac/Linux
export QUIP_TOKEN='your_token_here'
python3 quip_sync_cli.py \
  --source "https://quip-mycompany.com/ABC123XYZ/My-Team-Folder" \
  --target "/Users/username/Library/CloudStorage/OneDrive-SharedLibraries-mycompany.com/Team - Documents/Backup" \
  --mode incremental

# Windows
set QUIP_TOKEN=your_token_here
python quip_sync_cli.py ^
  --source "https://quip-mycompany.com/ABC123XYZ/My-Team-Folder" ^
  --target "C:\Users\username\OneDrive - mycompany.com\Team - Documents\Backup" ^
  --mode incremental
```

### Options

| Option | Description |
|--------|-------------|
| `--token`, `-t` | Quip API token (or set `QUIP_TOKEN` env var) |
| `--source`, `-s` | Quip folder URL or ID (required) |
| `--target`, `-d` | Local destination directory (required) |
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
# Example: ~/Library/CloudStorage/OneDrive-SharedLibraries-mycompany.com/Team - Documents/
```

**Windows:**
```
C:\Users\USERNAME\OneDrive - mycompany.com\Team - Documents\
```

## Troubleshooting

**Token has special characters (|, &, etc.):**
- Mac/Linux: Wrap in single quotes: `export QUIP_TOKEN='token|with|pipes'`
- Windows: No quotes needed with `set`

**Permission denied (Mac/Linux):**
```bash
# Use bash to run directly (easiest fix)
bash quip_sync_interactive.sh

# Or make the script executable
chmod +x quip_sync_interactive.sh
./quip_sync_interactive.sh
```

**Python not found:**
- Mac: `brew install python` or download from python.org
- Windows: Download from python.org, check "Add to PATH"

**Path too long errors (especially Windows):**

Windows has a 260-character path limit by default. OneDrive/SharePoint paths can be long, and deeply nested Quip folders make it worse.

Solutions:
1. **Use a shorter target path:**
   - Instead of: `C:\Users\username\OneDrive - mycompany.com\Team - Documents\Quip Backup\`
   - Try: `C:\QuipSync\` then move/copy to SharePoint after

2. **Enable long paths on Windows 10/11:**
   ```
   # Run PowerShell as Administrator
   New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
   ```
   Then restart your computer.

3. **Shorten folder names in Quip** before syncing (if you have edit access)

4. **Sync a subfolder** instead of the entire root folder to reduce nesting depth

**Files fail to sync / API errors:**
- Check your API token is valid and not expired
- Verify you have access to the Quip folder
- Some documents may be restricted or corrupted in Quip

**OneDrive sync conflicts:**
- Close OneDrive temporarily while running the sync
- Or sync to a local folder first, then copy to OneDrive

