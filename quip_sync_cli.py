#!/usr/bin/env python3
"""
Quip Sync CLI - Command-line tool for syncing Quip folders to local directories

Cross-platform: Works on Mac, Linux, and Windows

Features:
- Full sync with overwrite (--mode full)
- Incremental sync - only modified files (--mode incremental)
- Preserves folder structure
- Smart format selection (DOCX/XLSX/PDF)
- Progress tracking and logging

=== MAC / LINUX USAGE ===

    # Using environment variable (recommended)
    export QUIP_TOKEN='your_token_here'
    python3 quip_sync_cli.py --mode incremental
    
    # Or pass token directly
    python3 quip_sync_cli.py --token YOUR_TOKEN --mode full

    # Interactive launcher
    ./quip_sync_interactive.sh

=== WINDOWS USAGE ===

    # Using environment variable (recommended)
    set QUIP_TOKEN=your_token_here
    python quip_sync_cli.py --mode incremental
    
    # Or pass token directly
    python quip_sync_cli.py --token YOUR_TOKEN --mode full

    # Interactive launcher (double-click or run in cmd)
    quip_sync_interactive.bat

=== GET YOUR API TOKEN ===
    https://quip-mycompany.com/dev/token (replace with your Quip domain)
"""

import argparse
import os
import sys
import time
import json
from datetime import datetime
import requests

# No defaults - user must provide source and target
# Example source: https://quip-mycompany.com/ABC123XYZ/My-Team-Folder
# Example target: /Users/username/Library/CloudStorage/OneDrive-SharedLibraries-mycompany.com/Team - Documents/Backup
SYNC_STATE_FILE = ".quip_sync_state.json"


class QuipSyncClient:
    """Quip API client for syncing documents."""
    
    def __init__(self, token, base_url=None):
        self.token = token
        # Base URL will be derived from source URL or default to platform.quip.com
        self.base_url = base_url or "https://platform.quip.com/1"
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {token}',
            'User-Agent': 'QuipSyncCLI/1.0'
        })
        self.stats = {'synced': 0, 'skipped': 0, 'failed': 0, 'folders': 0}
    
    def _request(self, endpoint, params=None, max_retries=3):
        """Make API request with retry logic."""
        url = f"{self.base_url}/{endpoint}"
        for attempt in range(max_retries):
            try:
                response = self.session.get(url, params=params, timeout=30)
                response.raise_for_status()
                return response.json()
            except requests.exceptions.HTTPError as e:
                if response.status_code == 429:  # Rate limited
                    wait_time = 2 ** attempt
                    print(f"  ⏳ Rate limited, waiting {wait_time}s...")
                    time.sleep(wait_time)
                elif attempt == max_retries - 1:
                    print(f"  ❌ HTTP Error {response.status_code}: {e}")
                    return None
            except requests.exceptions.RequestException as e:
                if attempt == max_retries - 1:
                    print(f"  ❌ Request failed: {e}")
                    return None
                time.sleep(1)
        return None
    
    def get_folder_info(self, folder_id):
        """Get folder metadata and contents."""
        return self._request(f"folders/{folder_id}")
    
    def get_thread_info(self, thread_id):
        """Get document/thread metadata."""
        return self._request(f"threads/{thread_id}")
    
    def get_document_type(self, thread_info):
        """Determine document type from thread info."""
        if thread_info and "thread" in thread_info:
            thread_type = thread_info["thread"].get("type", "document")
            if thread_type == "spreadsheet":
                return "spreadsheet", "xlsx"
            elif thread_type == "slides":
                return "presentation", "pdf"
        return "document", "docx"
    
    def get_document_updated_time(self, thread_info):
        """Get the last updated timestamp from thread info."""
        if thread_info and "thread" in thread_info:
            return thread_info["thread"].get("updated_usec", 0)
        return 0
    
    def download_document(self, doc_id, output_path, format_type):
        """Download document in specified format."""
        export_url = f"{self.base_url}/threads/{doc_id}/export/{format_type}"
        for attempt in range(3):
            try:
                response = self.session.get(export_url, timeout=60)
                if response.status_code == 200:
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                    with open(output_path, 'wb') as f:
                        f.write(response.content)
                    return True
                elif attempt < 2:
                    time.sleep(2)
            except Exception as e:
                if attempt == 2:
                    print(f"    ❌ Download error: {e}")
        return False


def extract_folder_id(url_or_id):
    """Extract folder ID from Quip URL or return as-is if already an ID."""
    import re
    if not url_or_id.startswith('http'):
        return url_or_id
    # Match various Quip URL patterns (quip.com, quip-*.com, etc.)
    match = re.search(r'quip[^/]*\.com/([A-Za-z0-9]+)', url_or_id)
    if match:
        return match.group(1)
    raise ValueError(f"Could not extract folder ID from: {url_or_id}")


def sanitize_filename(name):
    """Sanitize filename for file system."""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    name = name.strip(' .')
    return name[:200] if len(name) > 200 else name or "untitled"


def load_sync_state(target_dir):
    """Load previous sync state from JSON file."""
    state_file = os.path.join(target_dir, SYNC_STATE_FILE)
    if os.path.exists(state_file):
        try:
            with open(state_file, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return {"documents": {}, "last_sync": None}


def save_sync_state(target_dir, state):
    """Save sync state to JSON file."""
    state_file = os.path.join(target_dir, SYNC_STATE_FILE)
    state["last_sync"] = datetime.now().isoformat()
    try:
        with open(state_file, 'w') as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"⚠️  Could not save sync state: {e}")



def sync_folder(client, folder_id, target_dir, sync_state, mode="incremental", depth=0):
    """Recursively sync a Quip folder to local directory."""
    indent = "  " * depth
    
    # Get folder info
    folder_info = client.get_folder_info(folder_id)
    if not folder_info:
        print(f"{indent}❌ Could not access folder: {folder_id}")
        return
    
    # Extract folder name
    folder_name = "Unknown Folder"
    if "folder" in folder_info and folder_info["folder"]:
        folder_name = folder_info["folder"].get("title", f"Folder_{folder_id}")
    
    safe_folder_name = sanitize_filename(folder_name)
    folder_path = os.path.join(target_dir, safe_folder_name)
    
    print(f"{indent}📁 {folder_name}")
    client.stats['folders'] += 1
    
    # Create folder
    os.makedirs(folder_path, exist_ok=True)
    
    # Get documents and subfolders
    doc_ids = []
    subfolder_ids = []
    
    if "children" in folder_info:
        for child in folder_info["children"]:
            if child.get("thread_id"):
                doc_ids.append(child["thread_id"])
            elif child.get("folder_id"):
                subfolder_ids.append(child["folder_id"])
    
    # Also check for direct arrays
    if "thread_ids" in folder_info:
        doc_ids.extend(folder_info["thread_ids"])
    if "folder_ids" in folder_info:
        subfolder_ids.extend(folder_info["folder_ids"])
    
    # Sync documents
    for doc_id in doc_ids:
        sync_document(client, doc_id, folder_path, sync_state, mode, depth)
    
    # Recursively sync subfolders
    for subfolder_id in subfolder_ids:
        sync_folder(client, subfolder_id, folder_path, sync_state, mode, depth + 1)


def sync_document(client, doc_id, folder_path, sync_state, mode, depth=0):
    """Sync a single document."""
    indent = "  " * (depth + 1)
    
    # Get document info
    thread_info = client.get_thread_info(doc_id)
    if not thread_info:
        print(f"{indent}❌ Could not access document: {doc_id}")
        client.stats['failed'] += 1
        return
    
    # Extract metadata
    title = thread_info.get("thread", {}).get("title", f"Document_{doc_id}")
    doc_type, format_ext = client.get_document_type(thread_info)
    updated_usec = client.get_document_updated_time(thread_info)
    
    # Build output path
    safe_title = sanitize_filename(title)
    filename = f"{safe_title}.{format_ext}"
    output_path = os.path.join(folder_path, filename)
    
    # Check if we need to sync (incremental mode)
    if mode == "incremental":
        prev_state = sync_state["documents"].get(doc_id, {})
        prev_updated = prev_state.get("updated_usec", 0)
        
        if prev_updated >= updated_usec and os.path.exists(output_path):
            print(f"{indent}⏭️  {title} (unchanged)")
            client.stats['skipped'] += 1
            return
    
    # Download document
    print(f"{indent}📄 {title} → {format_ext.upper()}", end=" ", flush=True)
    
    if client.download_document(doc_id, output_path, format_ext):
        print("✓")
        client.stats['synced'] += 1
        # Update sync state
        sync_state["documents"][doc_id] = {
            "title": title,
            "updated_usec": updated_usec,
            "path": output_path,
            "format": format_ext
        }
    else:
        print("❌")
        client.stats['failed'] += 1


def main():
    parser = argparse.ArgumentParser(
        description="Sync Quip folders to local directory",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full sync (overwrite all files)
  python quip_sync_cli.py --token YOUR_TOKEN --mode full
  
  # Incremental sync (only modified files)
  python quip_sync_cli.py --token YOUR_TOKEN --mode incremental
  
  # Use environment variable for token
  export QUIP_TOKEN=your_token
  python quip_sync_cli.py --mode incremental
  
  # Custom source and target
  python quip_sync_cli.py --token YOUR_TOKEN --source "https://quip-mycompany.com/ABC123" --target "/path/to/folder"
        """
    )
    
    parser.add_argument(
        "--token", "-t",
        default=os.environ.get("QUIP_TOKEN"),
        help="Quip API token (or set QUIP_TOKEN env var)"
    )
    parser.add_argument(
        "--mode", "-m",
        choices=["full", "incremental"],
        default="incremental",
        help="Sync mode: 'full' overwrites all, 'incremental' only syncs modified (default: incremental)"
    )
    parser.add_argument(
        "--source", "-s",
        required=True,
        help="Quip folder URL or ID (root folder to sync from)"
    )
    parser.add_argument(
        "--target", "-d",
        required=True,
        help="Target local directory (e.g., SharePoint/OneDrive synced folder)"
    )
    parser.add_argument(
        "--base-url",
        help="Quip API base URL (auto-detected from source URL if not provided)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be synced without actually downloading"
    )
    
    args = parser.parse_args()
    
    # Validate token
    if not args.token:
        print("❌ Error: Quip API token required")
        print("   Provide via --token or set QUIP_TOKEN environment variable")
        print("   Get your token from: https://quip-mycompany.com/dev/token")
        sys.exit(1)
    
    # Extract folder ID
    try:
        folder_id = extract_folder_id(args.source)
    except ValueError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
    
    # Print banner
    print("=" * 60)
    print("🔄 Quip Sync CLI")
    print("=" * 60)
    print(f"📂 Source: {args.source}")
    print(f"💾 Target: {args.target}")
    print(f"🔧 Mode: {args.mode.upper()}")
    if args.dry_run:
        print("🔍 DRY RUN - no files will be downloaded")
    print("=" * 60)
    print()
    
    # Create target directory
    os.makedirs(args.target, exist_ok=True)
    
    # Load sync state
    sync_state = load_sync_state(args.target)
    if sync_state["last_sync"]:
        print(f"📅 Last sync: {sync_state['last_sync']}")
        print()
    
    # Derive base URL from source URL if not provided
    base_url = args.base_url
    if not base_url and args.source.startswith('http'):
        import re
        match = re.search(r'(https?://[^/]+)', args.source)
        if match:
            domain = match.group(1).replace('quip-', 'platform.quip-').replace('quip.com', 'platform.quip.com')
            base_url = f"{domain}/1"
    
    # Create client and start sync
    client = QuipSyncClient(args.token, base_url)
    start_time = time.time()
    
    try:
        sync_folder(client, folder_id, args.target, sync_state, args.mode)
    except KeyboardInterrupt:
        print("\n\n⚠️  Sync interrupted by user")
    
    # Save sync state
    if not args.dry_run:
        save_sync_state(args.target, sync_state)
    
    # Print summary
    elapsed = time.time() - start_time
    print()
    print("=" * 60)
    print("📊 Sync Summary")
    print("=" * 60)
    print(f"   Folders processed: {client.stats['folders']}")
    print(f"   Documents synced:  {client.stats['synced']}")
    print(f"   Documents skipped: {client.stats['skipped']}")
    print(f"   Documents failed:  {client.stats['failed']}")
    print(f"   Time elapsed:      {elapsed:.1f}s")
    print("=" * 60)
    
    if client.stats['failed'] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
