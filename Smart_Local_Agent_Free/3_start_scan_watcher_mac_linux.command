#!/usr/bin/env bash
WATCHER="$HOME/.gemini/doc_tools/scan_watcher.py"
if [ ! -f "$WATCHER" ]; then
  SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
  WATCHER="$SCRIPT_DIR/doc_tools/scan_watcher.py"
fi
if [ ! -f "$WATCHER" ]; then
  echo "[ERROR] scan_watcher.py not found. Please run 1_install_mac_linux.sh first."
  exit 1
fi
CFG="$HOME/.gemini/scan_watcher_config.json"

if [ ! -f "$CFG" ]; then
  echo "=== Auto Scanner Watcher - First Time Setup ==="
  read -r -p "Scanner Inbox folder [default: $HOME/Desktop/Smart_Scan/1_Scan_Inbox]: " INBOX
  read -r -p "Word/Excel Drafts folder [default: $HOME/Desktop/Smart_Scan/2_Drafts_Word_Excel]: " DRAFTS
  read -r -p "Organized Archive folder [default: $HOME/Desktop/Smart_Scan/3_Organized_Archive]: " ARCH
  read -r -p "Layout: [1] by_type  [2] by_project_then_type (default 1): " LC
  LAYOUT="by_type"
  [ "$LC" = "2" ] && LAYOUT="by_project_then_type"
  ARGS=(--config "$CFG" --layout "$LAYOUT" --setup-only)
  [ -n "$INBOX" ] && ARGS+=(--inbox "$INBOX")
  [ -n "$DRAFTS" ] && ARGS+=(--drafts "$DRAFTS")
  [ -n "$ARCH" ] && ARGS+=(--archive "$ARCH")
  python3 "$WATCHER" "${ARGS[@]}"
fi

exec python3 "$WATCHER" --config "$CFG"
