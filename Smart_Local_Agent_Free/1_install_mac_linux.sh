#!/usr/bin/env bash
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
GEMINI_HOME="$HOME/.gemini"

echo "======================================================================"
echo "  Smart Local Document Agent - 100% Free Setup (macOS / Linux)"
echo "======================================================================"

if ! command -v node >/dev/null 2>&1; then
  echo "[!] Please install Node.js 20+ first (https://nodejs.org/ or brew install node)"
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "[!] Please install Python 3 first (https://python.org/ or brew install python)"
  exit 1
fi

echo "[1/3] Installing Google Gemini CLI (@google/gemini-cli) [Optional if using Ollama/Manual]..."
GEMINI_CLI_VER="${GEMINI_CLI_VER:-0}"
npm install -g "@google/gemini-cli@${GEMINI_CLI_VER}" || echo "[WARN] Skipping Gemini CLI install; Ollama and API keys remain available."

echo "[2/3] Installing Python libraries from requirements.txt..."
python3 -m pip install --user --upgrade -r "$SCRIPT_DIR/requirements.txt" || \
python3 -m pip install --break-system-packages --upgrade -r "$SCRIPT_DIR/requirements.txt"

echo "[3/3] Installing GEMINI.md playbook, 12 slash commands, and doc_tools..."
mkdir -p "$GEMINI_HOME/commands" "$GEMINI_HOME/doc_tools" "$GEMINI_HOME/templates"
[ -f "$GEMINI_HOME/GEMINI.md" ] && cp "$GEMINI_HOME/GEMINI.md" "$GEMINI_HOME/GEMINI.md.bak"
cp "$SCRIPT_DIR/GEMINI.md" "$GEMINI_HOME/GEMINI.md"
cp "$SCRIPT_DIR/commands/"*.toml "$GEMINI_HOME/commands/"
cp "$SCRIPT_DIR/doc_tools/"*.py "$GEMINI_HOME/doc_tools/"
cp "$SCRIPT_DIR/doc_tools/"*.html "$GEMINI_HOME/doc_tools/"
cp "$SCRIPT_DIR/doc_tools/"*.json "$GEMINI_HOME/doc_tools/"
for ico in "$SCRIPT_DIR/doc_tools/"app_icon*; do
  [ -f "$ico" ] && cp "$ico" "$GEMINI_HOME/doc_tools/"
done
mkdir -p "$HOME/Desktop/Smart_Projects_Hub"
cp "$SCRIPT_DIR/Agent_Guide_And_Commands.html" "$HOME/Desktop/Smart_Projects_Hub/0_OPEN_GUIDE_AR.html"
chmod +x "$GEMINI_HOME/doc_tools/"*.py "$SCRIPT_DIR/2_start_agent.command" "$SCRIPT_DIR/3_start_scan_watcher_mac_linux.command"

echo ""
echo "[OK] Installed successfully into $GEMINI_HOME ($0 Cost)"
echo "     Run 'gemini' inside ANY project folder and pick 'Login with Google'!"
