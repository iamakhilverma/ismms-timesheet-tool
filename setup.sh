#!/bin/bash
# One-time setup. Creates an ISOLATED virtualenv (.venv) inside this folder and
# installs the tool's dependencies there. It does not touch system Python or any
# other tools. Safe to re-run. Works on Apple Silicon and Intel Macs.
set -euo pipefail
cd "$(dirname "$0")"

echo "==> Setting up timesheet tool in: $(pwd)"

# 1. Find a usable python3 (>= 3.9)
PY=""
for cand in python3 /usr/bin/python3 /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$cand" >/dev/null 2>&1; then
    if "$cand" -c 'import sys; raise SystemExit(0 if sys.version_info[:2] >= (3,9) else 1)' 2>/dev/null; then
      PY="$cand"; break
    fi
  fi
done
if [ -z "$PY" ]; then
  echo "ERROR: No Python 3.9+ found."
  echo "Install Apple's command line tools (includes python3):  xcode-select --install"
  echo "...then re-run ./setup.sh"
  exit 1
fi
echo "==> Using Python: $("$PY" -V) ($PY)"

# 2. Create the isolated venv
if [ ! -d .venv ]; then
  echo "==> Creating virtualenv (.venv)"
  "$PY" -m venv .venv
fi

# 3. Install dependencies into the venv only
echo "==> Installing dependencies into .venv (system Python untouched)"
./.venv/bin/python -m pip install --upgrade pip >/dev/null
./.venv/bin/python -m pip install -r requirements.txt

# 4. Create the ./timesheet launcher that always uses the venv
cat > timesheet <<EOF
#!/bin/bash
# Auto-generated launcher. Runs the tool inside its isolated venv.
exec "$(pwd)/.venv/bin/python" "$(pwd)/timesheet.py" "\$@"
EOF
chmod +x timesheet

# 5. Add a global 'timesheet' command (a zsh alias -- no sudo, runs from anywhere).
#    The tool reads its own config/template by location, so the working dir doesn't matter.
ZSHRC="$HOME/.zshrc"
MARKER="# ismms-timesheet-tool"
ALIAS_LINE="alias timesheet=\"$(pwd)/timesheet\"  $MARKER"
add_alias() {
  touch "$ZSHRC"
  # drop any prior line we added (keeps the path current if the repo moved), then append
  grep -v "$MARKER" "$ZSHRC" > "$ZSHRC.tmp" 2>/dev/null && mv "$ZSHRC.tmp" "$ZSHRC"
  printf '\n%s\n' "$ALIAS_LINE" >> "$ZSHRC"
  echo "==> Added a global 'timesheet' alias to ~/.zshrc."
  echo "    Activate it now with:  source ~/.zshrc   (or just open a new terminal)"
}
if [ -t 0 ]; then
  printf "Add a global 'timesheet' command to ~/.zshrc so you can run it from anywhere? [Y/n] "
  read ans
  case "$ans" in [Nn]*) echo "Skipped (use ./timesheet from this folder, or re-run setup to add it later)." ;; *) add_alias ;; esac
else
  add_alias
fi

echo
echo "==> Done. Verifying:"
./timesheet doctor || true
echo
echo "Next steps (from anywhere once the alias is active, or ./timesheet here):"
echo "  timesheet setup                 # choose Outlook or Graph for this machine"
echo "  timesheet draft                 # stage this week's email"
echo "  timesheet draft --week next      # next week's submission"
