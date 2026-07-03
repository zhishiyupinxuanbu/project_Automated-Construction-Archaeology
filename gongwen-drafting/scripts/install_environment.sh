#!/usr/bin/env bash
set -euo pipefail

INSTALL_HOMEBREW=0
WITH_TESSERACT=1
WITH_PADDLEOCR=0
PYTHON_BIN="${PYTHON_BIN:-python3}"

usage() {
  cat <<'EOF'
Install runtime dependencies for the gongwen-drafting skill on macOS.

Usage:
  ./scripts/install_environment.sh [options]

Options:
  --install-homebrew   Install Homebrew if brew is missing.
  --no-tesseract       Skip tesseract and tesseract-lang installation.
  --with-paddleocr     Try to install paddleocr Python packages with pip.
  --python PATH        Python executable to use, default: python3.
  -h, --help           Show this help.

Default installs:
  - Python packages: python-docx, PyYAML
  - Homebrew packages when brew exists: poppler, tesseract, tesseract-lang

Notes:
  - The skill itself contains rules, templates, reference documents, and scripts.
  - This installer prepares the local runtime only.
  - Paddle OCR is not installed by default because macOS/Python/CPU wheels vary.
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --install-homebrew)
      INSTALL_HOMEBREW=1
      shift
      ;;
    --no-tesseract)
      WITH_TESSERACT=0
      shift
      ;;
    --with-paddleocr)
      WITH_PADDLEOCR=1
      shift
      ;;
    --python)
      PYTHON_BIN="${2:-}"
      if [[ -z "$PYTHON_BIN" ]]; then
        echo "[FAIL] --python requires a path" >&2
        exit 2
      fi
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "[FAIL] unknown option: $1" >&2
      usage
      exit 2
      ;;
  esac
done

log() {
  echo "[gongwen-install] $*"
}

if [[ "$(uname -s)" != "Darwin" ]]; then
  log "This installer is intended for macOS. Continuing with Python packages only."
fi

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "[FAIL] Python executable not found: $PYTHON_BIN" >&2
  echo "Install Python 3 first, then rerun this script." >&2
  exit 1
fi

log "Using Python: $("$PYTHON_BIN" -c 'import sys; print(sys.executable)')"
log "Installing Python packages: python-docx PyYAML"
"$PYTHON_BIN" -m pip install --upgrade python-docx PyYAML

if ! command -v brew >/dev/null 2>&1; then
  if [[ "$INSTALL_HOMEBREW" == "1" ]]; then
    log "Homebrew not found. Installing Homebrew."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    if [[ -x /opt/homebrew/bin/brew ]]; then
      eval "$(/opt/homebrew/bin/brew shellenv)"
    elif [[ -x /usr/local/bin/brew ]]; then
      eval "$(/usr/local/bin/brew shellenv)"
    fi
  else
    log "Homebrew not found. Skipping poppler/tesseract."
    log "Install Homebrew first, or rerun with --install-homebrew."
  fi
fi

if command -v brew >/dev/null 2>&1; then
  log "Installing Homebrew package: poppler"
  brew install poppler
  if [[ "$WITH_TESSERACT" == "1" ]]; then
    log "Installing Homebrew packages: tesseract tesseract-lang"
    brew install tesseract tesseract-lang
  fi
fi

if [[ "$WITH_PADDLEOCR" == "1" ]]; then
  log "Trying to install Paddle OCR Python packages."
  "$PYTHON_BIN" -m pip install --upgrade paddleocr paddlepaddle
else
  log "Skipping Paddle OCR install. Use --with-paddleocr if this Mac supports the required wheels."
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [[ -f "$SCRIPT_DIR/check_environment.py" ]]; then
  log "Running environment check."
  "$PYTHON_BIN" "$SCRIPT_DIR/check_environment.py" || {
    log "Environment check reported warnings or failures. Review the lines above."
    exit 1
  }
fi

log "Done."
