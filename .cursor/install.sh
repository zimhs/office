#!/usr/bin/env bash
# Cloud Agent install script for the 통합 영업 분석 대시보드 (Streamlit) app.
# Idempotent: safe to run repeatedly and against cached/partially-prepared state.
set -euo pipefail

cd "$(dirname "$0")/.."

# opendartreader>=0.3.2 (see requirements.txt) requires Python >= 3.13, which is
# newer than Ubuntu 24.04's default python3.12. Install 3.13 from deadsnakes.
if ! command -v python3.13 >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y -qq software-properties-common
  sudo add-apt-repository -y ppa:deadsnakes/ppa
  sudo apt-get update -qq
  sudo apt-get install -y -qq python3.13 python3.13-venv python3.13-dev
fi

if [ ! -x ".venv/bin/python" ]; then
  python3.13 -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt

echo "✅ Environment ready: $(python --version) with dependencies installed."
