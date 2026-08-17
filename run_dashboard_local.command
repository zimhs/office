#!/bin/bash
# 로컬 Streamlit 대시보드 (macOS 메모 연동용)
# Cloud URL이 아니라 이 맥에서 app.py 를 실행합니다.
set -euo pipefail
DIR="$(cd "$(dirname "$0")" && pwd)"
# Desktop 바로가기에서 실행해도 프로젝트 폴더로 이동
if [ ! -f "$DIR/app.py" ]; then
  DIR="/Users/maegbugpeulom1/Desktop/dashboard"
fi
cd "$DIR"
PORT="${STREAMLIT_PORT:-8501}"
URL="http://127.0.0.1:${PORT}"

if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display alert "python3 를 찾을 수 없습니다." message "Python 3를 설치한 뒤 다시 실행하세요."'
  exit 1
fi

if curl -sf "${URL}/_stcore/health" >/dev/null 2>&1; then
  open "${URL}"
  exit 0
fi

osascript -e 'display notification "로컬 대시보드를 시작합니다." with title "영업 대시보드"'

if command -v streamlit >/dev/null 2>&1; then
  exec streamlit run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif [ -x "$DIR/.venv/bin/streamlit" ]; then
  exec "$DIR/.venv/bin/streamlit" run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
else
  exec python3 -m streamlit run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
fi
