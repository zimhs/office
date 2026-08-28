#!/bin/bash
# 업무일지·공문 전용 Streamlit (포트 8502 — 메인 대시보드 8501 과 분리)
# 사용: dashboard/업무일지 폴더에서 더블클릭
#   chmod +x dashboard_Local.command

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$DIR/.." && pwd)"
APP="$DIR/app.py"

if [ ! -f "$APP" ]; then
  osascript -e 'display alert "업무일지 app.py 없음" message "dashboard/업무일지/app.py 가 있는지 확인하세요."'
  exit 1
fi

PORT="${STREAMLIT_PORT:-8502}"
URL="http://127.0.0.1:${PORT}"

fail() {
  osascript -e "display alert \"업무일지·공문 시작 실패\" message \"$1\""
  echo "$1"
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
  fail "python3 를 찾을 수 없습니다."
fi

if curl -sf "${URL}/_stcore/health" >/dev/null 2>&1; then
  open "${URL}"
  exit 0
fi

osascript -e 'display notification "업무일지·공문 대시보드를 시작합니다 (8502)." with title "업무일지"'

cd "$ROOT" || fail "dashboard 루트로 이동 실패: $ROOT"

REQ="$DIR/requirements.txt"
if [ ! -f "$REQ" ] && [ -f "$ROOT/requirements.txt" ]; then
  REQ="$ROOT/requirements.txt"
fi

if [ -x "$ROOT/.venv/bin/streamlit" ]; then
  exec "$ROOT/.venv/bin/streamlit" run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif command -v streamlit >/dev/null 2>&1; then
  exec streamlit run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif python3 -m streamlit --version >/dev/null 2>&1; then
  exec python3 -m streamlit run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
else
  fail "streamlit 이 없습니다.\ncd \"$ROOT\" && python3 -m pip install -r \"$REQ\""
fi
