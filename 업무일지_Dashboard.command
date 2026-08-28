#!/bin/bash
# 바탕화면 업무일지·공문 대시보드 바로가기 (macOS, 포트 8502)
# 사용: Desktop 에 두고 더블클릭 (dashboard 폴더와 같은 Desktop 이면 자동 탐색)
#   chmod +x ~/Desktop/업무일지_Dashboard.command
#   xattr -d com.apple.quarantine ~/Desktop/업무일지_Dashboard.command 2>/dev/null

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT=""
APP=""

_resolve_app() {
  local base="$1"
  if [ -f "${base}/업무일지/app.py" ]; then
    ROOT="$base"
    APP="${base}/업무일지/app.py"
    return 0
  fi
  return 1
}

# 1) 스크립트가 dashboard 폴더 안에 있을 때
_resolve_app "$HERE" || true

# 2) 바탕화면·일반 경로에서 dashboard 찾기
if [ -z "$APP" ]; then
  for cand in \
    "${HOME}/Desktop/dashboard" \
    "/Users/maegbugpeulom1/Desktop/dashboard" \
    "${HOME}/Desktop/dashboard-main" \
    "${HERE}/dashboard"
  do
    if _resolve_app "$cand"; then
      break
    fi
  done
fi

if [ -z "$APP" ] || [ ! -f "$APP" ]; then
  osascript -e 'display alert "업무일지 app.py 없음" message "Desktop/dashboard/업무일지/app.py 가 있는지 확인하세요. git pull 또는 curl로 업무일지 폴더를 받으세요."'
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

cd "$ROOT" || fail "dashboard 폴더로 이동 실패: $ROOT"

REQ="${ROOT}/업무일지/requirements.txt"
if [ ! -f "$REQ" ] && [ -f "${ROOT}/requirements.txt" ]; then
  REQ="${ROOT}/requirements.txt"
fi

if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
  exec "${ROOT}/.venv/bin/streamlit" run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif command -v streamlit >/dev/null 2>&1; then
  exec streamlit run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif python3 -m streamlit --version >/dev/null 2>&1; then
  exec python3 -m streamlit run "$APP" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
else
  fail "streamlit 이 없습니다.\ncd \"$ROOT\" && python3 -m pip install -r \"$REQ\""
fi
