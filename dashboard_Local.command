#!/bin/bash
# 메인(8501) + 업무일지·공문(8502) — 새 창에 탭 2개만
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음"'; exit 1; }

MAIN="${ROOT}/app.py"
WORK="${ROOT}/업무일지/app.py"
URL1="http://127.0.0.1:8501"
URL2="http://127.0.0.1:8502"
LOCKDIR="${HOME}/.dashboard_local_running"

if ! mkdir "$LOCKDIR" 2>/dev/null; then
  exit 0
fi
trap 'rmdir "$LOCKDIR" 2>/dev/null' EXIT

_up() {
  curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1
}

_kill_ports() {
  lsof -ti:8501 2>/dev/null | xargs kill 2>/dev/null || true
  lsof -ti:8502 2>/dev/null | xargs kill 2>/dev/null || true
  pkill -f "streamlit run.*8501" 2>/dev/null || true
  pkill -f "streamlit run.*8502" 2>/dev/null || true
  sleep 2
}

_run_bg() {
  local port="$1" app="$2"
  (
    cd "$ROOT" || exit 1
    export STREAMLIT_SERVER_HEADLESS=true
    export BROWSER=
    if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
      exec "${ROOT}/.venv/bin/streamlit" run "$app" \
        --server.port="$port" \
        --server.headless=true \
        --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    elif command -v streamlit >/dev/null 2>&1; then
      exec streamlit run "$app" \
        --server.port="$port" \
        --server.headless=true \
        --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    else
      exec python3 -m streamlit run "$app" \
        --server.port="$port" \
        --server.headless=true \
        --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    fi
  ) &
  local i
  for i in $(seq 1 45); do
    _up "$port" && return 0
    sleep 1
  done
  return 1
}

_open_one_window_two_tabs() {
  local chrome="/Applications/Google Chrome.app"
  if [ ! -d "$chrome" ]; then
    osascript -e 'display alert "Google Chrome 없음" message "Chrome 설치 후 다시 실행하세요."'
    return 1
  fi
  # AppleScript (Chrome 전용) — 실패 시 open -a Chrome 으로 fallback
  if /usr/bin/osascript <<'APPLESCRIPT' 2>/dev/null
tell application "Google Chrome"
  activate
  set w to make new window with properties {URL:"http://127.0.0.1:8501"}
  tell w to make new tab with properties {URL:"http://127.0.0.1:8502"}
end tell
APPLESCRIPT
  then
    return 0
  fi
  open -na "Google Chrome" --args --new-window "$URL1"
  sleep 0.4
  open -a "Google Chrome" "$URL2"
}

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음"'; exit 1; }

# 이미 둘 다 살아 있으면 탭 추가 안 함
if _up 8501 && _up 8502; then
  osascript -e 'display notification "이미 실행 중 — 탭 추가 안 함" with title "영업 대시보드"'
  exit 0
fi

# 꼬인 프로세스·중복 streamlit 정리 후 깨끗이 1개씩만 시작
_kill_ports

_run_bg 8502 "$WORK" || { osascript -e 'display alert "8502 시작 실패"'; exit 1; }
_run_bg 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패"'; exit 1; }

_open_one_window_two_tabs
osascript -e 'display notification "새 창 · 탭 2개 (8501+8502)" with title "영업 대시보드"'
