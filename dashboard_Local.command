#!/bin/bash
# 영업 대시보드(8501) — 업무일지·공문 통합 12탭
# v2026-09-03 — 8502 분리 종료 · Chrome 탭 1개
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음"'; exit 1; }

# 바탕화면 복사본이면 dashboard 폴더 최신본으로 넘김
HERE="$(cd "$(dirname "$0")" && pwd)"
CANON="$(cd "${ROOT}" && pwd)"
if [ "$HERE" != "$CANON" ] && [ -x "${CANON}/dashboard_Local.command" ]; then
  exec "${CANON}/dashboard_Local.command"
fi

MAIN="${ROOT}/app.py"
URL1="http://127.0.0.1:8501"
STAMP="${HOME}/.dashboard_browser_opened"
LAUNCH_LOCK="${HOME}/.dashboard_local_launching"

if ! mkdir "$LAUNCH_LOCK" 2>/dev/null; then
  osascript -e 'display notification "이미 실행 중" with title "영업 대시보드"'
  exit 0
fi
trap 'rmdir "$LAUNCH_LOCK" 2>/dev/null' EXIT

_up() { curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1; }

_mtime() {
  local f="$1"
  [ -f "$f" ] || { echo 0; return; }
  stat -f "%m" "$f" 2>/dev/null || stat -c "%Y" "$f" 2>/dev/null || echo 0
}

_code_stamp() {
  printf "%s|%s|%s|%s|%s|%s" \
    "$(_mtime "${MAIN}")" \
    "$(_mtime "${ROOT}/market_research_tab.py")" \
    "$(_mtime "${ROOT}/worklog_tab.py")" \
    "$(_mtime "${ROOT}/price_increase_tab.py")" \
    "$(_mtime "${ROOT}/drive_autoload.py")" \
    "$(_mtime "${ROOT}/cache_remote_sync.py")"
}

_save_code_stamp() {
  _code_stamp >"${ROOT}/.dash_code_stamp"
}

_code_changed_since_start() {
  local cur saved
  cur="$(_code_stamp)"
  saved=""
  [ -f "${ROOT}/.dash_code_stamp" ] && saved="$(cat "${ROOT}/.dash_code_stamp" 2>/dev/null || true)"
  [ -n "$saved" ] && [ "$cur" != "$saved" ]
}

_chrome_ensure_tab() {
  osascript <<'APPLESCRIPT' 2>/dev/null
tell application "Google Chrome"
  set u1 to "http://127.0.0.1:8501"
  set win1 to missing value
  set idx1 to 0

  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      set theURL to URL of t
      if theURL starts with u1 then
        set win1 to w
        set idx1 to ti
      end if
      set ti to ti + 1
    end repeat
  end repeat

  if win1 is not missing value then
    set index of win1 to 1
    set active tab index of win1 to idx1
    activate
    return
  end if

  make new window
  set URL of active tab of window 1 to u1
  activate
end tell
APPLESCRIPT
  if [ $? -ne 0 ]; then
    osascript -e 'display alert "Chrome 탭 열기 실패" message "Chrome이 실행 중인지, Automation 권한이 허용됐는지 확인하세요."'
    return 1
  fi
}

# 서버가 이미 떠 있으면 재시작 없음 — 단, 코드 변경 시에는 재시작
if _up 8501; then
  if _code_changed_since_start; then
    osascript -e 'display notification "코드 변경 감지 — 서버 재시작합니다" with title "영업 대시보드"'
  else
    _chrome_ensure_tab
    touch "$STAMP"
    osascript -e 'display notification "8501 대시보드로 이동 (통합 12탭)" with title "영업 대시보드"'
    exit 0
  fi
fi

_kill_ports() {
  lsof -ti:8501 2>/dev/null | xargs kill -9 2>/dev/null || true
  lsof -ti:8502 2>/dev/null | xargs kill -9 2>/dev/null || true
  pkill -f "streamlit run.*8501" 2>/dev/null || true
  pkill -f "streamlit run.*8502" 2>/dev/null || true
  sleep 2
}

_run_bg() {
  local port="$1" app="$2"
  (
    cd "$ROOT" || exit 1
    export STREAMLIT_SERVER_HEADLESS=true
    export BROWSER=/usr/bin/true
    if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
      exec "${ROOT}/.venv/bin/streamlit" run "$app" \
        --server.port="$port" --server.headless=true --server.address=127.0.0.1 \
        --browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1 \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    elif command -v streamlit >/dev/null 2>&1; then
      exec streamlit run "$app" \
        --server.port="$port" --server.headless=true --server.address=127.0.0.1 \
        --browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1 \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    else
      exec python3 -m streamlit run "$app" \
        --server.port="$port" --server.headless=true --server.address=127.0.0.1 \
        --browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1 \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    fi
  ) &
  local i
  for i in $(seq 1 45); do _up "$port" && return 0; sleep 1; done
  return 1
}

_kill_ports
_run_bg 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패"'; exit 1; }

_chrome_ensure_tab
touch "$STAMP"
_save_code_stamp
osascript -e 'display notification "Chrome · 8501 통합 대시보드 (업무일지·공문 포함)" with title "영업 대시보드"'
