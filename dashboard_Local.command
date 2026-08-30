#!/bin/bash
# 메인(8501) + 업무일지(8502) — Chrome 창 1개 · 탭 2개 · 재실행 시 탭 추가 없음
# v2026-08-30 — worklog_tab 등 변경 시 자동 재시작
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
WORK="${ROOT}/업무일지/app.py"
URL1="http://127.0.0.1:8501"
URL2="http://127.0.0.1:8502"
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
    "$(_mtime "${WORK}")" \
    "$(_mtime "${ROOT}/worklog_tab.py")" \
    "$(_mtime "${ROOT}/price_increase_tab.py")" \
    "$(_mtime "${ROOT}/drive_autoload.py")"
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

# Chrome CLI --new-window 는 '이전 세션 복원'으로 네이버·캘린더만 뜨는 경우가 있어 AppleScript 로만 탭 제어
_chrome_ensure_tabs() {
  osascript <<'APPLESCRIPT' 2>/dev/null
tell application "Google Chrome"
  set u1 to "http://127.0.0.1:8501"
  set u2 to "http://127.0.0.1:8502"
  set win1 to missing value
  set idx1 to 0
  set win2 to missing value
  set idx2 to 0

  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      set theURL to URL of t
      if theURL starts with u1 then
        set win1 to w
        set idx1 to ti
      end if
      if theURL starts with u2 then
        set win2 to w
        set idx2 to ti
      end if
      set ti to ti + 1
    end repeat
  end repeat

  if win1 is not missing value and win2 is not missing value then
    set index of win1 to 1
    set active tab index of win1 to idx1
    activate
    return
  end if

  if win1 is missing value and win2 is missing value then
    make new window
    set URL of active tab of window 1 to u1
    make new tab at end of window 1 with properties {URL:u2}
    set active tab index of window 1 to 1
    activate
    return
  end if

  if win1 is not missing value then
    set targetWin to win1
    if win2 is missing value then
      make new tab at end of targetWin with properties {URL:u2}
    end if
    set index of targetWin to 1
    set active tab index of targetWin to idx1
    activate
    return
  end if

  set targetWin to win2
  make new tab at end of targetWin with properties {URL:u1}
  set index of targetWin to 1
  repeat with ti from 1 to count of tabs of targetWin
    if URL of tab ti of targetWin starts with u1 then
      set active tab index of targetWin to ti
      exit repeat
    end if
  end repeat
  activate
end tell
APPLESCRIPT
  if [ $? -ne 0 ]; then
    osascript -e 'display alert "Chrome 탭 열기 실패" message "Chrome이 실행 중인지, Automation 권한이 허용됐는지 확인하세요."'
    return 1
  fi
}

# ★ 서버가 이미 떠 있으면 재시작 없음 — 단, app.py 등 코드 변경 시에는 재시작
if _up 8501 && _up 8502; then
  if _code_changed_since_start; then
    osascript -e 'display notification "코드 변경 감지 — 서버 재시작합니다" with title "영업 대시보드"'
  else
    _chrome_ensure_tabs
    touch "$STAMP"
    osascript -e 'display notification "대시보드 탭으로 이동 (서버 유지 · 코드 변경 시 자동 재시작)" with title "영업 대시보드"'
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

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음"'; exit 1; }

_kill_ports
_run_bg 8502 "$WORK" || { osascript -e 'display alert "8502 시작 실패"'; exit 1; }
_run_bg 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패"'; exit 1; }

_chrome_ensure_tabs
touch "$STAMP"
_save_code_stamp
osascript -e 'display notification "Chrome · 8501+8502 탭 2개 (코드 반영됨)" with title "영업 대시보드"'
