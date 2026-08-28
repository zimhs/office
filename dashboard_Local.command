#!/bin/bash
# 메인(8501) + 업무일지(8502) — Chrome 창 1개 · 탭 2개 · 재실행 시 탭 추가 없음
# v2026-08-28d
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

_chrome_has_tab() {
  local needle="$1"
  osascript 2>/dev/null <<APPLESCRIPT
tell application "Google Chrome"
  repeat with w in windows
    repeat with t in tabs of w
      if URL of t starts with "$needle" then return "yes"
    end repeat
  end repeat
end tell
return "no"
APPLESCRIPT
}

_chrome_focus_tabs() {
  osascript 2>/dev/null <<'APPLESCRIPT'
tell application "Google Chrome"
  set targetWin to missing value
  set targetTab to 0
  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      if URL of t starts with "http://127.0.0.1:8501" then
        set targetWin to w
        set targetTab to ti
        exit repeat
      end if
      set ti to ti + 1
    end repeat
    if targetWin is not missing value then exit repeat
  end repeat
  if targetWin is not missing value then
    set index of targetWin to 1
    set active tab index of targetWin to targetTab
    activate
  end if
end tell
APPLESCRIPT
}

_chrome_open_once() {
  local has1 has2
  has1="$(_chrome_has_tab "http://127.0.0.1:8501")"
  has2="$(_chrome_has_tab "http://127.0.0.1:8502")"
  if [ "$has1" = "yes" ] && [ "$has2" = "yes" ]; then
    _chrome_focus_tabs
    return 0
  fi
  local bin="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
  [ -x "$bin" ] || { osascript -e 'display alert "Google Chrome 없음"'; return 1; }
  "$bin" --new-window "$URL1" "$URL2" >/dev/null 2>&1 &
  sleep 0.5
}

# ★ 서버가 이미 떠 있으면 재시작·탭 추가 없이 Chrome만 포커스
if _up 8501 && _up 8502; then
  _chrome_focus_tabs
  touch "$STAMP"
  osascript -e 'display notification "이미 실행 중 (탭 추가 안 함)" with title "영업 대시보드"'
  exit 0
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

_chrome_open_once
touch "$STAMP"
osascript -e 'display notification "Chrome 한 창 · 탭 2개 (다시 누르면 탭 추가 안 함)" with title "영업 대시보드"'
