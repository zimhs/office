#!/bin/bash
# Cloud 메인 + 로컬 업무일지(8502) — Chrome 창 1개 · 탭 2개 · 재실행 시 탭 추가 없음
# v2026-08-28e
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
WORK="${ROOT}/업무일지/app.py"
CLOUD="https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app"
URL2="http://127.0.0.1:8502"
STAMP="${HOME}/.dashboard_cloud_browser_opened"
LAUNCH_LOCK="${HOME}/.dashboard_cloud_launching"

HERE="$(cd "$(dirname "$0")" && pwd)"
CANON="$(cd "${ROOT}" && pwd)"
if [ "$HERE" != "$CANON" ] && [ -x "${CANON}/dashboard_Cloud.command" ]; then
  exec "${CANON}/dashboard_Cloud.command"
fi

if ! mkdir "$LAUNCH_LOCK" 2>/dev/null; then
  osascript -e 'display notification "이미 실행 중" with title "영업 대시보드"'
  exit 0
fi
trap 'rmdir "$LAUNCH_LOCK" 2>/dev/null' EXIT

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음"'; exit 1; }

_up() { curl -sf "${URL2}/_stcore/health" >/dev/null 2>&1; }

_chrome_ensure_cloud_tabs() {
  osascript <<APPLESCRIPT 2>/dev/null
tell application "Google Chrome"
  set uCloud to "$CLOUD"
  set uLocal to "$URL2"
  set winCloud to missing value
  set idxCloud to 0
  set winLocal to missing value
  set idxLocal to 0

  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      set theURL to URL of t
      if theURL contains "streamlit.app" then
        set winCloud to w
        set idxCloud to ti
      end if
      if theURL starts with uLocal then
        set winLocal to w
        set idxLocal to ti
      end if
      set ti to ti + 1
    end repeat
  end repeat

  if winCloud is not missing value and winLocal is not missing value then
    set index of winCloud to 1
    set active tab index of winCloud to idxCloud
    activate
    return
  end if

  if winCloud is missing value and winLocal is missing value then
    make new window
    set URL of active tab of window 1 to uCloud
    make new tab at end of window 1 with properties {URL:uLocal}
    set active tab index of window 1 to 1
    activate
    return
  end if

  if winCloud is not missing value then
    set targetWin to winCloud
    if winLocal is missing value then
      make new tab at end of targetWin with properties {URL:uLocal}
    end if
    set index of targetWin to 1
    set active tab index of targetWin to idxCloud
    activate
    return
  end if

  set targetWin to winLocal
  make new tab at end of targetWin with properties {URL:uCloud}
  set index of targetWin to 1
  repeat with ti from 1 to count of tabs of targetWin
    if URL of tab ti of targetWin contains "streamlit.app" then
      set active tab index of targetWin to ti
      exit repeat
    end if
  end repeat
  activate
end tell
APPLESCRIPT
}

if _up; then
  _chrome_ensure_cloud_tabs
  touch "$STAMP"
  osascript -e 'display notification "Cloud+8502 탭으로 이동" with title "영업 대시보드"'
  exit 0
fi

(
  cd "$ROOT"
  export STREAMLIT_SERVER_HEADLESS=true
  export BROWSER=/usr/bin/true
  if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
    "${ROOT}/.venv/bin/streamlit" run "$WORK" \
      --server.port=8502 --server.headless=true --server.address=127.0.0.1 \
      --browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1 \
      >>"${ROOT}/.dash_8502.log" 2>&1
  else
    streamlit run "$WORK" \
      --server.port=8502 --server.headless=true --server.address=127.0.0.1 \
      --browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1 \
      >>"${ROOT}/.dash_8502.log" 2>&1
  fi
) &
for _ in $(seq 1 40); do _up && break; sleep 1; done

_chrome_ensure_cloud_tabs
touch "$STAMP"
osascript -e 'display notification "Chrome · Cloud+8502 탭 2개" with title "영업 대시보드"'
