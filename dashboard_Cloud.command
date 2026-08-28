#!/bin/bash
# Cloud 영업일보만 — Chrome 창 1개 · 탭 1개 (로컬 8502·업무 탭 없음)
# v2026-08-28f
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
CLOUD="https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app"
secrets="${ROOT}/.streamlit/secrets.toml"
if [ -f "$secrets" ]; then
  _url="$(grep -E '^[[:space:]]*dashboard_cloud_url[[:space:]]*=' "$secrets" 2>/dev/null | head -1 | sed -E 's/^[^=]*=[[:space:]]*["'\'' ]*([^"'\'']+)["'\'' ]*/\1/' | tr -d '\r')"
  [ -n "$_url" ] && CLOUD="${_url%/}"
fi
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

[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음"'; exit 1; }

# Chrome CLI --new-window 는 세션 복원으로 네이버·캘린더 창이 따로 뜨는 경우가 있어 AppleScript 로만 탭 제어
_chrome_ensure_cloud_tab() {
  osascript <<APPLESCRIPT 2>/dev/null
tell application "Google Chrome"
  set uCloud to "$CLOUD"
  set uLocal to "http://127.0.0.1:8502"
  set targetWin to missing value
  set idxCloud to 0

  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      set theURL to URL of t
      if theURL contains "streamlit.app" then
        set targetWin to w
        set idxCloud to ti
        exit repeat
      end if
      set ti to ti + 1
    end repeat
    if targetWin is not missing value then exit repeat
  end repeat

  if targetWin is missing value then
    if (count of windows) > 0 then
      set targetWin to window 1
      set URL of active tab of targetWin to uCloud
      set idxCloud to 1
    else
      make new window
      set targetWin to window 1
      set URL of active tab of targetWin to uCloud
      set idxCloud to 1
    end if
  end if

  -- 같은 창의 로컬 업무(8502) 탭 제거
  repeat with ti from (count of tabs of targetWin) to 1 by -1
    if URL of tab ti of targetWin starts with uLocal then
      close tab ti of targetWin
      if ti < idxCloud then set idxCloud to idxCloud - 1
    end if
  end repeat

  set index of targetWin to 1
  set active tab index of targetWin to idxCloud
  activate

  -- 네이버·캘린더만 있는 창 닫기 (영업일보 streamlit 없음)
  repeat with w in windows
    if w is targetWin then
      -- keep
    else
      set hasStreamlit to false
      set hasNaverCal to false
      repeat with t in tabs of w
        set theURL to URL of t
        if theURL contains "streamlit.app" then set hasStreamlit to true
        if theURL contains "naver.com" or theURL contains "calendar.google.com" then set hasNaverCal to true
      end repeat
      if hasNaverCal and not hasStreamlit then
        try
          close w
        end try
      end if
    end if
  end repeat
end tell
APPLESCRIPT
  if [ $? -ne 0 ]; then
    osascript -e 'display alert "Chrome 탭 열기 실패" message "Chrome이 실행 중인지, Automation 권한이 허용됐는지 확인하세요."'
    return 1
  fi
}

_chrome_ensure_cloud_tab
touch "$STAMP"
osascript -e 'display notification "Chrome · Cloud 영업일보 탭 1개" with title "영업 대시보드"'
