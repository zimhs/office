#!/bin/bash
# Cloud 영업일보만 — Chrome 창 1개 · 탭 1개 (로컬 8502·업무 탭 없음)
# v2026-09-08a
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

# Chrome만 기동. URL을 여기서 열면 AppleScript와 겹쳐 탭이 2개 생김.
if ! pgrep -qx "Google Chrome" 2>/dev/null; then
  open -a "Google Chrome" 2>/dev/null || true
  sleep 1.2
fi

# Chrome CLI --new-window 는 세션 복원으로 네이버·캘린더 창이 따로 뜨는 경우가 있어 AppleScript 로만 탭 제어
_chrome_ensure_cloud_tab() {
  osascript <<APPLESCRIPT 2>/dev/null
tell application "Google Chrome"
  set uCloud to "$CLOUD"
  set uLocal to "http://127.0.0.1:8502"
  set keepWin to missing value
  set keepIdx to 0

  repeat with w in windows
    set ti to 1
    repeat with t in tabs of w
      set theURL to URL of t
      if theURL contains "streamlit.app" then
        if keepWin is missing value then
          set keepWin to w
          set keepIdx to ti
        else
          try
            close t
          end try
        end if
      end if
      set ti to ti + 1
    end repeat
  end repeat

  if keepWin is missing value then
    if (count of windows) > 0 then
      set keepWin to window 1
      make new tab at end of tabs of keepWin with properties {URL:uCloud}
      set keepIdx to (count of tabs of keepWin)
    else
      make new window
      set keepWin to window 1
      set URL of active tab of keepWin to uCloud
      set keepIdx to 1
    end if
  end if

  -- 같은 창의 로컬 업무(8502) 탭 제거
  repeat with ti from (count of tabs of keepWin) to 1 by -1
    if URL of tab ti of keepWin starts with uLocal then
      close tab ti of keepWin
      if ti < keepIdx then set keepIdx to keepIdx - 1
    end if
  end repeat

  -- 같은 창에 남은 streamlit.app 중복 탭 제거
  set seenCloud to false
  repeat with ti from (count of tabs of keepWin) to 1 by -1
    if URL of tab ti of keepWin contains "streamlit.app" then
      if seenCloud then
        try
          close tab ti of keepWin
        end try
        if ti < keepIdx then set keepIdx to keepIdx - 1
      else
        set seenCloud to true
        set keepIdx to ti
      end if
    end if
  end repeat

  set index of keepWin to 1
  set active tab index of keepWin to keepIdx
  activate
end tell
APPLESCRIPT
  if [ $? -ne 0 ]; then
    osascript -e 'display alert "Chrome 탭 열기 실패" message "Chrome이 실행 중인지, Automation 권한이 허용됐는지 확인하세요."'
    return 1
  fi
}

# 세션 복원이 늦게 붙는 두 번째 streamlit 탭을 한 번 더 접음
_chrome_ensure_cloud_tab
sleep 1.2
_chrome_ensure_cloud_tab
touch "$STAMP"
osascript -e 'display notification "Chrome · Cloud 영업일보 탭 1개" with title "영업 대시보드"'
