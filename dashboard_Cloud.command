#!/bin/bash
# Cloud 메인 + 로컬 업무일지(8502) — Chrome 창 1개 · 탭 2개 · 재실행 시 탭 추가 없음
# v2026-08-28d
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

_chrome_has_tab() {
  local needle="$1"
  osascript 2>/dev/null <<APPLESCRIPT
tell application "Google Chrome"
  repeat with w in windows
    repeat with t in tabs of w
      if URL of t contains "$needle" then return "yes"
    end repeat
  end repeat
end tell
return "no"
APPLESCRIPT
}

_chrome_open_cloud_pair() {
  local hasCloud hasLocal
  hasCloud="$(_chrome_has_tab "streamlit.app")"
  hasLocal="$(_chrome_has_tab "127.0.0.1:8502")"
  if [ "$hasCloud" = "yes" ] && [ "$hasLocal" = "yes" ]; then
    osascript -e 'tell application "Google Chrome" to activate' 2>/dev/null || true
    return 0
  fi
  local bin="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
  [ -x "$bin" ] || { osascript -e 'display alert "Google Chrome 없음"'; exit 1; }
  "$bin" --new-window "$CLOUD" "$URL2" >/dev/null 2>&1 &
  sleep 0.5
}

if _up && [ -f "$STAMP" ]; then
  osascript -e 'tell application "Google Chrome" to activate' 2>/dev/null || true
  osascript -e 'display notification "이미 실행 중 (탭 추가 안 함)" with title "영업 대시보드"'
  exit 0
fi

if ! _up; then
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
fi

_chrome_open_cloud_pair
touch "$STAMP"
osascript -e 'display notification "Chrome 한 창 · Cloud+8502 (다시 누르면 탭 추가 안 함)" with title "영업 대시보드"'
