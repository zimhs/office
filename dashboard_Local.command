#!/bin/bash
# 메인(8501) + 업무일지·공문(8502) — 브라우저 탭 정확히 2개
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음"'; exit 1; }

MAIN="${ROOT}/app.py"
WORK="${ROOT}/업무일지/app.py"
URL1="http://127.0.0.1:8501"
URL2="http://127.0.0.1:8502"
LOCKDIR="${HOME}/.dashboard_local_running"

# 동시에 두 번 실행되면 즉시 종료
if ! mkdir "$LOCKDIR" 2>/dev/null; then
  exit 0
fi
trap 'rmdir "$LOCKDIR" 2>/dev/null' EXIT

_up() {
  curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1
}

_run_bg() {
  local port="$1" app="$2"
  _up "$port" && return 0
  (
    cd "$ROOT" || exit 1
    export STREAMLIT_SERVER_HEADLESS=true
    export BROWSER=
    if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
      "${ROOT}/.venv/bin/streamlit" run "$app" \
        --server.port="$port" \
        --server.headless=true \
        --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    elif command -v streamlit >/dev/null 2>&1; then
      streamlit run "$app" \
        --server.port="$port" \
        --server.headless=true \
        --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    else
      python3 -m streamlit run "$app" \
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

_open_exactly_two_tabs() {
  /usr/bin/osascript <<'APPLESCRIPT'
set u1 to "http://127.0.0.1:8501"
set u2 to "http://127.0.0.1:8502"
try
  tell application "Google Chrome"
    if (count of windows) = 0 then
      make new window with properties {URL:u1}
    else
      tell window 1 to make new tab with properties {URL:u1}
    end if
    tell window 1 to make new tab with properties {URL:u2}
    activate
    return
  end tell
end try
try
  tell application "Safari"
    if (count of windows) = 0 then
      make new document with properties {URL:u1}
    else
      tell window 1 to set current tab to (make new tab with properties {URL:u1})
    end if
    tell window 1 to make new tab with properties {URL:u2}
    activate
    return
  end tell
end try
do shell script "open " & quoted form of u1
delay 0.5
do shell script "open " & quoted form of u2
APPLESCRIPT
}

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음"'; exit 1; }

ALREADY1=0
ALREADY2=0
_up 8501 && ALREADY1=1
_up 8502 && ALREADY2=1

# 이미 둘 다 켜져 있으면 브라우저 안 열음
if [ "$ALREADY1" = 1 ] && [ "$ALREADY2" = 1 ]; then
  osascript -e 'display notification "이미 실행 중 (탭 추가 안 함)" with title "영업 대시보드"'
  exit 0
fi

_run_bg 8502 "$WORK" || { osascript -e 'display alert "8502 시작 실패"'; exit 1; }
_run_bg 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패"'; exit 1; }

_open_exactly_two_tabs
osascript -e 'display notification "8501 + 8502 탭 2개" with title "영업 대시보드"'
