#!/bin/bash
# 메인(8501) + 업무일지·공문(8502) — Chrome 탭 2개 (자동화/Safari 불필요)
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음"'; exit 1; }

MAIN="${ROOT}/app.py"
WORK="${ROOT}/업무일지/app.py"
URL1="http://127.0.0.1:8501"
URL2="http://127.0.0.1:8502"
LOCKDIR="${HOME}/.dashboard_local_running"

# 이전 실행이 비정상 종료되면 잠금 폴더가 남아 아무 것도 안 함 → 자동 삭제
if [ -d "$LOCKDIR" ]; then
  if ! lsof -ti:8501 >/dev/null 2>&1 && ! lsof -ti:8502 >/dev/null 2>&1; then
    rmdir "$LOCKDIR" 2>/dev/null || rm -rf "$LOCKDIR" 2>/dev/null
  fi
fi

if ! mkdir "$LOCKDIR" 2>/dev/null; then
  osascript -e 'display notification "잠시 후 다시 눌러 주세요 (실행 중)" with title "영업 대시보드"'
  exit 0
fi
trap 'rmdir "$LOCKDIR" 2>/dev/null' EXIT

_up() {
  curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1
}

_kill_ports() {
  lsof -ti:8501 2>/dev/null | xargs kill 2>/dev/null || true
  lsof -ti:8502 2>/dev/null | xargs kill 2>/dev/null || true
  sleep 2
}

_run_bg() {
  local port="$1" app="$2"
  _up "$port" && return 0
  (
    cd "$ROOT" || exit 1
    export STREAMLIT_SERVER_HEADLESS=true
    export BROWSER=
    if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
      exec "${ROOT}/.venv/bin/streamlit" run "$app" \
        --server.port="$port" --server.headless=true --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    elif command -v streamlit >/dev/null 2>&1; then
      exec streamlit run "$app" \
        --server.port="$port" --server.headless=true --browser.gatherUsageStats=false \
        >>"${ROOT}/.dash_${port}.log" 2>&1
    else
      exec python3 -m streamlit run "$app" \
        --server.port="$port" --server.headless=true --browser.gatherUsageStats=false \
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

_open_chrome_two_tabs() {
  local u1="$1" u2="$2"
  local bin="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
  if [ ! -x "$bin" ]; then
    osascript -e 'display alert "Google Chrome 없음" message "Chrome 설치 후 다시 실행하세요."'
    return 1
  fi
  # Chrome 한 창 · 탭 2개 (open -na 는 창이 2개 뜸)
  "$bin" --new-window "$u1" "$u2" >/dev/null 2>&1 &
  sleep 0.3
}

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음" message "dashboard/업무일지/app.py 확인"'; exit 1; }

if _up 8501 && _up 8502; then
  _open_chrome_two_tabs "$URL1" "$URL2"
  osascript -e 'display notification "Chrome 한 창 · 탭 2개" with title "영업 대시보드"'
  exit 0
fi

_kill_ports
_run_bg 8502 "$WORK" || { osascript -e 'display alert "8502 시작 실패" message "dashboard/.dash_8502.log 확인"'; exit 1; }
_run_bg 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패" message "dashboard/.dash_8501.log 확인"'; exit 1; }

_open_chrome_two_tabs "$URL1" "$URL2"
osascript -e 'display notification "Chrome 한 창 · 8501+8502" with title "영업 대시보드"'
