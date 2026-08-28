#!/bin/bash
# 메인(8501) + 업무일지·공문(8502) — 탭 2개만. 이 파일 하나로 동작.
set -e
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || { osascript -e 'display alert "dashboard 폴더 없음" message "Desktop/dashboard 확인"'; exit 1; }

MAIN="${ROOT}/app.py"
WORK="${ROOT}/업무일지/app.py"
URL1="http://127.0.0.1:8501"
URL2="http://127.0.0.1:8502"
LOCK="${HOME}/.dashboard_local_once.lock"

# 15초 안에 다시 누르면 무시 (탭 폭탄 방지)
if [ -f "$LOCK" ] && [ "$(($(date +%s) - $(stat -f %m "$LOCK" 2>/dev/null || echo 0)))" -lt 15 ]; then
  exit 0
fi
touch "$LOCK"

_run() {
  if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
    "${ROOT}/.venv/bin/streamlit" run "$@"
  elif command -v streamlit >/dev/null 2>&1; then
    streamlit run "$@"
  else
    python3 -m streamlit run "$@"
  fi
}

_up() {
  curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1
}

_start() {
  local port="$1" app="$2"
  _up "$port" && return 0
  (
    cd "$ROOT"
    _run "$app" --server.port="$port" --server.headless=true --browser.gatherUsageStats=false \
      >>"${ROOT}/.dash_${port}.log" 2>&1
  ) &
  for _ in $(seq 1 40); do _up "$port" && return 0; sleep 1; done
  return 1
}

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음" message "dashboard/업무일지/app.py 받으세요"'; exit 1; }

_start 8502 "$WORK" || { osascript -e 'display alert "8502 시작 실패"'; exit 1; }
_start 8501 "$MAIN" || { osascript -e 'display alert "8501 시작 실패"'; exit 1; }

open "$URL1"
sleep 0.5
open "$URL2"

osascript -e 'display notification "8501 + 8502 탭 2개" with title "영업 대시보드"'
