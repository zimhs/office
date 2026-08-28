#!/bin/bash
# Cloud 메인 + 로컬 업무일지(8502) — 탭 2개만
set -e
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
WORK="${ROOT}/업무일지/app.py"
CLOUD="https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app"
LOCK="${HOME}/.dashboard_cloud_once.lock"

if [ -f "$LOCK" ] && [ "$(($(date +%s) - $(stat -f %m "$LOCK" 2>/dev/null || echo 0)))" -lt 15 ]; then
  exit 0
fi
touch "$LOCK"

[ -f "$WORK" ] || { osascript -e 'display alert "업무일지 없음"'; exit 1; }

_up() { curl -sf "http://127.0.0.1:8502/_stcore/health" >/dev/null 2>&1; }

if ! _up; then
  (
    cd "$ROOT"
    if [ -x "${ROOT}/.venv/bin/streamlit" ]; then
      "${ROOT}/.venv/bin/streamlit" run "$WORK" --server.port=8502 --server.headless=true --browser.gatherUsageStats=false
    else
      streamlit run "$WORK" --server.port=8502 --server.headless=true --browser.gatherUsageStats=false
    fi
  ) &
  for _ in $(seq 1 40); do _up && break; sleep 1; done
fi

open "$CLOUD"
sleep 0.5
open "http://127.0.0.1:8502"
osascript -e 'display notification "Cloud + 8502 탭 2개" with title "영업 대시보드"'
