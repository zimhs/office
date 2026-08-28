#!/bin/bash
# 완전 초기화 — 서버 종료 + 다음 Local 실행 때만 Chrome 탭 열림
export PATH="/usr/local/bin:/opt/homebrew/bin:${PATH}"

for P in 8501 8502; do
  lsof -ti:"$P" 2>/dev/null | xargs kill -9 2>/dev/null || true
done
pkill -f "streamlit run" 2>/dev/null || true

rm -f "${HOME}/.dashboard_browser_opened"
rm -f "${HOME}/.dashboard_cloud_browser_opened"
rm -f "${HOME}/Desktop/dashboard/.dash_code_stamp"
rm -f "${HOME}/.dashboard_local.flock"
rm -rf "${HOME}/.dashboard_local_launching" 2>/dev/null
rm -rf "${HOME}/.dashboard_cloud_launching" 2>/dev/null
rm -rf "${HOME}/.dashboard_local_running" 2>/dev/null

osascript -e 'display notification "초기화 완료 — Chrome localhost 탭은 직접 닫고 Local 한 번만 실행" with title "영업 대시보드"'
