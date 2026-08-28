#!/bin/bash
# 8501·8502·streamlit 전부 종료 (탭 폭탄 초기화)
export PATH="/usr/local/bin:/opt/homebrew/bin:${PATH}"

for P in 8501 8502; do
  lsof -ti:"$P" 2>/dev/null | xargs kill -9 2>/dev/null || true
done
pkill -f "streamlit run" 2>/dev/null || true
rm -rf "${HOME}/.dashboard_local_running" 2>/dev/null
rm -f "${HOME}/.dashboard_local_once.lock" 2>/dev/null

osascript -e 'display notification "8501·8502 종료 — 브라우저 탭은 직접 닫아 주세요" with title "영업 대시보드"'
