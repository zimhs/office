#!/bin/bash
# 8501·8502 전부 종료
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/opt/homebrew/bin:${PATH}"
for P in 8501 8502; do
  lsof -ti:"$P" 2>/dev/null | xargs kill 2>/dev/null || true
done
rm -f "${HOME}/.dashboard_local_once.lock" 2>/dev/null
pkill -f "streamlit run.*8501" 2>/dev/null || true
pkill -f "streamlit run.*8502" 2>/dev/null || true
osascript -e 'display notification "대시보드 종료" with title "영업 대시보드"'
