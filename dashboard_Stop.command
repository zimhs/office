#!/bin/bash
# 로컬 대시보드(8501·8502) streamlit 프로세스 종료

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

for PORT in 8501 8502; do
  PIDS="$(lsof -ti:"${PORT}" 2>/dev/null || true)"
  if [ -n "$PIDS" ]; then
    echo "종료: 포트 ${PORT} → ${PIDS}"
    kill $PIDS 2>/dev/null || true
  else
    echo "포트 ${PORT}: 실행 중 아님"
  fi
done

rm -rf "${HOME}/.dashboard_launch" 2>/dev/null
rm -f "${HOME}/Desktop/dashboard/.dashboard_"*.starting 2>/dev/null

osascript -e 'display notification "8501·8502 대시보드를 종료했습니다." with title "영업 대시보드"'
