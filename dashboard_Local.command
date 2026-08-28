#!/bin/bash
# 로컬 Streamlit: 메인(8501) + 업무일지·공문(8502) — 브라우저 탭 2개만
# 바탕화면 복사본이어도 dashboard 폴더 최신본으로 자동 실행됩니다.

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

HERE="$(cd "$(dirname "$0")" && pwd)"

_dash_bootstrap_root() {
  if [ -f "${HERE}/app.py" ] && [ -f "${HERE}/dashboard_launch_common.sh" ]; then
    echo "$HERE"
    return 0
  fi
  local cand
  for cand in \
    "${HOME}/Desktop/dashboard" \
    "/Users/maegbugpeulom1/Desktop/dashboard" \
    "${HOME}/Desktop/dashboard-main" \
    "${HERE}/dashboard"
  do
    if [ -f "${cand}/app.py" ] && [ -f "${cand}/dashboard_launch_common.sh" ]; then
      echo "$cand"
      return 0
    fi
  done
  return 1
}

ROOT="$(_dash_bootstrap_root)" || {
  osascript -e 'display alert "프로젝트 폴더를 찾을 수 없습니다." message "Desktop/dashboard 에 app.py 와 dashboard_launch_common.sh 가 있는지 확인하세요."'
  exit 1
}

# shellcheck source=dashboard_launch_common.sh
. "${ROOT}/dashboard_launch_common.sh"

dash_exec_canonical "$HERE" "$ROOT" "dashboard_Local.command"

if ! dash_launch_lock "local_pair"; then
  exit 0
fi

fail() {
  osascript -e "display alert \"로컬 대시보드 시작 실패\" message \"$1\""
  echo "$1"
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
  fail "python3 를 찾을 수 없습니다."
fi

if ! dash_worklog_app "$ROOT" >/dev/null; then
  osascript -e 'display alert "업무일지 app.py 없음" message "dashboard/업무일지/app.py 를 받은 뒤 다시 실행하세요."'
fi

if ! dash_launch_local_pair "$ROOT"; then
  fail "대시보드 시작 실패. ${ROOT}/.dashboard_8501.log 또는 .dashboard_8502.log 를 확인하세요."
fi

osascript -e 'display notification "메인(8501) + 업무일지(8502) 탭 2개를 엽니다." with title "영업 대시보드"'
