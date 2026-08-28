#!/bin/bash
# Cloud 메인 + 로컬 업무일지·공문(8502) 동시 실행
#   chmod +x dashboard_Cloud.command
#   xattr -d com.apple.quarantine dashboard_Cloud.command 2>/dev/null

dash_export_path

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

fail() {
  osascript -e "display alert \"Cloud·업무일지 시작 실패\" message \"$1\""
  echo "$1"
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
  fail "python3 를 찾을 수 없습니다."
fi

if ! dash_worklog_app "$ROOT" >/dev/null; then
  fail "업무일지 app.py 없음 — dashboard/업무일지/app.py 를 받으세요."
fi

if ! dash_ensure_worklog "$ROOT"; then
  fail "업무일지(8502) 시작에 실패했습니다. ${ROOT}/.dashboard_8502.log 를 확인하세요."
fi

CLOUD_URL="$(dash_cloud_url "$ROOT")"
dash_open_cloud_and_worklog "$ROOT"

osascript -e "display notification \"Cloud 메인 + 업무일지(8502) 브라우저를 엽니다.\" with title \"영업 대시보드\""

# 8502 는 백그라운드 — 터미널 창 유지(로그 확인용)
echo "Cloud: ${CLOUD_URL}"
echo "업무일지: http://127.0.0.1:8502"
echo "종료: 이 창에서 Ctrl+C (8502 백그라운드 프로세스는 Activity Monitor에서 streamlit 종료)"
read -r -p "Enter 키를 누르면 이 안내 창만 닫습니다..."
