#!/bin/bash
# Cloud 메인 + 로컬 업무일지·공문(8502) — 브라우저 탭 2개만
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

if ! dash_launch_cloud_pair "$ROOT"; then
  fail "업무일지(8502) 시작 실패. ${ROOT}/.dashboard_8502.log 를 확인하세요."
fi

CLOUD_URL="$(dash_cloud_url "$ROOT")"
osascript -e 'display notification "Cloud + 업무일지(8502) 탭 2개를 엽니다." with title "영업 대시보드"'

echo "Cloud: ${CLOUD_URL}"
echo "업무일지: http://127.0.0.1:8502"
