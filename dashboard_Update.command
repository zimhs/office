#!/bin/bash
# Mac 로컬 — GitHub 최신 코드 받기 + 서버 재시작
# 더블클릭: git pull → Stop → Local (8501+8502)
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || {
  osascript -e 'display alert "dashboard 폴더 없음" message "~/Desktop/dashboard 가 있는지 확인하세요."'
  exit 1
}

HERE="$(cd "$(dirname "$0")" && pwd)"
CANON="$(cd "${ROOT}" && pwd)"
if [ "$HERE" != "$CANON" ] && [ -x "${CANON}/dashboard_Update.command" ]; then
  exec "${CANON}/dashboard_Update.command"
fi

cd "$ROOT" || exit 1

if [ ! -d .git ]; then
  osascript -e 'display alert "git 저장소 아님" message "터미널에서: cd ~/Desktop/dashboard && git clone https://github.com/zimhs/office.git ."'
  exit 1
fi

osascript -e 'display notification "GitHub에서 최신 코드 받는 중…" with title "영업 대시보드"'
if ! git fetch origin main 2>&1; then
  osascript -e 'display alert "git fetch 실패" message "인터넷·GitHub 연결을 확인하세요."'
  exit 1
fi

BEFORE="$(git rev-parse HEAD 2>/dev/null || echo none)"
if ! git pull origin main --ff-only 2>&1; then
  osascript -e 'display alert "git pull 실패" message "로컬 수정과 충돌했을 수 있습니다. 터미널에서 cd ~/Desktop/dashboard && git status 로 확인하세요."'
  exit 1
fi
AFTER="$(git rev-parse HEAD 2>/dev/null || echo none)"

if [ "$BEFORE" = "$AFTER" ]; then
  MSG="이미 최신입니다"
else
  MSG="코드 갱신됨 — 서버 재시작합니다"
fi
osascript -e "display notification \"${MSG}\" with title \"영업 대시보드\""

[ -x "${ROOT}/dashboard_Stop.command" ] && bash "${ROOT}/dashboard_Stop.command"
sleep 2
exec "${ROOT}/dashboard_Local.command"
