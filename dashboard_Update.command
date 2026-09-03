#!/bin/bash
# Mac 로컬 — GitHub 최신 코드 받기 + 서버 재시작
# 더블클릭: fetch → pull(또는 main 동기화) → Stop → Local (8501+8502)
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

_pull_ff_only() {
  git pull origin main --ff-only >/dev/null 2>&1
}

_sync_to_origin_main() {
  osascript -e 'display notification "로컬 git을 GitHub main과 맞춥니다" with title "영업 대시보드"'
  git reset --hard origin/main >/dev/null 2>&1
}

osascript -e 'display notification "GitHub에서 최신 코드 받는 중…" with title "영업 대시보드"'
if ! git fetch origin main 2>&1; then
  osascript -e 'display alert "git fetch 실패" message "인터넷·GitHub 연결을 확인하세요."'
  exit 1
fi

BEFORE="$(git rev-parse HEAD 2>/dev/null || echo none)"
SYNCED_HARD=false
STASHED=false

if ! _pull_ff_only; then
  if ! git diff --quiet 2>/dev/null || ! git diff --cached --quiet 2>/dev/null || [ -n "$(git ls-files --others --exclude-standard 2>/dev/null)" ]; then
    if git stash push -u -m "dashboard_Update auto $(date +%Y%m%d-%H%M%S)" >/dev/null 2>&1; then
      STASHED=true
    fi
  fi

  if ! _pull_ff_only; then
    if ! _sync_to_origin_main; then
      osascript -e 'display alert "git 동기화 실패" message "터미널에서 cd ~/Desktop/dashboard && git status 를 확인하세요."'
      exit 1
    fi
    SYNCED_HARD=true
  elif [ "$STASHED" = true ]; then
    if ! git stash pop >/dev/null 2>&1; then
      osascript -e 'display notification "로컬 수정은 stash에 보관됨 · git stash list" with title "영업 대시보드"'
    fi
  fi
fi

AFTER="$(git rev-parse HEAD 2>/dev/null || echo none)"
AFTER_SHORT="$(git rev-parse --short HEAD 2>/dev/null || echo "?")"
PI_BUILD="$(grep -E '^PI_UI_BUILD\s*=' "${ROOT}/price_increase_tab.py" 2>/dev/null | head -1 | sed -E 's/^PI_UI_BUILD\s*=\s*//; s/^["'\'']//; s/["'\'']\s*$//')"
[ -n "$PI_BUILD" ] || PI_BUILD="(공문빌드 확인불가)"

if [ "$SYNCED_HARD" = true ]; then
  osascript -e 'display alert "GitHub main과 동기화했습니다" message "로컬 git 커밋/수정은 제거되었을 수 있습니다. 필요하면 터미널에서 git stash list 로 확인하세요."'
  MSG="동기화 ${AFTER_SHORT} — 재시작 · ${PI_BUILD}"
elif [ "$BEFORE" = "$AFTER" ]; then
  MSG="이미 최신 ${AFTER_SHORT} · ${PI_BUILD}"
else
  MSG="갱신 ${AFTER_SHORT} — 재시작 · ${PI_BUILD}"
fi
osascript -e "display notification \"${MSG}\" with title \"영업 대시보드\""

# 코드 stamp 지워서 Local이 반드시 재시작하도록
rm -f "${ROOT}/.dash_code_stamp"

[ -x "${ROOT}/dashboard_Stop.command" ] && bash "${ROOT}/dashboard_Stop.command"
sleep 2
exec "${ROOT}/dashboard_Local.command"
