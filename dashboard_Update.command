#!/bin/bash
# Mac 로컬 — GitHub 최신 코드 받기 + 서버 재시작
# 더블클릭: fetch → pull(또는 main 동기화) → Chrome 주차 → 조용히 재시작
# 서버를 먼저 끄면 Chrome이 죽은 8501에 커넥트 에러를 띄우므로, 탭을 먼저 치운다.
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

ROOT="${HOME}/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || ROOT="/Users/maegbugpeulom1/Desktop/dashboard"
[ -f "${ROOT}/app.py" ] || {
  osascript -e 'display alert "dashboard 폴더 없음" message "~/Desktop/dashboard 가 있는지 확인하세요."' 2>/dev/null || true
  exit 1
}

HERE="$(cd "$(dirname "$0")" && pwd)"
CANON="$(cd "${ROOT}" && pwd)"
if [ "$HERE" != "$CANON" ] && [ -x "${CANON}/dashboard_Update.command" ]; then
  exec "${CANON}/dashboard_Update.command"
fi

cd "$ROOT" || exit 1

if [ ! -d .git ]; then
  osascript -e 'display alert "git 저장소 아님" message "터미널에서: cd ~/Desktop/dashboard && git clone https://github.com/zimhs/office.git ."' 2>/dev/null || true
  exit 1
fi

_notify() {
  # 알림 실패해도 업데이트·재시작은 계속 (AppleScript -2740 방지)
  NOTIFY_BODY="$1" osascript >/dev/null 2>&1 <<'APPLESCRIPT' || true
display notification (system attribute "NOTIFY_BODY") with title "영업 대시보드"
APPLESCRIPT
}

_pull_ff_only() {
  git pull origin main --ff-only >/dev/null 2>&1
}

_sync_to_origin_main() {
  _notify "로컬 git을 GitHub main과 맞춥니다"
  git reset --hard origin/main >/dev/null 2>&1
}

_fetch_origin_main() {
  local i
  for i in 1 2 3; do
    if git fetch origin main >/dev/null 2>&1; then
      return 0
    fi
    sleep 2
  done
  return 1
}

# 8501 탭을 빈 페이지로 — 서버를 끄기 전에 Chrome 커넥트 에러를 막음
_chrome_park_dashboard_tabs() {
  osascript <<'APPLESCRIPT' 2>/dev/null || true
tell application "Google Chrome"
  repeat with w in windows
    repeat with t in tabs of w
      set theURL to URL of t
      if (theURL starts with "http://127.0.0.1:8501") or (theURL starts with "http://localhost:8501") then
        set URL of t to "about:blank#office-dashboard-park"
      end if
    end repeat
  end repeat
end tell
APPLESCRIPT
}

_quiet_kill_dashboard() {
  lsof -ti:8501 2>/dev/null | xargs kill -9 2>/dev/null || true
  lsof -ti:8502 2>/dev/null | xargs kill -9 2>/dev/null || true
  pkill -f "streamlit run.*8501" 2>/dev/null || true
  pkill -f "streamlit run.*8502" 2>/dev/null || true
}

_notify "GitHub에서 최신 코드 받는 중…"
if ! _fetch_origin_main; then
  osascript -e 'display alert "git fetch 실패" message "인터넷·GitHub 연결을 확인하세요. 잠시 후 Update를 다시 실행해 보세요."' 2>/dev/null || true
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
      osascript -e 'display alert "git 동기화 실패" message "터미널에서 cd ~/Desktop/dashboard && git status 를 확인하세요."' 2>/dev/null || true
      exit 1
    fi
    SYNCED_HARD=true
  elif [ "$STASHED" = true ]; then
    if ! git stash pop >/dev/null 2>&1; then
      _notify "로컬 수정은 stash에 보관됨 · git stash list"
    fi
  fi
fi

AFTER="$(git rev-parse HEAD 2>/dev/null || echo none)"
AFTER_SHORT="$(git rev-parse --short HEAD 2>/dev/null || echo "?")"
# macOS BSD sed는 \s 미지원 → [[:space:]] (따옴표 남으면 예전 -2740 원인)
PI_BUILD="$(
  grep -E '^PI_UI_BUILD[[:space:]]*=' "${ROOT}/price_increase_tab.py" 2>/dev/null | head -1 \
    | sed -E 's/^PI_UI_BUILD[[:space:]]*=[[:space:]]*//; s/^["'\'']//; s/["'\''][[:space:]]*$//; s/\r$//'
)"
[ -n "$PI_BUILD" ] || PI_BUILD="(공문빌드 확인불가)"

if [ "$SYNCED_HARD" = true ]; then
  MSG="동기화 ${AFTER_SHORT} — 재시작 · ${PI_BUILD}"
  _notify "GitHub main과 맞췄습니다. 로컬 수정은 제거되었을 수 있습니다."
elif [ "$BEFORE" = "$AFTER" ]; then
  MSG="이미 최신 ${AFTER_SHORT} · ${PI_BUILD}"
else
  MSG="갱신 ${AFTER_SHORT} — 재시작 · ${PI_BUILD}"
fi
_notify "$MSG"

# Local이 8501이 살아 있어도 반드시 재시작하도록
: >"${ROOT}/.dash_force_restart"
rm -f "${ROOT}/.dash_code_stamp"

# Stop.command는 쓰지 않음 — 알림이 뜨고, Chrome이 죽은 8501에 남아 커넥트 에러가 깜빡임
_chrome_park_dashboard_tabs
_quiet_kill_dashboard
exec "${ROOT}/dashboard_Local.command"
