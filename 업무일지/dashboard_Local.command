#!/bin/bash
# 업무일지·공문만 (8502) — 메인 Local.command 가 8502도 켜므로 여기서는 브라우저만
# v2026-08-28d — 탭이 이미 있으면 포커스만
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/opt/homebrew/bin:${PATH}"

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
URL="http://127.0.0.1:8502"

curl -sf "${URL}/_stcore/health" >/dev/null 2>&1 || {
  osascript -e 'display alert "8502가 꺼져 있습니다" message "먼저 dashboard_Local.command 를 실행하세요."'
  exit 1
}

has_tab="$(osascript 2>/dev/null <<'APPLESCRIPT'
tell application "Google Chrome"
  repeat with w in windows
    repeat with t in tabs of w
      if URL of t starts with "http://127.0.0.1:8502" then return "yes"
    end repeat
  end repeat
end tell
return "no"
APPLESCRIPT
)"

if [ "$has_tab" = "yes" ]; then
  osascript -e 'tell application "Google Chrome" to activate' 2>/dev/null || true
  osascript -e 'display notification "8502 탭으로 이동 (새 탭 없음)" with title "업무일지"'
  exit 0
fi

bin="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
if [ -x "$bin" ]; then
  "$bin" "$URL" >/dev/null 2>&1 &
else
  open -a "Google Chrome" "$URL"
fi
osascript -e 'display notification "업무일지(8502) 탭 1개" with title "업무일지"'
