#!/bin/bash
# 업무일지·공문만 (8502) — 메인 Local.command 가 8502도 켜므로 여기서는 브라우저만
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/opt/homebrew/bin:${PATH}"

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
URL="http://127.0.0.1:8502"

curl -sf "${URL}/_stcore/health" >/dev/null 2>&1 || {
  osascript -e 'display alert "8502가 꺼져 있습니다" message "먼저 dashboard_Local.command 를 실행하세요."'
  exit 1
}

open "$URL"
osascript -e 'display notification "업무일지(8502) 탭 1개" with title "업무일지"'
