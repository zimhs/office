#!/bin/bash
# dashboard 전체 백업 — 더블클릭 원클릭 (macOS)
#   chmod +x dashboard_Backup.command
#   xattr -d com.apple.quarantine dashboard_Backup.command 2>/dev/null

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT="$DIR/backup_dashboard.sh"

notify() {
  osascript -e "display notification \"$2\" with title \"$1\"" 2>/dev/null || true
}

alert() {
  osascript -e "display alert \"$1\" message \"$2\"" 2>/dev/null || true
}

if [ ! -x "$SCRIPT" ]; then
  chmod +x "$SCRIPT" 2>/dev/null || true
fi

if [ ! -f "$SCRIPT" ]; then
  alert "백업 스크립트 없음" "backup_dashboard.sh 가 dashboard 폴더에 있어야 합니다."
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
fi

notify "영업 대시보드" "전체 백업을 시작합니다…"

OUT="$("/bin/bash" "$SCRIPT" 2>&1)" || {
  alert "백업 실패" "$OUT"
  echo "$OUT"
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
}

echo "$OUT"
TARGET="$(echo "$OUT" | tail -n 1)"
FOLDER="$DIR/백업/snapshots"

notify "영업 대시보드" "백업 완료 — 백업/snapshots 폴더"
alert "백업 완료" "저장 위치:\n${TARGET}\n\n최신: 백업/snapshots/latest"

open "$FOLDER" 2>/dev/null || open "$DIR/백업" 2>/dev/null || true
