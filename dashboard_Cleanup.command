#!/bin/bash
# dashboard 불필요 파일 정리 — 더블클릭 (macOS)
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/opt/homebrew/bin:${PATH}"
DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT="$DIR/cleanup_dashboard.sh"
chmod +x "$SCRIPT" 2>/dev/null || true
OUT="$("/bin/bash" "$SCRIPT" 2>&1)" || {
  osascript -e "display alert \"정리 실패\" message \"$OUT\"" 2>/dev/null || true
  echo "$OUT"
  read -r -p "Enter…"
  exit 1
}
echo "$OUT"
osascript -e 'display notification "불필요 파일 정리 완료" with title "영업 대시보드"' 2>/dev/null || true
osascript -e "display alert \"정리 완료\" message \"$OUT\"" 2>/dev/null || true
