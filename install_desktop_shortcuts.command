#!/bin/bash
# 바탕화면 바로가기를 dashboard 폴더 스크립트 심볼릭 링크로 교체 (한 번만 실행)
#   chmod +x install_desktop_shortcuts.command

DASH="${HOME}/Desktop/dashboard"

if [ ! -f "${DASH}/dashboard_Local.command" ]; then
  osascript -e 'display alert "dashboard 폴더 없음" message "Desktop/dashboard 가 있는지 확인하세요."'
  exit 1
fi

ln -sf "${DASH}/dashboard_Local.command" "${HOME}/Desktop/dashboard_Local.command"
ln -sf "${DASH}/dashboard_Cloud.command" "${HOME}/Desktop/dashboard_Cloud.command"

chmod +x "${DASH}/dashboard_Local.command"
chmod +x "${DASH}/dashboard_Cloud.command"
chmod +x "${DASH}/dashboard_launch_common.sh"

xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Local.command" 2>/dev/null
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Cloud.command" 2>/dev/null

osascript -e 'display notification "바탕화면 바로가기를 dashboard 폴더 최신본에 연결했습니다." with title "영업 대시보드"'

echo "완료: Desktop/dashboard_Local.command → 심볼릭 링크"
echo "완료: Desktop/dashboard_Cloud.command → 심볼릭 링크"
