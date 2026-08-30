#!/bin/bash
# 바탕화면 바로가기 = dashboard 폴더 파일 링크 (한 번만)
D="${HOME}/Desktop/dashboard"
ln -sf "${D}/dashboard_Local.command" "${HOME}/Desktop/dashboard_Local.command"
ln -sf "${D}/dashboard_Cloud.command" "${HOME}/Desktop/dashboard_Cloud.command"
ln -sf "${D}/dashboard_Stop.command" "${HOME}/Desktop/dashboard_Stop.command"
ln -sf "${D}/dashboard_Update.command" "${HOME}/Desktop/dashboard_Update.command"
chmod +x "${D}/dashboard_Local.command" "${D}/dashboard_Cloud.command" "${D}/dashboard_Stop.command" "${D}/dashboard_Update.command"
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Local.command" 2>/dev/null
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Cloud.command" 2>/dev/null
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Stop.command" 2>/dev/null
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Update.command" 2>/dev/null
osascript -e 'display notification "바탕화면 바로가기 연결 완료" with title "영업 대시보드"'
