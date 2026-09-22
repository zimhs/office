#!/bin/bash
# 바탕화면 바로가기 = dashboard 폴더 파일 링크 (한 번만)
D="${HOME}/Desktop/dashboard"
chmod +x "${D}/dashboard_Local.command" "${D}/dashboard_Cloud.command" "${D}/dashboard_Stop.command" "${D}/dashboard_Update.command"
rm -f "${HOME}/Desktop/dashboard_r.command" "${HOME}/Desktop/dashevoard_c.command"
osascript <<EOF
tell application "Finder"
  set dest to folder "Desktop" of home
  try
    delete (item "dashboard_r" of dest)
  end try
  try
    delete (item "dashevoard_c" of dest)
  end try
  set localCmd to POSIX file "${D}/dashboard_Local.command" as alias
  set cloudCmd to POSIX file "${D}/dashboard_Cloud.command" as alias
  make new alias file at dest to localCmd with properties {name:"dashboard_r"}
  make new alias file at dest to cloudCmd with properties {name:"dashevoard_c"}
end tell
EOF
osascript <<EOF
use framework "AppKit"
use scripting additions
set ws to current application's NSWorkspace's sharedWorkspace()
set imgR to current application's NSImage's alloc()'s initWithContentsOfFile:"${D}/icons/dashboard_r.png"
set imgC to current application's NSImage's alloc()'s initWithContentsOfFile:"${D}/icons/dashevoard_c.png"
ws's setIcon:imgR forFile:"${HOME}/Desktop/dashboard_r" options:0
ws's setIcon:imgC forFile:"${HOME}/Desktop/dashevoard_c" options:0
EOF
ln -sf "${D}/dashboard_Stop.command" "${HOME}/Desktop/dashboard_Stop.command"
ln -sf "${D}/dashboard_Update.command" "${HOME}/Desktop/dashboard_Update.command"
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Stop.command" 2>/dev/null
xattr -d com.apple.quarantine "${HOME}/Desktop/dashboard_Update.command" 2>/dev/null
osascript -e 'display notification "바탕화면 바로가기 연결 완료" with title "영업 대시보드"'
