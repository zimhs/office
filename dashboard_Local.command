#!/bin/bash
# 로컬 Streamlit 대시보드 바로가기 (macOS)
# 사용: app.py 와 같은 폴더에 두고 더블클릭
#   chmod +x dashboard_Local.command
#   xattr -d com.apple.quarantine dashboard_Local.command 2>/dev/null

export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"

DIR="$(cd "$(dirname "$0")" && pwd)"
if [ ! -f "$DIR/app.py" ]; then
  for cand in \
    "${HOME}/Desktop/dashboard" \
    "/Users/maegbugpeulom1/Desktop/dashboard" \
    "${HOME}/Desktop/dashboard-main" \
    "${DIR}/dashboard"
  do
    if [ -f "${cand}/app.py" ]; then
      DIR="$cand"
      break
    fi
  done
fi

cd "$DIR" || {
  osascript -e 'display alert "프로젝트 폴더를 찾을 수 없습니다." message "app.py 가 있는 폴더(예: Desktop/dashboard)에 이 파일을 넣어 주세요."'
  exit 1
}

if [ ! -f "$DIR/app.py" ]; then
  osascript -e "display alert \"app.py 없음\" message \"현재 폴더: $DIR\""
  exit 1
fi

PORT="${STREAMLIT_PORT:-8501}"
URL="http://127.0.0.1:${PORT}"

fail() {
  osascript -e "display alert \"로컬 대시보드 시작 실패\" message \"$1\""
  echo "$1"
  read -r -p "Enter 키를 누르면 종료합니다..."
  exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
  fail "python3 를 찾을 수 없습니다. Python을 설치한 뒤 다시 실행하세요."
fi

# 이미 실행 중이면 브라우저만 연다
if curl -sf "${URL}/_stcore/health" >/dev/null 2>&1; then
  open "${URL}"
  exit 0
fi

osascript -e 'display notification "로컬 대시보드를 시작합니다." with title "영업 대시보드"'

if [ -x "$DIR/.venv/bin/streamlit" ]; then
  exec "$DIR/.venv/bin/streamlit" run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif command -v streamlit >/dev/null 2>&1; then
  exec streamlit run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
elif python3 -m streamlit --version >/dev/null 2>&1; then
  exec python3 -m streamlit run "$DIR/app.py" --server.port="$PORT" --server.headless=false --browser.gatherUsageStats=false
else
  fail "streamlit 이 없습니다. 터미널에서:\ncd \"$DIR\" && python3 -m pip install -r requirements.txt"
fi
