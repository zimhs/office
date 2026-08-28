#!/bin/bash
# dashboard_Local.command · dashboard_Cloud.command 공통 실행 헬퍼

dash_export_path() {
  export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/Library/Frameworks/Python.framework/Versions/3.12/bin:/usr/local/bin:/opt/homebrew/bin:${HOME}/.local/bin:${PATH}"
}

dash_find_root() {
  local here="$1"
  if [ -f "${here}/app.py" ]; then
    echo "$here"
    return 0
  fi
  local cand
  for cand in \
    "${HOME}/Desktop/dashboard" \
    "/Users/maegbugpeulom1/Desktop/dashboard" \
    "${HOME}/Desktop/dashboard-main" \
    "${here}/dashboard"
  do
    if [ -f "${cand}/app.py" ]; then
      echo "$cand"
      return 0
    fi
  done
  return 1
}

dash_worklog_app() {
  local root="$1"
  if [ -f "${root}/업무일지/app.py" ]; then
    echo "${root}/업무일지/app.py"
    return 0
  fi
  return 1
}

dash_cloud_url() {
  local root="$1"
  local url=""
  local secrets="${root}/.streamlit/secrets.toml"
  if [ -f "$secrets" ]; then
    url="$(grep -E '^[[:space:]]*dashboard_cloud_url[[:space:]]*=' "$secrets" 2>/dev/null | head -1 | sed -E 's/^[^=]*=[[:space:]]*["'\'' ]*([^"'\'']+)["'\'' ]*/\1/' | tr -d '\r')"
  fi
  if [ -z "$url" ]; then
    url="${DASHBOARD_CLOUD_URL:-https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app}"
  fi
  echo "${url%/}"
}

dash_health() {
  curl -sf "http://127.0.0.1:$1/_stcore/health" >/dev/null 2>&1
}

dash_streamlit_run() {
  local root="$1"
  shift
  if [ -x "${root}/.venv/bin/streamlit" ]; then
    "${root}/.venv/bin/streamlit" run "$@"
  elif command -v streamlit >/dev/null 2>&1; then
    streamlit run "$@"
  else
    python3 -m streamlit run "$@"
  fi
}

dash_start_bg() {
  local root="$1"
  local port="$2"
  local app="$3"
  local log="${root}/.dashboard_${port}.log"

  if dash_health "$port"; then
    return 0
  fi

  (
    cd "$root" || exit 1
    dash_streamlit_run "$root" "$app" \
      --server.port="$port" \
      --server.headless=true \
      --browser.gatherUsageStats=false \
      >>"$log" 2>&1
  ) &

  local i
  for i in $(seq 1 45); do
    if dash_health "$port"; then
      return 0
    fi
    sleep 1
  done
  return 1
}

dash_ensure_main() {
  local root="$1"
  dash_start_bg "$root" 8501 "${root}/app.py"
}

dash_ensure_worklog() {
  local root="$1"
  local app
  app="$(dash_worklog_app "$root")" || return 1
  dash_start_bg "$root" 8502 "$app"
}

# Streamlit 자동 브라우저는 쓰지 않고, open 은 여기서만 (중복 탭 방지)
dash_open_both_local() {
  open "http://127.0.0.1:8501"
  open "http://127.0.0.1:8502"
}

dash_open_cloud_and_worklog() {
  local root="$1"
  open "$(dash_cloud_url "$root")"
  open "http://127.0.0.1:8502"
}

dash_launch_local_pair() {
  local root="$1"
  dash_ensure_worklog "$root" || return 1
  dash_ensure_main "$root" || return 1
  dash_open_both_local
  return 0
}

dash_launch_cloud_pair() {
  local root="$1"
  dash_ensure_worklog "$root" || return 1
  dash_open_cloud_and_worklog "$root"
  return 0
}
