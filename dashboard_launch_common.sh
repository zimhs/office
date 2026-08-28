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

# 바탕화면 복사본(구버전) → dashboard 폴더 최신 스크립트로 넘김
dash_exec_canonical() {
  local here="$1"
  local root="$2"
  local name="$3"
  local canon="${root}/${name}"
  if [ ! -f "$canon" ]; then
    return 1
  fi
  local here_abs canon_dir
  here_abs="$(cd "$here" && pwd)"
  canon_dir="$(cd "$(dirname "$canon")" && pwd)"
  if [ "$here_abs" != "$canon_dir" ]; then
    exec "$canon"
  fi
  return 0
}

dash_launch_lock() {
  local key="$1"
  local lock_dir="${HOME}/.dashboard_launch"
  local lock_file="${lock_dir}/${key}.lock"
  mkdir -p "$lock_dir"
  if ! mkdir "${lock_file}.d" 2>/dev/null; then
    return 1
  fi
  echo "$$" >"${lock_file}.d/pid"
  trap 'rm -rf "${lock_file}.d"' EXIT
  return 0
}

dash_start_bg() {
  local root="$1"
  local port="$2"
  local app="$3"
  local log="${root}/.dashboard_${port}.log"

  if dash_health "$port"; then
    return 0
  fi

  if [ -f "${root}/.dashboard_${port}.starting" ]; then
    local i
    for i in $(seq 1 30); do
      if dash_health "$port"; then
        return 0
      fi
      sleep 1
    done
  fi

  touch "${root}/.dashboard_${port}.starting"
  (
    cd "$root" || exit 1
    dash_streamlit_run "$root" "$app" \
      --server.port="$port" \
      --server.headless=true \
      --browser.gatherUsageStats=false \
      >>"$log" 2>&1
    rm -f "${root}/.dashboard_${port}.starting"
  ) &

  local i
  for i in $(seq 1 45); do
    if dash_health "$port"; then
      rm -f "${root}/.dashboard_${port}.starting"
      return 0
    fi
    sleep 1
  done
  rm -f "${root}/.dashboard_${port}.starting"
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

dash_open_two_urls() {
  local url1="$1"
  local url2="$2"
  open "$url1"
  sleep 0.4
  open "$url2"
}

dash_open_both_local() {
  dash_open_two_urls "http://127.0.0.1:8501" "http://127.0.0.1:8502"
}

dash_open_cloud_only() {
  open "$(dash_cloud_url "$1")"
}

dash_open_cloud_and_worklog() {
  dash_open_cloud_only "$1"
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
  dash_open_cloud_only "$root"
  return 0
}
