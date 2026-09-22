#!/bin/bash
# dashboard 폴더 불필요 파일 정리 (Mac / Linux)
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
if [ ! -f "$DIR/app.py" ]; then
  for cand in "${HOME}/Desktop/dashboard" "/Users/maegbugpeulom1/Desktop/dashboard"; do
    [ -f "${cand}/app.py" ] && DIR="$cand" && break
  done
fi
cd "$DIR" || exit 1

removed=0
rm_path() {
  for p in "$@"; do
    [ -e "$p" ] || continue
    rm -rf "$p"
    echo "삭제: $p"
    removed=$((removed + 1))
  done
}

# 구 백업·로그·스크린샷·캐시 (실행 데이터 uproad·uploaded_cache 는 유지)
rm_path \
  "app_test.py" \
  "dashboard_최종본_20260816_1954" \
  "통합 영업 분석 대시보드.html" \
  "통합 영업 분석 대시보드_files" \
  "엑셀원본.xlsx" \
  "탄산단가인상공문.xlsx" \
  "__pycache__" \
  ".dash_8501.log" \
  ".dash_8502.log" \
  "백업/dashboard_old_backup" \
  "스크린샷 2026-09-07 오후 10.59.31.png" \
  "*.pyc"

# __pycache__ 하위
find "$DIR" -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null || true
find "$DIR" -name "*.pyc" -delete 2>/dev/null || true
find "$DIR" -name ".DS_Store" -delete 2>/dev/null || true

echo "정리 완료 (${removed}개 항목). dashboard 실행에 필요한 파일만 남았습니다."
