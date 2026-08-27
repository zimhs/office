#!/bin/bash
# dashboard 전체 코드·설정·캐시 백업 (Mac / Linux)
# dashboard_Backup.command 또는 터미널에서 실행

set -euo pipefail

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

if [ ! -f "$DIR/app.py" ]; then
  echo "오류: app.py 가 있는 dashboard 폴더를 찾을 수 없습니다."
  exit 1
fi

BACKUP_ROOT="$DIR/백업"
SNAPSHOTS="$BACKUP_ROOT/snapshots"
STAMP="$(date +%Y-%m-%d_%H%M%S)"
TARGET="$SNAPSHOTS/$STAMP"
KEEP="${DASHBOARD_BACKUP_KEEP:-15}"

mkdir -p "$SNAPSHOTS"

echo "백업 시작: $TARGET"

if command -v rsync >/dev/null 2>&1; then
  rsync -a \
    --exclude '백업/' \
    --exclude '.git/' \
    --exclude '__pycache__/' \
    --exclude '*.pyc' \
    --exclude '.DS_Store' \
    --exclude '.agents/' \
    --exclude '.cursor/' \
    --exclude '.devcontainer/' \
    "$DIR/" "$TARGET/"
else
  mkdir -p "$TARGET"
  tar -C "$DIR" \
    --exclude='./백업' \
    --exclude='./.git' \
    --exclude='./__pycache__' \
    --exclude='*.pyc' \
    --exclude='.DS_Store' \
    -cf - . | tar -C "$TARGET" -xf -
fi

# 최신 백업 바로가기 (Mac Finder에서 latest 더블클릭)
LATEST="$SNAPSHOTS/latest"
rm -f "$LATEST"
ln -s "$STAMP" "$LATEST"

# 백업 정보
{
  echo "백업 시각: $(date '+%Y-%m-%d %H:%M:%S')"
  echo "원본 경로: $DIR"
  echo "스냅샷: $STAMP"
  du -sh "$TARGET" 2>/dev/null | awk '{print "용량: "$1}'
} > "$TARGET/BACKUP_INFO.txt"

# 오래된 스냅샷 정리 (latest 제외)
if [ "$KEEP" -gt 0 ] 2>/dev/null; then
  ls -1dt "$SNAPSHOTS"/[0-9][0-9][0-9][0-9]-* 2>/dev/null | tail -n +$((KEEP + 1)) | while read -r old; do
    [ -n "$old" ] || continue
    rm -rf "$old"
    echo "삭제(보관 한도): $(basename "$old")"
  done
fi

echo "백업 완료: $TARGET"
echo "$TARGET"
