영업 대시보드 — 로컬 백업 폴더
================================

■ 원클릭 백업 (Mac)
  dashboard 폴더의 dashboard_Backup.command 를 더블클릭

■ 터미널 백업
  cd ~/Desktop/dashboard
  ./backup_dashboard.sh

■ 저장 위치
  백업/snapshots/YYYY-MM-DD_HHMMSS/   ← 날짜·시간별 전체 복사본
  백업/snapshots/latest/              ← 가장 최근 백업 (바로가기)

■ 포함
  app.py, worklog_tab.py, 공문·시장조사 탭, requirements.txt,
  .streamlit 설정, uploaded_cache(캐시·템플릿), xlsx 등

■ 제외 (용량·중복 방지)
  백업/ 폴더 자체, .git, __pycache__, .DS_Store

■ 보관 개수
  기본 최근 15개 (환경변수 DASHBOARD_BACKUP_KEEP=20 등으로 변경)

■ 복원 예시
  cp -R "백업/snapshots/latest/worklog_tab.py" ./worklog_tab.py
