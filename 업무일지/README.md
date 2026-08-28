# 업무일지·공문 전용 대시보드

메인 **영업 분석 대시보드**(`app.py`, 포트 **8501**)와 별도로, **업무일지**와 **공문** 탭만 실행합니다.  
다른 브라우저 창·다른 포트(**8502**)에서 동시에 쓸 수 있습니다.

## 구성

| 항목 | 설명 |
|------|------|
| `app.py` | 진입점 (업무일지 + 공문 2탭) |
| `sales_loader.py` | 공문용 매출 CSV 로더 (`uploaded_cache/sales/`) |
| `dashboard_Local.command` | Mac 더블클릭 실행 (8502) |

**공유 모듈** (상위 `dashboard/` 폴더): `worklog_tab.py`, `price_increase_tab.py`, `dev_mode.py`, `drive_autoload.py`, `worklog_remote_sync.py`, `uploaded_cache/`

## Mac에서 실행

1. `dashboard` 폴더를 GitHub에서 받거나 `git pull`
2. `업무일지/dashboard_Local.command` 더블클릭  
   → `http://127.0.0.1:8502` 브라우저 열림
3. 메인 대시보드는 기존처럼 `dashboard_Local.command` (8501) 사용

터미널:

```bash
cd ~/Desktop/dashboard
python3 -m streamlit run 업무일지/app.py --server.port=8502
```

## 데이터·설정

- **업무일지:** `uploaded_cache/worklog/` (메인과 동일)
- **공문:** `uploaded_cache/price_increase/` + `uploaded_cache/sales/*.csv`
- **secrets:** 상위 `.streamlit/secrets.toml` (Gist, SMTP 등)

## Streamlit Cloud (선택)

별도 앱으로 배포하려면 **Main file path**를 `업무일지/app.py`로 지정합니다.

## curl로 Mac에 반영

```bash
cd ~/Desktop/dashboard
mkdir -p 업무일지
curl -L -o 업무일지/app.py "https://raw.githubusercontent.com/zimhs/office/cursor/worklog-letter-dashboard-4823/업무일지/app.py"
curl -L -o 업무일지/sales_loader.py "https://raw.githubusercontent.com/zimhs/office/cursor/worklog-letter-dashboard-4823/업무일지/sales_loader.py"
curl -L -o 업무일지/dashboard_Local.command "https://raw.githubusercontent.com/zimhs/office/cursor/worklog-letter-dashboard-4823/업무일지/dashboard_Local.command"
chmod +x 업무일지/dashboard_Local.command
```
