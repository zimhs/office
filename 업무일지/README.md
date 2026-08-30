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

### 코드 반영 (중요)

Cloud/PR에서 수정한 내용은 **Mac `~/Desktop/dashboard`에 자동으로 안 들어갑니다.** 아래 중 하나를 실행하세요.

**방법 A — 바로가기 (추천)**

1. `install_desktop_shortcuts.command` 한 번 실행 → 바탕화면에 `dashboard_Update.command` 생김
2. 수정 반영할 때마다 **`dashboard_Update.command` 더블클릭**
   - `git pull` → 서버 종료 → `dashboard_Local.command` 재시작

**방법 B — 터미널**

```bash
cd ~/Desktop/dashboard
git pull origin main
bash dashboard_Stop.command
bash dashboard_Local.command
```

**반영 확인:** 업무일지(8502) 상단 빌드 표시  
예) `2026-08-30 · 업무일지공문2탭 · 2026-08-30g · 익일빈줄·요약공백복구`  
`2026-08-28b`만 보이면 **아직 구버전**입니다.

> `dashboard_Local.command`만 누르면 **이미 켜진 서버를 그대로 쓰는 경우**가 많습니다.  
> `worklog_tab.py`만 바뀐 경우 예전에는 재시작을 안 했으나, 지금은 변경 감지에 포함됩니다.  
> 그래도 안 바뀌면 **`dashboard_Update.command`** 또는 **Stop → Local** 순서로 실행하세요.

### 실행

1. `dashboard` 폴더를 GitHub에서 받거나 `git pull`
2. **바탕화면 바로가기** (추천):
   ```bash
   cp ~/Desktop/dashboard/업무일지_Dashboard.command ~/Desktop/
   chmod +x ~/Desktop/업무일지_Dashboard.command
   xattr -d com.apple.quarantine ~/Desktop/업무일지_Dashboard.command 2>/dev/null
   ```
   → Desktop **`업무일지_Dashboard.command`** 더블클릭 · `http://127.0.0.1:8502`
3. 또는 `업무일지/dashboard_Local.command` 더블클릭
4. 메인 대시보드는 기존처럼 `dashboard_Local.command` (8501) 사용

**한 번에 두 개:** `dashboard_Local.command` → **메인 8501 + 업무일지 8502** 동시 실행 · `dashboard_Cloud.command` → **Cloud 메인 + 업무일지 8502**

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
