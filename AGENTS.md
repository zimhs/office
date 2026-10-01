# Agent notes

## PR · 머지 · 로컬 Update · Cloud/iPad (필수)

이 저장소 작업이 끝나면 **항상** 아래까지 진행한다. 사용자가 매번 「머지」를 말할 필요 없다.

1. 수정 완료 → 테스트 → 커밋 → 푸시 → PR 생성/갱신
2. 테스트가 통과하면 draft를 해제하고 **`main`에 머지**한다 (`gh pr merge --merge`)
3. 머지 후 **Streamlit Cloud 재배포가 돌아가게** 한다  
   - `main` 푸시만으로 Community Cloud가 자동 재배포하는 것이 기본  
   - 배포가 안 보이면 `main`에 재배포용 커밋을 추가해 트리거한다  
   - Cloud URL: `https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app`  
   - **아이패드는 같은 Cloud URL** (별도 빌드 없음) → Cloud 배포 = iPad 반영
4. 사용자 안내 (짧게):
   - 맥 로컬: **`dashboard_Update`** (Deploy·`dashboard_r` 아님)
   - Cloud / 아이패드: 위 URL에서 **새로고침** (배포 1–3분 후)
5. 같은 세션에서 연 관련 PR이 남아 있으면 함께 머지한다 (충돌·실패만 보고)

예외: 사용자가 명시적으로 「머지하지 마」「draft만」이라고 한 경우에만 머지를 건너뛴다.
