"""클라우드 부팅 Drive 자동로드 신선도 검증.

Streamlit Cloud는 리붓마다 시드 파일을 새로 체크아웃해 로컬 mtime이 항상 "방금"이
된다. 이 경우 mtime 기반 비교(force_refresh=False)는 Drive의 실제 최신본을 "오래된
것"으로 오판해 반영하지 못한다(예전 데이터 잔존). 부팅 자동로드가 force_refresh=True
로 Drive를 강제 반영해야 최신본이 자동으로 로드된다는 것을 검증한다.
"""
import os
import tempfile
import unittest
from unittest import mock

import drive_remote_fetch as drf

_STALE = b"OLD,data\n1,old\n"
_LATEST = b"NEW,data\n1,latest\n"

# Drive 쪽 modifiedTime을 과거로 둬서, 방금 체크아웃된 로컬 시드(mtime=now)보다
# "오래된 것"처럼 보이게 만든다 (실제 클라우드 리붓 상황 재현).
_DRIVE_META = {
    "id": "file-addr",
    "name": "주소.csv",
    "mimeType": "text/csv",
    "modifiedTime": "2020-01-01T00:00:00Z",
}


class DriveBootAutoloadFreshnessTest(unittest.TestCase):
    def _make_cache_with_stale_seed(self, tmp: str) -> str:
        os.makedirs(tmp, exist_ok=True)
        dst = os.path.join(tmp, "address.csv")
        with open(dst, "wb") as f:
            f.write(_STALE)
        # 방금 체크아웃된 시드처럼 mtime을 현재로 (기본값이 곧 now)
        return dst

    def _run_sync(self, tmp: str, *, force_refresh: bool):
        with mock.patch.object(drf, "resolve_drive_uproad_folder_id", return_value="folder-123"), \
             mock.patch.object(drf, "drive_remote_configured", return_value=True), \
             mock.patch.object(drf, "_list_children", return_value=([_DRIVE_META], None)), \
             mock.patch.object(drf, "_drive_download", return_value=(_LATEST, None)):
            return drf.sync_drive_copy_from_remote(tmp, force_refresh=force_refresh)

    def test_stale_seed_not_updated_without_force(self):
        """버그 재현: mtime이 최신인 시드는 force_refresh=False로는 갱신 안 됨."""
        with tempfile.TemporaryDirectory() as tmp:
            dst = self._make_cache_with_stale_seed(tmp)
            res = self._run_sync(tmp, force_refresh=False)
            self.assertTrue(res.get("ok"))
            with open(dst, "rb") as f:
                self.assertEqual(f.read(), _STALE, "force_refresh=False 는 예전 데이터를 그대로 둔다(버그)")

    def test_force_refresh_pulls_latest(self):
        """수정 검증: force_refresh=True 는 Drive 최신본을 강제 반영."""
        with tempfile.TemporaryDirectory() as tmp:
            dst = self._make_cache_with_stale_seed(tmp)
            res = self._run_sync(tmp, force_refresh=True)
            self.assertTrue(res.get("ok"))
            self.assertIn("주소.csv", res.get("copied") or [])
            with open(dst, "rb") as f:
                self.assertEqual(f.read(), _LATEST, "force_refresh=True 는 Drive 최신본으로 갱신한다")


class BootPathUsesForceRefreshTest(unittest.TestCase):
    """app.py 의 클라우드 부팅 자동로드가 force_refresh=True 를 쓰는지 소스 확인."""

    def test_followup_fragment_uses_force_refresh_true(self):
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            src = f.read()
        # 부팅 지연 동기화 블록에 force_refresh=False 가 남아 있으면 안 됨
        marker = "def _dash_consume_deferred_cloud_drive_sync"
        self.assertIn(marker, src)
        seg = src.split(marker, 1)[1].split("def ", 1)[0]
        # 주석(설명문)에 등장하는 문자열은 제외하고 실제 코드 라인만 검사
        code_lines = [ln for ln in seg.splitlines() if not ln.lstrip().startswith("#")]
        code = "\n".join(code_lines)
        self.assertIn("force_refresh=True", code)
        self.assertNotIn("force_refresh=False", code)


if __name__ == "__main__":
    unittest.main(verbosity=2)
