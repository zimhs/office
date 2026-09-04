"""클라우드 부팅 Drive 자동로드 신선도/성능/레이아웃 회귀 검증.

Streamlit Cloud는 리붓마다 시드 파일을 새로 체크아웃해 로컬 mtime이 항상 "방금"이
된다. 그래서 mtime 비교로는 Drive 최신본이 "오래된 것"으로 오판돼 반영되지 않았다
(예전 데이터 잔존). 이를 mtime 대신 md5(내용) 비교로 바꿔서:
  1) force_refresh 없이도 변경된 파일은 최신본으로 받아오고,
  2) 내용이 같으면 다운로드/재렌더를 건너뛰어(성능) 부팅 시 고정바 강제 재주입도 없앤다.
"""
import hashlib
import os
import tempfile
import unittest
from unittest import mock

import drive_remote_fetch as drf

_STALE = b"OLD,data\n1,old\n"
_LATEST = b"NEW,data\n1,latest\n"


def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _meta(content: bytes) -> dict:
    # 실제 Drive 메타처럼 md5Checksum/size 포함. modifiedTime은 과거(mtime 오판 유발용).
    return {
        "id": "file-addr",
        "name": "주소.csv",
        "mimeType": "text/csv",
        "modifiedTime": "2020-01-01T00:00:00Z",
        "md5Checksum": _md5(content),
        "size": str(len(content)),
    }


class DriveBootMd5FreshnessTest(unittest.TestCase):
    def _seed(self, tmp: str, content: bytes) -> str:
        os.makedirs(tmp, exist_ok=True)
        dst = os.path.join(tmp, "address.csv")
        with open(dst, "wb") as f:
            f.write(content)  # 방금 체크아웃된 시드처럼 mtime=now
        return dst

    def _run(self, tmp: str, *, drive_content: bytes, force_refresh: bool):
        dl = mock.Mock(return_value=(drive_content, None))
        with mock.patch.object(drf, "resolve_drive_uproad_folder_id", return_value="folder-123"), \
             mock.patch.object(drf, "drive_remote_configured", return_value=True), \
             mock.patch.object(drf, "_list_children", return_value=([_meta(drive_content)], None)), \
             mock.patch.object(drf, "_drive_download", dl):
            res = drf.sync_drive_copy_from_remote(tmp, force_refresh=force_refresh)
        return res, dl

    def test_stale_seed_updated_by_md5_without_force(self):
        """mtime이 최신인 시드라도, md5가 다르면 force 없이 최신본을 받아온다(예전 데이터 수정)."""
        with tempfile.TemporaryDirectory() as tmp:
            dst = self._seed(tmp, _STALE)
            res, dl = self._run(tmp, drive_content=_LATEST, force_refresh=False)
            self.assertTrue(res.get("ok"))
            self.assertIn("주소.csv", res.get("copied") or [])
            dl.assert_called()  # 실제 다운로드 발생
            with open(dst, "rb") as f:
                self.assertEqual(f.read(), _LATEST)

    def test_same_content_skips_download(self):
        """내용(md5)이 같으면 다운로드/갱신을 건너뛴다(부팅 가벼움 · 불필요 rerun 방지)."""
        with tempfile.TemporaryDirectory() as tmp:
            dst = self._seed(tmp, _LATEST)
            res, dl = self._run(tmp, drive_content=_LATEST, force_refresh=False)
            self.assertTrue(res.get("ok"))
            self.assertNotIn("주소.csv", res.get("copied") or [])
            dl.assert_not_called()  # 동일 → 다운로드 안 함
            with open(dst, "rb") as f:
                self.assertEqual(f.read(), _LATEST)


class BootPathRegressionTest(unittest.TestCase):
    """app.py 클라우드 부팅 자동로드가 (a) force_refresh=False (md5 위임) 이고
    (b) 고정바를 강제 재주입하지 않는지 소스로 확인 (레이아웃/성능 회귀 방지)."""

    def _segment(self) -> str:
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            src = f.read()
        marker = "def _dash_consume_deferred_cloud_drive_sync"
        self.assertIn(marker, src)
        seg = src.split(marker, 1)[1].split("\ndef ", 1)[0]
        # 주석 라인 제외 (설명문에 등장하는 키워드가 검사에 걸리지 않도록)
        return "\n".join(ln for ln in seg.splitlines() if not ln.lstrip().startswith("#"))

    def test_boot_uses_force_refresh_false(self):
        code = self._segment()
        self.assertIn("force_refresh=False", code)
        self.assertNotIn("force_refresh=True", code)

    def test_boot_does_not_force_sticky_reinject(self):
        code = self._segment()
        self.assertNotIn("_dash_after_drive_boot", code)
        self.assertNotIn("_dash_sticky_inject_ver", code)


if __name__ == "__main__":
    unittest.main(verbosity=2)
