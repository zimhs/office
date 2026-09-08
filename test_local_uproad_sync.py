"""로컬 dashboard/uproad → 캐시 동기화."""
import os
import tempfile
import unittest

from drive_autoload import (
    resolve_local_uproad_dir,
    sync_local_uproad_into_cache,
)


class ResolveLocalUproadDirTest(unittest.TestCase):
    def test_finds_trailing_space_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            spaced = os.path.join(tmp, "uproad ")
            os.makedirs(spaced)
            self.assertEqual(resolve_local_uproad_dir(tmp), spaced)

    def test_prefers_exact_uproad(self):
        with tempfile.TemporaryDirectory() as tmp:
            exact = os.path.join(tmp, "uproad")
            os.makedirs(exact)
            self.assertEqual(resolve_local_uproad_dir(tmp), exact)

    def test_missing_returns_none(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(resolve_local_uproad_dir(tmp))


class SyncLocalUproadIntoCacheTest(unittest.TestCase):
    def test_copies_debt_and_stripped_sales_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "uproad ")
            cache = os.path.join(tmp, "uploaded_cache")
            os.makedirs(src)
            with open(os.path.join(src, "채권.csv"), "w", encoding="utf-8") as f:
                f.write("거래처,구분,1월\nA,잔액,10\n")
            with open(os.path.join(src, " 202609.csv"), "w", encoding="utf-8") as f:
                f.write("x\n")
            res = sync_local_uproad_into_cache(cache, uproad_dir=src, include_worklog=False)
            self.assertTrue(res.get("ok"), res)
            self.assertIn("채권.csv", res.get("copied") or [])
            self.assertTrue(os.path.isfile(os.path.join(cache, "debt.csv")))
            self.assertTrue(os.path.isfile(os.path.join(cache, "sales", "202609.csv")))
            self.assertTrue(os.path.isfile(os.path.join(cache, ".debt_upload_stamp.json")))

    def test_missing_folder_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            missing = os.path.join(tmp, "nope")
            res = sync_local_uproad_into_cache(cache, uproad_dir=missing)
            self.assertFalse(res.get("ok"))
            self.assertTrue(res.get("skipped"))


class LocalSidebarButtonSourceTest(unittest.TestCase):
    def test_local_hides_drive_pull_button(self):
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            src = f.read()
        marker = 'Desktop/dashboard/uproad 의 CSV'
        self.assertIn(marker, src)
        self.assertIn("sync_local_uproad_into_cache", src)
        self.assertEqual(src.count('"☁️ Drive 복사본으로 동기화"'), 1)
        # 가져오기 버튼은 Cloud 전용 분기에만
        pull_idx = src.find("Drive 복사본에서 가져오기")
        self.assertGreater(pull_idx, 0)
        cloud_guard = src.rfind("else:", 0, pull_idx)
        self.assertGreater(cloud_guard, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
