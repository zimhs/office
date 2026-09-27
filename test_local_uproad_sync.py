"""로컬 dashboard/uproad → 캐시 동기화."""
import os
import tempfile
import unittest

from unittest.mock import patch

from drive_autoload import (
    _prefer_newer_source,
    load_sidebar_slot_from_connected_path,
    resolve_local_uproad_dir,
    sync_cache_to_drive_copy,
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


class SidebarSlotLoadTest(unittest.TestCase):
    def test_loads_address_and_sales_from_connected_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "uproad")
            cache = os.path.join(tmp, "cache")
            os.makedirs(src)
            with open(os.path.join(src, "주소.csv"), "w", encoding="utf-8") as f:
                f.write("거래처,주소\nA,서울\n")
            with open(os.path.join(src, "202609.csv"), "w", encoding="utf-8") as f:
                f.write("x\n")
            with patch("drive_autoload._connected_data_roots", return_value=[src]):
                addr = load_sidebar_slot_from_connected_path("address", cache)
                sales = load_sidebar_slot_from_connected_path("sales", cache)
            self.assertTrue(addr.get("ok"), addr)
            self.assertTrue(os.path.isfile(os.path.join(cache, "address.csv")))
            self.assertTrue(sales.get("ok"), sales)
            self.assertTrue(os.path.isfile(os.path.join(cache, "sales", "202609.csv")))

    def test_sales_prefers_larger_same_month_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            local = os.path.join(tmp, "uproad")
            drive = os.path.join(tmp, "drive")
            cache = os.path.join(tmp, "cache")
            os.makedirs(local)
            os.makedirs(drive)
            older = "거래처,매출일,매출액\nA,09/16,100\n".encode("utf-8-sig")
            newer = older + b"B,09/16,200\n"
            with open(os.path.join(drive, "202609.csv"), "wb") as f:
                f.write(older)
            with open(os.path.join(local, "202609.csv"), "wb") as f:
                f.write(newer)
            with patch("drive_autoload._connected_data_roots", return_value=[drive, local]):
                sales = load_sidebar_slot_from_connected_path("sales", cache)
            self.assertTrue(sales.get("ok"), sales)
            with open(os.path.join(cache, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(f.read(), newer)

    def test_prefer_newer_source_picks_larger_sales(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "a.csv")
            b = os.path.join(tmp, "b.csv")
            older = "거래처,매출일,매출액\nA,09/16,100\n".encode("utf-8-sig")
            newer = older + b"B,09/16,200\n"
            with open(a, "wb") as f:
                f.write(older)
            with open(b, "wb") as f:
                f.write(newer)
            self.assertEqual(
                _prefer_newer_source(a, b, kind="sales", name="202609.csv"),
                b,
            )

    def test_missing_file_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "uproad")
            cache = os.path.join(tmp, "cache")
            os.makedirs(src)
            with patch("drive_autoload._connected_data_roots", return_value=[src]):
                res = load_sidebar_slot_from_connected_path("debt", cache)
            self.assertFalse(res.get("ok"))
            self.assertIn("채권.csv", res.get("error") or "")


class SyncCacheToDriveTabDataTest(unittest.TestCase):
    def test_pushes_mail_visit_market_research(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            drive = os.path.join(tmp, "drive")
            os.makedirs(os.path.join(cache, "price_increase"))
            os.makedirs(os.path.join(cache, "visit_calendar"))
            os.makedirs(os.path.join(cache, "market_research", "uploads"))
            os.makedirs(drive)
            with open(os.path.join(cache, "price_increase", "mail_contacts.csv"), "w", encoding="utf-8") as f:
                f.write("거래처,이메일\nA,a@x.com\n")
            with open(os.path.join(cache, "visit_calendar", "store.json"), "w", encoding="utf-8") as f:
                f.write('{"visits":[]}\n')
            with open(os.path.join(cache, "market_research", "manual_entries.json"), "w", encoding="utf-8") as f:
                f.write("[]\n")
            with open(os.path.join(cache, "market_research", "uploads", "조사.xlsx"), "w", encoding="utf-8") as f:
                f.write("x\n")
            with patch("drive_autoload.resolve_drive_dashboard_copy", return_value=drive):
                res = sync_cache_to_drive_copy(cache, force=True)
            self.assertTrue(res.get("ok"), res)
            self.assertTrue(os.path.isfile(os.path.join(drive, "메일연락처.csv")))
            self.assertTrue(os.path.isfile(os.path.join(drive, "방문할일.json")))
            self.assertTrue(os.path.isfile(os.path.join(drive, "시장조사_직접입력.json")))
            self.assertTrue(os.path.isfile(os.path.join(drive, "시장조사업로드", "조사.xlsx")))


class LocalSidebarButtonSourceTest(unittest.TestCase):
    def test_local_hides_drive_pull_button(self):
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            src = f.read()
        marker = "로컬 데이터는 Drive로, 코드는 Cloud(office/main)로 같이 올립니다."
        self.assertIn(marker, src)
        self.assertIn("sync_local_uproad_into_cache", src)
        self.assertIn("_sidebar_slot_load_button", src)
        self.assertIn('file_uploader("거래처 주소록 (CSV)"', src)
        self.assertIn('불러오기', src)
        self.assertEqual(src.count('"☁️ Drive 복사본으로 동기화"'), 1)
        local_btn = src.split('"☁️ Drive 복사본으로 동기화"', 1)[1].split(
            "Drive 복사본에서 가져오기", 1
        )[0]
        self.assertIn("push_office_code", local_btn)
        self.assertIn("code_push", src)
        # 가져오기 버튼은 Cloud 전용 분기에만
        pull_idx = src.find("Drive 복사본에서 가져오기")
        self.assertGreater(pull_idx, 0)
        cloud_guard = src.rfind("else:", 0, pull_idx)
        self.assertGreater(cloud_guard, 0)


class OfficeCodePushAllowlistTest(unittest.TestCase):
    def test_code_list_skips_cache_secrets_and_unit_price(self):
        from office_code_push import list_office_code_files, repo_root

        names = list_office_code_files(repo_root())
        joined = "\n".join(names)
        self.assertIn("app.py", names)
        self.assertIn("office_code_push.py", names)
        self.assertIn("drive_autoload.py", names)
        self.assertNotIn("visit_calendar_tab.py", names)
        self.assertNotIn("test_visit_calendar.py", names)
        self.assertNotIn("uploaded_cache", joined)
        self.assertNotIn("secrets.toml", joined)
        self.assertNotIn("client_extra_sites.json", joined)


if __name__ == "__main__":
    unittest.main(verbosity=2)
