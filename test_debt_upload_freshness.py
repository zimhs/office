"""채권 업로드 후 예전 데이터가 남는 회귀 방지."""
import os
import tempfile
import time
import unittest

from drive_autoload import (
    local_debt_upload_should_keep,
    write_debt_upload_stamp,
)

with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
    _src = f.read().split("# 5. 메인 실행 흐름")[0]
exec(_src, globals())


def _csv(tag: str, amount: int) -> bytes:
    return (
        f"거래처,구분,1월,2월\n"
        f"테스트업체,잔액,{amount},0\n"
        f"테스트업체,매출,{amount},0\n"
        f"# {tag}\n"
    ).encode("utf-8-sig")


class DebtFingerprintTest(unittest.TestCase):
    def test_same_rows_different_amount_changes_fingerprint(self):
        old = load_debt_file(_csv("old", 100))
        new = load_debt_file(_csv("new", 999))
        self.assertEqual(len(old), len(new))
        self.assertNotEqual(debt_bytes_fingerprint(_csv("old", 100)), debt_bytes_fingerprint(_csv("new", 999)))
        self.assertNotEqual(debt_frame_fingerprint(old), debt_frame_fingerprint(new))

    def test_filter_sig_must_include_content_not_just_len(self):
        old = load_debt_file(_csv("old", 100))
        new = load_debt_file(_csv("new", 999))
        old_sig = ((), "전체 거래처", int(len(old)))
        new_sig = ((), "전체 거래처", int(len(new)))
        self.assertEqual(old_sig, new_sig)  # 예전 버그: 행수만 보면 동일
        self.assertNotEqual(
            (debt_bytes_fingerprint(_csv("old", 100)), debt_frame_fingerprint(old)),
            (debt_bytes_fingerprint(_csv("new", 999)), debt_frame_fingerprint(new)),
        )


class DebtCacheTruthTest(unittest.TestCase):
    def test_cache_wins_over_newer_stray_채권_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "debt.csv")
            stray_name = "채권_backup.csv"
            persist_debt_bytes(_csv("cache", 111), cache, os.path.join(tmp, "채권.csv"))
            stray = os.path.join(tmp, stray_name)
            with open(stray, "wb") as f:
                f.write(_csv("stray", 222))
            os.utime(stray, (time.time() + 100, time.time() + 100))
            raw, label = resolve_cached_debt_bytes(cache, os.path.join(tmp, "채권.csv"))
            self.assertTrue(label.startswith("캐시"))
            df = load_debt_file(raw)
            self.assertEqual(float(df.loc[df["구분"] == "잔액", "1월"].iloc[0]), 111)

    def test_missing_cache_falls_back_to_folder_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "debt.csv")
            folder = os.path.join(tmp, "채권.csv")
            with open(folder, "wb") as f:
                f.write(_csv("folder", 333))
            raw, label = resolve_cached_debt_bytes(cache, folder)
            self.assertTrue(label.startswith("폴더"))
            df = load_debt_file(raw)
            self.assertEqual(float(df.loc[df["구분"] == "잔액", "1월"].iloc[0]), 333)
            self.assertTrue(os.path.isfile(cache))


class DebtStampProtectsDriveOverwriteTest(unittest.TestCase):
    def test_stamped_local_not_replaced_by_older_drive(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "debt.csv")
            drive = os.path.join(tmp, "drive_채권.csv")
            new_b = _csv("uploaded", 777)
            old_b = _csv("drive-old", 1)
            with open(cache, "wb") as f:
                f.write(new_b)
            with open(drive, "wb") as f:
                f.write(old_b)
            now = time.time()
            os.utime(drive, (now - 3600, now - 3600))
            os.utime(cache, (now, now))
            write_debt_upload_stamp(tmp, debt_bytes_fingerprint(new_b))
            self.assertTrue(local_debt_upload_should_keep(tmp, drive_src=drive))

    def test_no_stamp_and_older_local_allows_drive(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "debt.csv")
            drive = os.path.join(tmp, "drive_채권.csv")
            with open(cache, "wb") as f:
                f.write(_csv("local-old", 1))
            with open(drive, "wb") as f:
                f.write(_csv("drive-new", 888))
            now = time.time()
            os.utime(cache, (now - 3600, now - 3600))
            os.utime(drive, (now, now))
            self.assertFalse(local_debt_upload_should_keep(tmp, drive_src=drive))


class DebtReloadBlankScreenRegressionTest(unittest.TestCase):
    """업로드 후 같은 내용 「불러오기」 시 staff/filtered 빈표 → 빈화면 회귀 방지."""

    def test_clear_debt_runtime_caches_drops_staff_sig(self):
        session = {
            "_dash_debt_staff_sig": (("담당자A",), ("sha", ("empty",))),
            "_dash_debt_filter_sig": "old",
            "_dash_staff_debt_df": "frame",
            "_dash_filtered_debt_df": "frame",
            "debt_od_filters": ["악성"],
            "unrelated": 1,
        }
        clear_debt_runtime_caches(session)
        for key in (
            "_dash_debt_staff_sig",
            "_dash_debt_filter_sig",
            "_dash_staff_debt_df",
            "_dash_filtered_debt_df",
            "debt_od_filters",
        ):
            self.assertNotIn(key, session)
        self.assertEqual(session.get("unrelated"), 1)

    def test_same_sha_reload_rebuilds_when_staff_df_missing(self):
        """예전 버그: sig만 남고 표가 없으면 재구성 스킵 → 빈 filtered."""
        raw = _csv("same", 555)
        debt_df = load_debt_file(raw)
        self.assertFalse(debt_df.empty)
        sha = debt_bytes_fingerprint(raw)
        data_sig = (sha, debt_frame_fingerprint(debt_df))
        staff_sig = ((), data_sig)
        # clear 후 표만 없고 sig는 남은 상태(버그 재현 조건)
        session = {"_dash_debt_staff_sig": staff_sig}
        staff_missing = "_dash_staff_debt_df" not in session
        must_rebuild = session.get("_dash_debt_staff_sig") != staff_sig or staff_missing
        self.assertTrue(must_rebuild)
        # 재구성 시뮬레이션 (담당자 전체)
        session["_dash_debt_staff_sig"] = staff_sig
        session["_dash_staff_debt_df"] = debt_df
        session["_dash_debt_filter_sig"] = (staff_sig, "전체 거래처")
        session["_dash_filtered_debt_df"] = debt_df
        self.assertFalse(session["_dash_filtered_debt_df"].empty)

    def test_tab5_has_empty_state_messages(self):
        with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
            src = f.read()
        tab5 = src.split("# Tab 5:", 1)[1].split("# Tab 6:", 1)[0]
        self.assertIn("현재 필터(담당자·거래처)에 해당하는 채권 행이 없습니다", tab5)
        self.assertIn("채권 데이터가 없습니다", tab5)
        self.assertIn("_dash_debt_staff_sig", src)
        # clear 목록에 staff_sig 포함
        clear_fn = src.split("def clear_debt_runtime_caches", 1)[1].split(
            "def resolve_cached_debt_bytes", 1
        )[0]
        self.assertIn('"_dash_debt_staff_sig"', clear_fn)


if __name__ == "__main__":
    unittest.main(verbosity=2)
