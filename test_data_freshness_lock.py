"""매출·채권 최신 잠금: 최신은 적용, 예전은 차단, git 되돌리기 복구."""
import os
import tempfile
import unittest

from data_freshness_lock import (
    copy_file_if_newer,
    debt_generation_from_bytes,
    restore_locked_latest,
    sales_generation_from_bytes,
    write_bytes_if_newer,
)
from drive_autoload import sync_local_uproad_into_cache


def _sales(end_day: int, extra: str = "") -> bytes:
    return (
        f",,,,,,,,,,,,,,거래처 매출내역,,,,,,,,,,,,\n"
        f",,,,,,,,,,,,,,일자 : 2026년 9월 1일 ~ 2026년 9월 {end_day}일,,,,,,,,,,,,\n"
        f"거래처,매출일,매출액\nA,2026-09-{end_day:02d},100{extra}\n"
    ).encode("utf-8-sig")


def _debt(last_month: int, amount: int) -> bytes:
    months = ",".join(f"{i}월" for i in range(1, 13))
    vals = ",".join(str(amount if i == last_month else 0) for i in range(1, 13))
    return f"거래처,구분,{months}\n테스트,잔액,{vals}\n".encode("utf-8-sig")


class GenerationParseTest(unittest.TestCase):
    def test_sales_header_end_date(self):
        g = sales_generation_from_bytes("202609.csv", _sales(11))
        self.assertEqual(g[0], 202609)
        self.assertEqual(g[1], 20260911)

    def test_sales_mmdd_uses_filename_year(self):
        raw = "거래처,매출일,매출액\nA,09/05,100\nB,09/16,200\n".encode("utf-8-sig")
        g = sales_generation_from_bytes("202609.csv", raw)
        self.assertEqual(g[0], 202609)
        self.assertEqual(g[1], 20260916)

    def test_sales_mmdd_reads_file_tail(self):
        pad = ("x" * 80000 + "\n").encode("utf-8")
        tail = "거래처,매출일,매출액\nA,09/16,200\n".encode("utf-8")
        g = sales_generation_from_bytes("202609.csv", pad + tail)
        self.assertEqual(g[1], 20260916)

    def test_later_day_is_newer(self):
        old = sales_generation_from_bytes("202609.csv", _sales(5))
        new = sales_generation_from_bytes("202609.csv", _sales(11))
        self.assertGreater(new, old)

    def test_debt_last_month(self):
        g4 = debt_generation_from_bytes(_debt(4, 10))
        g9 = debt_generation_from_bytes(_debt(9, 10))
        self.assertEqual(g4[0], 4)
        self.assertEqual(g9[0], 9)
        self.assertGreater(g9, g4)


class WriteIfNewerTest(unittest.TestCase):
    def test_newer_sales_applies_older_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            wrote, reason = write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(11), kind="sales", name="202609.csv"
            )
            self.assertTrue(wrote)
            self.assertEqual(reason, "wrote")
            wrote2, reason2 = write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(5), kind="sales", name="202609.csv"
            )
            self.assertFalse(wrote2)
            self.assertEqual(reason2, "blocked_older")
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertIn(b"9\xec\x9b\x94 11\xec\x9d\xbc", f.read())

    def test_even_newer_sales_updates_lock(self):
        with tempfile.TemporaryDirectory() as tmp:
            write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(11), kind="sales", name="202609.csv"
            )
            wrote, reason = write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(14, extra="9"), kind="sales", name="202609.csv"
            )
            self.assertTrue(wrote, reason)
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                raw = f.read()
            self.assertGreater(
                sales_generation_from_bytes("202609.csv", raw)[1],
                20260911,
            )

    def test_newer_debt_applies_older_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertTrue(
                write_bytes_if_newer(tmp, "debt.csv", _debt(9, 100), kind="debt")[0]
            )
            blocked = write_bytes_if_newer(tmp, "debt.csv", _debt(4, 1), kind="debt")
            self.assertEqual(blocked[1], "blocked_older")
            with open(os.path.join(tmp, "debt.csv"), "rb") as f:
                self.assertEqual(debt_generation_from_bytes(f.read())[0], 9)

    def test_same_month_debt_lower_balance_applies(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertTrue(
                write_bytes_if_newer(tmp, "debt.csv", _debt(9, 500), kind="debt")[0]
            )
            wrote, reason = write_bytes_if_newer(tmp, "debt.csv", _debt(9, 80), kind="debt")
            self.assertTrue(wrote, reason)
            with open(os.path.join(tmp, "debt.csv"), "rb") as f:
                self.assertEqual(debt_generation_from_bytes(f.read())[1], 80)

    def test_same_day_sales_rewrite_applies(self):
        with tempfile.TemporaryDirectory() as tmp:
            write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(17), kind="sales", name="202609.csv"
            )
            wrote, reason = write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(17, extra="updated"), kind="sales", name="202609.csv"
            )
            self.assertTrue(wrote, reason)
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertIn(b"updated", f.read())

    def test_smaller_same_day_sales_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            bigger = _sales(16, extra="xxxxxxxxxx")
            smaller = _sales(16)
            self.assertGreater(len(bigger), len(smaller))
            self.assertTrue(
                write_bytes_if_newer(
                    tmp, "sales/202609.csv", bigger, kind="sales", name="202609.csv"
                )[0]
            )
            wrote, reason = write_bytes_if_newer(
                tmp, "sales/202609.csv", smaller, kind="sales", name="202609.csv"
            )
            self.assertFalse(wrote)
            self.assertEqual(reason, "blocked_older")
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(f.read(), bigger)

    def test_force_applies_smaller_same_day_sales(self):
        with tempfile.TemporaryDirectory() as tmp:
            bigger = _sales(16, extra="xxxxxxxxxx")
            smaller = _sales(16)
            self.assertTrue(
                write_bytes_if_newer(
                    tmp, "sales/202609.csv", bigger, kind="sales", name="202609.csv"
                )[0]
            )
            wrote, reason = write_bytes_if_newer(
                tmp,
                "sales/202609.csv",
                smaller,
                kind="sales",
                name="202609.csv",
                force=True,
            )
            self.assertTrue(wrote, reason)
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(f.read(), smaller)


class RestoreAfterGitRevertTest(unittest.TestCase):
    def test_restore_snapshot_over_older_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(11), kind="sales", name="202609.csv"
            )
            write_bytes_if_newer(tmp, "debt.csv", _debt(9, 50), kind="debt")
            with open(os.path.join(tmp, "sales", "202609.csv"), "wb") as f:
                f.write(_sales(3))
            with open(os.path.join(tmp, "debt.csv"), "wb") as f:
                f.write(_debt(2, 1))
            restored = restore_locked_latest(tmp)
            self.assertIn("sales/202609.csv", restored)
            self.assertIn("debt.csv", restored)
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(
                    sales_generation_from_bytes("202609.csv", f.read())[1],
                    20260911,
                )
            with open(os.path.join(tmp, "debt.csv"), "rb") as f:
                self.assertEqual(debt_generation_from_bytes(f.read())[0], 9)

    def test_restore_keeps_same_month_newer_bytes(self):
        with tempfile.TemporaryDirectory() as tmp:
            write_bytes_if_newer(
                tmp, "sales/202609.csv", _sales(17), kind="sales", name="202609.csv"
            )
            write_bytes_if_newer(tmp, "debt.csv", _debt(9, 200), kind="debt")
            with open(os.path.join(tmp, "sales", "202609.csv"), "wb") as f:
                f.write(_sales(17, extra="storage"))
            with open(os.path.join(tmp, "debt.csv"), "wb") as f:
                f.write(_debt(9, 40))
            restored = restore_locked_latest(tmp)
            self.assertNotIn("sales/202609.csv", restored)
            self.assertNotIn("debt.csv", restored)
            with open(os.path.join(tmp, "sales", "202609.csv"), "rb") as f:
                self.assertIn(b"storage", f.read())
            with open(os.path.join(tmp, "debt.csv"), "rb") as f:
                self.assertEqual(debt_generation_from_bytes(f.read())[1], 40)


class LocalUproadRespectsLockTest(unittest.TestCase):
    def test_older_uproad_does_not_replace_newer_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            os.makedirs(os.path.join(cache, "sales"), exist_ok=True)
            write_bytes_if_newer(
                cache, "sales/202609.csv", _sales(11), kind="sales", name="202609.csv"
            )
            write_bytes_if_newer(cache, "debt.csv", _debt(9, 80), kind="debt")
            src = os.path.join(tmp, "uproad")
            os.makedirs(src)
            with open(os.path.join(src, "202609.csv"), "wb") as f:
                f.write(_sales(4))
            with open(os.path.join(src, "채권.csv"), "wb") as f:
                f.write(_debt(3, 1))
            res = sync_local_uproad_into_cache(cache, uproad_dir=src, include_worklog=False)
            self.assertTrue(res.get("ok"), res)
            with open(os.path.join(cache, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(
                    sales_generation_from_bytes("202609.csv", f.read())[1],
                    20260911,
                )
            with open(os.path.join(cache, "debt.csv"), "rb") as f:
                self.assertEqual(debt_generation_from_bytes(f.read())[0], 9)

    def test_force_apply_replaces_locked_sales(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            os.makedirs(os.path.join(cache, "sales"), exist_ok=True)
            write_bytes_if_newer(
                cache, "sales/202609.csv", _sales(11), kind="sales", name="202609.csv"
            )
            src = os.path.join(tmp, "uproad")
            os.makedirs(src)
            with open(os.path.join(src, "202609.csv"), "wb") as f:
                f.write(_sales(4))
            res = sync_local_uproad_into_cache(
                cache, uproad_dir=src, include_worklog=False, force_apply=True
            )
            self.assertTrue(res.get("ok"), res)
            with open(os.path.join(cache, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(
                    sales_generation_from_bytes("202609.csv", f.read())[1],
                    20260904,
                )

    def test_newer_uproad_updates_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            os.makedirs(os.path.join(cache, "sales"), exist_ok=True)
            write_bytes_if_newer(
                cache, "sales/202609.csv", _sales(5), kind="sales", name="202609.csv"
            )
            src = os.path.join(tmp, "uproad")
            os.makedirs(src)
            with open(os.path.join(src, "202609.csv"), "wb") as f:
                f.write(_sales(12))
            res = sync_local_uproad_into_cache(cache, uproad_dir=src, include_worklog=False)
            self.assertTrue(res.get("ok"), res)
            self.assertTrue(
                any(str(x).endswith("202609.csv") for x in (res.get("copied") or []))
            )
            with open(os.path.join(cache, "sales", "202609.csv"), "rb") as f:
                self.assertEqual(
                    sales_generation_from_bytes("202609.csv", f.read())[1],
                    20260912,
                )


class CopyHelperTest(unittest.TestCase):
    def test_copy_if_newer_skips_older_src(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            os.makedirs(cache)
            write_bytes_if_newer(cache, "debt.csv", _debt(8, 20), kind="debt")
            src = os.path.join(tmp, "old.csv")
            with open(src, "wb") as f:
                f.write(_debt(1, 1))
            dst = os.path.join(cache, "debt.csv")
            wrote, reason = copy_file_if_newer(src, dst, cache, kind="debt", name="채권.csv")
            self.assertFalse(wrote)
            self.assertEqual(reason, "blocked_older")


if __name__ == "__main__":
    unittest.main(verbosity=2)
