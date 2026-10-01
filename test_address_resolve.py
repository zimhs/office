"""주소록 로드·거래처 주소 매칭 (ERP 다중열·가스코아산(대창))."""
from __future__ import annotations

import os
import unittest

with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
    _src = f.read().split("# 5. 메인 실행 흐름")[0]
exec(_src, globals())


def _simple_csv() -> bytes:
    return (
        "거래처,주소,전화번호\n"
        "가스코아산(대창),경기도 시흥시 공단1대로 391,010-9077-2486\n"
        "가스코아산(몰드스틸),경기도 화성시 팔탄면 밤뒤길 92-72,\n"
    ).encode("utf-8-sig")


def _erp_csv() -> bytes:
    """원본 ERP 사업체 목록 형식 — 앞열이 코드·사업체명, 주소는 뒤쪽."""
    return (
        "사업체코드,사업체명,납품지역,매입/매출공유,영업담당자,사업체구분,"
        "거래처구분,상호,대표자명,사업자등록번호,업태,업종,우편번호,주소\n"
        "00540,가스코아산(대창),,[아산] 가스코아산,담당자없음,거래처 본사,"
        "거래업체,(주)대창,,,,,,경기도 시흥시 공단1대로 391\n"
        "00528,가스코아산(몰드스틸),,[아산] 가스코아산,담당자없음,거래처 본사,"
        "거래업체,몰드스틸,,,,,,경기도 화성시 팔탄면 밤뒤길 92-72\n"
    ).encode("utf-8-sig")


class AddressBookColumnPickTest(unittest.TestCase):
    def test_simple_two_col_book(self):
        d = load_address_file(_simple_csv())
        self.assertEqual(d.get("가스코아산(대창)"), "경기도 시흥시 공단1대로 391")

    def test_erp_export_uses_사업체명_and_주소(self):
        """예전 버그: 0·1열만 써서 00540→가스코아산(대창) 으로 매핑됨."""
        d = load_address_file(_erp_csv())
        self.assertNotIn("00540", d)
        self.assertNotIn("540", d)
        self.assertEqual(d.get("가스코아산(대창)"), "경기도 시흥시 공단1대로 391")
        self.assertEqual(
            resolve_client_address("가스코아산(대창)", d),
            "경기도 시흥시 공단1대로 391",
        )


class ResolveClientAddressTest(unittest.TestCase):
    def test_exact_and_nfc_parens(self):
        d = load_address_file(_simple_csv())
        self.assertEqual(
            resolve_client_address("가스코아산(대창)", d),
            "경기도 시흥시 공단1대로 391",
        )
        # 전각 괄호
        self.assertEqual(
            resolve_client_address("가스코아산（대창）", d),
            "경기도 시흥시 공단1대로 391",
        )

    def test_workspace_cache_has_대창(self):
        path = os.path.join(os.path.dirname(__file__), "uploaded_cache", "address.csv")
        if not os.path.isfile(path):
            self.skipTest("uploaded_cache/address.csv 없음")
        with open(path, "rb") as f:
            d = load_address_file(f.read())
        self.assertEqual(
            resolve_client_address("가스코아산(대창)", d),
            "경기도 시흥시 공단1대로 391",
        )


class MergeAddressSourcesTest(unittest.TestCase):
    def test_merge_fills_missing_from_second_source(self):
        a = {"가스코아산(몰드스틸)": "화성시"}
        b = {
            "가스코아산(대창)": "경기도 시흥시 공단1대로 391",
            "가스코아산(몰드스틸)": "옛주소",
        }
        m = merge_address_dicts(a, b)
        self.assertEqual(m["가스코아산(몰드스틸)"], "화성시")  # 앞 출처 우선
        self.assertEqual(m["가스코아산(대창)"], "경기도 시흥시 공단1대로 391")

    def test_load_merged_reads_uproad_gap(self):
        import tempfile
        from unittest.mock import patch

        with tempfile.TemporaryDirectory() as tmp:
            cache = os.path.join(tmp, "cache")
            uproad = os.path.join(tmp, "uproad")
            os.makedirs(cache)
            os.makedirs(uproad)
            # 캐시에는 대창 없음
            cache_csv = (
                "거래처,주소\n가스코아산(몰드스틸),화성시\n"
            ).encode("utf-8-sig")
            with open(os.path.join(cache, "address.csv"), "wb") as f:
                f.write(cache_csv)
            with open(os.path.join(uproad, "주소.csv"), "wb") as f:
                f.write(_simple_csv())
            with patch(
                "drive_autoload.resolve_local_uproad_dir", return_value=uproad
            ):
                d = load_merged_address_dict(cache_csv, cache)
            self.assertEqual(
                resolve_client_address("가스코아산(대창)", d),
                "경기도 시흥시 공단1대로 391",
            )

    def test_tab2_hint_and_clear_helpers_exist(self):
        with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
            src = f.read()
        self.assertIn("def clear_address_runtime_caches", src)
        self.assertIn("def load_merged_address_dict", src)
        self.assertIn('slot == "address"', src)
        self.assertIn("주소록", src)
        self.assertIn("사이드바 주소.csv를 업로드하거나", src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
