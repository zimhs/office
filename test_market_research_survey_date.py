"""시장조사 — 새 입력 조사일이 전체 조회(목록·검색·기간)에 남는지."""
from __future__ import annotations

import unittest
from datetime import date
from unittest.mock import patch

import pandas as pd


class MarketResearchSurveyDateTest(unittest.TestCase):
    def setUp(self):
        import market_research_tab as mr

        self.mr = mr

    def test_normalize_and_entry_fallback(self):
        self.assertEqual(self.mr._normalize_survey_date(date(2026, 9, 16)), "2026-09-16")
        self.assertEqual(self.mr._normalize_survey_date("2026/9/3"), "2026-09-03")
        self.assertEqual(self.mr._normalize_survey_date("2026-09-16 08:21:00"), "2026-09-16")
        self.assertEqual(self.mr._normalize_survey_date(""), "")
        self.assertEqual(
            self.mr._entry_survey_date({"saved_at": "2026-08-01 10:00:00"}),
            "2026-08-01",
        )
        self.assertEqual(
            self.mr._entry_survey_date({"조사일": "2026-09-16", "saved_at": "2026-08-01"}),
            "2026-09-16",
        )

    def test_add_manual_entry_stores_survey_date(self):
        saved: list[list] = []

        def _save(entries):
            saved.append(list(entries))

        with patch.object(self.mr, "load_manual_entries", return_value=[]):
            with patch.object(self.mr, "save_manual_entries", side_effect=_save):
                ent = self.mr.add_manual_entry(
                    {"업체명": "테스트가스", "조사일": date(2026, 9, 16)}
                )
        self.assertEqual(ent["조사일"], "2026-09-16")
        rec = self.mr._manual_to_record(ent)
        self.assertEqual(rec["조사일"], "2026-09-16")
        self.assertEqual(saved[0][0]["조사일"], "2026-09-16")

    def test_merge_keeps_latest_survey_date(self):
        df = pd.DataFrame(
            [
                {"업체명": "갑산업", "지역": "화성", "산업단지": "미분류", "주소": "", "업종": "", "사용가스": "", "공급사": "", "담당자": "", "연락처": "", "비고": "", "출처": "직접입력", "파일": "a", "시트": "직접입력", "조사일": "2026-08-01"},
                {"업체명": "갑산업", "지역": "화성", "산업단지": "미분류", "주소": "", "업종": "", "사용가스": "", "공급사": "", "담당자": "", "연락처": "", "비고": "", "출처": "직접입력", "파일": "b", "시트": "직접입력", "조사일": "2026-09-16"},
            ]
        )
        merged, removed = self.mr.merge_duplicate_rows(df)
        self.assertEqual(removed, 1)
        self.assertEqual(merged.iloc[0]["조사일"], "2026-09-16")

    def test_filter_and_search_use_survey_date(self):
        df = pd.DataFrame(
            {
                "조사일": ["2026-09-16", "2026-08-01", ""],
                "지역": ["화성", "평택", "오산"],
                "산업단지": ["미분류", "미분류", "미분류"],
                "공급사": ["", "", ""],
                "_factory_only": [False, False, False],
                "_search": ["테스트가스 2026-09-16", "을산업 2026-08-01", "병산업"],
            }
        )
        only_sep = self.mr._filter_frame(
            df, regions=(), complexes=(), suppliers=(), query="",
            include_factory=False, date_from="2026-09-01", date_to="2026-09-30",
        )
        self.assertEqual(list(only_sep["조사일"]), ["2026-09-16"])
        by_q = self.mr._filter_frame(
            df, regions=(), complexes=(), suppliers=(), query="2026-09-16",
            include_factory=False,
        )
        self.assertEqual(list(by_q["조사일"]), ["2026-09-16"])

    def test_ui_creates_survey_date_and_global_query(self):
        with open(self.mr.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn('survey_date_in = st.date_input(', src)
        self.assertIn('"조사일": survey_date_in', src)
        self.assertIn('key="mr_w_date_from"', src)
        self.assertIn('key="mr_w_date_to"', src)
        self.assertIn('"조사일",', src)
        self.assertIn("date_from=app_d0", src)
        self.assertIn("검색 (업체·주소·단지·가스·비고·조사일)", src)


if __name__ == "__main__":
    unittest.main()
