"""담당자별 월 매출 — 전체매출비중 / 당해매출비중 컬럼 회귀 테스트."""
from __future__ import annotations

import os
import unittest

import pandas as pd

with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
    _src = f.read().split("# 5. 메인 실행 흐름")[0]
exec(_src, globals())


class StaffYearShareTest(unittest.TestCase):
    def _sample(self):
        # A: 2025 heavy, 2026 light / B: 2025 light, 2026 heavy
        return pd.DataFrame(
            {
                "담당자": ["A", "A", "B", "B"],
                "연도": ["2025", "2026", "2025", "2026"],
                "월": [1, 1, 1, 1],
                "연도월_정렬": ["25년 1", "26년 1", "25년 1", "26년 1"],
                "매출액": [3_000_000, 1_000_000, 1_000_000, 3_000_000],
            }
        )

    def test_column_names_and_order(self):
        df = self._sample()
        desired = ["26년 1", "25년 1"]
        out = cached_staff_pivot.__wrapped__(df, desired)
        cols = list(out.columns)
        self.assertEqual(cols[0], "전체매출비중")
        self.assertEqual(cols[1], "당해매출비중")
        self.assertEqual(cols[2], "총 매출 합계 (만원)")
        self.assertNotIn("매출 비중 (%)", cols)

    def test_current_year_share_differs_from_all(self):
        df = self._sample()
        out = cached_staff_pivot.__wrapped__(df, ["26년 1", "25년 1"])
        # 전체: A=B=50% (각 400만원 환산 전 동일 합)
        self.assertAlmostEqual(float(out.loc["A", "전체매출비중"]), 50.0, places=4)
        self.assertAlmostEqual(float(out.loc["B", "전체매출비중"]), 50.0, places=4)
        # 당해(2026): A=25%, B=75%
        self.assertAlmostEqual(float(out.loc["A", "당해매출비중"]), 25.0, places=4)
        self.assertAlmostEqual(float(out.loc["B", "당해매출비중"]), 75.0, places=4)
        self.assertAlmostEqual(float(out["당해매출비중"].sum()), 100.0, places=4)
        self.assertAlmostEqual(float(out["전체매출비중"].sum()), 100.0, places=4)


if __name__ == "__main__":
    unittest.main(verbosity=2)
