"""설비재고 지사 필터 — [울산] regex 문자클래스 버그 회귀 방지."""
from __future__ import annotations

import os
import unittest

import pandas as pd

with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
    _src = f.read().split("# 5. 메인 실행 흐름")[0]
exec(_src, globals())


class EquipBranchFilterTest(unittest.TestCase):
    def test_ulsan_does_not_include_asan(self):
        df = pd.DataFrame(
            {
                "지사": ["[아산]", "[울산]", "[아산]", "[울산]"],
                "일련(제조)번호": ["A1", "U1", "A2", "U2"],
            }
        )
        out = filter_equipment_by_branch(df, "[울산]")
        self.assertEqual(list(out["지사"]), ["[울산]", "[울산]"])
        self.assertEqual(list(out["일련(제조)번호"]), ["U1", "U2"])

    def test_asan_only(self):
        df = pd.DataFrame({"지사": ["[아산]", "[울산]", "[아산]"]})
        out = filter_equipment_by_branch(df, "[아산]")
        self.assertTrue((out["지사"] == "[아산]").all())
        self.assertEqual(len(out), 2)

    def test_all_branches(self):
        df = pd.DataFrame({"지사": ["[아산]", "[울산]"]})
        out = filter_equipment_by_branch(df, "전체 지사")
        self.assertEqual(len(out), 2)

    def test_legacy_contains_regex_was_wrong(self):
        """과거 버그: str.contains('[울산]') 은 [아산]도 통과했다."""
        df = pd.DataFrame({"지사": ["[아산]", "[울산]"]})
        buggy = df[df["지사"].astype(str).str.contains("[울산]")]
        self.assertIn("[아산]", list(buggy["지사"]))
        fixed = filter_equipment_by_branch(df, "[울산]")
        self.assertNotIn("[아산]", list(fixed["지사"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
