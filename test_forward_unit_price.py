"""단가 이월: int64 월 열에 소수 단가를 넣어도 TypeError가 나면 안 된다."""
from __future__ import annotations

import os
import unittest

import numpy as np
import pandas as pd


def _load_apply_forward_unit_price():
    path = os.path.join(os.path.dirname(__file__), "app.py")
    with open(path, encoding="utf-8") as f:
        src = f.read()
    start = src.find("def apply_forward_unit_price")
    end = src.find("\ndef build_unit_price_change_pivot")
    ns = {"np": np, "pd": pd}
    exec(src[start:end], ns)
    return ns["apply_forward_unit_price"]


class ForwardUnitPriceTest(unittest.TestCase):
    def test_decimal_price_fills_int64_month_column(self):
        apply = _load_apply_forward_unit_price()
        price = pd.DataFrame(
            {"26년 1월": [278.55], "26년 2월": [0]},
            index=["거래처A"],
        )
        price["26년 2월"] = price["26년 2월"].astype("int64")
        qty = pd.DataFrame(
            {"26년 1월": [10.0], "26년 2월": [0.0]},
            index=["거래처A"],
        )
        out = apply(price, qty, ["2026"], ["1월", "2월"])
        self.assertAlmostEqual(float(out.at["거래처A", "26년 2월"]), 278.55)


if __name__ == "__main__":
    unittest.main()
