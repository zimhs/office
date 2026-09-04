"""모든 탭은 시작부터 펼침 — 화면 불러오기 stub 로 접히지 않아야 한다."""
from __future__ import annotations

import ast
import unittest


_DEFER_FNS = (
    "_dash_should_defer_light_tab",
    "_dash_should_defer_heavy_tab",
)


class DashTabExpandTest(unittest.TestCase):
    def test_defer_helpers_always_return_false(self):
        with open("app.py", encoding="utf-8") as f:
            tree = ast.parse(f.read())
        found: set[str] = set()
        for node in tree.body:
            if not isinstance(node, ast.FunctionDef) or node.name not in _DEFER_FNS:
                continue
            found.add(node.name)
            returns = [n for n in ast.walk(node) if isinstance(n, ast.Return)]
            self.assertTrue(returns, msg=f"{node.name} has no return")
            for ret in returns:
                self.assertIsInstance(ret.value, ast.Constant, msg=f"{node.name} return is not constant")
                self.assertIs(ret.value.value, False, msg=f"{node.name} must always return False")
        self.assertEqual(found, set(_DEFER_FNS))


if __name__ == "__main__":
    unittest.main()
