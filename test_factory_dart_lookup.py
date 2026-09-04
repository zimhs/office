"""공장등록·DART 조회 — 일시 오류와 키 오류를 구분해야 한다."""
from __future__ import annotations

import ast
import unittest


def _load_err_helpers():
    with open("app.py", encoding="utf-8") as f:
        tree = ast.parse(f.read())
    wanted = ("_factory_err_is_fatal", "_factory_err_is_transient")
    nodes = [
        n for n in tree.body
        if isinstance(n, ast.FunctionDef) and n.name in wanted
    ]
    names = {n.name for n in nodes}
    if set(wanted) != names:
        raise AssertionError(f"missing helpers: {set(wanted) - names}")
    ns: dict = {}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
    return ns


class FactoryDartLookupTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.h = _load_err_helpers()

    def test_empty_body_is_transient_not_fatal(self):
        err = "EMPTY_BODY (해외IP 차단/타임아웃)"
        self.assertFalse(self.h["_factory_err_is_fatal"](err))
        self.assertTrue(self.h["_factory_err_is_transient"](err))

    def test_timeout_is_transient(self):
        self.assertTrue(self.h["_factory_err_is_transient"]("Read timed out"))
        self.assertFalse(self.h["_factory_err_is_fatal"]("Read timed out"))

    def test_bad_key_is_fatal(self):
        err = "등록되지 않은 서비스키"
        self.assertTrue(self.h["_factory_err_is_fatal"](err))
        self.assertFalse(self.h["_factory_err_is_transient"](err))


if __name__ == "__main__":
    unittest.main()
