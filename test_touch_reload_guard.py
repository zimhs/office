import re
import unittest
from pathlib import Path

SRC = Path(__file__).with_name("app.py").read_text(encoding="utf-8")


def _func(name: str) -> str:
    start = SRC.index(f"def {name}(")
    nxt = re.search(r"^def ", SRC[start + 4 :], re.M)
    return SRC[start : start + 4 + nxt.start()] if nxt else SRC[start:]


class TouchReloadGuardTest(unittest.TestCase):
    def test_reload_skipped_when_server_already_touch(self):
        for name in ("inject_top30_month_bridge", "inject_ipad_plotly_controls"):
            body = _func(name)
            self.assertEqual(body.count("parentWin.location.replace(url.toString())"), 1, name)
            self.assertIn('if (!__SERVER_TOUCH__ && parentWin.sessionStorage.getItem("__dash_touch_boot")', body)
            self.assertIn('.replace("__SERVER_TOUCH__", "true" if is_touch_ui() else "false")', body)
            self.assertLess(body.index("__SERVER_TOUCH__ &&"), body.index("parentWin.location.replace(url.toString())"))


if __name__ == "__main__":
    unittest.main()
