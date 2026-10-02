"""Cloud 데스크톱 업무입력은 로컬과 같은 CCv2 줄칸을 쓴다 (textarea quiet 아님)."""
from __future__ import annotations

import unittest
from unittest import mock


class WorklogCloudLinedUiTest(unittest.TestCase):
    def setUp(self):
        import worklog_tab as wt

        self.wt = wt

    def test_quiet_ui_only_on_ipad_not_cloud_linux(self):
        with mock.patch.object(self.wt, "_wl_is_ipad_ui", return_value=False), mock.patch.object(
            self.wt.platform, "system", return_value="Linux"
        ):
            self.assertFalse(self.wt._wl_quiet_ui())
        with mock.patch.object(self.wt, "_wl_is_ipad_ui", return_value=True), mock.patch.object(
            self.wt.platform, "system", return_value="Linux"
        ):
            self.assertTrue(self.wt._wl_quiet_ui())

    def test_mount_uses_ccv2_on_cloud_desktop(self):
        sentinel = object()

        class _SS(dict):
            pass

        ss = _SS()
        with mock.patch.object(self.wt, "_wl_quiet_ui", return_value=False), mock.patch.object(
            self.wt, "_WL_LINES_EDITOR", return_value=sentinel
        ) as editor, mock.patch.object(self.wt.st, "session_state", ss), mock.patch.object(
            self.wt.st, "text_area"
        ) as ta:
            got = self.wt._wl_mount_v2_lines_editor(
                key="k",
                data={"lines": ["a"]},
                default={"lines": ["a"]},
                on_lines_change=None,
                changed_key="chg",
                session_lines=["a"],
            )
        self.assertIs(got, sentinel)
        editor.assert_called_once()
        ta.assert_not_called()

    def test_mount_uses_textarea_on_ipad_quiet(self):
        class _SS(dict):
            pass

        ss = _SS()
        with mock.patch.object(self.wt, "_wl_quiet_ui", return_value=True), mock.patch.object(
            self.wt, "_WL_LINES_EDITOR"
        ) as editor, mock.patch.object(self.wt.st, "session_state", ss), mock.patch.object(
            self.wt.st, "text_area", return_value="x\ny"
        ) as ta:
            got = self.wt._wl_mount_v2_lines_editor(
                key="k",
                data={"lines": ["a"]},
                default={"lines": ["a"]},
                on_lines_change=None,
                changed_key="chg",
                session_lines=["a"],
            )
        self.assertEqual(got.lines, ["x", "y"])
        ta.assert_called_once()
        editor.assert_not_called()
        self.assertTrue(ss.get("chg"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
