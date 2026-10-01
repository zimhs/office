"""Drive 동기화 → office/main 코드 push (fetch first 거절 방지)."""
from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from unittest.mock import patch

import office_code_push as ocp


class _Proc:
    def __init__(self, code=0, stdout="", stderr=""):
        self.returncode = code
        self.stdout = stdout
        self.stderr = stderr


class OfficeCodePushTest(unittest.TestCase):
    def test_push_fetches_before_push(self):
        calls: list[list[str]] = []

        def fake_run(cmd, cwd, text=True, capture_output=True, timeout=120):
            calls.append(list(cmd))
            if cmd[:2] == ["git", "fetch"]:
                return _Proc(0)
            if cmd[:2] == ["git", "add"]:
                return _Proc(0)
            if cmd[:3] == ["git", "diff", "--cached"]:
                return _Proc(0)  # clean
            if cmd[:2] == ["git", "rebase"]:
                return _Proc(0)
            if cmd[:3] == ["git", "rev-list", "--count"]:
                return _Proc(0, stdout="0\n")
            if cmd[:2] == ["git", "push"]:
                return _Proc(0)
            return _Proc(0)

        with tempfile.TemporaryDirectory() as tmp:
            open(os.path.join(tmp, "app.py"), "w").write("#x\n")
            with patch.object(ocp, "repo_root", return_value=tmp), patch.object(
                ocp.subprocess, "run", side_effect=fake_run
            ):
                res = ocp.push_office_code()
        self.assertTrue(res.get("ok"))
        self.assertTrue(res.get("skipped"))
        self.assertEqual(calls[0][:3], ["git", "fetch", ocp._OFFICE_PUSH_URL])
        self.assertIn("main", calls[0])
        # ahead=0 이면 push 호출 없음
        self.assertFalse(any(c[:2] == ["git", "push"] for c in calls))

    def test_push_after_commit_rebases_then_pushes(self):
        calls: list[list[str]] = []

        def fake_run(cmd, cwd, text=True, capture_output=True, timeout=120):
            calls.append(list(cmd))
            if cmd[:2] == ["git", "fetch"]:
                return _Proc(0)
            if cmd[:2] == ["git", "add"]:
                return _Proc(0)
            if cmd[:3] == ["git", "diff", "--cached"]:
                return _Proc(1)  # dirty
            if cmd[:2] == ["git", "commit"]:
                return _Proc(0)
            if cmd[:2] == ["git", "rebase"]:
                return _Proc(0)
            if cmd[:3] == ["git", "rev-list", "--count"]:
                return _Proc(0, stdout="1\n")
            if cmd[:2] == ["git", "push"]:
                return _Proc(0)
            return _Proc(0)

        with tempfile.TemporaryDirectory() as tmp:
            open(os.path.join(tmp, "app.py"), "w").write("#y\n")
            with patch.object(ocp, "repo_root", return_value=tmp), patch.object(
                ocp.subprocess, "run", side_effect=fake_run
            ):
                res = ocp.push_office_code(message="test commit")
        self.assertTrue(res.get("ok"))
        self.assertTrue(res.get("committed"))
        kinds = [c[1] for c in calls if c and c[0] == "git"]
        self.assertEqual(kinds[:5], ["fetch", "add", "diff", "commit", "rebase"])
        self.assertIn("push", kinds)

    def test_old_rejected_message_not_returned_when_rebase_ok(self):
        """예전: fetch 없이 push → rejected (fetch first)."""
        def fake_run(cmd, cwd, text=True, capture_output=True, timeout=120):
            if cmd[:2] == ["git", "fetch"]:
                return _Proc(0)
            if cmd[:2] == ["git", "add"]:
                return _Proc(0)
            if cmd[:3] == ["git", "diff", "--cached"]:
                return _Proc(0)
            if cmd[:2] == ["git", "rebase"]:
                return _Proc(0)
            if cmd[:3] == ["git", "rev-list", "--count"]:
                return _Proc(0, stdout="0\n")
            return _Proc(1, stderr="! [rejected] HEAD -> main (fetch first)")

        with tempfile.TemporaryDirectory() as tmp:
            open(os.path.join(tmp, "app.py"), "w").write("#z\n")
            with patch.object(ocp, "repo_root", return_value=tmp), patch.object(
                ocp.subprocess, "run", side_effect=fake_run
            ):
                res = ocp.push_office_code()
        self.assertTrue(res.get("ok"))
        self.assertNotIn("rejected", str(res.get("error") or "").lower())


if __name__ == "__main__":
    unittest.main(verbosity=2)
