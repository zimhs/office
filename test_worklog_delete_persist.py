"""업무일지 날짜 삭제가 새로고침/Drive 동기화 후에도 유지되는지."""
from __future__ import annotations

import json
import os
import tempfile
import unittest
from datetime import date
from unittest import mock


class WorklogDeletePersistTest(unittest.TestCase):
    def test_copy_worklog_tree_skips_deleted_day_file(self):
        import drive_autoload as da

        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "drive_worklog")
            dst = os.path.join(td, "cache_worklog")
            os.makedirs(src)
            os.makedirs(dst)
            day = "2026-10-01.xlsx"
            with open(os.path.join(src, day), "wb") as f:
                f.write(b"PK\x03\x04fake")
            with open(os.path.join(src, "template.xlsx"), "wb") as f:
                f.write(b"PK\x03\x04tmpl")
            with open(os.path.join(dst, "_worklog_deleted.json"), "w", encoding="utf-8") as f:
                json.dump({"2026-10-01": 1.0}, f)

            copied = da._copy_worklog_tree(src, dst, force=True)
            self.assertFalse(os.path.isfile(os.path.join(dst, day)))
            self.assertTrue(os.path.isfile(os.path.join(dst, "template.xlsx")))
            self.assertTrue(any("template" in c for c in copied) or os.path.isfile(os.path.join(dst, "template.xlsx")))

    def test_list_saved_excludes_deleted_manifest(self):
        import worklog_tab as wt

        with tempfile.TemporaryDirectory() as td:
            day = os.path.join(td, "2026-10-01.xlsx")
            with open(day, "wb") as f:
                f.write(b"PK\x03\x04fake")
            with open(os.path.join(td, "_worklog_deleted.json"), "w", encoding="utf-8") as f:
                json.dump({"2026-10-01": 123.0}, f)
            with mock.patch.object(wt, "WORKLOG_DIR", td), mock.patch.object(
                wt, "_list_archive_saved_dates", return_value=set()
            ), mock.patch.object(wt.st, "session_state", {}):
                got = wt.list_saved_worklog_dates()
            self.assertNotIn("2026-10-01", got)

    def test_read_cells_empty_when_marked_deleted(self):
        import worklog_tab as wt

        d = date(2026, 10, 1)
        with mock.patch.object(wt, "_worklog_day_marked_deleted", return_value=True):
            cells = wt.read_worklog_cells(d)
            arch = wt.read_worklog_cells_from_archive(d)
        # 아카이브/파일 복원 금지 + 빈 초안만
        self.assertIsNone(arch)
        self.assertFalse(wt._worklog_cells_have_draft(cells))

    def test_try_pull_skips_deleted(self):
        import worklog_tab as wt

        d = date(2026, 10, 1)
        with mock.patch.object(wt, "_worklog_day_marked_deleted", return_value=True), mock.patch.object(
            wt.st, "session_state", {}
        ):
            self.assertFalse(wt._try_pull_remote_worklog_day(d))


if __name__ == "__main__":
    unittest.main(verbosity=2)
