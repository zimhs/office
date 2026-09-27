"""맥 Drive 동기화 버튼 → 코드만 GitHub office/main 에 올린다.

캐시·매출·비밀·방문할일 단가 파일은 넣지 않는다.
"""
from __future__ import annotations

import os
import subprocess
from typing import Dict, List

_OFFICE_PUSH_URL = "https://github.com/zimhs/office.git"
_DENY_NAMES = {
    "visit_calendar_tab.py",
    "test_visit_calendar.py",
}


def repo_root() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def list_office_code_files(root: str | None = None) -> List[str]:
    base = root or repo_root()
    out: List[str] = []
    try:
        names = os.listdir(base)
    except OSError:
        return out
    for name in sorted(names):
        if name in _DENY_NAMES:
            continue
        if name.endswith(".py") or name == "requirements.txt":
            out.append(name)
    return out


def _run(cmd: List[str], cwd: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        cwd=cwd,
        text=True,
        capture_output=True,
        timeout=120,
    )


def push_office_code(*, message: str = "") -> Dict[str, object]:
    """코드 변경이 있으면 커밋 후 origin/main 에 push. 데이터 파일은 제외."""
    root = repo_root()
    files = [n for n in list_office_code_files(root) if os.path.isfile(os.path.join(root, n))]
    if not files:
        return {"ok": False, "skipped": True, "error": "올릴 코드 파일 없음", "files": []}
    add = _run(["git", "add", "--"] + files, root)
    if add.returncode != 0:
        return {
            "ok": False,
            "error": (add.stderr or add.stdout or "git add 실패")[:240],
            "files": files,
        }
    dirty = _run(["git", "diff", "--cached", "--quiet", "--"] + files, root)
    committed = False
    if dirty.returncode == 1:
        msg = (message or "chore: 맥 Drive 동기화에서 코드를 Cloud에 반영한다").strip()
        commit = _run(["git", "commit", "-m", msg], root)
        if commit.returncode != 0:
            return {
                "ok": False,
                "error": (commit.stderr or commit.stdout or "git commit 실패")[:240],
                "files": files,
            }
        committed = True
    push = _run(["git", "push", _OFFICE_PUSH_URL, "HEAD:main"], root)
    if push.returncode != 0:
        return {
            "ok": False,
            "committed": committed,
            "error": (push.stderr or push.stdout or "git push 실패")[:240],
            "files": files,
        }
    return {
        "ok": True,
        "committed": committed,
        "skipped": not committed,
        "files": files,
    }
