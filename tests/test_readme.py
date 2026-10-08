"""docs/common-tasks.md (once the README's "Common tasks") runs as written (ROADMAP.md, "Usability (end-to-end pilot)").

Each ``python`` block of that section runs on its own, in a fresh folder holding the files
the section names: ``agreement.docx`` (the pilot's input), ``charts.docx``, ``brand.dotx``
and ``logo.png``.  A recipe that stops working fails here, so the docs cannot rot.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

from conftest import FIXTURE_DIR

COMMON_TASKS = Path(__file__).parent.parent / "docs" / "common-tasks.md"
_BLOCK = re.compile(r"^### (?P<title>.+?)$|^```python\n(?P<code>.*?)^```$", re.M | re.S)


def common_tasks() -> list[tuple[str, str]]:
    text = COMMON_TASKS.read_text(encoding="utf-8")
    out, title = [], None
    for match in _BLOCK.finditer(text):
        if match.group("title"):
            title = match.group("title")
        else:
            out.append((title, match.group("code")))
    return out


TASKS = common_tasks()


def test_the_section_covers_the_tasks_an_agent_does_most():
    titles = " ".join(title for title, _ in TASKS).lower()
    for word in ("read", "replace", "tracked", "comments", "revisions", "markdown", "tables", "pictures",
                 "charts", "sections", "headers", "contents", "render", "templates", "validate"):
        assert word in titles, word


@pytest.mark.parametrize("title,code", TASKS, ids=[title for title, _ in TASKS])
def test_common_task_runs_as_written(title, code, tmp_path, monkeypatch):
    shutil.copy(FIXTURE_DIR / "generated" / "pilot" / "agreement-summary.docx", tmp_path / "agreement.docx")
    shutil.copy(FIXTURE_DIR / "generated" / "charts" / "charts.docx", tmp_path / "charts.docx")
    shutil.copy(FIXTURE_DIR / "generated" / "templates" / "brand.dotx", tmp_path / "brand.dotx")
    with zipfile.ZipFile(FIXTURE_DIR / "generated" / "lists-and-styles.docx") as package:
        (tmp_path / "logo.png").write_bytes(package.read("word/media/image1.png"))
    monkeypatch.chdir(tmp_path)
    exec(compile(code, f"docs/common-tasks.md: {title}", "exec"), {"__name__": "__readme__"})
