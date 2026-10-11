"""Rendering and layout in a daemonic process (a Celery prefork worker's child):
multiprocessing may not start workers there, so ooxml-edit's pool uses subprocess workers,
which run only registered tasks -- the document's renderer and layout are."""

from __future__ import annotations

import multiprocessing
import os

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.worker import task_reference

from docx_agent import Document
from docx_agent.layout import convert_bytes
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.tools.shared import render_pages

import docx2svg


def _document() -> bytes:
    document = Document.new()
    document.append_paragraph("Laid out and rendered elsewhere.")
    return document.to_bytes()


def test_the_renderer_and_the_layout_are_registered_tasks():
    assert task_reference(render_pages) == "docx_agent.tools.shared:render_pages"
    assert task_reference(convert_bytes) == "docx_agent.layout:convert_bytes"


def work_in_a_daemonic_process(queue, data: bytes) -> None:
    try:
        with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=[]) as box:
            session = box.session()
            session.open(data, "report.docx")
            render = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 300})
            check = box.dispatch(session, "check", {"doc": "d1", "include": ["reflow"]})
            queue.put({"ok": render.ok and check.ok, "runner": box.pool.runner,
                       "pids": box.pool.pids(), "pid": os.getpid(),
                       "image": render.images[0].data if render.ok else None,
                       "error": [r.error.message for r in (render, check) if r.error]})
    except BaseException as exc:  # noqa: BLE001 -- reported to the test
        queue.put({"ok": False, "error": f"{type(exc).__name__}: {exc}"})


def test_a_daemonic_process_works_in_subprocess_workers():
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")
    data = _document()
    context = multiprocessing.get_context("spawn")
    queue = context.Queue()
    process = context.Process(target=work_in_a_daemonic_process, args=(queue, data),
                              daemon=True)
    process.start()
    try:
        got = queue.get(timeout=240)
    finally:
        process.join(30)
    assert got["ok"], got["error"]
    assert got["runner"] == "subprocesses"
    assert got["pids"] and got["pid"] not in got["pids"]
    assert got["image"] == render_pages(data, [1], 300, [])[0][0]
