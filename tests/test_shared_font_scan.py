"""The render workers' font scan, made once (round 25, R1): the first worker to start scans
the font folders and saves the scan; the others take it to ``docx2svg.warm(scan=...)``,
which reads only the files changed or added since.  Renders are the same bytes with and
without it.  The fonts here are made up in memory (a name, a head and an OS/2 table: what a
scan reads); the end-to-end test renders a committed fixture with the machine's faces, in
subprocesses, so each starts cold.
"""

from __future__ import annotations

import json
import os
import pickle
import struct
import subprocess
import sys
from pathlib import Path

import pytest

from docx2svg import fonts

from docx_agent.tools import render

DOCUMENT = Path(__file__).resolve().parent / "fixtures" / "wordto" / "sample-with-table.docx"
_HEAD = struct.pack(">IIIIHH", 0x00010000, 0, 0, 0x5F0F3CF5, 0, 1000) + b"\0" * 34


def _face(family: str) -> bytes:
    """An sfnt of ``family`` with what a scan reads: ``name``, ``head`` and ``OS/2``."""
    text = family.encode("utf-16-be")
    name = struct.pack(">3H", 0, 1, 18) + struct.pack(">6H", 3, 1, 0x409, 1, len(text), 0) + text
    os2 = bytearray(78)
    struct.pack_into(">HH", os2, 4, 400, 5)
    tables = {b"OS/2": bytes(os2), b"head": _HEAD, b"name": name}
    directory, body, start = b"", b"", 12 + 16 * len(tables)
    for tag in sorted(tables):
        directory += struct.pack(">4sIII", tag, 0, start + len(body), len(tables[tag]))
        body += tables[tag] + b"\0" * (-len(tables[tag]) % 4)
    return struct.pack(">IHHHH", 0x00010000, len(tables), 0, 0, 0) + directory + body


@pytest.fixture
def folder(tmp_path, monkeypatch):
    """A font folder of two made-up faces, the only one scanned; the scan kept in
    ``tmp_path``; the font caches as in a new process."""
    fonts_dir = tmp_path / "fonts"
    fonts_dir.mkdir()
    (fonts_dir / "a.ttf").write_bytes(_face("Shared Scan A"))
    (fonts_dir / "b.ttf").write_bytes(_face("Shared Scan B"))
    place = str(tmp_path / "scan.pickle")
    monkeypatch.setenv("OOXML_FONT_DIRS", str(fonts_dir))
    monkeypatch.setattr(fonts, "default_font_dirs", lambda: ())
    monkeypatch.setattr(render, "_scan_place", lambda: place)

    def fresh() -> list[str]:
        """Forget what this process read, and count the files read from now on."""
        monkeypatch.setattr(fonts, "_SCANNED", {})
        monkeypatch.setattr(fonts, "_RASTERISER", {})
        fonts.installed_index.cache_clear()
        fonts._rasteriser_families_index.cache_clear()
        read: list[str] = []
        real = fonts._scanned_faces
        monkeypatch.setattr(fonts, "_scanned_faces", lambda path: read.append(path.name) or real(path))
        return read

    monkeypatch.setattr("docx2svg.png.svg_to_png", lambda *a, **k: b"")
    yield fonts_dir, place, fresh
    fonts.installed_index.cache_clear()
    fonts._rasteriser_families_index.cache_clear()


def test_the_first_worker_scans_and_the_next_reads_no_font_file(folder):
    fonts_dir, place, fresh = folder
    read = fresh()
    render.RENDERER.warm_up()                                 # the first worker
    assert sorted(read) == ["a.ttf", "b.ttf"]
    with open(place, "rb") as handle:
        saved = pickle.load(handle)
    assert set(saved.files) == {str(fonts_dir / "a.ttf"), str(fonts_dir / "b.ttf")}

    read = fresh()                                            # the next, a new process
    render.RENDERER.warm_up()
    assert read == []
    assert {"shared scan a", "shared scan b"} <= {key[0] for key in fonts.installed_index((fonts_dir,))}


def test_a_changed_font_folder_is_still_seen(folder):
    fonts_dir, place, fresh = folder
    fresh()
    render.RENDERER.warm_up()
    (fonts_dir / "c.ttf").write_bytes(_face("Shared Scan C"))
    (fonts_dir / "b.ttf").write_bytes(_face("Shared Scan B, changed"))
    read = fresh()
    render.RENDERER.warm_up()
    assert sorted(read) == ["b.ttf", "c.ttf"]                 # what changed, nothing else
    families = {key[0] for key in fonts.installed_index((fonts_dir,))}
    assert {"shared scan a", "shared scan b, changed", "shared scan c"} <= families
    read = fresh()                                            # saved again: the next reads none
    render.RENDERER.warm_up()
    assert read == []


def test_no_place_to_keep_it_scans_in_each_worker(folder, monkeypatch):
    _, place, fresh = folder
    monkeypatch.setattr(render, "_scan_place", lambda: None)
    read = fresh()
    render.RENDERER.warm_up()
    assert sorted(read) == ["a.ttf", "b.ttf"] and not os.path.exists(place)


def test_a_damaged_scan_is_made_again(folder):
    _, place, fresh = folder
    Path(place).write_bytes(b"not a scan")
    read = fresh()
    render.RENDERER.warm_up()
    assert sorted(read) == ["a.ttf", "b.ttf"]
    with open(place, "rb") as handle:
        assert isinstance(pickle.load(handle), fonts.FontScan)


_SCRIPT = """
import hashlib, json, sys
import docx2svg
from docx_agent.tools import render
from docx_agent.tools.shared import render_pages
mode, place = sys.argv[1], sys.argv[2]
render._scan_place = lambda: place
if mode != "cold":
    render.RENDERER.warm_up()
data = open({document!r}, "rb").read()
options = docx2svg.ConvertOptions()
svgs = docx2svg.convert_docx_to_svg(data, options)
images, count, coverage = render_pages(data, [1], 300)
print(json.dumps({{"svg": [hashlib.sha256(s.encode()).hexdigest() for s in svgs],
                  "png": [hashlib.sha256(i).hexdigest() for i in images], "count": count,
                  "warnings": [str(w) for w in options.warnings]}}))
"""


def test_renders_are_the_same_with_and_without_the_shared_scan(tmp_path):
    """Cold, the first worker (it scans and saves), the next (it takes the scan): three
    processes, the same pages, images and warnings."""
    pytest.importorskip("resvg_py")
    script = _SCRIPT.format(document=str(DOCUMENT))
    place = str(tmp_path / "scan.pickle")
    outputs = []
    for mode in ("cold", "first", "next"):
        result = subprocess.run([sys.executable, "-c", script, mode, place], capture_output=True, text=True,
                                check=False, env=dict(os.environ))
        assert result.returncode == 0, result.stderr
        outputs.append(json.loads(result.stdout.strip().splitlines()[-1]))
        if mode == "first":
            assert os.path.exists(place)
    assert outputs[0]["svg"] and outputs[1] == outputs[0] and outputs[2] == outputs[0]
