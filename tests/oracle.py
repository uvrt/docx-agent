"""Driving Microsoft Word -- one instance per machine -- as the oracle.

Word is the only authoritative answer to "will this file open?", and there is one of it.
docx2svg's measuring work drives the same Word, and its recovery path kills Word and deletes
``~$`` lock files, which would destroy an export of ours in flight, and ours would destroy
one of its.  So (ROADMAP.md, "Word is one instance per machine"):

* **One machine-wide lock.**  Every launch, export, save and recovery happens while holding
  an advisory ``flock`` on :data:`LOCK`, a file in the Office group container.  docx2svg's
  oracle takes the same lock (blocking, for up to 600 s) around its own Word work, and never
  quits a Word it did not start.  Here it is not waited for: if someone holds it, the
  oracle is skipped.
* **Word already running means someone is using it** -- a person, or a tool that does not
  take the lock -- so :func:`session` still refuses to start when Word is running, rather
  than wait, kill or share, and the tests skip with "Word is in use".
* **Recovery only under the lock**, and only of what this helper started: quitting the Word
  it launched, deleting ``~$`` files in its own staging directory.
* **Cached by content.**  Exports and re-saves are cached by the SHA-256 of the input, in
  :data:`ORACLE_DIR` (outside the repository: a Word PDF carries font subsets).
* **Paths under the sandbox.**  Word may read and write the Office group container without
  a grant; documents are staged there.

The AppleScripts are docx2svg's (``tools/word_export_pdf.applescript``, whose third argument
``final`` exports Word's final view of a document with revisions or comments, and
``tools/word_save_docx.applescript``), found in the sibling checkout or through
``DOCX2SVG_ORACLE_SCRIPT`` (the export script's path; the save script is beside it), as
pptx-agent finds pptx2svg's.
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

try:
    import fcntl
except ImportError:  # Windows: there is no Word oracle there (it drives Word on macOS)
    fcntl = None

WORD_APP = Path("/Applications/Microsoft Word.app")
EXPORT_SCRIPT = Path(
    os.environ.get("DOCX2SVG_ORACLE_SCRIPT")
    or Path(__file__).resolve().parents[2] / "docx2svg" / "tools" / "word_export_pdf.applescript"
)
SAVE_SCRIPT = EXPORT_SCRIPT.with_name("word_save_docx.applescript")

GROUP_CONTAINER = Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office"
#: The machine-wide Word lock (ROADMAP.md, "One machine-wide lock").
LOCK = GROUP_CONTAINER / "word-oracle.lock"
#: Where documents are staged for Word and its answers cached.
ORACLE_DIR = GROUP_CONTAINER / "docx-agent-oracle"

_MARKUP = (b"<w:ins ", b"<w:del ", b"<w:moveFrom ", b"<w:moveTo ", b"<w:commentReference",
           b"Change>", b"<w:commentRangeStart")


class WordBusy(RuntimeError):
    """Word is in use by someone else (or the lock is held): the oracle does not run."""


def available() -> bool:
    return WORD_APP.exists() and EXPORT_SCRIPT.exists() and SAVE_SCRIPT.exists()


def word_running() -> bool:
    return subprocess.run(["pgrep", "-x", "Microsoft Word"], capture_output=True).returncode == 0


@contextlib.contextmanager
def session() -> Iterator["Session"]:
    """Hold the machine-wide lock for a run of Word work; :class:`WordBusy` if Word or the
    lock is taken.  Word is left not running on exit."""
    if not available():
        raise WordBusy("Microsoft Word or docx2svg's oracle scripts are not available")
    GROUP_CONTAINER.mkdir(parents=True, exist_ok=True)
    handle = open(LOCK, "a+")
    try:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise WordBusy("another job holds the Word lock") from None
        if word_running():
            raise WordBusy("Word is in use")
        ORACLE_DIR.mkdir(exist_ok=True)
        current = Session()
        try:
            yield current
        finally:
            current.recover()
    finally:
        fcntl.flock(handle, fcntl.LOCK_UN)
        handle.close()


@dataclass
class Outcome:
    ok: bool
    #: ``done``, ``rejected`` (Word blocked: a dialog, which for a file means a repair
    #: prompt) or ``error``.
    outcome: str
    path: Path | None = None
    detail: str = ""
    seconds: float = 0.0

    def __bool__(self) -> bool:
        return self.ok


class Session:
    """Word work under the lock.  Every method stages its input in :data:`ORACLE_DIR`."""

    def export_pdf(self, data: bytes, *, name: str = "doc", timeout: int = 150, suffix: str = ".docx") -> Outcome:
        """Word's PDF of ``data``, in its final view when the document has revisions or
        comments; cached by content.  ``suffix``: the extension the file is staged under
        (``.dotx``, ``.docm``...: Word holds a file's kind to its extension)."""
        final = _zip_has_markup(data)
        digest = hashlib.sha256(data + suffix.encode()).hexdigest()[:16] if suffix != ".docx" \
            else hashlib.sha256(data).hexdigest()[:16]
        pdf = ORACLE_DIR / f"{name}-{digest}{'-final' if final else ''}.pdf"
        if pdf.exists():
            return Outcome(True, "done", pdf)
        return self._run([str(EXPORT_SCRIPT)], data, pdf, name, digest, timeout,
                         extra=["final"] if final else [], suffix=suffix)

    def resave(self, data: bytes, *, name: str = "doc", timeout: int = 150) -> Outcome:
        """``data`` opened and saved again by Word as a .docx; cached by content."""
        digest = hashlib.sha256(data).hexdigest()[:16]
        saved = ORACLE_DIR / f"{name}-{digest}-saved.docx"
        if saved.exists():
            return Outcome(True, "done", saved)
        return self._run([str(SAVE_SCRIPT)], data, saved, name, digest, timeout)

    def run_script(self, script: str, data: bytes | None, *, name: str, tag: str,
                   args: list[str] = (), timeout: int = 150) -> Outcome:
        """Run an AppleScript given as text with ``(input .docx, output .docx, *args)``;
        with ``data`` ``None`` there is no input (the script makes its own document)."""
        digest = hashlib.sha256((data or b"") + script.encode() + "\0".join(args).encode()).hexdigest()[:16]
        out = ORACLE_DIR / f"{name}-{digest}-{tag}.docx"
        if out.exists():
            return Outcome(True, "done", out)
        return self._run(["-e", script], data, out, name, digest, timeout, extra=list(args))

    def _run(self, command: list[str], data: bytes | None, out: Path, name: str, digest: str,
             timeout: int, extra: list[str] = (), suffix: str = ".docx") -> Outcome:
        staged = ORACLE_DIR / f"{name}-{digest}-in{suffix}"
        if data is not None:
            staged.write_bytes(data)
        # The previous run's Word may still be quitting; launching into it fails with -609
        # or -1712 and nothing else to show for it (docx2svg's tools/oracle.py measured).
        for _ in range(150):
            if not word_running():
                break
            time.sleep(0.1)
        start = time.monotonic()
        try:
            completed = subprocess.run(["osascript", *command, str(staged), str(out), *extra],
                                       capture_output=True, text=True, timeout=timeout)
            detail, timed_out = (completed.stderr or completed.stdout).strip(), False
        except subprocess.TimeoutExpired:
            detail, timed_out = f"osascript did not return within {timeout}s", True
        seconds = time.monotonic() - start
        ok = not timed_out and out.exists() and out.stat().st_size > 0
        if not ok:
            self.recover()
        staged.unlink(missing_ok=True)
        self._clear_locks()
        if ok:
            return Outcome(True, "done", out, detail, seconds)
        blocked = timed_out or "-1712" in detail
        out.unlink(missing_ok=True)
        return Outcome(False, "rejected" if blocked else "error", None, detail, seconds)

    def recover(self) -> None:
        """Quit the Word this session started, and remove the ``~$`` locks it left."""
        if word_running():
            try:
                subprocess.run(["osascript", "-e", 'tell application "Microsoft Word" to quit saving no'],
                               capture_output=True, timeout=30)
            except subprocess.TimeoutExpired:
                pass  # blocked on a dialog: killed below
            for _ in range(50):
                if not word_running():
                    break
                time.sleep(0.2)
            if word_running():
                subprocess.run(["pkill", "-x", "Microsoft Word"], check=False)
                for _ in range(100):
                    if not word_running():
                        break
                    time.sleep(0.1)
            if word_running():
                # Seen once (E5): a Word hung on a dialog ignores the polite signal.
                subprocess.run(["pkill", "-9", "-x", "Microsoft Word"], check=False)
                for _ in range(100):
                    if not word_running():
                        break
                    time.sleep(0.1)
        self._clear_locks()

    @staticmethod
    def _clear_locks() -> None:
        for lock in ORACLE_DIR.glob("~$*"):
            lock.unlink(missing_ok=True)


def _zip_has_markup(data: bytes) -> bool:
    import io
    import zipfile

    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                part = archive.read(name)
                if any(mark in part for mark in _MARKUP):
                    return True
    return False


def pdf_pages(pdf: Path) -> list[str]:
    """The text of every page of a PDF (pypdfium2, docx2svg's oracle extra)."""
    import pypdfium2

    document = pypdfium2.PdfDocument(str(pdf))
    try:
        return [page.get_textpage().get_text_range() for page in document]
    finally:
        document.close()


@dataclass(frozen=True)
class PdfChar:
    """One character of a PDF page as pdfium reads it: what it is, the size it is drawn at
    in points, weight (400 regular, 700 bold; 0 or -1 when the font does not say -- Word's
    subsets do not, their names do), fill colour and font name."""

    char: str
    size: float
    weight: int
    color: tuple[int, int, int]
    font: str
    #: Filled and stroked (text render mode 2): how Word draws bold in a face without a bold
    #: variant (Calibri Light).
    stroked: bool = False


def pdf_chars(pdf: Path) -> list[list[PdfChar]]:
    """Every character of every page, with how it is drawn (pypdfium2)."""
    import ctypes

    import pypdfium2
    import pypdfium2.raw as raw

    document = pypdfium2.PdfDocument(str(pdf))
    pages: list[list[PdfChar]] = []
    try:
        for page in document:
            textpage = page.get_textpage()
            chars = []
            for index in range(textpage.count_chars()):
                code = raw.FPDFText_GetUnicode(textpage.raw, index)
                red, green, blue, alpha = (ctypes.c_uint(), ctypes.c_uint(), ctypes.c_uint(), ctypes.c_uint())
                raw.FPDFText_GetFillColor(textpage.raw, index, red, green, blue, alpha)
                buffer = ctypes.create_string_buffer(256)
                flags = ctypes.c_int()
                raw.FPDFText_GetFontInfo(textpage.raw, index, buffer, 256, flags)
                matrix = raw.FS_MATRIX()
                raw.FPDFText_GetMatrix(textpage.raw, index, matrix)
                # Word's PDF sets text at 1 pt and scales it by the text matrix.
                scale = (matrix.a ** 2 + matrix.b ** 2) ** 0.5 or 1.0
                size = raw.FPDFText_GetFontSize(textpage.raw, index) * scale
                obj = raw.FPDFText_GetTextObject(textpage.raw, index)
                mode = raw.FPDFTextObj_GetTextRenderMode(obj) if obj else 0
                chars.append(PdfChar(chr(code) if code else "", round(size, 2),
                                     raw.FPDFText_GetFontWeight(textpage.raw, index),
                                     (red.value, green.value, blue.value), buffer.value.decode("latin-1"),
                                     mode == 2))
            pages.append(chars)
    finally:
        document.close()
    return pages


def find_in_pdf(pages: list[list[PdfChar]], text: str) -> list[PdfChar] | None:
    """The characters of the first place ``text`` is drawn, or ``None``.  Whitespace is
    collapsed: Word's PDF separates justified words with tabs and lines with CRLF."""
    wanted = " ".join(text.split())
    for chars in pages:
        kept: list[PdfChar] = []
        joined = []
        for c in chars:
            if c.char.isspace():
                if joined and joined[-1] != " ":
                    joined.append(" ")
                    kept.append(c)
            else:
                joined.append(c.char)
                kept.append(c)
        at = "".join(joined).find(wanted)
        if at >= 0:
            return kept[at:at + len(wanted)]
    return None
