"""Sample documents installed with docx-agent, for shapes its own edits never make.

``accept_loses_text.docx`` (round 55, R119): the shape ``check``'s ``accept_loses_text``
names -- a row's last cell, and a table's last cell, whose last paragraph's mark is a pending
tracked deletion.  Open it as any document::

    from docx_agent import samples
    data = samples.read("accept_loses_text")          # bytes
    path = samples.path("accept_loses_text")          # a file on disk, for Word or a session

``tools/make_samples.py`` writes it and says how to make the shape in a document of one's
own; docs/common-tasks.md says what check and accept give on it.
"""

from __future__ import annotations

from importlib import resources
from pathlib import Path

#: The samples, by name.
NAMES = ("accept_loses_text",)


def _resource(name: str):
    if name not in NAMES:
        raise KeyError(f"no sample {name!r}; the samples are {', '.join(NAMES)}")
    return resources.files(__name__).joinpath(f"{name}.docx")


def read(name: str) -> bytes:
    """The sample's bytes."""
    return _resource(name).read_bytes()


def path(name: str) -> Path:
    """The sample's file on disk (installed as a plain file, as setuptools installs
    package data)."""
    return Path(str(_resource(name)))


__all__ = ["NAMES", "path", "read"]
