"""The oracle's machine-wide Word lock, without Word: a held lock means "skip", never wait."""

from __future__ import annotations

import sys

import pytest

fcntl = pytest.importorskip("fcntl")  # flock is POSIX's; the Word oracle runs on macOS

import oracle

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="Word's group container is macOS's")


def test_a_held_lock_is_word_in_use(monkeypatch):
    monkeypatch.setattr(oracle, "available", lambda: True)
    oracle.GROUP_CONTAINER.mkdir(parents=True, exist_ok=True)
    with open(oracle.LOCK, "a+") as other:
        fcntl.flock(other, fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            with pytest.raises(oracle.WordBusy, match="lock"):
                with oracle.session():
                    pass
        finally:
            fcntl.flock(other, fcntl.LOCK_UN)


def test_a_running_word_is_word_in_use(monkeypatch):
    monkeypatch.setattr(oracle, "available", lambda: True)
    monkeypatch.setattr(oracle, "word_running", lambda: True)
    with pytest.raises(oracle.WordBusy, match="Word is in use"):
        with oracle.session():
            pass
