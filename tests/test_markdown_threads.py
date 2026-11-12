"""The Markdown parsers are safe to share between threads from their first call (0.1.98,
found looking into the serial replay's flake): markdown-it compiles each rule chain lazily,
on the first parse, into a dict it publishes before filling -- a second thread parsing
meanwhile ran with some rules or none ("none of the block rules matched", a ``KeyError``, or
a table or heading read as a paragraph).  A process's first Markdown calls from several
threads at once could hit it.  The parsers are now compiled before any thread is given
them."""

from __future__ import annotations

import sys
import threading

import pytest

from docx_agent.markdown import parse, styled

TEXT = ("### Note 2.4\n\nMorning quiet *thread* valley[^1].\n\n| A | B |\n|---|---|\n| Evening | 1 |\n\n"
        "- a\n- ~~b~~\n\n[^1]: A note.\n")
SPANS = r"Plain *emphasis* and \*escaped\* text"


def _rulers(md):
    return (md.core.ruler, md.block.ruler, md.inline.ruler, md.inline.ruler2)


@pytest.mark.parametrize("make", [parse.parser, styled._parser], ids=["block", "inline"])
def test_a_parser_is_compiled_before_it_is_shared(make):
    make.cache_clear()
    try:
        md = make()
        assert all(ruler.__cache__ is not None for ruler in _rulers(md))
    finally:
        make.cache_clear()


def _first_parses_agree(read, threads: int = 4, rounds: int = 150) -> list:
    """``read()`` in ``threads`` threads at once, each round on parsers made afresh -- by the
    first caller, as a process's first call makes them, then shared -- with the interpreter
    switching threads as often as it can: what any thread read otherwise than one thread
    alone."""
    parse.parser.cache_clear()
    styled._parser.cache_clear()
    alone = read()
    wrong: list = []
    interval = sys.getswitchinterval()
    sys.setswitchinterval(1e-6)
    try:
        for _ in range(rounds):
            parse.parser.cache_clear()
            styled._parser.cache_clear()
            parse.parser()
            styled._parser()
            barrier = threading.Barrier(threads)
            found: list = []

            def work() -> None:
                barrier.wait()
                try:
                    found.append(read())
                except Exception as exc:  # noqa: BLE001 -- what the race raised
                    found.append(repr(exc))

            workers = [threading.Thread(target=work) for _ in range(threads)]
            for worker in workers:
                worker.start()
            for worker in workers:
                worker.join()
            wrong += [value for value in found if value != alone]
            if wrong:
                break
    finally:
        sys.setswitchinterval(interval)
        parse.parser.cache_clear()
        styled._parser.cache_clear()
    return wrong


def test_first_parses_in_threads_read_as_one_alone():
    assert not _first_parses_agree(lambda: parse.parse_ast(TEXT))


def test_first_span_parses_in_threads_read_as_one_alone():
    assert not _first_parses_agree(lambda: [token.type for token in styled._parser().parseInline(SPANS, {})[0].children])
