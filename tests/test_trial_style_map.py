"""``insert_markdown(style_map=...)`` as a plain dict (ROADMAP.md, "Trial findings", 2): the
keys, ``StyleMap``, ``Rule`` and the default table exported, and ``paragraph`` scoped to the
body's paragraphs -- a table cell's and a footnote's keep the document's defaults unless
named."""

from __future__ import annotations

import pytest

import docx_agent
from docx_agent import DEFAULT_STYLE_MAP, Document, EditError, Rule, StyleMap
from docx_agent.markdown.stylemap import DICT_KEYS

DRAFT = """# Annual report

Prepared by the team.[^1]

## Summary

Body text.

| Region | Total |
| --- | ---: |
| North | 4.1 |

> A quote.

```
code
```

[^1]: A note.
"""


def _styles(document: Document, story: str = "body") -> list[tuple[str, str | None]]:
    return [(p.text, p.style_name) for p in document.paragraphs(story)]


def report() -> Document:
    document = Document.new()
    document.styles.add("Report Title", based_on="Title", size=28)
    document.styles.add("Report Body", based_on="Normal", size=10.5)
    return document


def test_exported_from_the_package():
    assert {"StyleMap", "Rule", "DEFAULT_STYLE_MAP"} <= set(docx_agent.__all__)
    assert StyleMap.DEFAULT is DEFAULT_STYLE_MAP
    assert StyleMap.DEFAULT.style_for("heading", 1).style == "heading 1"
    assert isinstance(StyleMap.DEFAULT.rules[0], Rule)


def test_a_plain_dict_maps_headings_body_quote_and_code():
    document = report()
    document.insert_markdown(DRAFT, style_map={"h1": "Report Title", "h2": "Heading 1", "paragraph": "Report Body",
                                               "quote": "Intense Quote", "code": "Plain Text"})
    styles = dict(_styles(document))
    assert styles["Annual report"] == "Report Title"
    assert styles["Summary"] == "heading 1"
    assert styles["Prepared by the team.\ufffc"] == styles["Body text."] == "Report Body"   # \ufffc: the note
    assert styles["A quote."] == "Intense Quote"
    assert styles["code"] == "Plain Text"
    assert document.validate() == []


def test_paragraph_is_body_text_only():
    """The trial: mapping paragraph restyled table cells and the footnote's text too."""
    document = report()
    document.insert_markdown(DRAFT, style_map={"paragraph": "Report Body"})
    styles = dict(_styles(document))
    assert styles["North"] == styles["4.1"] == "Normal"          # a table cell: the default
    note = document.notes("footnote")[0]
    assert [document.paragraph(p).style_name for p in note.paragraph_ids] == ["footnote text"]


def test_cells_and_notes_take_a_style_when_named():
    document = report()
    document.insert_markdown(DRAFT, style_map={"paragraph": "Report Body", "table_cell": "Report Body",
                                               "footnote": "Report Body"})
    assert dict(_styles(document))["North"] == "Report Body"
    note = document.notes("footnote")[0]
    assert [document.paragraph(p).style_name for p in note.paragraph_ids] == ["Report Body"]


@pytest.mark.parametrize("key,construct", [("heading1", ("heading", 1)), ("Heading 2", ("heading", 2)),
                                           ("blockquote", ("blockquote", 0)), ("bullet2", ("bullet_list", 2)),
                                           ("number", ("ordered_list", 1)), ("inline_code", ("code", 0))])
def test_the_keys_and_their_other_spellings(key, construct):
    mapped = StyleMap.from_dict({key: "Custom"})
    assert mapped.style_for(*construct).style == "Custom"
    assert mapped.style_for("paragraph").style == "Normal"       # the rest is the default


def test_every_key_round_trips_through_to_dict():
    assert set(StyleMap.DEFAULT.to_dict()) == set(DICT_KEYS)
    assert StyleMap.from_dict(StyleMap.DEFAULT.to_dict()).to_dict() == StyleMap.DEFAULT.to_dict()


def test_a_bad_key_or_value_says_what_is_taken():
    document = Document.new()
    with pytest.raises(ValueError, match="the keys are paragraph, h1"):
        document.insert_markdown("x", style_map={"heading": "Title"})
    with pytest.raises(TypeError, match="a style name"):
        document.insert_markdown("x", style_map={"h1": 1})
    with pytest.raises(TypeError, match="a dict such as"):
        document.insert_markdown("x", style_map=[("h1", "Title")])
    with pytest.raises(EditError, match="no paragraph style 'Nope'"):
        document.insert_markdown("x", style_map={"paragraph": "Nope"})


def test_a_built_in_style_named_is_added_not_dropped():
    """A mapped body style the document lacks was silently Normal; a built-in one is added as
    Word writes it."""
    document = Document.new()
    result = document.insert_markdown("Body text.", style_map={"paragraph": "Body Text"})
    assert [s for s in _styles(document) if s[0]] == [("Body text.", "Body Text")]
    assert any("Body Text" in warning for warning in result.warnings)


def test_a_style_map_object_still_works():
    style_map = StyleMap.DEFAULT.with_rules(Rule("heading", "Title", 1))
    document = Document.new()
    document.insert_markdown("# Title\n\nText.", style_map=style_map)
    assert ("Title", "Title") in _styles(document)
