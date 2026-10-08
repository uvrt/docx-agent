"""Find and replace held to an independent model over a randomised sweep (ROADMAP.md, Phase
E1, "Done when").

Paragraphs are generated from random pieces -- runs of random formatting, hyperlinks,
insertions and deletions by another author, complex and simple fields, proofing marks and
bookmarks -- and random spans of their current text are replaced.  The model is a list of
``(character, formatting)``; formatting is a run's properties together with the containers
it sits in (a hyperlink, an insertion, a field's result).  After every replacement:

* the text is the model's: the span replaced, everything else as it was;
* every character outside the span keeps its formatting, and every new character takes the
  formatting of the span's first character (Word's rule);
* a span that cuts a field, or runs from a field's result into other text, is refused and
  nothing changes;
* text deleted by another author is still there, in the original view;
* no run is left empty, the validity checks find nothing new, and undo gives back the
  original bytes.
"""

from __future__ import annotations

import io
import random
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.edit import text as _text
from docx_agent.validate import check

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
from make_fixtures import DECL, NS, REL, WML, content_types, relationships, settings, styles  # noqa: E402

from test_roundtrip import entries  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WORDS = ["net", "revenue", "grew", "by", "four", "teen", "per", "cent", "in", "Q3", "a", "the", "and"]
FORMATS = ["", "<w:b/>", "<w:i/>", "<w:b/><w:i/>", '<w:color w:val="C00000"/>', '<w:u w:val="single"/>',
           '<w:sz w:val="28"/><w:szCs w:val="28"/>']
AUTHOR = 'w:author="Other" w:date="2026-10-01T09:00:00Z"'


def chunk(rng: random.Random) -> str:
    text = " ".join(rng.choice(WORDS) for _ in range(rng.randint(1, 3)))
    if rng.random() < 0.5:
        text = text[: rng.randint(1, len(text))]
    return text


def run(rng: random.Random, text: str | None = None) -> str:
    props = rng.choice(FORMATS)
    body = text if text is not None else chunk(rng)
    return f'<w:r>{f"<w:rPr>{props}</w:rPr>" if props else ""}<w:t xml:space="preserve">{body}</w:t></w:r>'


def piece(rng: random.Random, ids: list[int]) -> str:
    kind = rng.choices(["run", "link", "ins", "del", "field", "simple", "proof", "bookmark"],
                       [10, 2, 2, 1, 1, 1, 1, 1])[0]
    if kind == "run":
        return run(rng)
    if kind == "link":
        return '<w:hyperlink r:id="rIdLink" w:history="1">' + "".join(run(rng) for _ in range(rng.randint(1, 3))) + "</w:hyperlink>"
    if kind == "ins":
        ids[0] += 1
        return f'<w:ins w:id="{ids[0]}" {AUTHOR}>' + run(rng) + "</w:ins>"
    if kind == "del":
        ids[0] += 1
        return f'<w:del w:id="{ids[0]}" {AUTHOR}><w:r><w:delText xml:space="preserve">{chunk(rng)}</w:delText></w:r></w:del>'
    if kind == "field":
        return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(rng, str(rng.randint(1, 99)))
                + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    if kind == "simple":
        return '<w:fldSimple w:instr=" AUTHOR ">' + run(rng) + "</w:fldSimple>"
    if kind == "proof":
        return '<w:proofErr w:type="spellStart"/>' + run(rng) + '<w:proofErr w:type="spellEnd"/>'
    ids[0] += 1
    return f'<w:bookmarkStart w:id="{ids[0]}" w:name="b{ids[0]}"/>' + run(rng) + f'<w:bookmarkEnd w:id="{ids[0]}"/>'


def generated(seed: int, paragraphs: int = 12) -> bytes:
    rng = random.Random(seed)
    ids = [100]
    body = "".join(
        f'<w:p w14:paraId="{0x10000000 + k:08X}" w14:textId="77777777">'
        + "".join(piece(rng, ids) for _ in range(rng.randint(3, 9))) + "</w:p>"
        for k in range(paragraphs))
    body += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False),
                                                       ("rIdLink", REL + "hyperlink", "https://example.com/", True)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def model(paragraph) -> list[tuple[str, str, object]]:
    """(character, formatting, field) for every character of the current view, read
    independently of the edit's own code: the run's properties and its containers."""
    out = []
    for atom in _text.atoms(paragraph._element):
        run_element = atom.run
        properties = run_element.find(W + "rPr")
        key = etree.tostring(properties, method="c14n", exclusive=True).decode() if properties is not None else ""
        chain = []
        node = run_element.getparent()
        while node is not None and node.tag != W + "p":
            chain.append(node.tag.rpartition("}")[2])
            node = node.getparent()
        field = "complex" if atom.field is not None and atom.field.tag == W + "fldChar" else None
        out.append((atom.char, key + "|" + "/".join(chain) + ("|in-field" if field else ""), atom.field))
    return out


@pytest.mark.parametrize("seed", range(8))
def test_replace_against_the_model(seed):
    data = generated(seed)
    document = Document.open(data)
    baseline = set(check(document.package))
    rng = random.Random(1000 + seed)
    replaced = refused = 0
    for _ in range(60):
        paragraphs = [p for p in document.paragraphs() if p.text]
        paragraph = rng.choice(paragraphs)
        before = model(paragraph)
        text = paragraph.text
        start = rng.randrange(len(text))
        end = rng.randint(start + 1, min(len(text), start + 25))
        replacement = rng.choice(["", "X", "new words", " ", chunk(rng)])
        deletions = [node.text for node in paragraph._element.iter(W + "delText")]
        span_fields = {id(f) if f is not None else None for _, _, f in before[start:end]}
        snapshot = entries(document.to_bytes())
        found = document.range(f"{paragraph.id}@{start}:{end}")
        if len(span_fields) > 1:
            with pytest.raises(EditError):
                found.replace(replacement)
            assert entries(document.to_bytes()) == snapshot
            refused += 1
            continue
        found.replace(replacement)
        replaced += 1
        paragraph = document.paragraph(paragraph.id)
        after = model(paragraph)
        assert paragraph.text == text[:start] + replacement + text[end:]
        assert [(c, f) for c, f, _ in after[:start]] == [(c, f) for c, f, _ in before[:start]]
        tail = len(text) - end
        if tail:
            assert [(c, f) for c, f, _ in after[-tail:]] == [(c, f) for c, f, _ in before[end:]]
        first = before[start][1]
        assert all(f == first for _, f, _ in after[start:start + len(replacement)])
        assert [node.text for node in paragraph._element.iter(W + "delText")] == deletions
        assert all(run.text for run in paragraph.runs if run._element.find(W + "fldChar") is None
                   and run._element.find(W + "instrText") is None and run._element.find(W + "delText") is None)
        assert document.undo()
        assert entries(document.to_bytes()) == snapshot
        assert document.redo()
    assert replaced > 20
    print(f"seed {seed}: {replaced} replaced, {refused} refused")
    assert set(check(Document.open(document.to_bytes()).package)) - baseline == set()


@pytest.mark.parametrize("seed", range(4))
def test_document_replace_matches_every_occurrence_like_str_replace(seed):
    data = generated(50 + seed, paragraphs=10)
    document = Document.open(data)
    rng = random.Random(seed)
    for needle in rng.sample(WORDS, 4):
        texts = [p.text for p in document.paragraphs()]
        expected_refusal = False
        for p in document.paragraphs():
            for found in p.find(needle):
                owners = {id(a.field) if a.field is not None else None
                          for a in _text.atoms(p._element)[found.start:found.end]}
                expected_refusal |= len(owners) > 1
        if expected_refusal:
            with pytest.raises(EditError):
                document.replace(needle, needle.upper() + "!")
            continue
        result = document.replace(needle, needle.upper() + "!")
        assert result.count == sum(t.count(needle) for t in texts)
        assert [p.text for p in document.paragraphs()] == [t.replace(needle, needle.upper() + "!") for t in texts]
