"""W5: accept Alice's revisions, reject Bob's, keep Chen's pending; add a tracked summary section."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w5-selective-review/input/remote-working-guideline.docx"


def texts(d):
    return [norm(p.text) for p in d.paragraphs()]


def expected(chen: str):
    d = Document.open(INPUT)
    for r in d.revisions(author="Alice Moreau"):
        pass
    d.accept([r.id for r in d.revisions(author="Alice Moreau")])
    d.reject([r.id for r in d.revisions(author="Bob Lindqvist")])
    if chen == "accept":
        d.accept_all()
    else:
        d.reject_all()
    return texts(d)


def body(c, out: Path):
    doc = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    revs = doc.revisions()
    authors = sorted({r.author for r in revs})
    c.check("no Alice or Bob revisions left", not [r for r in revs if r.author in ("Alice Moreau", "Bob Lindqvist")],
            authors)
    chen_src = sorted((r.kind, norm(r.text)) for r in src.revisions(author="Chen Wei"))
    chen_out = sorted((r.kind, norm(r.text)) for r in doc.revisions(author="Chen Wei"))
    c.check("Chen's revisions still pending, unchanged", chen_out == chen_src, chen_out)
    c.check("only Chen and Claude revisions remain", set(authors) <= {"Chen Wei", "Claude"}, authors)

    rej = Document.open(out)
    rej.reject_all()
    exp_rej = expected("reject")
    c.check("reject_all(): input with Alice accepted, Bob rejected, Chen rejected, no summary",
            texts(rej) == exp_rej, [x for x in zip(texts(rej), exp_rej) if x[0] != x[1]][:4] or
            (len(texts(rej)), len(exp_rej), texts(rej)[-4:]))
    acc = Document.open(out)
    acc.accept_all()
    exp_acc = expected("accept")
    got = texts(acc)
    c.check("accept_all(): the same plus Chen's changes, before the summary",
            got[:len(exp_acc)] == exp_acc, [x for x in zip(got, exp_acc) if x[0] != x[1]][:4])
    tail = acc.paragraphs()[len(exp_acc):]
    if tail and not norm(tail[0].text) and len(exp_acc) and not exp_acc[-1]:
        tail = tail[1:]
    heads = [p for p in tail if norm(p.text) == "Review summary"]
    c.check("'Review summary' heading added at the end", heads, [norm(p.text) for p in tail][:3])
    if heads:
        # TASK.md says "Heading 1 like the other section headings", but those are Heading 2 (a task-text
        # defect found by run1): either is accepted.
        c.check("summary heading is a heading (1 or 2)", (heads[0].style_name or "").lower() in ("heading 1",
                "heading 2"), heads[0].style_name)
    bullets = [p for p in tail if p.list is not None and norm(p.text)]
    c.check("summary has one bullet per accepted/rejected change (6 changes: 3 Alice, 3 Bob)",
            len(bullets) >= 6, [norm(p.text) for p in bullets])
    fmts = {p.list.format for p in bullets}
    c.check("summary list is bulleted", fmts == {"bullet"}, fmts)
    joined = " ".join(norm(p.text).lower() for p in bullets)
    for key in ("must", "lost device", "outside core hours", "five days", "750", "vpn"):
        c.check(f"summary mentions '{key}'", key in joined)
    c.check("summary says accepted and rejected", "accepted" in joined and "rejected" in joined)
    claude = doc.revisions(author="Claude")
    kinds = {r.kind for r in claude}
    c.check("summary is a tracked insertion by Claude", claude and kinds <= {"insertion", "paragraph-insertion",
                                                                             "paragraph insertion"} or
            (claude and "insertion" in " ".join(kinds)), kinds)
    c.check("Chen's text untouched", "a laptop, a headset and a monitor" in " ".join(got) and "EUR 30" in " ".join(got))


if __name__ == "__main__":
    main("w5-selective-review", body)
