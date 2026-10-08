"""W11: every content control filled as the brief says, still a content control, the rest of
the template unchanged."""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w11-contract-form/input/services-agreement.docx"
WANT = {"Client": "Halden Group AS", "Term": "24 months", "Monthly fee": "18,400"}


def body(c, out: Path):
    doc, src = Document.open(out), Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    controls = {ctl.alias: ctl for ctl in doc.content_controls()}
    c.check("the five controls are still there", sorted(controls) == sorted(
        ctl.alias for ctl in src.content_controls()), sorted(controls))
    for alias, value in WANT.items():
        got = controls.get(alias)
        c.check(f"{alias}: {value}", got is not None and norm(str(got.value)) == value,
                got and got.value)
    start = controls.get("Start date")
    value = start.value if start is not None else None
    c.check("Start date: 3 November 2026", value == dt.date(2026, 11, 3)
            or str(value)[:10] == "2026-11-03", value)
    shown = next((norm(p.text) for p in doc.paragraphs() if p.text.startswith("Start date")), "")
    c.check("the date shows in the control's format", "3-11-2026" in shown, shown)
    support = controls.get("Support")
    c.check("Support ticked", support is not None and support.value is True,
            support and support.value)
    labels = [norm(p.text).split(":")[0] for p in doc.paragraphs()]
    c.check("the template's text unchanged", labels == [norm(p.text).split(":")[0]
                                                         for p in src.paragraphs()], labels)
    c.check("no tracked changes", not doc.changes())


if __name__ == "__main__":
    main("w11-contract-form", body)
