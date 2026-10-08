"""Do what each reviewer's comment asks, as tracked changes by "Claude"; reply to each
comment saying what changed, and resolve it.

    python examples/review_comments.py agreement-summary.docx reviewed.docx

The end-to-end pilot's task (ROADMAP.md, "Usability (end-to-end pilot)"), written again
from the "Common tasks" (docs/common-tasks.md) alone.  Each comment's anchor says where the change goes;
what the change is was read from the comments first (``to_markdown(view="markup")``).
"""

import sys

from docx_agent import Document

AUTHOR = "Claude"

# A phrase of the comment's text -> (the edit, on the text the comment is attached to; the reply).
EDITS = {
    "36 months": (lambda doc, at: at.replace("36 months"),
                  "Changed the term from 24 to 36 months."),
    "excludes VAT": (lambda doc, at: at.insert_after(", excluding VAT"),
                     "Added that the fee excludes VAT."),
    # Attached to "at most 3% per year": only "3%" changes, so the tracked change is just that.
    "2.5%": (lambda doc, at: doc.anchor("3%", within=at.paragraph_ids()[0]).replace("2.5%"),
             "Changed the indexation cap from 3% to 2.5%."),
    "five working days": (lambda doc, at: at.replace("five working days"),
                          "Changed the notice period from 48 hours to five working days."),
    "procurement@example.com": (lambda doc, at: at.insert_after(" (procurement@example.com)"),
                                "Added the shared mailbox procurement@example.com."),
}


def review(source, target) -> Document:
    doc = Document.open(source)
    with doc.tracking(author=AUTHOR):
        for comment in doc.comments(replies=False):
            phrase = next(phrase for phrase in EDITS if phrase in comment.text)
            edit, note = EDITS[phrase]
            edit(doc, comment.anchor)
            comment.reply(note, author=AUTHOR)
            comment.resolve()
    doc.save(target, validate=True)            # refuses to write if validate() finds a problem
    return doc


if __name__ == "__main__":
    review(sys.argv[1], sys.argv[2])
    print(Document.open(sys.argv[2]).to_markdown(view="markup"))
