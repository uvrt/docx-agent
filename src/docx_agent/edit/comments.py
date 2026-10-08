"""Comments: add, reply, resolve and reopen, edit and delete -- in the modern parts, as Word
16.106 writes them (``tools/e3_probe.py``).

A comment is ``w:comment`` in ``comments.xml`` (``w:id``, ``w:author``, ``w:date``,
``w:initials``), its paragraphs in the Comment Text style ("annotation text"), the first
opening with a run in the Comment Reference style ("annotation reference") holding
``w:annotationRef``.  Beside it Word keeps:

* ``commentsExtended.xml``: ``w15:commentEx`` per comment, keyed by the ``w14:paraId`` of
  its last paragraph -- ``w15:done`` (resolved) and ``w15:paraIdParent`` (the comment it
  replies to);
* ``commentsIds.xml``: ``w16cid:commentId`` -- the same paraId and the comment's
  ``w16cid:durableId``, the id Word keeps across saves (it renumbers ``w:id``);
* ``commentsExtensible.xml``: ``w16cex:commentExtensible`` -- the durable id and the UTC
  date;
* ``people.xml``: ``w15:person`` per author, ``w15:presenceInfo`` with provider ``None``
  and the author's name as user id.

In the story the comment covers a range between ``w:commentRangeStart`` and
``w:commentRangeEnd``, followed by a run holding ``w:commentReference`` in the Comment
Reference style at the size of the text it follows (Word writes the size; measured).  A
reply is anchored where its parent is, its markers and reference beside the parent's.

Comments are addressed by their durable id, ``c:<durableId>`` (``c#<w:id>`` for one with
none); comments are not revisions: they are written the same with tracking on or off.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids
from . import inline as _inline
from . import text as _text
from .ranges import TextRange

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_W15 = "{%s}" % W15
_W16CID = "{%s}" % W16CID
_W16CEX = "{%s}" % W16CEX

REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
REL_COMMENTS_EXTENSIBLE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
REL_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

#: Each part: (relationship type, default name, content type, root).
_PARTS = {
    "comments": (REL_COMMENTS, "word/comments.xml", _WML + ".comments+xml",
                 f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14"/>'),
    "extended": (REL_COMMENTS_EXTENDED, "word/commentsExtended.xml", _WML + ".commentsExtended+xml",
                 f'<w15:commentsEx xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"/>'),
    "ids": (REL_COMMENTS_IDS, "word/commentsIds.xml", _WML + ".commentsIds+xml",
            f'<w16cid:commentsIds xmlns:w16cid="{W16CID}" xmlns:mc="{MC}" mc:Ignorable="w16cid"/>'),
    "extensible": (REL_COMMENTS_EXTENSIBLE, "word/commentsExtensible.xml", _WML + ".commentsExtensible+xml",
                   f'<w16cex:commentsExtensible xmlns:w16cex="{W16CEX}" xmlns:mc="{MC}" mc:Ignorable="w16cex"/>'),
    "people": (REL_PEOPLE, "word/people.xml", _WML + ".people+xml",
               f'<w15:people xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"/>'),
}
_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


class CommentOps:
    """Comment editing on :class:`docx_agent.Document` (reading is
    :class:`~docx_agent.edit.annotations.AnnotationOps`')."""

    # -- parts ------------------------------------------------------------------------------

    def _comment_part(self: "Document", key: str, create: bool = False) -> str | None:
        rel_type, default, content_type, root = _PARTS[key]
        main = self.package.document_part()
        found = self.package.related_parts_of_type(main, rel_type)
        if found and self.package.has_part(found[0]):
            return found[0]
        if not create:
            return None
        part = default if not self.package.has_part(default) else \
            self.package.unused_part_name(default.replace(".xml", "{n}.xml"))
        self.package.add_part(part, _DECL + root.encode(), content_type, override=True)
        self.package.add_relationship(main, rel_type, part)
        return part

    def _comment_tree(self: "Document", key: str, create: bool = False) -> Element | None:
        part = self._comment_part(key, create)
        return self.package.tree(part) if part is not None else None

    # -- adding -----------------------------------------------------------------------------

    def add_comment(self: "Document", where: "str | TextRange", text: str, *, author: str | None = None,
                    initials: str | None = None, date=None) -> "EditResult":
        """A comment on a range (a range id or range; an empty range comments on a place),
        by ``author`` (the tracking author, else :data:`~docx_agent.revisions.stamp.DEFAULT_AUTHOR`)
        at ``date`` (now by default).  ``text`` may hold ``\\n`` between paragraphs.  Returns
        the comment's id, ``c:<durableId>``."""
        from .document import EditError, EditResult

        where = where if isinstance(where, TextRange) else self.range(where)
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        first, last = entries[0], entries[-1]
        for entry, offset in ((first, where.start), (last, where.end)):
            if not 0 <= offset <= len(_text.atoms(entry.element)):
                raise EditError(f"{where.id} is outside its paragraphs' text")
        who = self._comment_author(author, initials, date)
        size = self._reference_size(last.element, where.end)
        with self._edit():
            renames = self._prepare_many([(part, list(entries))])
            part, entries = where._entries()
            comment_id = self._next_annotation_id()
            start_marker = make("w:commentRangeStart")
            start_marker.set(_W + "id", str(comment_id))
            end_marker = make("w:commentRangeEnd")
            end_marker.set(_W + "id", str(comment_id))
            reference = self._reference_run(comment_id, size)
            # The start first: placed before the character at the start, it moves no
            # character, so the end (after the last one) lands after it, even on a place.
            start_parent, start_index = _inline.position(entries[0].element, where.start)
            start_parent.insert(start_index, start_marker)
            if where.collapsed:
                start_marker.addnext(end_marker)
            else:
                end_parent, end_index = _inline.position(entries[-1].element, where.end)
                end_parent.insert(end_index, end_marker)
            end_marker.addnext(reference)
            durable = self._write_comment(comment_id, text, who, parent=None)
            self.package.mark_dirty(part)
        return EditResult(f"c:{durable}", created=[f"c:{durable}"], renamed=renames)

    def reply_to_comment(self: "Document", identifier: str, text: str, *, author: str | None = None,
                         initials: str | None = None, date=None) -> "EditResult":
        """A reply in the comment's thread (to its first comment when given a reply),
        anchored where it is, as Word anchors one."""
        from .document import EditResult

        record = self._comment_record(identifier)
        root = self._thread_root(record)
        who = self._comment_author(author, initials, date)
        with self._edit():
            self._prepare_many([])
            record = self._comment_record(root["id"])
            comment_id = self._next_annotation_id()
            parent_w_id = record["w_id"]
            for part in self._parts():
                tree = self.package.tree(part)
                for tag, new in (("commentRangeStart", "commentRangeStart"), ("commentRangeEnd", "commentRangeEnd")):
                    for node in tree.iter(_W + tag):
                        if node.get(_W + "id") == parent_w_id:
                            marker = make(f"w:{new}")
                            marker.set(_W + "id", str(comment_id))
                            last = node
                            following = node.getnext()
                            while following is not None and following.tag == node.tag:
                                last, following = following, following.getnext()
                            last.addnext(marker)
                            self.package.mark_dirty(part)
                for node in list(tree.iter(_W + "commentReference")):
                    if node.get(_W + "id") == parent_w_id:
                        run = node.getparent()
                        size = run.find(f"{_W}rPr/{_W}sz")
                        reference = self._reference_run(comment_id, int(size.get(_W + "val")) if size is not None
                                                        and (size.get(_W + "val") or "").isdigit() else None)
                        last = run
                        following = run.getnext()
                        while following is not None and following.find(_W + "commentReference") is not None:
                            last, following = following, following.getnext()
                        last.addnext(reference)
                        self.package.mark_dirty(part)
            durable = self._write_comment(comment_id, text, who, parent=record)
        return EditResult(f"c:{durable}", created=[f"c:{durable}"])

    def _comment_author(self: "Document", author, initials, date):
        from ..revisions.stamp import Tracking

        tracking = self._active_tracking() or self._tracking
        if author is None and tracking is not None:
            return Tracking.make(tracking.author, date if date is not None else tracking.date,
                                 initials or tracking.initials)
        return Tracking.make(author, date, initials)

    def _reference_size(self: "Document", paragraph: Element, offset: int) -> int | None:
        """The size, in half-points, of the text a comment's reference follows: Word writes
        it on the reference run (measured), so the mark sits at the text's size."""
        from . import effective as _effective

        items = _text.atoms(paragraph)
        run = items[offset - 1].run if 0 < offset <= len(items) else (items[0].run if items else None)
        try:
            effective = _effective.run(self, paragraph, run)
        except Exception:  # pragma: no cover - a paragraph docx2svg cannot resolve
            return None
        return int(round(effective.size * 2))

    def _reference_run(self: "Document", comment_id: int, size: int | None) -> Element:
        run = make("w:r")
        properties = make("w:rPr")
        style = make("w:rStyle")
        style.set(_W + "val", self.styles._ensure("annotation reference", "character"))
        properties.append(style)
        if size is not None:
            for tag in ("w:sz", "w:szCs"):
                node = make(tag)
                node.set(_W + "val", str(size))
                insert_in_order(properties, node)
        run.append(properties)
        reference = make("w:commentReference")
        reference.set(_W + "id", str(comment_id))
        run.append(reference)
        return run

    def _write_comment(self: "Document", comment_id: int, text: str, who, parent: dict | None) -> str:
        """The comment in ``comments.xml`` and its entries in the modern parts; returns its
        durable id."""
        comments = self._comment_tree("comments", create=True)
        part = self._comment_part("comments")
        node = make("w:comment")
        node.set(_W + "id", str(comment_id))
        node.set(_W + "author", who.author)
        node.set(_W + "date", who.date)
        node.set(_W + "initials", who.initials or "")
        comments.append(node)
        used = self._used()
        last_para = self._comment_paragraphs(node, text, used, part)
        _ids.ensure_w14(comments)
        extended = self._comment_tree("extended", create=True)
        entry = etree.SubElement(extended, _W15 + "commentEx")
        entry.set(_W15 + "paraId", last_para)
        if parent is not None:
            entry.set(_W15 + "paraIdParent", _last_para_id(parent["element"]))
        entry.set(_W15 + "done", "0")
        durable = self._new_durable_id()
        ids = self._comment_tree("ids", create=True)
        item = etree.SubElement(ids, _W16CID + "commentId")
        item.set(_W16CID + "paraId", last_para)
        item.set(_W16CID + "durableId", durable)
        extensible = self._comment_tree("extensible", create=True)
        item = etree.SubElement(extensible, _W16CEX + "commentExtensible")
        item.set(_W16CEX + "durableId", durable)
        item.set(_W16CEX + "dateUtc", who.date)
        self._ensure_person(who.author)
        self._complete_comment_parts()
        for key in _PARTS:
            found = self._comment_part(key)
            if found is not None:
                self.package.mark_dirty(found)
        return durable

    def _comment_paragraphs(self: "Document", node: Element, text: str, used: set[int], part: str,
                            last_id: str | None = None) -> str:
        """Write ``text``'s paragraphs into a comment; returns the last one's paraId (given
        ``last_id``, the last paragraph takes it)."""
        _text.check_text(text.replace("\n", " "))
        style = self.styles._ensure("annotation text", "paragraph")
        reference_style = self.styles._ensure("annotation reference", "character")
        lines = text.split("\n")
        para_id = None
        for k, line in enumerate(lines):
            paragraph = make("w:p")
            para_id, text_id = _ids.generate(part + "\0comment", used, 2)
            if k == len(lines) - 1 and last_id is not None:
                para_id = last_id
            _ids.stamp(paragraph, para_id, text_id)
            properties = make("w:pPr")
            pstyle = make("w:pStyle")
            pstyle.set(_W + "val", style)
            properties.append(pstyle)
            paragraph.append(properties)
            if k == 0:
                mark_run = make("w:r")
                run_properties = make("w:rPr")
                rstyle = make("w:rStyle")
                rstyle.set(_W + "val", reference_style)
                run_properties.append(rstyle)
                mark_run.append(run_properties)
                mark_run.append(make("w:annotationRef"))
                paragraph.append(mark_run)
            if line:
                paragraph.append(_text.make_run(line, None))
            node.append(paragraph)
        return para_id

    def _complete_comment_parts(self: "Document") -> None:
        """Give every comment its entries in the modern parts -- a comment written by
        another tool may have none -- as Word gives them on saving: ``w15:commentEx``
        (not done), a durable id, and its extensible entry."""
        comments = self._comment_tree("comments")
        if comments is None:
            return
        extended = self._comment_tree("extended", create=True)
        ids = self._comment_tree("ids", create=True)
        extensible = self._comment_tree("extensible", create=True)
        have_ex = {(n.get(_W15 + "paraId") or "").upper() for n in extended.findall(_W15 + "commentEx")}
        have_ids = {(n.get(_W16CID + "paraId") or "").upper(): n.get(_W16CID + "durableId")
                    for n in ids.findall(_W16CID + "commentId")}
        have_cex = {(n.get(_W16CEX + "durableId") or "").upper() for n in extensible}
        for comment in comments.findall(_W + "comment"):
            para = _last_para_id(comment)
            if not para:
                continue
            if para.upper() not in have_ex:
                entry = etree.SubElement(extended, _W15 + "commentEx")
                entry.set(_W15 + "paraId", para)
                entry.set(_W15 + "done", "0")
            durable = have_ids.get(para.upper())
            if durable is None:
                durable = self._new_durable_id()
                item = etree.SubElement(ids, _W16CID + "commentId")
                item.set(_W16CID + "paraId", para)
                item.set(_W16CID + "durableId", durable)
                have_ids[para.upper()] = durable
            if durable.upper() not in have_cex and comment.get(_W + "date"):
                item = etree.SubElement(extensible, _W16CEX + "commentExtensible")
                item.set(_W16CEX + "durableId", durable)
                item.set(_W16CEX + "dateUtc", comment.get(_W + "date"))
                have_cex.add(durable.upper())

    # -- people ------------------------------------------------------------------------------

    def _ensure_person(self: "Document", author: str) -> None:
        """List ``author`` in ``people.xml``, as Word lists everyone who made a comment or a
        revision (measured: a document with only tracked changes gets one too)."""
        people = self._comment_tree("people", create=True)
        if any(person.get(_W15 + "author") == author for person in people.findall(_W15 + "person")):
            return
        person = etree.SubElement(people, _W15 + "person")
        person.set(_W15 + "author", author)
        presence = etree.SubElement(person, _W15 + "presenceInfo")
        presence.set(_W15 + "providerId", "None")
        presence.set(_W15 + "userId", author)
        self.package.mark_dirty(self._comment_part("people"))

    def _authors_in_use(self: "Document") -> set[str]:
        from .annotations import REVISION_KINDS

        out: set[str] = set()
        for part in self._parts():
            for node in self.package.tree(part).iter():
                if isinstance(node.tag, str) and (node.tag in REVISION_KINDS or node.tag == _W + "comment"):
                    author = node.get(_W + "author")
                    if author is not None:
                        out.add(author)
        return out

    def _tidy_people(self: "Document", before: set[str]) -> None:
        """Drop the people an edit left with no revision or comment (of those it found with
        one), and ``people.xml`` once nobody is left -- as Word's save does (measured:
        accepting every change leaves no ``people.xml``)."""
        people = self._comment_tree("people")
        if people is None:
            return
        gone = before - self._authors_in_use()
        changed = False
        for person in people.findall(_W15 + "person"):
            if person.get(_W15 + "author") in gone:
                remove(person)
                changed = True
        if changed:
            self.package.mark_dirty(self._comment_part("people"))
        if not len(people.findall(_W15 + "person")):
            self._drop_comment_part("people")

    def _drop_comment_part(self: "Document", key: str) -> None:
        part = self._comment_part(key)
        if part is None:
            return
        main = self.package.document_part()
        for rid, relationship in self.package.relationships(main).items():
            if relationship.target_part == part:
                self.package.release(main, [rid])
                self.package.mark_dirty(main)
                break

    def _new_durable_id(self: "Document") -> str:
        ids = self._comment_tree("ids")
        used: set[int] = set()
        if ids is not None:
            for item in ids:
                try:
                    used.add(int(item.get(_W16CID + "durableId") or "", 16))
                except ValueError:
                    pass
        seed = f"durable\0{len(used)}"
        return _ids.generate(seed, used, 1)[0]

    # -- changing ---------------------------------------------------------------------------

    def _comment_record(self: "Document", identifier: str) -> dict:
        for record in self._comment_records():
            if record["id"] == identifier:
                return record
        raise KeyError(f"no comment {identifier!r}")

    def _thread_root(self: "Document", record: dict) -> dict:
        records = {r["id"]: r for r in self._comment_records()}
        seen = set()
        while record.get("parent") and record["parent"] in records and record["id"] not in seen:
            seen.add(record["id"])
            record = records[record["parent"]]
        return record

    def _set_done(self: "Document", identifier: str, done: bool) -> "EditResult":
        from .document import EditError, EditResult

        record = self._thread_root(self._comment_record(identifier))
        para = _last_para_id(record["element"])
        if not para:
            raise EditError(f"{identifier} has no paragraph id to mark resolved by")
        with self._edit():
            extended = self._comment_tree("extended", create=True)
            entry = next((e for e in extended.findall(_W15 + "commentEx")
                          if (e.get(_W15 + "paraId") or "").upper() == para.upper()), None)
            if entry is None:
                entry = etree.SubElement(extended, _W15 + "commentEx")
                entry.set(_W15 + "paraId", para)
            if entry.get(_W15 + "done") == ("1" if done else "0"):
                return EditResult(record["id"], changed=False)
            entry.set(_W15 + "done", "1" if done else "0")
            self.package.mark_dirty(self._comment_part("extended"))
        return EditResult(record["id"])

    def resolve_comment(self: "Document", identifier: str) -> "EditResult":
        """Mark a comment's thread resolved (``w15:done="1"`` on its first comment)."""
        return self._set_done(identifier, True)

    def reopen_comment(self: "Document", identifier: str) -> "EditResult":
        """Mark a resolved comment's thread open again (``w15:done="0"``).  One undo step.
        ``doc.reopen_comment("c:76FCC91F")``."""
        return self._set_done(identifier, False)

    def edit_comment(self: "Document", identifier: str, text: str) -> "EditResult":
        """Replace a comment's text; its paragraphs are written anew, the last keeping its
        paraId (which the modern parts key the comment by)."""
        from .document import EditResult

        record = self._comment_record(identifier)
        with self._edit():
            record = self._comment_record(identifier)
            node = record["element"]
            last = _last_para_id(node)
            for paragraph in node.findall(_W + "p"):
                remove(paragraph)
            used = self._used()
            self._comment_paragraphs(node, text, used, record["part"], last_id=last)
            self.package.mark_dirty(record["part"])
        return EditResult(identifier)

    def delete_comment(self: "Document", identifier: str) -> "EditResult":
        """Delete a comment with its replies: the comments, their range markers and
        reference runs, and their entries in the modern parts."""
        from .document import EditResult

        record = self._comment_record(identifier)
        with self._edit():
            removed = self._delete_comment_now(record["id"])
        return EditResult(None, removed=removed)

    def _delete_comment_now(self: "Document", identifier: str) -> list[str]:
        before = self._authors_in_use()
        removed = self._delete_comments(identifier)
        comments = self._comment_tree("comments")
        if comments is not None and comments.find(_W + "comment") is None:
            # No comment left: no comment parts, as Word writes it.
            for key in ("comments", "extended", "ids", "extensible"):
                self._drop_comment_part(key)
        self._tidy_people(before)
        return removed

    def _delete_comments(self: "Document", identifier: str) -> list[str]:
        records = self._comment_records()
        by_id = {r["id"]: r for r in records}
        if identifier not in by_id:
            return []
        doomed = [identifier]
        changed = True
        while changed:
            changed = False
            for record in records:
                if record["parent"] in doomed and record["id"] not in doomed:
                    doomed.append(record["id"])
                    changed = True
        w_ids = {by_id[d]["w_id"] for d in doomed}
        paras = {_last_para_id(by_id[d]["element"]).upper() for d in doomed}
        for part in self._parts():
            tree = self.package.tree(part)
            touched = False
            for node in list(tree.iter(_W + "commentRangeStart", _W + "commentRangeEnd", _W + "commentReference")):
                if node.get(_W + "id") in w_ids:
                    parent = node.getparent()
                    remove(node)
                    touched = True
                    if parent is not None and parent.tag == _W + "r" and not any(
                            isinstance(c.tag, str) and c.tag != _W + "rPr" for c in parent):
                        remove(parent)
            if touched:
                self.package.mark_dirty(part)
        for d in doomed:
            remove(by_id[d]["element"])
        self.package.mark_dirty(by_id[identifier]["part"])
        for key, attribute in (("extended", _W15 + "paraId"), ("ids", _W16CID + "paraId")):
            tree = self._comment_tree(key)
            if tree is None:
                continue
            for item in list(tree):
                if (item.get(attribute) or "").upper() in paras:
                    remove(item)
            self.package.mark_dirty(self._comment_part(key))
        durables = {d.partition(":")[2] for d in doomed if d.startswith("c:")}
        tree = self._comment_tree("extensible")
        if tree is not None:
            for item in list(tree):
                if item.get(_W16CEX + "durableId") in durables:
                    remove(item)
            self.package.mark_dirty(self._comment_part("extensible"))
        return doomed

    def _describing_comments(self: "Document"):
        """``(comment id, table)`` for every comment docx-agent wrote to describe a
        structural table change (its text opens with
        :data:`~docx_agent.revisions.review.DESCRIBES`), with the table it is anchored in."""
        from ..revisions.review import DESCRIBES

        out = []
        for record in self._comment_records():
            node = record["element"]
            text = "".join(t.text or "" for t in node.iter(_W + "t"))
            if not text.startswith(DESCRIBES):
                continue
            table = None
            for part in self._parts():
                for marker in self.package.tree(part).iter(_W + "commentRangeStart"):
                    if marker.get(_W + "id") == record["w_id"]:
                        parent = marker.getparent()
                        while parent is not None and parent.tag != _W + "tbl":
                            parent = parent.getparent()
                        table = parent
            out.append((record["id"], table))
        return out


def _last_para_id(comment: Element) -> str:
    paragraphs = comment.findall(".//" + _W + "p")
    return (paragraphs[-1].get(_ids.PARA_ID) or "") if paragraphs else ""
