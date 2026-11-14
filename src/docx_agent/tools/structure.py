"""Structure: ``word_move``, ``word_copy_from``, ``word_sections``, ``word_headers_footers``,
``word_fields``, ``word_notes`` and ``word_links``."""

from __future__ import annotations

import re
import time
from typing import Any

from ooxml_edit.tools import REVIEW, Result, ToolError, array, number, obj, string

from ..edit.sections import STARTS
from . import positions
from ._base import (Timing, ahead_seconds, ahead_timed, lay_out_ahead, need, update_fields_ahead, one_of, outcome,
                    remember, resolve, short, word_tool)
from .text import AT, BLOCKS_REF, find_range, name_blocks, named, text_range

DOC = string("Document id.")
CURSOR = string("next_cursor of the previous page.", optional=True)


def _place(call: Any, after: str | None, before: str | None, after_section: str | None = None) -> dict:
    """``after=``/``before=`` for the library, an ``after_section`` heading resolved to the
    last block of its section.  Round 40: ``after`` may say its side as ``at`` does (an
    ``at`` is read as ``after``: :mod:`.positions`) -- ``before:X``, ``after:X``, ``end``
    (after the body's last block)."""
    given = {k: v for k, v in (("after", after), ("before", before), ("after_section", after_section)) if v}
    if len(given) != 1:
        raise ToolError("invalid_arguments", "give exactly one of after, before" +
                        (", after_section" if after_section is not None or "after_section" in given else ""),
                        valid_options=["after", "before", "after_section"])
    if after_section:
        return {"after": call.document.section_blocks(resolve(call, after_section))[-1]}
    if after:
        side, anchor = positions.side_of(after)
        if side == "end":
            ids = [b["id"] for b in call.document.blocks() if b["id"]]
            return {"after": ids[-1]}
        return {side: resolve(call, anchor)}
    return {"before": resolve(call, before)}


# -- W11 word_move -----------------------------------------------------------------------------


@word_tool("word_move", "Move a heading's whole section, or a span of blocks, keeping ids; tracked when tracking is on.",
           {"doc": DOC,
            "section_heading": string("Heading whose section moves.", optional=True),
            "blocks": string("Or a block or span p:A..t:B.", optional=True),
            "after": string("Move after this block.", optional=True),
            "before": string("Move before this block.", optional=True),
            "after_section": string("Move after this heading's section.", optional=True)},
           refs=('section_heading', 'blocks', 'after', 'before', 'after_section'), group="word_structure", mutates=True, exactly_one=[("section_heading", "blocks")],
           # Round 40: at (after:X, before:X) read as after; target is what moves, so no place.
           aliases=positions.aliases("word_move", "sides", ("section_heading", "blocks", "after", "before",
                                                         "after_section", "target", "where")))
def word_move(call: Any, doc: str, section_heading: str | None = None, blocks: str | None = None,
              after: str | None = None, before: str | None = None, after_section: str | None = None) -> Result:
    document = call.document
    moving = document.section_blocks(resolve(call, section_heading)) if section_heading else resolve(call, blocks)
    place = _place(call, after, before, after_section)
    edit = document.move_blocks(moving, **place)
    ids = moving if isinstance(moving, list) else [moving]
    result = outcome([edit], f"Moved {len(ids)} block(s)", data={"moved": ids})
    result.changed = list(dict.fromkeys(result.changed + [edit.renamed.get(i, i) for i in ids]))
    return result


# -- W12 word_copy_from ------------------------------------------------------------------------


@word_tool("word_copy_from", "Copy blocks from an open document, this one too, with styles, lists, pictures, notes, comments; style_map renames source styles.",
           {"doc": DOC,
            "source_doc": string("Open document to copy from, e.g. d2."),
            "range": string("Source block or span p:A..t:B.", optional=True),
            "section_heading": string("Or a source heading: its whole section.", optional=True),
            "at": AT,
            "styles": string("Default use_destination.", enum=["use_destination", "keep_source", "merge"],
                             optional=True),
            "style_map": array(obj({"source": string(),
                                    "destination": string()}),
                               "Source style -> this document's style.", optional=True),
            "unmapped": string("Source styles this document lacks: import them (default) or use body.",
                               enum=["import", "body"], optional=True),
            "ref": BLOCKS_REF},
           group="word_structure", mutates=True, documents=("doc", "source_doc"),
           exactly_one=[("range", "section_heading")],
           aliases=positions.aliases("word_copy_from", "blocks", ("at", "range", "section_heading")))
def word_copy_from(call: Any, doc: str, source_doc: str, at: str, range: str | None = None,
                   section_heading: str | None = None, styles: str = "use_destination",
                   style_map: list[dict] | None = None, unmapped: str = "import", ref: str | None = None) -> Result:
    from .text import _at

    named(call, "word_copy_from", ref, at)

    # source_doc may be doc: copy_blocks then copies from the document as it was before the call.
    source = call.entries[source_doc].document
    ids = source.section_blocks(section_heading) if section_heading else range
    mapping = {m["source"]: m["destination"] for m in style_map} if style_map else None
    edit = call.document.copy_blocks(source, ids, at=_at(call, at), styles=styles, style_map=mapping,
                                     unmapped=unmapped)
    copied = dict(list(edit.copied.items())[:60])
    return name_blocks(call, ref, outcome([edit], f"Copied {len(edit.blocks) or len(edit.copied)} block(s) from "
                                          f"{source_doc}", data={"blocks": edit.blocks, "copied": copied}))


# -- W16 word_sections -------------------------------------------------------------------------


MARGINS = obj({side: number(minimum=0, maximum=1584, optional=True)
               for side in ("top", "right", "bottom", "left")}, "set.", optional=True)


@word_tool("word_sections", "Sections: list, insert or remove a break, or set page setup: orientation, size, margins, columns, start.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "insert_break", "remove_break", "set"]),
            "after": string("insert_break: the block before.", optional=True),
            "before": string("insert_break: the block after.", optional=True),
            "section": string("remove_break, set: section (s:...; s:body the last).", optional=True),
            "start": string("Where the next section starts. Default nextPage.", enum=list(STARTS),
                            optional=True),
            "orientation": string("set.", enum=["portrait", "landscape"], optional=True),
            "page_width": number("set.", minimum=72, maximum=1584, optional=True),
            "page_height": number("set.", minimum=72, maximum=1584, optional=True),
            "margins": MARGINS,
            "columns": number("set: text columns.", minimum=1, maximum=45, optional=True)},
           refs=('after', 'before', 'section'), group="word_structure", mutates=True,
           action_roles={"list": REVIEW},
           aliases=positions.aliases("word_sections", "sides", ("after", "before", "section", "target", "where")))
def word_sections(call: Any, doc: str, action: str, after: str | None = None, before: str | None = None,
                  section: str | None = None, start: str | None = None, orientation: str | None = None,
                  page_width: float | None = None, page_height: float | None = None,
                  margins: dict | None = None, columns: float | None = None) -> Result:
    from .read import _sections

    document = call.document
    if action == "list":
        return Result(summary="Sections", data=_sections(document))
    if action == "insert_break":
        place = _place(call, after, before)
        edit = document.insert_section_break(**place, kind=start or "nextPage")
        return outcome([edit], f"Inserted a section break; the section before it is {edit.id}",
                       data={"section": edit.id})
    section = resolve(call, need(section, "section"))
    if action == "remove_break":
        return outcome([document.remove_section_break(section)], f"Removed the break ending {section}")
    values: dict[str, Any] = {}
    if start:
        values["start"] = start
    if orientation:
        values["orientation"] = orientation
    if page_width is not None:
        values["page_width"] = page_width
    if page_height is not None:
        values["page_height"] = page_height
    for side, value in (margins or {}).items():
        values[f"margin_{side}"] = value
    if columns is not None:
        values["columns"] = int(columns)
    if not values:
        raise ToolError("invalid_arguments", "set needs a property to set",
                        valid_options=["start", "orientation", "page_width", "page_height", "margins", "columns"])
    return outcome([document.set_section(section, **values)], f"Set {', '.join(values)} of {section}",
                   changed=[section])


# -- W17 word_headers_footers ------------------------------------------------------------------

#: What page_x_of_y puts between the text and "Page X of Y".
PAGE_SEPARATOR = " | "


@word_tool("word_headers_footers", "Set a section's header or footer text, with an optional page number; remove it, link it to the previous section's, or unlink. Returns its story name.",
           {"doc": DOC,
            "section": string("Section; s:body is the last."),
            "which": string("Which story.", enum=["header", "footer"]),
            "type": string("Pages it is for. Default default (most).",
                           enum=["default", "first", "even"], optional=True),
            "action": string("What to do.", enum=["set", "remove", "link", "unlink"]),
            "text": string("set: the text; \\n between paragraphs.", optional=True),
            "page_number": string("set: after_text, the bare number; page_x_of_y, 'Page X of Y' "
                                  "(' | ' after text).",
                                  enum=["none", "after_text", "page_x_of_y"],
                                  optional=True)},
           refs=('section',), group="word_structure", mutates=True)
def word_headers_footers(call: Any, doc: str, section: str, which: str, action: str, type: str = "default",
                         text: str | None = None, page_number: str = "none") -> Result:
    document = call.document
    section = resolve(call, section)
    footer = which == "footer"
    if action in ("remove", "link"):
        edit = document.link_to_previous(section, type, footer=footer)
        if not edit.changed:
            # Round 27: no own header (footer) to remove or link is as asked; said so.
            note = f"{section} had no {type} {which} of its own: no change"
            result = outcome([edit], note)
            result.warnings.append(note)
            return result
        return outcome([edit], f"Removed the {type} {which} of {section}" if action == "remove" else
                       f"Linked the {type} {which} of {section} to the previous section's")
    if action == "unlink":
        edit = document.unlink_from_previous(section, type, footer=footer)
        result = outcome([edit], f"Unlinked: {edit.id}", data={"story": edit.id})
        if not edit.changed:
            result.summary = f"{section} had its own {type} {which} already ({edit.id}): no change"
            result.warnings.append(result.summary)
        return result
    text = (text or "").replace("\r\n", "\n")
    lines = text.split("\n")
    own = document.section(section).stories()[which][type]
    edits = []
    if own["own"]:
        story = own["story"]
        paragraphs = document.paragraphs(story)
        edits.append(paragraphs[0].set_text(lines[0]))
        for extra in paragraphs[1:]:
            edits.append(document.delete_block(extra.id))
    else:
        edit = (document.add_footer if footer else document.add_header)(section, type, lines[0])
        edits.append(edit)
        story = edit.id
    last = document.paragraphs(story)[0].id
    for line in lines[1:]:
        edit = document.insert_paragraph(line, after=last)
        edits.append(edit)
        last = edit.id
    if page_number != "none":
        paragraph = document.paragraphs(story)[-1]
        end = len(paragraph.text)
        if page_number == "page_x_of_y":
            # "Page X of Y", after a separator when the text does not end in a space.
            before = PAGE_SEPARATOR if paragraph.text and not paragraph.text[-1].isspace() else ""
            edits.append(document.insert_text(f"{paragraph.id}@{end}", f"{before}Page  of "))
            paragraph = document.paragraphs(story)[-1]
            edits.append(document.insert_page_number(f"{paragraph.id}@{len(paragraph.text)}", "NUMPAGES"))
            edits.append(document.insert_page_number(f"{paragraph.id}@{end + len(before) + 5}", "PAGE"))
        else:
            edits.append(document.insert_page_number(f"{paragraph.id}@{end}", "PAGE"))
    result = outcome(edits, f"Set the {type} {which} of {section}", data={"story": story})
    return result


# -- W18 word_fields ---------------------------------------------------------------------------


_LEVELS = re.compile(r"^([1-9])-([1-9])$")


def field_json(field: Any) -> dict[str, Any]:
    return {"id": field.id, "keyword": field.keyword, "instruction": field.instruction.strip(),
            "result": short(field.result, 160), "paragraph": field.paragraph_id}


#: The fields whose results ``update`` reads from the layout (``Document._compute``).
_LAID_OUT = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "TOC", "STYLEREF"})


def _fields_ahead(call: Any, doc: str, action: str, kind: str | None = None,
                  ids: list[str] | None = None, **_: Any) -> None:
    """``word_fields``' work ahead of its change (round 33, R16c), with the document's lock
    let go so that other calls do not wait for it: an update of page numbers or contents
    done on a copy, its layouts kept (:func:`~._base.update_fields_ahead`); the layout a
    page cross-reference reads (:func:`~._base.lay_out_ahead`)."""
    if action == "update":
        started = time.perf_counter()
        if any(found.keyword in _LAID_OUT for found in call.document._all_fields()):
            from .owner import layout_first

            layout_first(call)      # lane O: the layout its document's checks read, first
            update_fields_ahead(call, ids)
        ahead_timed(time.perf_counter() - started)    # round 45: the update's timing says it
    elif action == "insert_cross_reference" and kind == "page":
        lay_out_ahead(call)


@word_tool("word_fields", "Fields: list; insert a TOC, caption, cross-reference, date or other field; update all results (TOC, page numbers, REF, SEQ) after changes.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "insert_toc", "insert_caption",
                                                   "insert_cross_reference", "insert_date", "insert_field",
                                                   "update"]),
            "after": string("insert_toc, insert_caption: the block before.", optional=True),
            "before": string("insert_toc, insert_caption: the block after.", optional=True),
            "at": string("Other inserts: position or range (p:A@5), or find.", optional=True),
            "find": string("Other inserts: text occurring once; goes after it.", optional=True),
            "levels": string("insert_toc: heading levels. Default 1-3.", optional=True),
            "bookmark": string("insert_cross_reference: the bookmark.", optional=True),
            "kind": string("insert_cross_reference: its text (default) or page.", enum=["text", "page"],
                           optional=True),
            "label": string("insert_caption: Figure (default), Table...", optional=True),
            "text": string("insert_caption: text after the number.", optional=True),
            "instruction": string("insert_field: the instruction, e.g. NUMPAGES.", optional=True),
            "ids": array(string(), "update: only these fields (fld:...).", optional=True)},
           refs=('after', 'before', 'at', 'ids[]'), group="word_structure", mutates=True,
           action_roles={"list": REVIEW}, ahead=_fields_ahead,
           aliases=positions.aliases("word_fields", "mark", ("after", "before", "at"), find=True))
def word_fields(call: Any, doc: str, action: str, after: str | None = None, before: str | None = None,
                at: str | None = None, find: str | None = None, levels: str | None = None,
                bookmark: str | None = None, kind: str = "text", label: str | None = None,
                text: str | None = None, instruction: str | None = None, ids: list[str] | None = None) -> Result:
    document = call.document
    if action == "list":
        return Result(summary="Fields", data=[field_json(f) for f in document.fields()][:100])
    if action == "update":
        # Round 45 (R67): where the time went -- the work ahead (the update done on a copy, its
        # lock let go, with its layouts), the update under the lock, and the checks after.
        timing = Timing()
        ahead = ahead_seconds()
        if ahead is not None:
            timing.add("ahead", ahead)
        call.touch(timing)
        with timing.part("update"):
            edit = document.update_fields(resolve(call, ids) if ids else None, date=call.now().date())
        return outcome([edit], f"Updated fields: {edit.count} result(s) changed",
                       data={"changed_results": edit.count, "timing": timing.parts})
    if action == "insert_toc":
        first, last = 1, 3
        if levels:
            match = _LEVELS.match(levels)
            if not match or int(match.group(1)) > int(match.group(2)):
                raise ToolError("invalid_arguments", "levels is first-last, e.g. 1-3", field="levels")
            first, last = int(match.group(1)), int(match.group(2))
        edit = document.insert_toc(**_place(call, *_beside(after, before, at)), levels=(first, last))
        return outcome([edit], f"Inserted a table of contents {edit.id}")
    if action == "insert_caption":
        edit = document.insert_caption(**_place(call, *_beside(after, before, at)), label=label or "Figure",
                                       text=text or "")
        return outcome([edit], "Inserted a caption")
    if at is None and find is None and (after or before):
        # Round 40: a field after a block goes at its end; before it, at its start.
        at = positions._write("mark", "after" if after else "before", after or before)
        if at is None:
            raise ToolError("invalid_arguments", f"{action} goes at a position: at p:<id>@<offset>, or a range "
                            "(at its end)", field="at")
    where = _position(call, at, find)
    if action == "insert_cross_reference":
        edit = document.insert_cross_reference(where, need(bookmark, "bookmark"), kind=kind)
    elif action == "insert_date":
        edit = document.insert_date(where, date=call.now().date())
    else:
        edit = document.insert_field(where, need(instruction, "instruction"))
    return outcome([edit], f"Inserted field {edit.id}")


def _beside(after: str | None, before: str | None, at: str | None) -> tuple[str | None, str | None]:
    """``word_fields``' blocks (a TOC, a caption) beside a block: ``after``/``before``, or
    (round 40) ``at`` as the placing tools write it, read as ``after`` -- ``before:X`` too."""
    if at is not None and after is None and before is None:
        return at, None
    return after, before


def _position(call: Any, at: str | None, find: str | None) -> Any:
    """A position: an id given, or the end of text found once."""
    key = one_of({"at": at, "find": find}, "at", "find")
    if key == "find":
        found = find_range(call, find, None)
        if not found.single:
            return found
        base, span = found.id.rsplit("@", 1)
        return f"{base}@{span.split(':')[-1]}"
    return resolve(call, at)


# -- W19 word_notes ----------------------------------------------------------------------------


@word_tool("word_notes", "Footnotes and endnotes: list, insert after text found once or at a position, edit, delete, move.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "insert", "edit", "delete", "move"]),
            "kind": string("insert. Default footnote.", enum=["footnote", "endnote"], optional=True),
            "at": string("insert, move: position or range; the mark goes at its end.", optional=True),
            "find": string("insert, move: or text occurring once; the mark goes after it.", optional=True),
            "note": string("edit, delete, move: the note (fn:2, en:1).", optional=True),
            "text": string("insert, edit: the text.", optional=True),
            "ref": string("insert: ref name for the note.", optional=True)},
           refs=('at', 'note'), group="word_structure", mutates=True,
           action_roles={"list": REVIEW},
           aliases=positions.aliases("word_notes", "mark", ("at", "note"), find=True))
def word_notes(call: Any, doc: str, action: str, kind: str = "footnote", at: str | None = None,
               find: str | None = None, note: str | None = None, text: str | None = None,
               ref: str | None = None) -> Result:
    document = call.document
    if action == "list":
        return Result(summary="Notes", data=[{"id": n.id, "kind": n.kind, "text": short(n.text, 200)}
                                             for n in document.notes()][:100])
    if action == "insert":
        where = _position(call, at, find)
        edit = (document.insert_endnote if kind == "endnote" else document.insert_footnote)(where, text or "")
        remember(call, ref, edit.id)
        return outcome([edit], f"Inserted {edit.id}")
    note = resolve(call, need(note, "note"))
    if action == "edit":
        return outcome([document.edit_note(note, need(text, "text"))], f"Edited {note}")
    if action == "delete":
        return outcome([document.delete_note(note)], f"Deleted {note}")
    return outcome([document.move_note(note, _position(call, at, find))], f"Moved {note}")


# -- W20 word_links ----------------------------------------------------------------------------


@word_tool("word_links", "Hyperlinks and bookmarks: list, link text to a URL or bookmark, bookmark text, rename or remove a bookmark.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "add_hyperlink", "add_bookmark", "rename_bookmark",
                                                   "remove_bookmark"]),
            "target": string("add_*: range or paragraph; add_bookmark: or blocks (t:X, p:A..t:B).", optional=True),
            "find": string("add_*: or text occurring once.", optional=True),
            "url": string("add_hyperlink: external address.", optional=True),
            "bookmark": string("Bookmark name; add_hyperlink: link to it instead of url.",
                               optional=True),
            "new_name": string("rename_bookmark: new name.", optional=True)},
           refs=('target',), group="word_structure", mutates=True,
           action_roles={"list": REVIEW})
def word_links(call: Any, doc: str, action: str, target: str | None = None, find: str | None = None,
               url: str | None = None, bookmark: str | None = None, new_name: str | None = None) -> Result:
    document = call.document
    if action == "list":
        return Result(summary="Links and bookmarks", data={
            "bookmarks": [{"name": b.name, "text": short(b.text, 80)} for b in document.bookmarks()][:100],
            "hyperlinks": [{"id": h.id, "text": short(h.text, 80), "url": h.address, "bookmark": h.anchor}
                           for h in document.hyperlinks()][:100]})
    if action in ("rename_bookmark", "remove_bookmark"):
        name = need(bookmark, "bookmark")
        if action == "remove_bookmark":
            return outcome([document.remove_bookmark(name)], f"Removed bookmark {name}")
        return outcome([document.rename_bookmark(name, need(new_name, "new_name"))],
                       f"Renamed bookmark {name} to {new_name}")
    found = one_of({"target": target, "find": find}, "target", "find") == "find"
    if action == "add_bookmark":
        # Round 49 (R91): blocks too -- a table, a span of blocks (p:A..t:B), a heading's range.
        where = find_range(call, find, None) if found else resolve(call, target)
        edit = document.add_bookmark(where, need(bookmark, "bookmark"))
        return outcome([edit], f"Bookmarked as {bookmark}" if edit.changed else f"No change: {edit.warnings[0]}")
    where = find_range(call, find, None) if found else text_range(document, resolve(call, target))
    if (url is None) == (bookmark is None):
        raise ToolError("invalid_arguments", "give exactly one of url, bookmark", valid_options=["url", "bookmark"])
    return outcome([document.add_hyperlink(where, url, anchor=bookmark)], "Made a hyperlink")


TOOLS = [word_move, word_copy_from, word_sections, word_headers_footers, word_fields, word_notes, word_links]
