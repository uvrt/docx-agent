"""Review: ``word_set_tracking``, ``word_changes`` (list, accept and reject tracked changes)
and ``word_comments``."""

from __future__ import annotations

from collections import Counter
from typing import Any

from ooxml_edit.tools import ACCEPT, REVIEW, Result, ToolError, array, boolean, obj, shared, string

from ..revisions.changes import CHANGE_KINDS
from ._base import author_of, need, one_of, outcome, page, remember, resolve, short, word_tool
from .text import find_range

DOC = string("Document id.")
CURSOR = string("next_cursor of the previous page.", optional=True)
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# -- W4 word_set_tracking ----------------------------------------------------------------------


@word_tool("word_set_tracking", "Tracking mode: while on, every edit is a tracked change by author. word_switch also sets Word's own Track Changes.",
           {"doc": DOC,
            "on": boolean("Track later edits."),
            "author": string("Revisions' author; needed to turn it on.", optional=True),
            "word_switch": boolean("Also set Word's Track Changes to on.", optional=True)},
           batchable=False, group="word_review", mutates=True, track=False)
def word_set_tracking(call: Any, doc: str, on: bool, author: str | None = None,
                      word_switch: bool | None = None) -> Result:
    entry = call.entry
    if on:
        need(author, "author", " to turn tracking on")
        entry.tracking = {"on": True, "author": author}
    else:
        entry.tracking = None
    edits = []
    if word_switch is not None:
        edits.append(call.document.set_word_tracks_changes(word_switch))
    result = outcome(edits, f"Tracking {'on, as ' + author if on else 'off'}"
                     + ("" if word_switch is None else f"; Word's switch {'on' if word_switch else 'off'}"),
                     data={"tracking": bool(on), "author": author if on else None,
                           "word_switch": call.document.word_tracks_changes})
    return result


# -- W13, W14 word_changes --------------------------------------------------------------------------


def _within_paragraphs(call: Any, within: str | None, document: Any = None) -> set[str] | None:
    if not within:
        return None
    from .text import paragraphs_of

    return {p.id for p in paragraphs_of(document if document is not None else call.document, resolve(call, within))}


def change_json(change: Any) -> dict[str, Any]:
    data = {"id": change.id, "kind": change.kind, "author": change.author, "date": change.date}
    if change.structure:
        data["structure"] = change.structure
    if change.old_text:
        data["old_text"] = short(change.old_text, 200)
    if change.text:
        data["text"] = short(change.text, 200)
    data["paragraphs"] = change.paragraph_ids[:5]
    if len(change.revisions) > 1:
        data["records"] = change.revisions
    return data


def revision_json(revision: Any) -> dict[str, Any]:
    data = {"id": revision.id, "kind": revision.kind, "author": revision.author, "date": revision.date,
            "paragraph": revision.paragraph_id}
    if revision.text:
        data["text"] = short(revision.text, 200)
    return data


def _listing(arguments: Any) -> bool:
    return arguments.get("action") in ("list", "preview")


@word_tool("word_changes", "Tracked changes: list them as a reviewer reads them (a replacement is one change, old and new text) or as Word's records; accept, reject or preview them by id, author, kind, within, or all, reporting what remains.",
           {"doc": DOC,
            "action": string("What to do; list changes nothing.", enum=["list", "accept", "reject", "preview"]),
            "ids": array(string(), "Not list: these changes (rev: ids).", optional=True),
            "author": string("Only this author's.", optional=True),
            "kind": string("Only this kind.", enum=list(CHANGE_KINDS), optional=True),
            "within": string("Only in this block or span.", optional=True),
            "all": boolean("Not list: every change.", optional=True),
            "detail": string("list. Default grouped.", enum=["grouped", "records"], optional=True),
            "cursor": CURSOR},
           refs=('ids[]', 'within'), group="word_review", mutates=True, track=False, reads=_listing,
           action_roles={"list": REVIEW, "preview": ACCEPT, "accept": ACCEPT, "reject": ACCEPT}, review_core=True,
           action_fields={"ids": ("accept", "reject", "preview"), "all": ("accept", "reject", "preview")})
def word_changes(call: Any, doc: str, action: str, ids: list[str] | None = None, author: str | None = None,
                 kind: str | None = None, within: str | None = None, all: bool = False,
                 detail: str = "grouped", cursor: str | None = None) -> Result:
    if action == "list":
        if ids or all:
            raise ToolError("invalid_arguments", "list filters by author, kind and within",
                            field="ids" if ids else "all")
        return _list_changes(call, author, kind, within, detail, cursor)
    if detail != "grouped" or cursor is not None:
        raise ToolError("invalid_arguments", "detail and cursor are for list",
                        field="detail" if detail != "grouped" else "cursor")
    if action == "preview":
        return _preview_changes(call, ids, author, kind, within, all)
    return _review_changes(call, action, ids, author, kind, within, all)


def _list_changes(call: Any, author: str | None, kind: str | None, within: str | None, detail: str,
                  cursor: str | None) -> Result:
    document = call.document
    inside = _within_paragraphs(call, within)
    if detail == "records":
        items = [revision_json(r) for r in document.revisions(author=author, within=resolve(call, within))
                 if kind is None or r.kind.startswith(kind) or kind in r.kind]
    else:
        items = [change_json(c) for c in document.changes(author=author, kind=kind)
                 if inside is None or inside & set(c.paragraph_ids)]
    shown, total, next_cursor = page(items, cursor, call.limits.max_list_items)
    result = Result(summary=f"{total} {'record' if detail == 'records' else 'change'}(s)", data=shown,
                    next_cursor=next_cursor)
    result.total = total
    return result


def _recorded_tables(document: Any) -> dict[str, tuple[Any, bytes, Any]]:
    """The tables the document keeps a source record of: each one's element, its XML now
    and its record."""
    from lxml import etree

    kept = document.provenance()
    out = {}
    for key in document._provenance_keys():
        found = document._direct(key) if key.rpartition("/")[2].startswith("t:") else None
        if found is not None and hasattr(found[1], "rows"):
            out[key] = (found[1].element, etree.tostring(found[1].element), kept.get(key))
    return out


def _follow_review(call: Any, action: str, recorded: dict[str, tuple[Any, bytes, Any]]) -> None:
    """A record describes its table with the tracked changes accepted.  Accepting keeps it
    -- under the table's new id when accepting renamed the table (a tracked row at the top
    inserted or deleted).  Rejecting a change in a recorded table may put back values that
    are not the source's, so its record is dropped (a stale address is worse than none)."""
    from lxml import etree

    document = call.document
    for key, (element, before, record) in recorded.items():
        now = next((table.id for story in document.stories for table in story.tables
                    if document._direct(table.id) is not None
                    and document._direct(table.id)[1].element is element), None)
        if action == "reject":
            if now is None or etree.tostring(element) != before:
                call.store_provenance(key, None)
        elif now is not None and now != key and record is not None:
            call.store_provenance(key, None)
            call.store_provenance(now, record)


def _selection(call: Any, document: Any, ids: list[str] | None, author: str | None, kind: str | None,
               within: str | None, all: bool) -> tuple[str, Any]:
    """Which changes an accept, reject or preview settles: ``("ids", record ids)``, ``("all",
    None)`` or ``("filter", (paragraphs, records))`` -- refused as the call that would work."""
    if not (ids or author or kind or within or all):
        raise ToolError("invalid_arguments", "say which changes: ids, author, kind, within or all",
                        valid_options=["ids", "author", "kind", "within", "all"])
    if ids and (author or kind or within or all):
        raise ToolError("invalid_arguments", "give ids, or filters, not both", field="ids")
    if ids:
        # Every id named, resolved before anything is accepted, and accepted (or rejected) in one
        # go: two records sharing one w:id are rev:95 and rev:95#1, and accepting the first
        # renamed the second, so the next id named was refused (round 54, lane A open 4).
        grouped = {c.id: c for c in document.changes()}
        known = {identifier for _, identifier, _ in document._revision_elements()}
        wanted: list[str] = []
        missing: list[str] = []
        for identifier in resolve(call, ids):
            change = grouped.get(identifier)
            members = list(change.revisions) if change is not None else [identifier] if identifier in known else []
            if not members:
                missing.append(identifier)
            wanted += [m for m in members if m not in wanted]
        if missing:
            listed = list(grouped)
            raise ToolError("not_found", f"no change or record {', '.join(repr(m) for m in missing)}: the ids "
                            f"are word_changes list's (changes) or list detail records' (records), as they "
                            f"stand now", field="ids", valid_options=listed[:50],
                            details={"changes": len(listed),
                                     "call": {"tool": "word_changes",
                                              "arguments": {"doc": call.entry.doc_id, "action": "list"}}})
        return "ids", wanted
    if all and not (author or kind or within):
        return "all", None
    # Round 54 (lane M): and the records inside it -- a paragraph whose mark is deleted is
    # named by the one it joins in the final view, so a table's deleted cells' changes
    # were missed by paragraph ids alone.
    records = {identifier for _, identifier, _ in document._select_revisions(None, within=resolve(call, within))} \
        if within else set()
    return "filter", (_within_paragraphs(call, within, document), records)


def _matching(document: Any, author: str | None, kind: str | None, place: tuple) -> list[Any]:
    inside, records = place
    return [c for c in document.changes(author=author, kind=kind)
            if inside is None or inside & set(c.paragraph_ids) or records & set(c.revisions)]


def _settle(document: Any, action: str, how: str, what: Any, author: str | None, kind: str | None) -> list[Any]:
    """Accept or reject what :func:`_selection` chose; the edits made."""
    if how == "ids":
        return [document.accept(what) if action == "accept" else document.reject(what)]
    if how == "all":
        return [document.accept_all() if action == "accept" else document.reject_all()]
    edits = []
    seen: set[tuple] = set()
    while True:
        pending = [c for c in _matching(document, author, kind, what)
                   if (c.kind, c.author, tuple(c.revisions), c.text, c.old_text) not in seen]
        if not pending:
            break
        change = pending[0]
        seen.add((change.kind, change.author, tuple(change.revisions), change.text, change.old_text))
        edits.append(change.accept() if action == "accept" else change.reject())
    return edits


# -- R121: what a review removed, and whose changes it settled -----------------------------------


def _record_marks(document: Any) -> Counter:
    """Every revision record as ``(part, tag, w:id, author, date)``, counted."""
    return Counter((part, element.tag, element.get(_W + "id"), element.get(_W + "author"), element.get(_W + "date"))
                   for part, _, element in document._revision_elements())


def _opened_stamps(document: Any, part: str) -> set[tuple[str | None, str | None]]:
    """The ``(author, date)`` of every revision record ``part`` had when the document was
    opened (kept on the document: what was read never changes)."""
    from ..edit.annotations import REVISION_KINDS
    from ..edit.ids import live_elements

    kept = document.__dict__.setdefault("_opened_stamps", {})
    found = kept.get(part)
    if found is None:
        root = document.package.opened().tree(part)
        found = set() if root is None else {(node.get(_W + "author"), node.get(_W + "date"))
                                            for node in live_elements(root, frozenset(REVISION_KINDS))}
        kept[part] = found
    return found


def _this_session(document: Any, settled: Counter) -> int:
    """How many of the records ``settled`` were written since the document was opened: a
    record whose part is as it was read, or whose author and date a record had then, was
    there before (a record an edit split keeps its author and date); every other one an
    edit of this session wrote (its date the edit's time)."""
    live = document.__dict__.get("_live", document)       # a read copy: the document's
    changed = live.package.changed_parts()
    return sum(count for (part, _, _, author, date), count in settled.items()
               if part in changed and date is not None and (author, date) not in _opened_stamps(live, part))


def _merged_blocks(edits: list[Any]) -> dict:
    """The blocks a review's edits removed, by kind (each edit's ``removed_blocks``)."""
    merged: dict = {}
    for edit in edits:
        for key, value in (getattr(edit, "removed_blocks", None) or {}).items():
            if isinstance(value, dict):
                merged.setdefault(key, {}).update(value)
            else:
                merged.setdefault(key, []).extend(value)
    joins = merged.get("merged", {})
    for gone in joins:                      # merged into one a later step merged in turn
        seen = {gone}
        while joins[gone] in joins and joins[gone] not in seen:
            seen.add(joins[gone])
            joins[gone] = joins[joins[gone]]
    tables = {t["id"] for t in merged.get("tables", [])}
    if "rows" in merged:
        merged["rows"] = [r for r in merged["rows"] if r.get("table") not in tables]
        if not merged["rows"]:
            del merged["rows"]
    return merged


#: How many removed blocks a summary names; the rest are counted (data.removed has them all).
NAMED = 6


def _plural(count: int, word: str) -> str:
    return f"{count} {word}{'' if count == 1 else 's'}"


def _removed_text(blocks: dict) -> str:
    """``removed table t:A (3 rows, 6 paragraphs), paragraph p:B`` -- or ``removed no block``."""
    named = [f"table {t['id']} ({_plural(t['rows'], 'row')}, {_plural(t['paragraphs'], 'paragraph')})"
             for t in blocks.get("tables", [])]
    named += [f"row {r['id']}" + (f" of {r['table']}" if r.get("table") else "") for r in blocks.get("rows", [])]
    named += [f"paragraph {p}" for p in blocks.get("paragraphs", [])]
    named += [f"paragraph {p} (merged into {into})" for p, into in blocks.get("merged", {}).items()]
    named += [f"content control {c}" for c in blocks.get("controls", [])]
    if not named:
        return "removed no block"
    more = len(named) - NAMED
    return "removed " + ", ".join(named[:NAMED]) + (f" and {more} more (data.removed)" if more > 0 else "")


def _session_text(count: int, records: int) -> str:
    if not count:
        return ""
    return f" ({'all' if count == records else count} written this session, since the document was opened)"


def _review_changes(call: Any, action: str, ids: list[str] | None, author: str | None, kind: str | None,
                    within: str | None, all: bool) -> Result:
    document = call.document
    how, what = _selection(call, document, ids, author, kind, within, all)
    recorded = _recorded_tables(document)
    marks = _record_marks(document)
    before = sum(marks.values())
    edits = _settle(document, action, how, what, author, kind)
    count = before - len(document.revisions())
    settled = marks - _record_marks(document)
    session = _this_session(document, settled)
    _follow_review(call, action, recorded)
    left = document.changes()
    blocks = _merged_blocks(edits)
    # Round 56 (R121): the blocks removed named, and -- for a call by author, kind, within or
    # all -- how many of the records were this session's own.
    summary = f"{action.capitalize()}ed {count} revision record(s)"
    if how != "ids":
        summary += _session_text(session, count)
    if blocks or count:
        summary += "; " + _removed_text(blocks)
    data: dict[str, Any] = {"records": count, "this_session": session, "changes_left": len(left),
                            "left_by_author": _by_author(left)}
    if blocks:
        data["removed"] = blocks
    result = outcome(edits, summary, data=data)
    result.changed = [e.id for e in edits if e is not None and e.id]
    return result


def _scratch(document: Any) -> Any:
    """An editable copy of the document as it is now, with its ids (a preview's)."""
    from .. import Document
    from ..edit import ids as _ids

    copy = Document.open(document.to_bytes(), stamping=document.stamping)
    copy._ledger = _ids.copy_ledger(document._ledger)
    copy._aliases.update(document._aliases)
    return copy


def _preview_changes(call: Any, ids: list[str] | None, author: str | None, kind: str | None,
                     within: str | None, all: bool) -> Result:
    """R121: what accept and what reject would settle and remove, on copies: nothing changed."""
    document = call.document
    how, what = _selection(call, document, ids, author, kind, within, all)
    if how == "all":
        changes = document.changes()
    elif how == "ids":
        wanted = set(what)
        changes = [c for c in document.changes() if wanted & set(c.revisions)]
    else:
        changes = _matching(document, author, kind, what)
    data: dict[str, Any] = {}
    parts = []
    for action in ("accept", "reject"):
        copy = _scratch(document)
        if how == "filter":
            # The places as the copy names them: its own paragraphs and records.
            what = _selection(call, copy, ids, author, kind, within, all)[1]
        marks = _record_marks(copy)
        before = sum(marks.values())
        edits = _settle(copy, action, how, what, author, kind)
        count = before - len(copy.revisions())
        settled = marks - _record_marks(copy)
        blocks = _merged_blocks(edits)
        outcome_data = {"records": count, "this_session": _this_session(document, settled),
                        "changes_left": len(copy.changes())}
        if blocks:
            outcome_data["removed"] = blocks
        data[action] = outcome_data
        parts.append(f"{action} would settle {count} record(s)"
                     + _session_text(outcome_data["this_session"], count) + " and " + _removed_text(blocks)
                     .replace("removed no block", "remove no block").replace("removed ", "remove ", 1))
    shown, total, _ = page([change_json(c) for c in changes], None, call.limits.max_list_items)
    data["changes"] = shown
    if total > len(shown):
        data["changes_not_shown"] = {"count": total - len(shown),
                                     "call": {"tool": "word_changes",
                                              "arguments": {"doc": call.entry.doc_id, "action": "list"}}}
    result = Result(summary=f"Preview, nothing changed: {len(changes)} change(s); " + "; ".join(parts), data=data)
    return result


def _by_author(changes: list[Any]) -> dict[str, int]:
    out: dict[str, int] = {}
    for change in changes:
        out[change.author] = out.get(change.author, 0) + 1
    return out


# -- W15 word_comments -------------------------------------------------------------------------


COMMENT_ITEM = obj({
    "target": string("add: range, paragraph, table, row or cc:.", optional=True),
    "find": string("add: or text occurring once.", optional=True),
    "within": string("add: block or span to search for find.", optional=True),
    "comment": string("The comment (c:...), except for add.", optional=True),
    "text": string("add, reply, edit: the text.", optional=True),
    "resolve": boolean("reply: also resolve the thread.", optional=True),
    "ref": string("add, reply: ref name for the new comment.", optional=True),
})


def comment_json(comment: Any) -> dict[str, Any]:
    anchor = comment.anchor
    data = {"id": comment.id, "author": comment.author, "done": comment.done,
            "text": short(comment.text, 300)}
    if anchor is not None:
        data["on"] = {"id": anchor.id, "text": short(anchor.text, 120)}
    if comment.replies:
        data["replies"] = [{"id": r.id, "author": r.author, "text": short(r.text, 200)} for r in comment.replies]
    return data


@word_tool("word_comments", "Comment threads: list, add, reply, resolve, reopen, edit, delete, many at once. Comments are never tracked changes.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "add", "reply", "resolve", "reopen", "edit",
                                                   "delete"]),
            "items": array(COMMENT_ITEM, "The comments to act on (not list).", optional=True),
            "author": string("add, reply: author. Default the tracking author.", optional=True),
            "open_only": boolean("list: only unresolved threads.", optional=True),
            "cursor": CURSOR},
           refs=('items[].target', 'items[].within', 'items[].comment'), group="word_review", mutates=True, track=False,
           action_roles=shared.COMMENT_ACTION_ROLES, call_role=shared.comment_role,
           review_core=True)
def word_comments(call: Any, doc: str, action: str, items: list[dict] | None = None,
                  author: str | None = None, open_only: bool = False, cursor: str | None = None) -> Result:
    document = call.document
    if action == "list":
        threads = [c for c in document.comments(replies=False) if not (open_only and c.done)]
        shown, total, next_cursor = page([comment_json(c) for c in threads], cursor, call.limits.max_list_items)
        result = Result(summary=f"{total} thread(s)", data=shown, next_cursor=next_cursor)
        result.total = total
        return result
    if items is None:
        raise ToolError("invalid_arguments", f"{action} needs items", field="items")
    if not items:
        return Result.nothing_given("changed")
    who = author_of(call, author)
    edits, made = [], []
    for item in items:
        if action == "add":
            text = need(item.get("text"), "items[].text")
            where = (find_range(call, item["find"], item.get("within"))
                     if one_of(item, "find", "target") == "find"
                     else resolve(call, item["target"]))   # a range's id, or a table's or row's (R68)
            edit = document.add_comment(where, text, author=who, date=call.now())
            made.append(edit.id)
            remember(call, item.get("ref"), edit.id)
        else:
            identifier = resolve(call, need(item.get("comment"), "items[].comment"))
            if action == "reply":
                edit = document.reply_to_comment(identifier, need(item.get("text"), "items[].text"),
                                                  author=who, date=call.now())
                made.append(edit.id)
                remember(call, item.get("ref"), edit.id)
                if item.get("resolve"):
                    edits.append(edit)
                    edit = document.resolve_comment(identifier)
            elif action == "resolve":
                edit = document.resolve_comment(identifier)
            elif action == "reopen":
                edit = document.reopen_comment(identifier)
            elif action == "edit":
                edit = document.edit_comment(identifier, need(item.get("text"), "items[].text"))
            else:
                edit = document.delete_comment(identifier)
        edits.append(edit)
    result = outcome(edits, f"{action}: {len(items)} comment(s)", data={"comments": made} if made else None)
    return result


TOOLS = [word_set_tracking, word_changes, word_comments]
