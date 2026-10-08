"""Review: ``word_set_tracking``, ``word_changes`` (list, accept and reject tracked changes)
and ``word_comments``."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, obj, string

from ..revisions.changes import CHANGE_KINDS
from ._base import author_of, need, one_of, outcome, page, remember, resolve, short, word_tool
from .text import find_range, text_range

DOC = string("Document id.")
CURSOR = string("next_cursor of the previous page.", optional=True)


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


def _within_paragraphs(call: Any, within: str | None) -> set[str] | None:
    if not within:
        return None
    from .text import paragraphs_of

    return {p.id for p in paragraphs_of(call.document, resolve(call, within))}


def change_json(change: Any) -> dict[str, Any]:
    data = {"id": change.id, "kind": change.kind, "author": change.author, "date": change.date}
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
    return arguments.get("action") == "list"


@word_tool("word_changes", "Tracked changes: list them as a reviewer reads them (a replacement is one change, old and new text) or as Word's records; accept or reject them by id, author, kind, within, or all, reporting what remains.",
           {"doc": DOC,
            "action": string("What to do; list changes nothing.", enum=["list", "accept", "reject"]),
            "ids": array(string(), "accept, reject: these changes (rev: ids).", optional=True),
            "author": string("Only this author's.", optional=True),
            "kind": string("Only this kind.", enum=list(CHANGE_KINDS), optional=True),
            "within": string("Only in this block or span.", optional=True),
            "all": boolean("accept, reject: every change.", optional=True),
            "detail": string("list. Default grouped.", enum=["grouped", "records"], optional=True),
            "cursor": CURSOR},
           refs=('ids[]', 'within'), group="word_review", mutates=True, track=False, reads=_listing)
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


def _review_changes(call: Any, action: str, ids: list[str] | None, author: str | None, kind: str | None,
                    within: str | None, all: bool) -> Result:
    document = call.document
    if not (ids or author or kind or within or all):
        raise ToolError("invalid_arguments", "say which changes: ids, author, kind, within or all",
                        valid_options=["ids", "author", "kind", "within", "all"])
    if ids and (author or kind or within or all):
        raise ToolError("invalid_arguments", "give ids, or filters, not both", field="ids")
    edits = []
    count = 0
    if ids:
        grouped = {c.id: c for c in document.changes()}
        for identifier in resolve(call, ids):
            change = grouped.get(identifier)
            before = len(document.revisions())
            if change is not None:
                edits.append(change.accept() if action == "accept" else change.reject())
            else:
                edits.append(document.accept(identifier) if action == "accept" else document.reject(identifier))
            count += before - len(document.revisions())
    elif all and not (author or kind or within):
        before = len(document.revisions())
        edits.append(document.accept_all() if action == "accept" else document.reject_all())
        count = before - len(document.revisions())
    else:
        before = len(document.revisions())
        inside = _within_paragraphs(call, within)

        def matching() -> list[Any]:
            return [c for c in document.changes(author=author, kind=kind)
                    if inside is None or inside & set(c.paragraph_ids)]

        seen: set[tuple] = set()
        while True:
            pending = [c for c in matching() if (c.kind, c.author, tuple(c.revisions), c.text, c.old_text) not in seen]
            if not pending:
                break
            change = pending[0]
            seen.add((change.kind, change.author, tuple(change.revisions), change.text, change.old_text))
            edits.append(change.accept() if action == "accept" else change.reject())
        count = before - len(document.revisions())
    left = document.changes()
    result = outcome(edits, f"{action.capitalize()}ed {count} revision record(s)",
                     data={"records": count, "changes_left": len(left),
                           "left_by_author": _by_author(left)})
    result.changed = [e.id for e in edits if e is not None and e.id]
    return result


def _by_author(changes: list[Any]) -> dict[str, int]:
    out: dict[str, int] = {}
    for change in changes:
        out[change.author] = out.get(change.author, 0) + 1
    return out


# -- W15 word_comments -------------------------------------------------------------------------


COMMENT_ITEM = obj({
    "target": string("add: range or paragraph to attach to.", optional=True),
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
           refs=('items[].target', 'items[].within', 'items[].comment'), group="word_review", mutates=True, track=False)
def word_comments(call: Any, doc: str, action: str, items: list[dict] | None = None,
                  author: str | None = None, open_only: bool = False, cursor: str | None = None) -> Result:
    document = call.document
    if action == "list":
        threads = [c for c in document.comments(replies=False) if not (open_only and c.done)]
        shown, total, next_cursor = page([comment_json(c) for c in threads], cursor, call.limits.max_list_items)
        result = Result(summary=f"{total} thread(s)", data=shown, next_cursor=next_cursor)
        result.total = total
        return result
    if not items:
        raise ToolError("invalid_arguments", f"{action} needs items", field="items")
    who = author_of(call, author)
    edits, made = [], []
    for item in items:
        if action == "add":
            text = need(item.get("text"), "items[].text")
            where = (find_range(call, item["find"], item.get("within"))
                     if one_of(item, "find", "target") == "find"
                     else text_range(document, resolve(call, item["target"])))
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
