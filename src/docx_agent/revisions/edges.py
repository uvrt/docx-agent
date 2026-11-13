"""What Word's Accept All makes of a deleted paragraph mark at a table's edge (round 54,
lane R): the reading docx2svg 0.11.13 draws as the final view (its ROADMAP, "Revisions --
measured", R.3; measured by lane J with Word 16.106's Accept All, ``feedback-round54/probe-J``),
shared by ``word_read``'s final view, the tools' positions and ``check``.

* A paragraph whose mark is deleted just before a table joins the first paragraph of the
  table's first cell (a nested table's, where the cell begins with one), its text first,
  under the cell paragraph's properties -- a paragraph ending a section stays.
* A cell whose last paragraph's mark is deleted takes the next cell of its row in: its last
  paragraph joined to the next cell's first, the next cell's other blocks after it, the next
  cell gone (its columns too, where no other row covers them); a row's first cell's other
  blocks leave the table before that row (the table split there in a later row).
* A cell with no next cell in its row (a row's last cell, the table's last cell) whose last
  mark is deleted is removed **with its text**: Word's Accept All draws neither (measured;
  lane R found that the document Word then saves still holds the text, split off after the
  row in a broken table -- a zero-span cell -- which Word draws once reopened).

Only the final view: ``w:cellDel`` cells are passed over first (:func:`cell_deleted`)."""

from __future__ import annotations

from ..oxml.xml import Element

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_GONE = frozenset({_W + "del", _W + "moveFrom"})


def mark_removed(paragraph: Element) -> bool:
    """Whether a paragraph's mark is deleted (or moved away)."""
    mark = paragraph.find(f"{_W}pPr/{_W}rPr")
    return mark is not None and any(child.tag in _GONE for child in mark)


def last_mark_removed(container: Element) -> bool:
    """Whether a cell's last block is a paragraph whose mark is deleted (a block-level content
    control's or ``w:customXml``'s content seen through)."""
    for node in reversed([n for n in container if isinstance(n.tag, str)]):
        if node.tag == _W + "p":
            return mark_removed(node)
        if node.tag == _W + "tbl":
            return False
        if node.tag in (_W + "sdt", _W + "customXml"):
            content = node.find(_W + "sdtContent") if node.tag == _W + "sdt" else node
            if content is not None and any(n.tag in (_W + "p", _W + "tbl", _W + "sdt", _W + "customXml")
                                           for n in content if isinstance(n.tag, str)):
                return last_mark_removed(content)
    return False


def cell_ends(cells: list[Element]) -> tuple[list[list[int]], set[int]]:
    """A row's shown cells (``w:cellDel`` ones left out) as Accept All leaves them: groups of
    indexes, the first of each the cell kept and the others taken into it, and the indexes of
    the cells removed (a run of cells ending the row, each ending in a deleted mark, removed
    with their text; a cell taken in is among them too)."""
    groups: list[list[int]] = []
    gone: set[int] = set()
    n = 0
    while n < len(cells):
        group = [n]
        while last_mark_removed(cells[group[-1]]) and group[-1] + 1 < len(cells):
            group.append(group[-1] + 1)
        if last_mark_removed(cells[group[-1]]):
            gone.update(group)
        else:
            groups.append(group)
            gone.update(group[1:])
        n = group[-1] + 1
    return groups, gone


def cell_deleted(cell: Element) -> bool:
    return cell.find(f"{_W}tcPr/{_W}cellDel") is not None


def lost_cells(cells: list[Element]) -> list[int]:
    """The indexes (into ``cells``, a row's shown cells) of the cells Accept All removes with
    their text: :func:`cell_ends`' removed cells that are not taken into a cell kept."""
    groups, gone = cell_ends(cells)
    taken = {n for group in groups for n in group[1:]}
    return sorted(n for n in gone if n not in taken)
