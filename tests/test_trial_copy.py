"""Copying a section in the destination's styles (ROADMAP.md, "Trial findings", 3), on the
trial's merge-a-section task (``fixtures/generated/trial``: a handbook and a safety bulletin
docx-agent wrote): ``copy_blocks(style_map=..., unmapped="body")``, ``styles="merge"``
keeping headings and lists, ``Styles.remove`` and ``Styles.purge_unused``."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_agent import Document, EditError
from docx_agent.edit.styles import StyleError
from docx_agent.validate import check

TRIAL = Path(__file__).parent / "fixtures" / "generated" / "trial"
SECTION = "p:3B212964"          # the bulletin's "Working at height"
VEHICLE_CHECKS = "p:69E36C9F"   # the handbook's heading the section goes before


def inputs() -> tuple[Document, Document]:
    return Document.open(TRIAL / "field-handbook.docx"), Document.open(TRIAL / "safety-bulletin.docx")


def names(document: Document) -> set[str]:
    return {style.name for style in document.styles}


def section(handbook: Document) -> list:
    paragraphs = handbook.paragraphs()
    start = next(k for k, p in enumerate(paragraphs) if p.text == "Working at height")
    end = next(k for k, p in enumerate(paragraphs) if p.text == "Vehicle checks")
    return paragraphs[start:end]


def test_the_task_copies_the_section_in_the_handbook_styles():
    """The trial's W6: the section after "Site access" with the handbook's Heading 1 and
    "Handbook Body", its lists, picture and footnote -- and none of the bulletin's styles."""
    handbook, bulletin = inputs()
    before = names(handbook)
    result = handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}",
                                  unmapped="body")
    copied = section(handbook)
    assert [(p.style_name, p.list.format if p.list else None) for p in copied] == [
        ("heading 1", None), ("Handbook Body", None),
        ("List Bullet", "bullet"), ("List Bullet", "bullet"), ("List Bullet", "bullet"),
        ("Handbook Body", None),
        ("List Number", "decimal"), ("List Number", "decimal"), ("List Number", "decimal"),
        ("Handbook Body", None), ("Handbook Body", None)]
    assert names(handbook) - before == {"List Number"}            # a built-in, as Word writes it
    assert not any("Bulletin" in name for name in names(handbook))
    notes = handbook.notes("footnote")
    assert [n.text.strip() for n in notes][-1] == "Both incidents involved stepladders on uneven ground."
    assert len(notes) == 3 and result.copied["fn:1"] == "fn:3"   # numbered after the handbook's
    assert [p.alt_text for p in handbook.pictures()] == ["Ladder at a 75 degree angle against a wall"]
    assert check(handbook.package) == []
    handbook.undo()
    assert names(handbook) == before


def test_style_map_names_the_destination_style_by_source_name():
    handbook, bulletin = inputs()
    handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}",
                         style_map={"Bulletin Body": "Handbook Body", "Bulletin Note": "Quote"})
    styles = {p.text[:20]: p.style_name for p in section(handbook)}
    assert styles["Falls from ladders c"] == "Handbook Body" and styles["Figure: the safe lad"] == "Quote"
    assert "Bulletin Body" not in names(handbook)
    with pytest.raises(EditError, match="no paragraph style 'Nope'"):
        handbook.copy_blocks(bulletin, SECTION, style_map={"heading 1": "Nope"})


def test_unmapped_import_is_word_s_default_paste():
    handbook, bulletin = inputs()
    result = handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}")
    assert {"Bulletin Body", "Bulletin Note"} <= names(handbook)
    assert any("imported" in warning for warning in result.warnings)
    with pytest.raises(EditError, match="unmapped is one of"):
        handbook.copy_blocks(bulletin, SECTION, unmapped="drop")


def test_merge_keeps_headings_and_lists():
    """The trial: merge flattened the heading and the list items into Normal."""
    handbook, bulletin = inputs()
    handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}",
                         styles="merge")
    copied = section(handbook)
    assert copied[0].style_name == "heading 1"
    assert [p.style_name for p in copied if p.list] == ["List Bullet"] * 3 + ["List Number"] * 3
    assert [p.list.format for p in copied if p.list] == ["bullet"] * 3 + ["decimal"] * 3
    assert not any("Bulletin" in name for name in names(handbook))
    assert check(handbook.package) == []


def test_remove_refuses_a_style_in_use_unless_given_a_replacement():
    handbook, bulletin = inputs()
    handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}")
    assert handbook.styles.usage("Bulletin Body") == {"content": 3, "styles": 0}
    with pytest.raises(EditError, match="used 3 time"):
        handbook.styles.remove("Bulletin Body")
    result = handbook.styles.remove("Bulletin Body", replacement="Handbook Body")
    assert result.removed == ["style:BulletinBody"] and "Bulletin Body" not in names(handbook)
    assert "Handbook Body" in {p.style_name for p in section(handbook)}
    with pytest.raises(EditError, match="default paragraph style"):
        handbook.styles.remove("Normal")
    with pytest.raises(StyleError):
        handbook.styles.remove("No such style")
    assert check(handbook.package) == []
    handbook.undo()
    assert "Bulletin Body" in names(handbook)


def test_purge_unused_removes_the_custom_styles_nothing_uses():
    handbook, bulletin = inputs()
    assert not handbook.styles.purge_unused().changed            # the handbook uses all of its own
    handbook.copy_blocks(bulletin, bulletin.section_blocks(SECTION), at=f"before:{VEHICLE_CHECKS}")
    for paragraph in section(handbook):
        if paragraph.style_name in ("Bulletin Body", "Bulletin Note"):
            paragraph.style = "Handbook Body"
    result = handbook.styles.purge_unused()
    assert sorted(result.removed) == ["style:BulletinBody", "style:BulletinNote"]
    assert "Heading 2 Char" in names(handbook)                   # linked to a built-in style kept
    assert check(handbook.package) == []
