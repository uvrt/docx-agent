"""TEMPORARY (s3): this repository's copy of the Office theme is ooxml-common's.

Run once on the pull request that reads ``ooxml_common.drawingml.theme`` instead; removed
with the copy.
"""

from ooxml_common.drawingml import theme

from docx_agent.edit import blank


def test_the_theme_is_the_shared_one():
    assert blank.THEME_COLORS == theme.OFFICE_THEME_COLORS
    assert (blank.MAJOR_FONT, blank.MINOR_FONT) == (theme.OFFICE_MAJOR_FONT,
                                                    theme.OFFICE_MINOR_FONT)
    assert blank._APTOS_PANOSE == theme.APTOS_PANOSE
    assert blank.SCRIPT_FONTS == theme.WORD_SCRIPT_FONTS
