"""The language of an existing document (round 35, R24): :meth:`Document.set_language`.

**Measured** in Word 16 for Mac (``tools/language_probe.py``; the facts in
``tests/observations/language-word.json``):

* *What decides the proofing language.*  A run's own ``w:lang``, else its styles', else
  ``w:docDefaults``' (``docx_agent.edit.typography``'s *Language*, measured there):
  ``w:themeFontLang`` is not it.  A document opened with only ``w:docDefaults`` changed is
  proofed in that language wherever no style or run says otherwise.
* *What Word writes.*  A new document in a language carries it in ``w:docDefaults``'
  ``w:lang`` and in the settings' ``w:themeFontLang`` (``blank.py``, measured).  Setting the
  language of an open document's Normal style writes ``w:lang`` on the Normal style alone;
  setting the language of all the text writes ``w:lang`` on every run and paragraph mark;
  setting the Normal template's language writes the template (Normal.dotm: its Normal
  style) and leaves the open document as it was.  None of them rewrites ``w:docDefaults``,
  ``w:themeFontLang`` or another style's own ``w:lang`` (Heading 1 in en-GB, a German body
  style, an Italian character style all stay).
* *Tags.*  Word writes a language in one of three places of ``w:lang``: ``w:val`` for the
  ordinary (Latin, Cyrillic, Greek...) languages, ``w:eastAsia`` for Chinese, Japanese,
  Korean and Yi, ``w:bidi`` for the complex scripts (Arabic, Hebrew, the Indic languages,
  Thai...) -- :data:`WORD_TAGS`, every language of Word's dictionary but one (Scottish
  Gaelic, which it writes as nothing).  It reads a tag in any case and with ``_`` for
  ``-``, and a script it does not write (``fr-Latn-BE``) as the tag without it; a bare
  language (``fr``, ``nl``) or an unknown region (``fr-XX``) as a language with no region
  its dictionary names; an unknown language (``xx-YY``) as none at all -- the text is
  proofed in Word's own editing language.

**Writing.**  :meth:`~LanguageOps.set_language` writes what a new document in that language
has: ``w:docDefaults``' ``w:lang`` and ``w:themeFontLang``, in the place Word writes the
language, the other two left as they are.  Styles that state the language that was the
default -- a template restating its own default, which would keep their paragraphs in the
old language -- are given the new one too (``styles="matching"``, the default); a style in
another language is a deliberate exception (an English quotation style, code) and stays.
Runs keep their own ``w:lang`` unless asked (``runs=True``): a run's own tag is where an
author marked a passage as another language, and Word's whole-text language command is
what writes one on every run -- which then outweighs every later default.
"""

from __future__ import annotations

import difflib
from typing import TYPE_CHECKING, Any, Iterable

from ..oxml.xml import Element, insert_in_order, make
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: The tags Word writes, by the ``w:lang`` attribute it writes each in (measured: every
#: language of Word's dictionary set on text and saved).
WORD_TAGS: dict[str, frozenset[str]] = {
    "val": frozenset("""
        af-ZA am-ET az-Cyrl-AZ az-Latn-AZ be-BY bg-BG bin-NG ca-ES chr-Cher-US cs-CZ cy-GB da-DK de-AT de-CH
        de-DE de-LI de-LU el-GR en-029 en-AU en-BZ en-CA en-GB en-ID en-IE en-JM en-NZ en-PH en-TT en-US
        en-ZA en-ZW es-AR es-BO es-CL es-CO es-CR es-DO es-EC es-ES es-ES_tradnl es-GT es-HN es-MX es-NI
        es-PA es-PE es-PR es-PY es-SV es-UY es-VE et-EE eu-ES ff-NG fi-FI fil-PH fo-FO fr-029 fr-BE fr-CA
        fr-CD fr-CH fr-CI fr-CM fr-FR fr-HT fr-LU fr-MA fr-MC fr-ML fr-RE fr-SN fy-NL ga-IE gl-ES gn-PY
        ha-Latn-NG haw-US hr-HR hsb-DE hu-HU hy-AM ibb-NG id-ID ig-NG is-IS it-CH it-IT iu-Cans-CA ka-GE
        kk-KZ kr-NG ky-KG la-Latn lt-LT lv-LV mk-MK mn-MN ms-BN ms-MY mt-MT nb-NO nl-BE nl-NL nn-NO om-ET
        pl-PL pt-BR pt-PT rm-CH ro-MD ro-RO ru-MD ru-RU se-NO sk-SK sl-SI so-SO sq-AL sr-Cyrl-CS sr-Latn-CS
        st-ZA sv-FI sv-SE sw-KE tg-Cyrl-TJ ti-ER ti-ET tk-TM tn-ZA tr-TR ts-ZA tt-RU tzm-Latn-DZ uk-UA
        uz-Cyrl-UZ uz-Latn-UZ ve-ZA vi-VN xh-ZA yo-NG zu-ZA""".split()),
    "eastAsia": frozenset("ii-CN ja-JP ko-KR zh-CN zh-HK zh-MO zh-SG zh-TW".split()),
    "bidi": frozenset("""
        ar-AE ar-BH ar-DZ ar-EG ar-IQ ar-JO ar-KW ar-LB ar-LY ar-MA ar-OM ar-QA ar-SA ar-SY ar-TN ar-YE
        as-IN bn-IN bo-CN dv-MV fa-IR gu-IN he-IL hi-IN km-KH kn-IN kok-IN ks-Arab lo-LA ml-IN mni-IN
        mr-IN my-MM ne-NP or-IN pa-IN ps-AF sa-IN sd-Arab-PK sd-Deva-IN si-LK syr-SY ta-IN te-IN th-TH
        tzm-Arab-MA ur-PK yi-Hebr""".split()),
}
_BY_KEY: dict[str, tuple[str, str]] = {tag.lower(): (tag, slot) for slot, tags in WORD_TAGS.items() for tag in tags}

class LanguageError(EditError):
    """A language or style choice refused; ``options`` holds what would do."""

    def __init__(self, message: str, options: list[str] | None = None) -> None:
        super().__init__(message)
        self.options = list(options or [])


#: Tags offered when a refused one is like none Word knows.
_COMMON = ("fr-FR", "fr-BE", "nl-BE", "nl-NL", "en-US", "en-GB", "de-DE")

#: ``styles=`` values besides a list of style names.
STYLE_MODES = ("matching", "all", "none")


def word_tag(language: Any) -> tuple[str, str, str | None]:
    """``(tag, attribute, note)``: the tag Word writes for ``language`` and the ``w:lang``
    attribute it goes in (``val``, ``eastAsia``, ``bidi``); ``note`` when the tag was read
    as another spelling of it.  :class:`EditError` -- with the tags that would do -- for
    what Word does not know as one language and region."""
    if not isinstance(language, str) or not language.strip():
        raise LanguageError(f"language is a tag such as 'fr-BE', 'nl-BE' or 'en-US', not {language!r}",
                            options=list(_COMMON))
    given = language.strip()
    found = _BY_KEY.get(given.lower()) or _BY_KEY.get(given.replace("_", "-").lower())
    if found is None:
        parts = given.replace("_", "-").split("-")
        if len(parts) == 3 and len(parts[1]) == 4:          # fr-Latn-BE: Word reads fr-BE
            found = _BY_KEY.get(f"{parts[0]}-{parts[2]}".lower())
    if found is not None:
        tag, slot = found
        return tag, slot, (None if tag == given else f"{given!r} is written as Word writes it: {tag}")
    primary = given.replace("_", "-").split("-")[0].lower()
    # The language's own country first (fr-FR, nl-NL), then the others.
    same = sorted((tag for tag in _BY_KEY.values() if tag[0].lower().split("-")[0] == primary),
                  key=lambda found: (found[0].lower() != f"{primary}-{primary}", found[0]))
    if same:
        options = [tag for tag, _ in same]
        what = ("a language without a region" if "-" not in given.replace("_", "-")
                else "not a region Word knows for it")
        raise LanguageError(f"{given!r} is {what}: Word proofs a language with its region, one of "
                        f"{', '.join(options[:30])}" + (" ..." if len(options) > 30 else ""), options=options)
    close = difflib.get_close_matches(given.lower(), list(_BY_KEY), n=6, cutoff=0.5)
    options = [_BY_KEY[key][0] for key in close]
    options += [tag for tag in _COMMON if tag not in options]
    raise LanguageError(f"{given!r} is not a language Word knows (it would proof the text in its own editing "
                    f"language); a tag such as {', '.join(options)}", options=options)


def _lang_of(holder: Element | None) -> Element | None:
    return holder.find(_W + "lang") if holder is not None else None


def _set(holder: Element, slot: str, tag: str) -> bool:
    """``tag`` in ``holder``'s ``w:lang`` ``slot`` (written where missing); whether it changed."""
    lang = holder.find(_W + "lang")
    if lang is None:
        lang = make("w:lang")
        insert_in_order(holder, lang)
    if lang.get(_W + slot) == tag:
        return False
    lang.set(_W + slot, tag)
    return True


def _same(a: str | None, b: str | None) -> bool:
    return a is not None and b is not None and a.replace("_", "-").lower() == b.replace("_", "-").lower()


class LanguageOps:
    """:meth:`Document.set_language` and :meth:`Document.languages`."""

    def languages(self: "Document") -> dict[str, Any]:
        """Where the document states languages: ``default`` (``w:docDefaults``' ``w:lang``,
        by attribute), ``theme`` (``w:themeFontLang``'s), ``styles`` (each style with its own:
        its name -> its attributes) and ``runs`` (each tag runs and paragraph marks carry
        themselves -> how many, in every story)."""
        package = self.package
        out: dict[str, Any] = {"default": {}, "theme": {}, "styles": {}, "runs": {}}
        part = package.styles_part()
        root = package.tree(part) if part else None
        if root is not None:
            lang = _lang_of(root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr"))
            out["default"] = _attributes(lang)
            for style in root.iter(_W + "style"):
                lang = _lang_of(style.find(_W + "rPr"))
                if lang is not None and _attributes(lang):
                    name = style.find(_W + "name")
                    out["styles"][name.get(_W + "val") if name is not None else style.get(_W + "styleId")] = \
                        _attributes(lang)
        settings = package.settings_part()
        theme = package.tree(settings).find(_W + "themeFontLang") if settings else None
        out["theme"] = _attributes(theme)
        for lang in self._run_langs():
            for value in _attributes(lang).values():
                out["runs"][value] = out["runs"].get(value, 0) + 1
        return out

    def _run_langs(self: "Document") -> list[Element]:
        """Every run's and paragraph mark's own ``w:lang``, in every story (not those in a
        revision's record of the formatting before)."""
        out = []
        for part in self._parts():
            for lang in self.package.tree(part).iter(_W + "lang"):
                holder = lang.getparent()
                if holder is None or holder.tag != _W + "rPr":
                    continue
                owner = holder.getparent()
                if owner is not None and owner.tag in (_W + "r", _W + "pPr"):
                    out.append(lang)
        return out

    def set_language(self: "Document", language: str, *, styles: "str | Iterable[str]" = "matching",
                     runs: bool = False, default: bool = True) -> "EditResult":
        """Set the document's language -- the one its text is proofed in and written for
        (round 35, R24) -- as a new document in it has it: ``w:docDefaults``' ``w:lang`` and
        the settings' ``w:themeFontLang``, in the attribute Word writes ``language`` in
        (:func:`word_tag`: ``fr-BE``, ``nl-BE``, ``en-US``...).  One undo step; the module
        has what was measured.  Text written afterwards gets the language's spacing
        (``docx_agent.edit.typography``); text already there is not changed (``check``'s
        ``punctuation_spacing`` lists it, with fixes)::

            doc.set_language("fr-BE")
            doc.set_language("nl-BE", styles="all")
            doc.set_language("en-GB", styles=["Quote"], default=False)

        ``styles``: ``"matching"`` (default) -- the styles whose own language is the default's
        before the change get the new one too, the others keep theirs; ``"all"`` -- every
        style with its own language; ``"none"``; or style names -- those styles (given
        their own language where they had none).  ``runs``: also the runs' own tags that are
        the old default's (never another language's).  ``default=False`` leaves the
        defaults as they are (a style's language alone).  The result's warnings say which
        styles keep another language and how many runs carry their own."""
        from .document import EditResult

        tag, slot, note = word_tag(language)
        ids: list[str] = []
        if isinstance(styles, str):
            mode = styles.strip().lower()
            if mode not in STYLE_MODES:
                raise LanguageError(f"styles is one of {', '.join(STYLE_MODES)} or a list of style names, "
                                f"not {styles!r}", options=list(STYLE_MODES))
        elif styles is None:
            mode = "none"
        else:
            mode = "names"
            ids = [self.styles.get(str(name)).id for name in styles]    # a missing one refused first
        if not default and mode in ("matching", "all", "none") and not runs:
            raise LanguageError("default=False changes only the styles named: give styles=[...] style names",
                            options=[s.name for s in self.styles if s.name][:30])
        package = self.package
        part = package.styles_part()
        if part is None:
            raise EditError("the document has no style sheet (styles.xml) to hold its default language")
        warnings = [note] if note else []
        old = _attributes(_lang_of(package.tree(part).find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr"))).get(slot)
        changed: list[str] = []
        kept: dict[str, str] = {}
        with self._edit():
            root = package.tree(part)
            dirty = False
            if default:
                defaults = root.find(_W + "docDefaults")
                if defaults is None:
                    defaults = make("w:docDefaults")
                    insert_in_order(root, defaults)
                run_defaults = defaults.find(_W + "rPrDefault")
                if run_defaults is None:
                    run_defaults = make("w:rPrDefault")
                    insert_in_order(defaults, run_defaults)
                properties = run_defaults.find(_W + "rPr")
                if properties is None:
                    properties = make("w:rPr")
                    run_defaults.append(properties)
                if _set(properties, slot, tag):
                    changed.append("default")
                    dirty = True
            for style in root.iter(_W + "style"):
                properties = style.find(_W + "rPr")
                lang = _lang_of(properties)
                own = lang.get(_W + slot) if lang is not None else None
                name_node = style.find(_W + "name")
                name = name_node.get(_W + "val") if name_node is not None else style.get(_W + "styleId")
                wanted = (mode == "all" and own is not None) \
                    or (mode == "matching" and _same(own, old)) \
                    or (mode == "names" and style.get(_W + "styleId") in ids)
                if wanted:
                    if properties is None:
                        properties = make("w:rPr")
                        insert_in_order(style, properties)
                    if _set(properties, slot, tag):
                        changed.append(f"style {name}")
                        dirty = True
                elif own is not None and not _same(own, tag):
                    kept[name] = own
            if dirty:
                package.mark_dirty(part)
            if default:
                settings = package.settings_part() or self._create_settings_part()
                settings_root = package.tree(settings)
                theme = settings_root.find(_W + "themeFontLang")
                if theme is None:
                    theme = make("w:themeFontLang")
                    insert_in_order(settings_root, theme)
                if theme.get(_W + slot) != tag:
                    theme.set(_W + slot, tag)
                    changed.append("theme")
                    package.mark_dirty(settings)
            rewritten = 0
            own_runs: dict[str, int] = {}
            for lang in self._run_langs():
                value = lang.get(_W + slot)
                if value is None:
                    continue
                if runs and _same(value, old) and value != tag:
                    lang.set(_W + slot, tag)
                    rewritten += 1
                    package.mark_dirty(_part_of(self, lang))
                elif not _same(value, tag):
                    own_runs[value] = own_runs.get(value, 0) + 1
            if rewritten:
                changed.append(f"{rewritten} run(s)")
        if kept:
            shown = sorted(kept.items())
            warnings.append("styles keeping their own language: " + ", ".join(f"{n} ({t})" for n, t in shown[:12])
                            + (f" and {len(shown) - 12} more" if len(shown) > 12 else ""))
        if own_runs:
            stale = sum(n for t, n in own_runs.items() if _same(t, old))
            warnings.append("runs keeping their own language tag: " + ", ".join(
                f"{t} ({n})" for t, n in sorted(own_runs.items(), key=lambda kv: -kv[1])[:8])
                + (f"; runs=True also sets the {stale} in {old}" if stale and not runs else ""))
        if not changed:
            warnings.insert(0, f"the document's language was {tag} already" if default
                            else f"the styles named were in {tag} already")
            return EditResult(None, changed=False, warnings=warnings)
        return EditResult(None, count=len(changed), warnings=warnings)


def _attributes(lang: Element | None) -> dict[str, str]:
    if lang is None:
        return {}
    return {slot: lang.get(_W + slot) for slot in ("val", "eastAsia", "bidi") if lang.get(_W + slot)}


def _part_of(document: "Document", node: Element) -> str:
    root = node.getroottree().getroot()
    for part in document._parts():
        if document.package.tree(part) is root:
            return part
    raise EditError("a run's language is in no story")  # pragma: no cover
