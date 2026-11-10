"""The language a text is written in, next to the language it is tagged with (round 49, R94).

``check``'s layout facts list it as ``text_language`` (:mod:`docx_agent.layout_facts`):
paragraphs whose words are another language than the tag Word proofs them in, and the
document's language by its words next to the language its default paragraph style gives.

**Detecting.**  :func:`detect` scores a text against small profiles of each language's
function words (:data:`PROFILES`: articles, pronouns, prepositions, conjunctions, auxiliary
verbs -- the words every sentence has whatever it is about), written for this module, no
list taken from elsewhere and no dependency.  A word in several languages' profiles counts
for each a share (``1/n``); a word that is none counts :data:`CUE_WEIGHT` for the
languages whose spellings it shows (:data:`CUES`: letters, endings).  The language with the
highest score wins; its ``confidence`` is how far it leads the next, ``(best - second) /
best`` (0: a tie, 1: no other language scored).  A text with fewer than :data:`MIN_WORDS` words, or
fewer than :data:`MIN_HITS` profile words (or under :data:`MIN_HIT_SHARE` of its words),
is not judged: a heading, a label, a list of names, placeholder Latin says too little.  Languages: :data:`LANGUAGES` -- English, French, Dutch, German,
Spanish, Italian, Portuguese.  A tag in another language is not compared.

**Measured** on made-up text in each language (``tests/data/language-samples.json``,
``tools/language_accuracy.py``): see :data:`MIN_WORDS` and :data:`MIN_CONFIDENCE`.

**The fact** (:func:`facts`):

``document`` (``checked`` for a check of some pages or a range: then of their words, and
no paragraph is folded into it, no fix)
    ``{"tag", "source", "detected", "share", "words", "paragraphs"}``: the language of
    the default paragraph style (``tag``; ``source`` where it comes from: ``default``,
    ``style:<name>``) and the language most of the words of the paragraphs judged at least
    :data:`MIN_CONFIDENCE` sure are in (``detected``, ``share`` of those ``words``;
    ``None`` when none is that sure); ``paragraphs`` counts the paragraphs judged.  Where they differ, ``mismatched`` counts the
    paragraphs in the detected language under that tag (they are not listed one by one)
    and ``fix`` sets the document's language (``word_styles`` ``set_language``), as
    ``punctuation_spacing``'s ``language_mismatch`` does.
``mismatched``
    The other paragraphs whose words are another language than their tag, at least
    :data:`MIN_CONFIDENCE` sure: ``{"para" | "range", "tag", "source", "detected",
    "confidence", "words"}`` -- ``range`` (``p:A..p:B``, ``paragraphs``) for consecutive
    top-level paragraphs alike (shorter ones between them, too short to judge, go with
    them).  The tag is the one most of the paragraph's text has; ``source`` where it comes
    from (``run``: the runs' own ``w:lang``).  At most :data:`~docx_agent.layout_facts.CAP`;
    ``mismatched_left_out`` counts the rest.  Paragraphs that agree are not listed:
    ``document``'s ``paragraphs`` counts those judged, ``too_short`` those with text too
    short to judge, ``unclear`` those with words enough but too few function words.
"""

from __future__ import annotations

import functools
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Callable, Iterable

from ..oxml.xml import make

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: Each language's function words (lower case; an elided form with its apostrophe).
PROFILES: dict[str, frozenset[str]] = {
    "en": frozenset("""
        the of and to in is are was were be been being it its that this these those with for on at by
        from as an a or not but have has had will would can could should shall which who whom whose what
        when where why how their they them we our us you your he she his her him there than then also
        into about after before all each more most some such only if do does did may might must no so
        very over between through under while both any other since until upon because whether""".split()),
    "fr": frozenset("""
        le la les l' un une des du de d' à et est sont était étaient être été a ont avait avaient au aux
        en dans par pour sur avec sans sous ce cet cette ces qui que qu' dont où il elle ils elles nous
        vous on je se s' son sa ses leur leurs notre nos votre vos ne n' pas plus mais ou donc car si
        aussi très comme tout tous toute toutes entre chez depuis après avant déjà encore même peut
        peuvent doit doivent c' cela ceci lors ainsi lorsque puis alors""".split()),
    "nl": frozenset("""
        de het een en van in is zijn was waren op te dat die dit deze met voor aan bij door om uit naar
        over tot als of maar ook niet geen nog wel al er hij zij ze wij we jullie u ik je hun haar
        wordt worden werd werden heeft hebben had kan kunnen moet moeten zal zullen wat wie waar
        wanneer hoe omdat dus want dan toch zo meer veel alle elke ieder onder tussen sinds na zonder
        daar hier mijn ons onze welke wanneer""".split()),
    "de": frozenset("""
        der die das den dem des ein eine einen einem einer eines und ist sind war waren sein wird
        werden wurde wurden hat haben hatte in im am an auf aus bei mit nach von vom zu zum zur für
        über unter durch gegen ohne um zwischen seit nicht kein keine auch noch schon nur sehr oder
        aber denn dass wenn weil ob als wie was wer wo er sie es wir ihr ich du man sich ihre unsere
        dieser diese dieses kann können muss müssen soll sollen alle mehr hier dort bereits""".split()),
    "es": frozenset("""
        el la los las un una unos unas y e o u de del al en con por para sin sobre entre hasta desde
        es son era eran fue ser está están estar ha han había hay que qué quien cual cuando donde como
        lo le les se su sus nuestro nuestra este esta estos estas ese esa pero también muy más no ni
        ya todo todos toda todas cada puede pueden debe deben mientras aunque porque sino nos yo él
        ella ellos""".split()),
    "it": frozenset("""
        il lo la i gli le l' un uno una un' e ed o di del dello della dei degli delle a al allo alla ai
        agli alle da dal dalla in nel nello nella nei negli nelle con su sul sulla per tra fra è sono
        era erano essere ha hanno aveva che chi cui quale quando dove come non ma anche più molto già
        ancora se si ci questo questa questi queste quello quella ogni tutto tutti suo sua loro nostro
        può possono deve devono dell' all' nell' sull' dall'""".split()),
    "pt": frozenset("""
        o a os as um uma uns umas e ou de do da dos das em no na nos nas ao aos à às com por pelo pela
        para sem sobre entre até desde é são era eram foi ser está estão estar tem têm tinha há que quem
        qual quando onde como não mas também muito mais já ainda se seu sua seus suas este esta isso
        isto esse essa cada todo todos pode podem deve devem nosso nossa ele ela eles você""".split()),
}

#: Spellings each language has and the others mostly lack, in a word (lower case): letters,
#: letter groups, endings.  A word that shows the cues of several languages counts a share
#: of :data:`CUE_WEIGHT` for each, as a profile word does.
CUES: dict[str, re.Pattern[str]] = {
    "en": re.compile(r"^th|^wh|w|ing$|ould$|ly$|ness$|ough"),
    "fr": re.compile(r"[çœ]|[éèêëâîôû]|eau|aient$|ais$|ait$|ez$|eux$|oir$"),
    "nl": re.compile(r"ij|oe|aa|uu|sch|cht|lijk|heid$|^ge|^ver"),
    "de": re.compile(r"[äöüß]|sch|ung$|keit$|heit$|tz|^ge|^ver|^zu|ich$|cht"),
    "es": re.compile(r"ñ|ci[oó]n$|ciones$|[áíóú]|ado$|ada$|ll|os$|mente$"),
    "it": re.compile(r"zion[ei]$|gli|cch|zz|[àèìòù]$|tt|sse$|are$|ere$|ire$"),
    "pt": re.compile(r"ção$|ções$|[ãõ]|lh|nh|[âêô]|mente$|ar$|em$"),
}
#: What a word's cues count, together (a profile word counts 1).
CUE_WEIGHT = 0.5

#: The languages detected, in the order a tie is decided.
LANGUAGES = tuple(PROFILES)

#: Each word's weight per language: 1 over the number of profiles it is in.
_WEIGHTS: dict[str, tuple[tuple[str, float], ...]] = {}
for _word in set().union(*PROFILES.values()):
    _in = [language for language in LANGUAGES if _word in PROFILES[language]]
    _WEIGHTS[_word] = tuple((language, 1.0 / len(_in)) for language in _in)
del _word, _in

#: Words a text needs to be judged (measured: see the module's notes).
MIN_WORDS = 8
#: Profile words a text needs to be judged.
MIN_HITS = 3
#: The share of a text's words its profile words must be at least: running text in these
#: languages has a function word in every three to five (measured: at least 0.2 in every
#: window of 15 words of the made-up text); Latin placeholder text ("lorem ipsum") about
#: 0.05-0.13.
MIN_HIT_SHARE = 0.15
#: A paragraph is listed as another language than its tag at least this sure.
MIN_CONFIDENCE = 0.5

_APOSTROPHES = str.maketrans({"’": "'", "ʼ": "'", "`": "'"})
#: A word: letters, an elided article or pronoun with its apostrophe (``l'``, ``dell'``).
_WORD = re.compile(r"[^\W\d_]+(?:'(?=[^\W\d_]))?")


@functools.lru_cache(maxsize=1 << 16)
def _cued(word: str) -> tuple[tuple[str, float], ...]:
    """A word's :data:`CUES`, as weights per language (a document repeats its words)."""
    cued = [language for language in LANGUAGES if CUES[language].search(word)]
    return tuple((language, CUE_WEIGHT / len(cued)) for language in cued)


@dataclass(frozen=True)
class Detected:
    """What :func:`detect` found: ``language`` (``None``: not judged), ``confidence`` (how far
    it leads the next language: ``(best - second) / best``), the text's ``words`` and the profile words among them
    (``hits``)."""

    language: str | None
    confidence: float
    words: int
    hits: int


def words_of(text: str) -> list[str]:
    """``text``'s words, lower case, an elision with its apostrophe (``l'``)."""
    return _WORD.findall(text.translate(_APOSTROPHES).casefold())


def detect(text: str, min_words: int = MIN_WORDS) -> Detected:
    """The language of ``text`` by its function words (the module's notes)."""
    words = words_of(text)
    if len(words) < min_words:
        return Detected(None, 0.0, len(words), 0)
    scores = dict.fromkeys(LANGUAGES, 0.0)
    hits = 0
    for word in words:
        weights = _WEIGHTS.get(word)
        if weights is not None:
            hits += 1
        else:
            weights = _cued(word)
        for language, weight in weights:
            scores[language] += weight
    total = sum(scores.values())
    if hits < MIN_HITS or hits < MIN_HIT_SHARE * len(words) or total <= 0:
        return Detected(None, 0.0, len(words), hits)
    ranked = sorted(LANGUAGES, key=lambda language: -scores[language])
    best, second = scores[ranked[0]], scores[ranked[1]]
    return Detected(ranked[0], round((best - second) / best, 2), len(words), hits)


def primary(tag: str | None) -> str | None:
    """A tag's language (``fr`` of ``fr-BE``), lower case."""
    return tag.replace("_", "-").split("-")[0].lower() if tag else None


#: The tag a fix proposes for a language the document names nowhere.
DEFAULT_TAGS = {"en": "en-US", "fr": "fr-FR", "nl": "nl-NL", "de": "de-DE", "es": "es-ES",
                "it": "it-IT", "pt": "pt-PT"}


def tag_for(document: "Document", language: str) -> str:
    """The tag a fix sets for ``language``: the document locale's when it is that language,
    else a tag of it the styles state, else :data:`DEFAULT_TAGS`'."""
    locale = getattr(getattr(document, "locale", None), "language", None)
    if primary(locale) == language and "-" in str(locale):
        return str(locale)
    package = document.package
    part = package.styles_part()
    root = package.tree(part) if part else None
    if root is not None:
        for lang in root.iter(_W + "lang"):
            value = lang.get(_W + "val")
            if primary(value) == language and "-" in (value or ""):
                return str(value)
    return DEFAULT_TAGS[language]


# -- the fact ---------------------------------------------------------------------------------


@dataclass
class _Judged:
    identifier: str
    tag: str | None
    source: str | None
    found: Detected


def _judge(cascade: Any, element: Any, identifier: str) -> tuple[_Judged | None, str | None]:
    """``(judged, why not)``: the paragraph's tag and detected language, or ``None`` and
    ``"too_short"`` (text of too few words), ``"unclear"`` (words enough, too few of them
    function words: names, placeholder Latin) or ``None`` (no text).  Its runs' text: ``w:t``
    (deleted text is ``w:delText``), a text box's paragraphs left to themselves."""
    raw = "".join(node.text or "" for node in element.iter(_W + "t"))
    if len(words_of(raw)) < MIN_WORDS:
        return None, "too_short" if raw.strip() else None
    if element.find(f".//{_W}txbxContent") is not None:
        from . import text as _text
        runs = _text.runs(element)
    else:
        runs = element.iter(_W + "r")
    chars: dict[tuple[str | None, str | None], int] = {}
    text: list[str] = []
    for run in runs:
        written = "".join(node.text or "" for node in run.findall(_W + "t"))
        if not written:
            continue
        language, proofed, code, source = cascade.of_with_source(element, run)
        if not proofed or code:
            text.append(" ")
            continue
        text.append(written)
        key = (language, source)
        chars[key] = chars.get(key, 0) + len(written)
    found = detect("".join(text))
    if found.language is None or not chars:
        return None, "too_short" if found.words < MIN_WORDS else "unclear"
    (tag, source), _ = max(chars.items(), key=lambda item: item[1])
    return _Judged(identifier, tag, source, found), None


def facts(document: "Document", *, keep: Callable[[str], bool] | None = None, whole: bool = True,
          top_level: Iterable[tuple[str, Any, str]] = (), cap: int = 20) -> dict[str, Any] | None:
    """``check``'s ``text_language`` (the module's notes); ``None`` when no paragraph has
    words enough to judge.  ``keep(identifier)``: whether a paragraph is in the check's
    pages and range; ``whole``: whether that is the whole document -- else the language of
    the words checked is ``checked``, not ``document``, and says nothing of the document's
    (no paragraph folded into it, no fix); ``top_level``: the body's blocks in order,
    ``(kind, element, id)`` (``layout_facts``' ``_top_level``), for the ranges."""
    from . import typography

    cascade = typography.Cascade(document)
    judged: list[_Judged] = []
    short: set[str] = set()
    unclear = 0
    for part in document._parts():
        for entry in document._index(part).paragraphs:
            if keep is not None and not keep(entry.id):
                continue
            found, why = _judge(cascade, entry.element, entry.id)
            if found is not None:
                judged.append(found)
            elif why == "too_short":
                short.add(entry.id)
            elif why == "unclear":
                unclear += 1
    if not judged:
        return None
    counts = {**({"too_short": len(short)} if short else {}), **({"unclear": unclear} if unclear else {})}
    default_tag, default_source = cascade.of_with_source(make("w:p"), None)[0::3]
    words: dict[str, int] = {}
    for item in judged:
        if item.found.confidence >= MIN_CONFIDENCE:
            words[item.found.language] = words.get(item.found.language, 0) + item.found.words  # type: ignore[index]
    total = sum(words.values())
    if not total:
        return {"document" if whole else "checked": {"tag": default_tag, "source": default_source or "none",
                                                     "detected": None, "paragraphs": len(judged)},
                **counts}
    dominant = max(LANGUAGES, key=lambda language: words.get(language, 0))
    document_fact: dict[str, Any] = {"tag": default_tag, "source": default_source or "none",
                                     "detected": dominant, "share": round(words[dominant] / total, 2),
                                     "words": total, "paragraphs": len(judged)}
    own = primary(default_tag)
    folded = 0
    differs = whole and own in PROFILES and own != dominant

    def mismatched(item: _Judged) -> bool:
        tag = primary(item.tag)
        return (tag in PROFILES and tag != item.found.language
                and item.found.confidence >= MIN_CONFIDENCE)

    listed: list[_Judged] = []
    for item in judged:
        if not mismatched(item):
            continue
        if differs and item.found.language == dominant and primary(item.tag) == own and \
                item.source != "run":
            folded += 1                                # the document's own mismatch
            continue
        listed.append(item)
    if differs:
        document_fact["mismatched"] = folded
        document_fact["fix"] = {"tool": "word_styles", "calls": [
            {"action": "set_language", "language": tag_for(document, dominant)}]}
        document_fact["note"] = (f"text mostly {dominant}, proofed as {default_tag}; the fix sets the "
                                 "document's language (the text's own region if another)")
    out: dict[str, Any] = {"document" if whole else "checked": document_fact}
    entries = _grouped(listed, list(top_level), short)
    if entries:
        out["mismatched"] = entries[:cap]
    if len(entries) > cap:
        out["mismatched_left_out"] = len(entries) - cap
    out.update(counts)
    return out


def _grouped(listed: list[_Judged], top_level: list[tuple[str, Any, str]],
             short: set[str]) -> list[dict[str, Any]]:
    """``listed`` as entries, consecutive top-level paragraphs alike as one ``range``."""
    order = {identifier: i for i, (kind, _, identifier) in enumerate(top_level) if kind == "p"}
    out: list[dict[str, Any]] = []
    run: list[_Judged] = []

    def alike(a: _Judged, b: _Judged) -> bool:
        if a.identifier not in order or b.identifier not in order:
            return False
        if (primary(a.tag), a.source, a.found.language) != (primary(b.tag), b.source, b.found.language):
            return False
        # Only paragraphs too short to judge (or without text) between them.
        between = top_level[order[a.identifier] + 1:order[b.identifier]]
        return all(kind == "p" and (identifier in short or not _has_text(element))
                   for kind, element, identifier in between)

    def flush() -> None:
        if not run:
            return
        words = sum(item.found.words for item in run)
        confidence = round(sum(item.found.confidence * item.found.words for item in run) / words, 2)
        first, last = run[0], run[-1]
        entry: dict[str, Any] = ({"para": first.identifier} if len(run) == 1 else
                                 {"range": f"{first.identifier}..{last.identifier}",
                                  "paragraphs": order[last.identifier] - order[first.identifier] + 1})
        entry.update({"tag": first.tag, "source": first.source or "none",
                      "detected": first.found.language, "confidence": confidence, "words": words})
        out.append(entry)
        run.clear()

    for item in listed:
        if run and not alike(run[-1], item):
            flush()
        run.append(item)
    flush()
    return out


def _has_text(element: Any) -> bool:
    return any((node.text or "").strip() for node in element.iter(_W + "t"))


__all__ = ["DEFAULT_TAGS", "Detected", "LANGUAGES", "MIN_CONFIDENCE", "MIN_HIT_SHARE", "MIN_HITS", "MIN_WORDS",
           "PROFILES", "detect", "facts", "primary", "tag_for", "words_of"]
