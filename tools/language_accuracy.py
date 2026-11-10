"""Measure :func:`docx_agent.edit.text_language.detect` on made-up text (round 49, R94).

Reads ``tests/fixtures/language-samples.json`` (paragraphs written for this in each language
the detector knows), cuts every paragraph into windows of ``n`` words (from every third word,
so a window starts mid-sentence too) and detects each, at several ``n``.  Prints, per length:
the windows judged (enough profile words), the judged ones right, and the windows that would
be listed as another language than their tag -- wrong and at least ``MIN_CONFIDENCE`` sure:
the false positives a check would report on text tagged right -- and the windows detected
right that sure (``sure``): those a check would list were they tagged wrong.

    python tools/language_accuracy.py [--held-out | --fresh | --all] [--json] [--confusion N]

The sets: ``tuning`` (the default) the texts the cues were chosen with; ``held_out`` written
after, looked at once (it showed French "à" missing from its profile); ``fresh`` written
after the last change and measured once; ``--all`` the three together.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from docx_agent.edit.text_language import LANGUAGES, MIN_CONFIDENCE, detect

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "tests" / "fixtures" / "language-samples.json"
LENGTHS = (3, 4, 5, 6, 8, 10, 12, 15, 20, 30)


def windows(text: str, n: int, step: int = 3) -> list[str]:
    words = text.split()
    if len(words) < n:
        return []
    return [" ".join(words[i:i + n]) for i in range(0, len(words) - n + 1, step)]


def measure(samples: dict[str, list[str]], n: int) -> dict:
    per: dict[str, dict] = {}
    confusion: dict[str, dict[str, int]] = {}
    for language in LANGUAGES:
        cut = [w for paragraph in samples[language] for w in windows(paragraph, n)]
        judged = right = listed_wrong = sure = 0
        for text in cut:
            found = detect(text, min_words=0)
            if found.language is None:
                continue
            judged += 1
            if found.language == language:
                right += 1
                sure += found.confidence >= MIN_CONFIDENCE
            else:
                confusion.setdefault(language, {}).setdefault(found.language, 0)
                confusion[language][found.language] += 1
                if found.confidence >= MIN_CONFIDENCE:
                    listed_wrong += 1
        per[language] = {"windows": len(cut), "judged": judged, "right": right,
                         "false_positives": listed_wrong, "sure": sure}
    total = {key: sum(v[key] for v in per.values()) for key in ("windows", "judged", "right",
                                                                  "false_positives", "sure")}
    return {"n": n, "per_language": per, "total": total, "confusion": confusion}


def main() -> None:
    sets = json.loads(SAMPLES.read_text(encoding="utf-8"))
    which = ("held_out" if "--held-out" in sys.argv else "fresh" if "--fresh" in sys.argv
             else "all" if "--all" in sys.argv else "tuning")
    if which == "all":
        samples = {language: [text for name in ("tuning", "held_out", "fresh")
                              for text in sets[name][language]] for language in LANGUAGES}
    else:
        samples = sets[which]
    results = [measure(samples, n) for n in LENGTHS]
    if "--json" in sys.argv:
        print(json.dumps(results, indent=1))
        return
    print(f"{'words':>5} {'windows':>7} {'judged':>7} {'right/judged':>12} {'FP/windows':>10} {'sure/windows':>12}  per language right/judged")
    for result in results:
        t = result["total"]
        per = "  ".join(f"{lang} {v['right']}/{v['judged']}" for lang, v in result["per_language"].items())
        print(f"{result['n']:>5} {t['windows']:>7} {t['judged']:>7} "
              f"{t['right'] / max(1, t['judged']):>12.3f} {t['false_positives'] / max(1, t['windows']):>10.3f} {t['sure'] / max(1, t['windows']):>12.3f}  {per}")
    if "--confusion" in sys.argv:
        n = int(sys.argv[sys.argv.index("--confusion") + 1])
        print(json.dumps(next(r for r in results if r["n"] == n)["confusion"], indent=1))


if __name__ == "__main__":
    main()
