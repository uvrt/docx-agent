"""ooxml-edit's child-order registry is last-wins, and its charts subpackage owns the
sequences of DrawingML text (``a:p``, ``a:r``, ``a:br``, ``a:fld``), charts (``c:*``) and
diagrams (``dgm:*``, ``dsp:txBody``).  docx-agent registers none of them: importing it leaves
every one of the subpackage's sequences as the subpackage wrote it, whichever is imported
first, and adds only WordprocessingML's own.

Each snapshot is taken in a fresh interpreter, so what one import registers is seen alone.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

#: docx-agent's vocabulary module, executed alone (importing it through the package would
#: import the package, charts adapter and all).
VOCABULARY = Path(__file__).resolve().parent.parent / "src" / "docx_agent" / "oxml" / "xml.py"

_DUMP = ("import json; from ooxml_edit.xml import CHILD_ORDER; "
         "print(json.dumps({k: repr(v) for k, v in CHILD_ORDER.items()}))")


def _snapshot(imports: str) -> dict[str, str]:
    done = subprocess.run([sys.executable, "-c", f"{imports}; {_DUMP}"], capture_output=True, text=True, check=True)
    return json.loads(done.stdout)


def test_docx_agent_leaves_the_charts_subpackages_orders_alone():
    own = _snapshot("import importlib.util as u; s = u.spec_from_file_location('vocabulary', "
                    f"{str(VOCABULARY)!r}); s.loader.exec_module(u.module_from_spec(s))")
    charts = _snapshot("import ooxml_edit.charts")
    assert not [key for key in own if key.split(":")[0] in ("a", "c", "dgm", "dsp", "x")]
    for imports in ("import docx_agent", "import ooxml_edit.charts; import docx_agent",
                    "import ooxml_edit.charts; import docx_agent.oxml.xml; import docx_agent.edit.charts"):
        both = _snapshot(imports)
        assert {key: both[key] for key in charts} == charts, imports
        # Nothing else: docx-agent's own sequences and the subpackage's, unchanged.
        assert both == {**own, **charts}, imports
