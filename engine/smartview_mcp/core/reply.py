"""The tool reply of an answer: save + activate the workbook, and the result dict (table, checks, warnings, sheet,
workbook) that sv_answer and sv_drill return. Also the per-workbook cache of a question's pulled grids."""
from __future__ import annotations

import logging

from .. import userexcel
from ..timing import stage

log = logging.getLogger("smartview")
_sessions: dict[str, dict[int, dict[str, "userexcel.Pull"]]] = {}   # workbook path -> q -> sheet -> Pull


def session(path: str, q: int) -> dict:
    return _sessions.setdefault(path.lower(), {}).setdefault(q, {})


def forget(path: str, q: int) -> None:
    """Drop a question's cached grids (they are reloaded from the workbook on next use)."""
    _sessions.get(path.lower(), {}).pop(q, None)


def load_session(t, q: int) -> dict:
    sess = _sessions.get(t.path.lower(), {}).get(q)
    if not sess:                                       # e.g. after a server restart: rebuild from the workbook
        with stage("reload pulls from workbook"):
            sess = userexcel.load_pulls(t, q)
        if sess:
            _sessions.setdefault(t.path.lower(), {})[q] = sess
    return sess or {}


def fin(v, kind: str = "heltall") -> str:
    """Financial text for chat, matching the sheet: 1 000 000 / (1 000 000) / – ; % with sign; FTE 1 decimal."""
    if v is None or v == "":
        return ""
    if not isinstance(v, (int, float)):
        return str(v)
    if kind.startswith("prosent"):
        return "–" if abs(v) < 5e-5 else f"{v * 100:+.1f} %".replace(".", ",")
    if abs(v) < 0.5 and kind != "tall1":
        return "–"
    s = f"{abs(v):,.1f}" if kind == "tall1" else f"{abs(v):,.0f}"
    s = s.replace(",", " ").replace(".", ",")
    return f"({s})" if v < 0 else s


def answer_markdown(spec: dict, res: dict) -> str:
    cols = spec.get("columns", [])
    head = ["Line"] + [c.get("header", "") for c in cols]
    out = ["| " + " | ".join(h.replace("|", "/") for h in head) + " |", "|---|" + "---:|" * len(cols)]
    n = 0
    for sec in spec.get("sections", []):
        out.append(f"| **{sec.get('title', '')}** |" + " |" * len(cols))
        for ln in sec.get("lines", []):
            k = ln.get("key")
            if not k or k not in res["values"]:
                continue
            vals = res["values"][k]
            cells = []
            for j, v in enumerate(vals):
                kind = ln.get("num") or cols[j].get("kind") or "heltall"
                if j < len(ln.get("cells", [])) and isinstance(ln["cells"][j], dict) and "varpct" in ln["cells"][j]:
                    kind = "prosent1"
                cells.append(fin(v, kind))
            label = ln.get("label", k)
            if ln.get("kind") in ("sum", "grand"):
                label = f"**{label}**"
            out.append("| " + label + " | " + " | ".join(cells) + " |")
            n += 1
            if n >= 45:
                out.append("| … see the answer sheet for the rest |" + " |" * len(cols))
                return "\n".join(out)
    return "\n".join(out)


def result(t, q: int, spec: dict, res: dict, pulls: list, warnings: list) -> dict:
    try:
        if userexcel.com(lambda: t.wb.ReadOnly, limit=10):
            warnings = [f"'{t.name}' is open READ-ONLY (someone or something else holds the file) - the analysis "
                        "is built but cannot be saved under this name."] + warnings
    except Exception:   # noqa: BLE001 - informational only
        pass
    return {
        "workbook": t.path, "question": f"Q{q}", "answer_sheet": res["sheet"],
        "data_sheets": [p.sheet for p in pulls] + ["Sources"],
        "key_points": spec.get("key_points", []),
        "table": answer_markdown(spec, res),
        "checks": res["checks"],
        "warnings": warnings + res["warnings"] + (res.get("error_warnings")
                                                  or [f"formula error on line '{k}'" for k in res["formula_errors"]]),
        "style_lint": res["lint"],
        # tn-formelstil: 'block' findings must be fixed before the answer is presented as done
        "formula_style": res.get("formula_style"),
        "note": "Your workbook was NOT saved - review and save it yourself." if not t.created
        else f"New analysis workbook saved: {t.path}",
    }


def tidy(t, sheet: str | None = None) -> None:
    """Fixed, neat sheet order (sheetorder: Answer first, other sheets, data per question, Sources, hidden) and land
    on `sheet`. Our own analysis files get the full order; in a user's own workbook only our sheets are reordered
    among themselves. Cosmetic: never fails the tool."""
    from .. import sheetorder
    from ..server import _is_our_analysis_file
    try:
        with stage("arrange sheets"):
            mode = "full" if (t.created or _is_our_analysis_file(t.path)) else "active"
            sheetorder.arrange_sheets(t.wb, mode, activate=sheet)
    except Exception as e:   # noqa: BLE001
        log.warning("arrange sheets failed in %s: %s", t.name, e)
        if sheet:
            try:
                userexcel.com(t.wb.Worksheets(sheet).Activate)
            except Exception:   # noqa: BLE001
                pass


def finish(t, q: int, spec: dict, res: dict, saved_new: bool):
    tidy(t, res["sheet"])
    t.status(None)
    if saved_new:
        userexcel.com(t.wb.Save)
