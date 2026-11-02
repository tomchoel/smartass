"""check (design §3.4), skeleton: reconciliation checks, formula errors and sheet = fetch, as a findings table
Critical / Warning / Info (structure after Anthropic financial-services audit-xls, Apache-2.0). Critical = a number
on the answer cannot be trusted; Warning = trustworthy but the user must know; Info = cosmetic."""
from __future__ import annotations

ORDER = {"Critical": 0, "Warning": 1, "Info": 2}


def _warned(k: str, labels: set[str]) -> bool:
    """A check row is '<label>' (static) or '<label> — <column header>' (answer.py)."""
    return any(k == lb or k.startswith(f"{lb} — ") or k.startswith(f"{lb} (") for lb in labels)


def review(res: dict, sources: list, extra: list[dict] | None = None, warn_checks: set[str] | None = None) -> dict:
    f: list[tuple[str, str]] = [(x["severity"], x["text"]) for x in extra or []]
    checks = res.get("checks") or {}
    f += [("Warning" if _warned(k, warn_checks or set()) else "Critical", f"check '{k}' is {v}")
          for k, v in checks.items() if v != "OK"]
    f += [("Critical", w) for w in (res.get("error_warnings")
                                    or [f"formula error on line '{k}'" for k in res.get("formula_errors") or []])]
    for s in sources:
        f += [("Critical", f"{s.pull.sheet}: the sheet differs from the fetch ({m})") for m in s.mismatches]
        if s.note:
            f.append(("Warning", f"{s.pull.sheet}: {s.note}"))
        f += [("Warning", f"{s.pull.sheet}: {w}") for w in s.pull.parsed.pov_problems]
        if s.pull.parsed.unknown_members:
            f.append(("Warning", f"{s.pull.sheet}: unknown members {', '.join(s.pull.parsed.unknown_members)}"))
    if not checks:
        f.append(("Warning", "the answer has no reconciliation checks"))
    f += [("Warning", f"formula style: {b.get('message')} ({b.get('cell')})")
          for b in (res.get("formula_style") or {}).get("block") or []]
    f += [("Warning", str(w)) for w in res.get("warnings") or []]
    f += [("Info", str(x)) for x in res.get("lint") or []]
    f.sort(key=lambda x: ORDER[x[0]])
    return {"ok": not any(s == "Critical" for s, _ in f), "findings": [{"severity": s, "text": t} for s, t in f]}
