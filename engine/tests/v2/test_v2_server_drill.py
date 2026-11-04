# engine/tests/v2/test_v2_server_drill.py
import asyncio
from types import SimpleNamespace

from smartass import server
from smartass.model.lines import LineSet
from smartass.model.spec import AnalysisSpec

CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}


def spec_with(*blocks):
    return AnalysisSpec.model_validate({
        "version": 2, "question": "Ledger lines", "plan": {"depth": "quick"}, "queries": [],
        "blocks": list(blocks)})


def wire(monkeypatch, results):
    """results: one fake DrillResult per _drill call, in order."""
    cells, prepared = [], {}

    async def fake_drill(cell):
        cells.append(cell)
        return results[len(cells) - 1]

    async def user(fn, what):
        return fn(None)

    def prep(raw, provider, lines=None):
        prepared.update(raw=raw, lines=lines)
        return "PRE"

    monkeypatch.setattr(server, "_drill", fake_drill)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "run_prepare", prep)
    monkeypatch.setattr(server, "excel_job", lambda h, pre, target="session": {"q": 1, "warnings": ["w0"],
                                                                                 "pre": pre, "target": target})
    return cells, prepared


def result(rows=None, columns=None, value=-10.0, warnings=()):
    return SimpleNamespace(rows=rows if rows is not None else [{"AMOUNT": -10.0, "AP_CUSTOMER_NAME": "A"}],
                           columns=columns or ["AP_CUSTOMER_NAME", "AMOUNT"], cell_value=value,
                           reconciliation={"gap": 0.0}, warnings=list(warnings))


def test_fin_analyze_drills_each_lines_block_once(monkeypatch):
    cells, prepared = wire(monkeypatch, [result(), result(value=-5.0, warnings=["thin"])])
    cell2 = dict(CELL, **{"All Time Periods": "Aug 26"})
    spec = spec_with({"type": "lines", "cell": CELL}, {"type": "lines", "cell": cell2, "group_by": "account"})
    out = asyncio.run(server.fin_analyze(spec=spec))
    assert cells == [CELL, cell2]
    ls = prepared["lines"]
    assert sorted(ls) == ["b0", "b1"] and all(isinstance(v, LineSet) for v in ls.values())
    assert ls["b0"].cube_value == -10.0 and ls["b0"].cell == CELL and ls["b0"].alias == "b0"
    assert ls["b1"].warnings == ["thin"] and ls["b1"].reconciliation == {"gap": 0.0}
    assert out["pre"] == "PRE" and "thin" in out["warnings"] and "w0" in out["warnings"]


def test_project_grouping_adds_the_initiative(monkeypatch):
    rows = [{"AMOUNT": -10.0, "PROJECT": "P1", "PROJECT_DESC": "Alpha", "AP_CUSTOMER_NAME": "A"}]
    _, prepared = wire(monkeypatch, [result(rows=rows, columns=["AMOUNT", "PROJECT"])])
    asyncio.run(server.fin_analyze(spec=spec_with({"type": "lines", "cell": CELL, "group_by": "project"})))
    ls = prepared["lines"]["b0"]
    assert ls.rows[0]["INITIATIVE"] == "Alpha" and "INITIATIVE" in ls.columns and "AMOUNT" in ls.columns


def test_zero_rows_still_make_a_line_set_with_its_warnings(monkeypatch):
    _, prepared = wire(monkeypatch, [result(rows=[], columns=[], value=-7.0, warnings=["no drill report"])])
    asyncio.run(server.fin_analyze(spec=spec_with({"type": "lines", "cell": CELL})))
    ls = prepared["lines"]["b0"]
    assert ls.rows == [] and ls.cube_value == -7.0 and ls.warnings == ["no drill report"]


def test_a_spec_without_lines_blocks_does_not_drill(monkeypatch):
    from v2util import spec_dict
    cells, prepared = wire(monkeypatch, [])
    asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict())))
    assert cells == [] and prepared["lines"] == {}


def test_fin_drill_builds_a_one_block_spec(monkeypatch):
    cells, prepared = wire(monkeypatch, [result()])
    out = asyncio.run(server.fin_drill(CELL, group_by="account", top=5, target="active"))
    raw = prepared["raw"]
    assert cells == [CELL] and out["target"] == "active"
    assert raw["version"] == 2 and raw["queries"] == [] and raw["plan"]["depth"] == "quick"
    assert raw["question"].startswith("Ledger lines behind")
    assert [(b["type"], b["group_by"], b["top"]) for b in raw["blocks"]] == [("lines", "account", 5)]
    assert list(prepared["lines"]) == ["b0"]


def test_followup_re_drills_the_patched_spec(monkeypatch):
    cells, prepared = wire(monkeypatch, [result()])
    stored = {3: {"spec": {"v2": spec_with({"type": "lines", "cell": CELL}).model_dump()}, "ids": {},
                  "owned": {"rows": 3, "cols": 2}}}
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, "Q3 Answer", "C:/a/x.xlsx"))
    monkeypatch.setattr(server, "redraw_job", lambda h, pre, q, owned, target="session": {"q": q, "warnings": []})
    asyncio.run(server.fin_analyze(patch={"blocks": [{"type": "lines", "cell": CELL, "top": 3}]}))
    assert cells == [CELL] and prepared["raw"]["blocks"][0]["top"] == 3 and list(prepared["lines"]) == ["b0"]


def test_tool_names_and_instructions():
    assert "fin_drill" in server.TOOL_NAMES
    assert "fin_drill" in server.INSTRUCTIONS and len(server.INSTRUCTIONS) < 2500


def test_text_fix_candidates_are_listed_apart_from_the_warnings(monkeypatch):
    broken = "Text broken in EPM (letters lost at load): Tlefon AS; Acme Cnsult - add corrections to x\text_fixes.json"
    wire(monkeypatch, [result(warnings=[broken, "thin"])])
    out = asyncio.run(server.fin_analyze(spec=spec_with({"type": "lines", "cell": CELL})))
    assert out["text_fixes"] == ["Tlefon AS", "Acme Cnsult"]
    assert "thin" in out["warnings"] and not any(w.startswith("Text broken in EPM") for w in out["warnings"])


def test_no_text_fixes_key_without_broken_texts():
    assert "text_fixes" not in server._with_notes({"warnings": []}, ["thin"])


def test_several_lines_blocks_are_drilled_concurrently(monkeypatch):
    log = []

    async def slow_drill(cell):
        log.append(("start", cell["All Time Periods"]))
        await asyncio.sleep(0.05)
        log.append(("end", cell["All Time Periods"]))
        return result()

    monkeypatch.setattr(server, "_drill", slow_drill)
    cell2 = dict(CELL, **{"All Time Periods": "Aug 26"})
    sets, _ = asyncio.run(server._line_sets(spec_with({"type": "lines", "cell": CELL},
                                                      {"type": "lines", "cell": cell2})))
    assert sorted(sets) == ["b0", "b1"] and sets["b1"].cell == cell2
    assert [k for k, _ in log[:2]] == ["start", "start"]        # both started before either ended
