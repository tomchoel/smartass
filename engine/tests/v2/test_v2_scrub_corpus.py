"""Every warning, caveat and note text the engine can put in a reply passes the scrub unchanged (Task 6b): the scrub
(logsetup.redact; svxml's provider error text goes through the same one) runs on reply warnings and log lines, so a
rule that matched ordinary text would garble them. The corpus: every caveat text of knowledge/semantic/facts, every
concept's notes and item notes (concept_caveats), the draft caveat, the pipeline's warnings (a grid neither engine
could fetch, the second engine, a member left out) as the offline fixtures make them with member codes, the provider
route's own warnings (a member Smart View did not know, no data, an error cell, a lost request retried), the drill's
warnings (text broken in EPM, a gap to the cube, a split that does not tie, no lines), Planning REST's member without
a label, the key points over the limit, the sign-in prompt (a made-up page and code), the atlas's "NOT replaced"
reasons, the Login refusal, the drill's next steps and the next steps of every error class - each made by the
engine's own code. Nothing reaches EPM, the identity domain or Excel.

Not a collector of every warning the suite happens to produce: there is no one funnel (each source returns its own
list), and tests send made-up secret shapes through those lists on purpose to see them scrubbed."""
import asyncio
from pathlib import Path
from types import SimpleNamespace

import pytest
from test_v2_fallback_flow import _Rest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.core import logsetup, serverkit
from smartass.data import provider
from smartass.data.provider import SvHttpProvider, facts_from_parsed
from smartass.epm import drill_lines, drillstep, oauth, rest, svhttp, svxml
from smartass.epm.svhttp import SvHttpError
from smartass.knowledge import atlas
from smartass.knowledge import resolver as kresolver
from smartass.knowledge.semantic import load_concepts, load_semantic
from smartass.model.blocks.concept import concept_caveats
from smartass.model.requests import GridRequest
from smartass.model.spec import CommentaryBlock, ConceptBlock
from smartass.ops import atlas_build
from smartass.pipeline import left_out, prepare
from smartass.render.keypoints import free_points
from smartass.xl import dialogs

SEMANTIC = Path(__file__).resolve().parents[3] / "knowledge" / "semantic"
REFUSED = ("Smart View provider: Grid is Invalid. - the second engine (Planning REST) could not fetch it either "
           "(query fm): Planning REST refused the FinMax grid (Periodic / NOK / Actual): All Entities member "
           "un.SY.2BUS does not exist there or you have no access to it; as the only All Entities member asked it "
           "cannot be left out - check the member or rebuild the atlas (`smartass atlas build`)")


def _knowledge() -> list[str]:
    _metrics, facts_ = load_semantic(SEMANTIC)
    texts = [f.text for f in facts_]
    by_id = {c.id: c for c in load_concepts(SEMANTIC / "concepts")}
    r = SimpleNamespace(concept=by_id.get)
    for cid in by_id:
        spec = SimpleNamespace(blocks=[ConceptBlock(type="concept", query="q1", concept=cid)])
        texts += [c["text"] for c in concept_caveats(spec, r)]
    texts.append(f"{kresolver.DRAFT_CAVEAT} (account_c30, abolished_ig; smartass knowledge approve).")
    return texts


def _pipeline() -> list[str]:
    class Flaky(FakeProvider):
        def fetch(self, req):
            if req.periods == ("Sep 25",):
                raise SvHttpError(REFUSED)
            return super().fetch(req)
    failed = prepare(spec_dict(), resolver(), Flaky(facts())).warnings
    assert any(REFUSED in w for w in failed)              # the provider's text reached the warning whole, scrubbed
    texts = [REFUSED, *failed]
    texts += prepare(spec_dict(), resolver(), _Rest(facts(), lambda req: True)).warnings
    got = [SimpleNamespace(dropped=[("FinMax", "Account", "ac.C50S"), ("FinMax", "Entity", "un.SY.B2")])]
    warnings, caveats = left_out(spec_dict_model(), got, resolver())
    return texts + warnings + [c["text"] for c in caveats]


def spec_dict_model():
    from smartass.model.spec import AnalysisSpec
    return AnalysisSpec.model_validate(spec_dict())


def _provider(monkeypatch) -> list[str]:
    """data/provider.py's own warnings, as facts_from_parsed and the lost request's retry make them."""
    req = GridRequest("fm", "FinMax", ("ac.C30", "ac.C50S", "ac.C60", "ac.C70"), ("un.SY.2TEC", "un.SY.B2"),
                      ("Jul 26", "Sep 26"), "Actual", "", "Periodic", ("pr.MOB1",))
    labels = {"Jul 26": "Jul 26", "Sep 26": "Sep 26"}
    cols = ["All Accounts", "All Entities", "All Products", "Jul 26", "Sep 26"]
    grids = [SimpleNamespace(unknown_members=["ac.C99", "un.SY.X1"], empty=False, columns=cols, rows=[]),
             SimpleNamespace(unknown_members=[], empty=True, columns=cols, rows=[]),
             SimpleNamespace(unknown_members=[], empty=False, columns=cols,
                             rows=[["ac.C30", "un.SY.2TEC", "pr.MOB1", "#NoAccess", 5.0],
                                   ["ac.C50S", "un.SY.B2", "pr.MOB1", None, "#Invalid"]])]
    texts = [w for g in grids for w in facts_from_parsed(req, g, labels)[1]]
    texts += facts_from_parsed(GridRequest("fm", "FinMax", ("ac.C30",), ("un.SY.2TEC",), (), "Actual", ""),
                               SimpleNamespace(unknown_members=[], empty=True, columns=cols[:2], rows=[]), {})[1]
    lost = [SvHttpError("Smart View provider at https://plan1.example: lost the request (TimeoutError)")]

    def retrieve(*args):
        if lost:
            raise lost.pop()
        return None
    monkeypatch.setattr(provider, "RETRY_WAIT_S", 0)
    out = SimpleNamespace(warnings=[])
    SvHttpProvider(SimpleNamespace(retrieve=retrieve))._retrieve(out, "FinMax", None)
    return texts + out.warnings


def _drill(monkeypatch) -> list[str]:
    """epm/drillstep.drill's warnings, its pieces and their lines faked: text broken in EPM (drill_lines' own mark of
    a lost letter), a gap to the cube, split pieces that do not tie, a cell without an entity; then a drill with no
    lines at all."""
    lost = drill_lines._LOST
    split = [{"cell": {"p": 1}, "cube_value": -100.0, "items": [], "view": {"v": 1}, "secs": 0},
             {"cell": {"p": 2}, "cube_value": -200.0, "items": [], "view": {"v": 1}, "secs": 0}]
    lines = {"rows": [], "cols": [], "profile": None, "piece_recs": [], "split_check": 1234.56,
             "stats": {"fixed_mojibake": 0, "fixed_by_list": 0,
                       "unfixed": [f"SYPR68 Wholesale H{lost}{lost}yhastighet", f"Leverand{lost}r AS"]},
             "rec": {"status": "gap", "gap": -250.0, "cube_value": -1000.0}}
    none = [{"cell": {"p": 1}, "cube_value": -5.0, "items": [], "view": None, "secs": 0, "note": "no report"},
            {"cell": {"p": 2}, "cube_value": 0.0, "items": [], "view": None, "secs": 0, "note": "empty cell (skipped)"}]
    texts = []
    for pieces in (split, none):
        async def cell(conn, c, max_lines, how, pieces=pieces):
            return pieces, -301.0
        monkeypatch.setattr(drillstep, "_drill_cell", cell)
        monkeypatch.setattr(drillstep, "_lines_of", lambda p, view, parent_cube: lines)
        got = asyncio.run(drillstep.drill({"All Accounts": "ac.C40S", "All Time Periods": "Q3 26"}, say=texts.append))
        texts += got.warnings
    return texts


def _rest(tmp_path, monkeypatch) -> list[str]:
    """epm/rest.py's warning for a member the atlas has no label for (test_rest's faked Planning REST)."""
    from test_rest import ATLAS, _Epm, _grid
    monkeypatch.setattr(rest, "_listed", {})
    d = atlas.connect(tmp_path / "atlas.db")
    try:
        for conn, dim, members in ATLAS:
            atlas.store(d, conn, dim, [{"name": n, "label": lab, "parent": None if i == 0 else dim,
                                        "depth": int(i > 0), "order": i} for i, (n, lab) in enumerate(members)], 1.0)
        return _grid(_Epm(), d, accounts=("ac.C30",), entities=("un.SY.E1", "un.SY.E9")).warnings
    finally:
        d.close()


def _key_points() -> list[str]:
    block = CommentaryBlock(type="commentary", title="What drove EBITDA?", key_points=list("abcdefg"))
    return free_points(SimpleNamespace(blocks=[block]))[1]


def _sign_in(monkeypatch) -> list[str]:
    monkeypatch.setattr(oauth, "pending_logins", lambda: [
        {"pod": "https://plan1-a123456.epm.eu-frankfurt-1.ocs.oraclecloud.com",
         "verification_uri": "https://idcs-0a1b2c3d4e.identity.oraclecloud.com/ui/v1/device", "user_code": "QWER-TYUI",
         "minutes_left": 9, "note": "The earlier sign-in did not complete (expired_token)."},
        {"pod": "https://stratmodel.example", "error": "the identity domain did not answer (URLError)"}])
    return [oauth.sign_in_message(), oauth.NOT_SET_UP]


def _atlas() -> list[str]:
    reasons = [atlas_build.NEW_DIMENSION.format(dim="Product"), atlas_build.NOT_SIGNED_IN,
               "another atlas build is running", "1 dimension or cube failed", "3 dimensions or cubes failed",
               "nothing was built (no such dimension?)", "atlas.db could not be read (database is locked)"]
    return ([f"atlas not replaced: {w}" for w in reasons]
            + [f"atlas: NOT replaced ({w}); 0 built, 1 failed (C:\\Users\\u\\AppData\\Local\\smartass\\atlas.db)"
               for w in reasons])


def _login_refusal(monkeypatch) -> list[str]:
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [])
    return [str(dialogs.login_refusal(22, ["Budget 2026.xlsx - Excel", "Book1"])), str(dialogs.login_refusal(7))]


def _next_steps() -> list[str]:
    errors = [svhttp.NoScope("x"), svhttp.AtlasStale("x"), svhttp.NoToken("x", transient=True), svhttp.NoToken("x"),
              svhttp.GridRefused("x"), svhttp.TokenRefused("x"), RuntimeError("x")]
    return ([drillstep._link_advice(e) for e in errors] + [drillstep.HTTP_OFF] + list(serverkit.NEXT_STEP.values())
            + [serverkit.GENERIC_NEXT])


@pytest.fixture
def corpus(monkeypatch, tmp_path) -> list[str]:
    texts = (_knowledge() + _pipeline() + _provider(monkeypatch) + _drill(monkeypatch) + _rest(tmp_path, monkeypatch)
             + _key_points() + _sign_in(monkeypatch) + _atlas() + _login_refusal(monkeypatch) + _next_steps())
    return [t for t in texts if t]


def test_the_corpus_holds_every_kind(corpus):
    joined = "\n".join(corpus)
    assert len(corpus) > 80
    for part in ("C30 REVENUES", "second engine", "could not be fetched", "was left out", "QWER-TYUI",
                 "NOT replaced", "Smart View's Login is open", "smartass login --force", "ac.C50S", "un.SY.B2",
                 "Smart View did not know ac.C99", "FinMax: no data for", "#NoAccess (left out)",
                 "lost (cold sign-in) and retried", "Text broken in EPM", "Lines differ from the cube by",
                 "split pieces' cube values differ", "No drill-through lines", "All Entities not given",
                 "has no label for All Entities member un.SY.E9", "7 key points, the first 5"):
        assert part in joined, part


def test_every_warning_caveat_and_note_passes_the_scrub_unchanged(corpus):
    changed = [(t, logsetup.redact(t)) for t in corpus if logsetup.redact(t) != t]
    assert changed == []
    assert [t for t in corpus if len(t) <= svxml.DESC_MAX and svxml._scrubbed(t) != t] == []
