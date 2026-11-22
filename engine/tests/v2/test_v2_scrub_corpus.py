"""Every warning, caveat and note text the engine can put in a reply passes the scrub unchanged (Task 6b): the scrub
(logsetup.redact; svxml's provider error text goes through the same one) runs on reply warnings and log lines, so a
rule that matched ordinary text would garble them. The corpus: every caveat text of knowledge/semantic/facts, every
concept's notes and item notes (concept_caveats), the draft caveat, the pipeline's warnings (a grid neither engine
could fetch, the second engine, a member left out) as the offline fixtures make them with member codes, the sign-in
prompt (a made-up page and code), the atlas's "NOT replaced" reasons, the Login refusal, the drill's next steps and
the next steps of every error class. Nothing reaches EPM, the identity domain or Excel."""
from pathlib import Path
from types import SimpleNamespace

import pytest
from test_v2_fallback_flow import _Rest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.core import logsetup, serverkit
from smartass.epm import drillstep, oauth, svhttp, svxml
from smartass.epm.svhttp import SvHttpError
from smartass.knowledge import resolver as kresolver
from smartass.knowledge.semantic import load_concepts, load_semantic
from smartass.model.blocks.concept import concept_caveats
from smartass.model.spec import ConceptBlock
from smartass.ops import atlas_build
from smartass.pipeline import left_out, prepare
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
def corpus(monkeypatch) -> list[str]:
    texts = _knowledge() + _pipeline() + _sign_in(monkeypatch) + _atlas() + _login_refusal(monkeypatch) + _next_steps()
    return [t for t in texts if t]


def test_the_corpus_holds_every_kind(corpus):
    joined = "\n".join(corpus)
    assert len(corpus) > 80
    for part in ("C30 REVENUES", "second engine", "could not be fetched", "was left out", "QWER-TYUI",
                 "NOT replaced", "Smart View's Login is open", "smartass login --force", "ac.C50S", "un.SY.B2"):
        assert part in joined, part


def test_every_warning_caveat_and_note_passes_the_scrub_unchanged(corpus):
    changed = [(t, logsetup.redact(t)) for t in corpus if logsetup.redact(t) != t]
    assert changed == []
    assert [t for t in corpus if len(t) <= svxml.DESC_MAX and svxml._scrubbed(t) != t] == []
