# engine/tests/v2/test_v2_display_labels.py
"""Lines built from a metric token show a display name: the semantic metric name for an id, alias or name; for a
raw member the metric whose only member in that cube it is, else the atlas label. Key values follow the labels."""
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.model.blocks import common
from smartass.model.resolver import StaticResolver
from smartass.pipeline import prepare


class IdResolver(StaticResolver):
    """A resolver that knows metric ids (as knowledge.db does): 'opex_id' -> Opex."""
    IDS = {"opex_id": "Opex", "revenue_id": "Revenue"}

    def metric(self, name):
        n = self.IDS.get(name, name).lower()
        return next((m for m in self.metrics if m.name.lower() == n), None)


def _by_id(fn):
    def method(self, metric, *a):
        return getattr(StaticResolver, fn)(self, self.IDS.get(metric, metric), *a)
    return method


for _fn in ("members", "sign", "status", "materiality", "pct_floor", "balance", "aggregation"):
    setattr(IdResolver, _fn, _by_id(_fn))


def _resolver():
    base = resolver()
    return IdResolver(metrics=base.metrics, tree=base.tree, labels={**base.labels, "ac.C60": "C60 OTHER OPEX"})


def test_metric_label_rules():
    r = _resolver()
    assert common.metric_label(r, "opex_id", "FinMax") == "Opex"              # a metric id
    assert common.metric_label(r, "opex", "FinMax") == "Opex"                 # a name in any case
    assert common.metric_label(r, "ac.C30", "FinMax") == "Revenue"            # the only member of Revenue in FinMax
    assert common.metric_label(r, "ac.C60", "FinMax") == "C60 OTHER OPEX"     # one of Opex's members: the atlas label
    assert common.metric_label(r, "ac.X1", "FinMax") == "ac.X1"               # unknown: as given


def test_answer_lines_and_key_values_use_display_names():
    spec = spec_dict(
        queries=[{"id": "fm", "cube": "FinMax", "metrics": ["ac.C30", "opex_id", "ac.C60", "ac.EBITDA"],
                  "entity": "un.SY.2BUS", "periods": "closed", "compare": ["ly"]}],
        blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["ac.EBITDA", "opex_id"]},
                {"type": "table", "query": "fm", "title": "P&L"},
                {"type": "bridge", "query": "fm", "total": "ac.EBITDA", "steps": ["ac.C30", "opex_id"]},
                {"type": "trend", "query": "fm", "metrics": ["ac.C30"]},
                {"type": "ranking", "query": "fm", "metric": "ac.C60", "top": 1},
                {"type": "drivers", "query": "fm", "metric": "opex_id", "depth": 1}])
    pre = prepare(spec, _resolver(), FakeProvider(facts()))
    outs = {o.type: o for o in pre.result.outs}
    assert [ln.label for ln in outs["kpi_tiles"].lines] == ["EBITDA", "Opex"]
    assert [ln.label for ln in outs["table"].lines] == ["Revenue", "Opex", "EBITDA", "C60 OTHER OPEX"]  # chain first
    assert outs["bridge"].title == "EBITDA bridge"
    assert [ln.label for ln in outs["bridge"].lines] == ["EBITDA Last year", "Revenue", "Opex",
                                                         "Other / unexplained", "EBITDA Actual"]
    assert outs["trend"].headers == ["Revenue Actual", "Revenue Last year"]
    assert outs["ranking"].title == "Top 1 by C60 OTHER OPEX"
    assert outs["drivers"].title == "Drivers of Opex"
    kv = pre.result.key_values
    assert "P&L · Revenue" in kv and "P&L · C60 OTHER OPEX" in kv and not any("ac.C30" in k for k in kv)
    assert not any("ac.EBITDA" in k or "opex_id" in k for k in kv)
    # the line still carries the token the data is picked by
    assert [ln.metric for ln in outs["table"].lines] == ["ac.C30", "opex_id", "ac.EBITDA", "ac.C60"]
