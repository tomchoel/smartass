# engine/tests/v2/test_v2_legacy_sheetorder.py
from smartass.xl import sheetorder

from smartass import legacy


def test_data_sheet_is_classified_as_ours(monkeypatch):
    import re
    # undo any earlier configure_v2() in this process, restored after the test
    monkeypatch.setattr(sheetorder, "_OURS", re.compile(r"^Q(d+) (Answer|FinMax|FinLite|FinPlan|Lines)(?: (.+))?$"))
    monkeypatch.setattr(sheetorder, "_RANK", {k: v for k, v in sheetorder._RANK.items() if k != "Data"})
    assert sheetorder.classify("Q3 Data")[0] == "other"
    legacy.teach_sheetorder()
    kind = sheetorder.classify("Q3 Data")[0]
    assert kind != "other"
    names = ["Sources", "Q3 Data", "Mine", "Q3 Answer"]
    order = sheetorder.sheet_order(names, set(names))                   # real API: (names, visible, mode=...)
    assert order.index("Q3 Answer") < order.index("Mine") < order.index("Q3 Data") < order.index("Sources")
