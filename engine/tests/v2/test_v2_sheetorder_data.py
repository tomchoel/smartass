# engine/tests/v2/test_v2_sheetorder_data.py
from smartass.xl import sheetorder


def test_data_sheet_is_ours_natively():
    kind, key = sheetorder.classify("Q3 Data")
    assert kind == "data" and key[1] < sheetorder._RANK["FinMax"]
    names = ["Sources", "Q3 FinMax", "Q3 Data", "Mine", "Q3 Answer"]
    order = sheetorder.sheet_order(names, set(names))
    assert order.index("Q3 Answer") < order.index("Mine") < order.index("Q3 Data") < order.index("Q3 FinMax") \
        < order.index("Sources")
