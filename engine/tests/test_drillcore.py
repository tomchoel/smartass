"""drillcore: the drill's pure parts (errors, URL building, parsing, redaction, paging), figures invented."""
import json

import pytest

from smartass.epm import drillcore as dc
from smartass.epm.drillcore import CellTooWide, DrillError, redact

URL = "aif/m/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=Drill&sso_token=T&attribute=Entities.id.un.SY"


def test_sui_url_drops_drillurl_and_asks_for_json():
    u = dc.sui_url("https://epm", "./" + URL, 123)
    assert u == ("https://epm/aif/m/FdmeeMain.jsp?module=aif.drilldown&sso_token=T&attribute=Entities.id.un.SY"
                 "&FORMAT=SUI&_=123")


def test_sui_url_passes_values_through_byte_for_byte():
    u = dc.sui_url("https://epm/", "./aif/x.jsp?a=b%20c&DRILLURL=Drill&t=A+B/c=%2F&e=", 5)
    assert u == "https://epm/aif/x.jsp?a=b%20c&t=A+B/c=%2F&e=&FORMAT=SUI&_=5"


def _sui(**kw):
    return json.dumps({"DRILL_LOG": "/tmp/x", "APPLICATION_ID": 52, "PARTITIONKEY": 39, "status": 0,
                       "QUERY": "SELECT 1", "PARAMSARRAY": ["39", "un.SY"], "PARAMSTYPEARRAY": ["BigDecimal", "String"],
                       **kw})


def test_parse_sui_gives_epms_query_unchanged():
    q = dc.parse_sui(_sui())
    assert q.app_id == 52 and q.partition == 39
    assert json.loads(q.body) == {"QUERY": "SELECT 1", "PARAMS": ["39", "un.SY"],
                                  "PARAMSTYPE": ["BigDecimal", "String"]}


def test_parse_sui_no_data_is_a_note_not_an_error():
    with pytest.raises(dc.NoDrillData, match="not loaded to this intersection"):
        dc.parse_sui(json.dumps({"status": 1, "DATA": [], "details": "Drill query does not return any data. "
                                 "Data was not loaded to this intersection using Data Integration."}))


def test_parse_sui_too_wide():
    with pytest.raises(CellTooWide, match="EPMFDM-140373"):
        dc.parse_sui(json.dumps({"status": 1, "details": "EPMFDM-140373:The drill from summary operation ..."}))


def test_parse_sui_not_json_means_signed_out():
    with pytest.raises(dc.SignedOut):
        dc.parse_sui("<html>epmexit.jsp</html>")


def test_parse_sui_other_refusal():
    with pytest.raises(DrillError, match="boom"):
        dc.parse_sui(json.dumps({"status": 1, "details": "boom"}))


@pytest.mark.parametrize("status, text, gone", [
    (0, "", True), (301, "", True), (302, "", True), (303, "", True), (307, "", True), (401, "", True),
    (403, "", True), (200, "<html>sign in</html>", True), (200, '{"responsePayload": {}}', False),
    (500, "", False), (404, "", False)])
def test_expired_takes_a_redirect_to_the_exit_page_as_a_lost_session(status, text, gone):
    assert dc.expired(status, dc.payload(text)) is gone


def test_redact_strips_queries_and_tokens():
    s = "Could not open https://x.oraclecloud.com/aif/FdmeeMain.jsp?module=a&sso_token=SECRET&attribute=b: timeout"
    out = redact(s)
    assert "SECRET" not in out and "module=" not in out and out.startswith("Could not open https://x.oraclecloud")
    assert "SECRET" not in redact("token sso_token=SECRET&x=1 in a console line")

# ------------------------------------------------------------------ Pager: getData paging bookkeeping
def _page_through(post, max_lines, progress=None, recover=None):
    """The paging loop drill_token runs around Pager, synchronously: post(offset) -> (status, text)."""
    pg = dc.Pager(max_lines, progress, recover is not None)
    while True:
        step = pg.take(*post(pg.off))
        if step == "recover":
            recover()
        elif step == "done":
            return pg.result()


def _pages(total, sizes, has_more_always=False, include_total=True):
    """Fake getData: answers per offset from a list of page sizes."""
    pages, off = {}, 0
    for i, n in enumerate(sizes):
        pay = {"items": [{"AMOUNTX": 1.0}] * n, "hasMore": has_more_always or i < len(sizes) - 1}
        if include_total:
            pay["totalResults"] = total
        pages[off] = json.dumps({"responsePayload": pay})
        off += n
    calls = []
    extra = {"items": [{"AMOUNTX": 1.0}], "hasMore": True}
    if include_total:
        extra["totalResults"] = total

    def post(o):
        calls.append(o)
        return 200, pages.get(o, json.dumps({"responsePayload": extra}))
    return post, calls


def test_pager_pages_until_total():
    post, calls = _pages(1231, [500, 500, 231])
    seen = []
    items, total = _page_through(post, 300000, lambda n, t: seen.append((n, t)))
    assert (len(items), total, calls) == (1231, 1231, [0, 500, 1000])
    assert seen[-1] == (1231, 1231)


def test_pager_follows_has_more_past_a_short_total(caplog):
    post, calls = _pages(81, [81, 1])             # spec §9.2: totalResults may be one short
    items, total = _page_through(post, 300000)
    assert len(items) == 82 and total == 81 and calls == [0, 81]
    assert "totalResults was 81" in caplog.text


def test_pager_safety_cap_when_has_more_never_ends(monkeypatch):
    from smartass.core import config
    monkeypatch.setattr(config, "DRILL_PAGE", 10)
    post, calls = _pages(81, [81], has_more_always=True)   # then 1-line pages with hasMore forever
    items, total = _page_through(post, 300000)
    assert len(items) == 81 + 10 and len(calls) == 11


def test_a_piece_up_to_the_split_size_is_one_getdata_call():
    """EPM's offset pages have no stable order: a 623-line cell paged by 500 came back with duplicates and a
    different sum each time (measured 02.11.2026). So any piece the split allows is fetched in ONE call."""
    from smartass.core import config
    n, offsets = config.DRILL_SPLIT_AT, []
    assert n <= config.DRILL_PAGE

    def post(off):
        limit = config.DRILL_PAGE                           # the getData limit drill_token asks for
        offsets.append(off)
        got = min(limit, n - off)
        return 200, json.dumps({"responsePayload": {"items": [{"AMOUNTX": "-1"}] * got, "totalResults": n,
                                                    "hasMore": off + got < n}})
    items, total = _page_through(post, config.DRILL_MAX_LINES)
    assert len(items) == n and offsets == [0]


def test_pager_limit_holds_without_total():
    post, _ = _pages(None, [500] * 10, has_more_always=True, include_total=False)
    with pytest.raises(dc.TooManyLines) as e:
        _page_through(post, 1200)
    assert e.value.total >= 1200


def test_pager_limit_from_total():
    post, calls = _pages(5000, [500], has_more_always=True)
    with pytest.raises(dc.TooManyLines) as e:
        _page_through(post, 1000)
    assert e.value.total == 5000 and calls == [0]


def test_pager_http_error_with_json_fails():
    body = json.dumps({"responsePayload": {"items": [], "totalResults": 0}})
    with pytest.raises(DrillError, match="HTTP 500"):
        _page_through(lambda o: (500, body), 1000, recover=lambda: None)


def test_pager_recovers_once_and_resumes_at_the_offset():
    post, calls = _pages(1231, [500, 500, 231])
    state = {"expired": True, "recovered": 0}

    def flaky(o):
        if o == 500 and state["expired"]:
            calls.append(o)
            return 0, "TypeError: Failed to fetch"
        return post(o)

    def recover():
        state["expired"] = False
        state["recovered"] += 1
    items, total = _page_through(flaky, 300000, recover=recover)
    assert len(items) == 1231 and state["recovered"] == 1 and calls == [0, 500, 500, 1000]


def test_pager_a_second_lost_session_fails():
    n = {"r": 0}
    with pytest.raises(DrillError, match="HTTP 401"):
        _page_through(lambda o: (401, "<html>sign in</html>"), 1000, recover=lambda: n.__setitem__("r", n["r"] + 1))
    assert n["r"] == 1

def test_pager_steps_one_answer_at_a_time():
    pg = dc.Pager(10, can_recover=True)
    page = lambda items, more, total=None: json.dumps({"responsePayload": {"items": items, "hasMore": more,  # noqa: E731
                                                                          "totalResults": total}})
    assert pg.take(200, page([1, 2], True, 3)) == "more" and pg.off == 2
    assert pg.take(200, "<html>exit</html>") == "recover" and pg.off == 2
    assert pg.take(200, page([3], False)) == "done"
    assert pg.result() == ([1, 2, 3], 3)
    with pytest.raises(DrillError, match="HTTP 200 without data at offset 0"):
        dc.Pager(10, can_recover=False).take(200, "<html/>")
    with pytest.raises(dc.TooManyLines):
        dc.Pager(2).take(200, page([1], True, 3))
