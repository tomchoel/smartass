"""Dialogs in the user's Excel: classification and the text that names one (no Win32 calls)."""
from smartass.xl import dialogs
from smartass.xl.dialogs import Dialog


def test_classify_login_window():
    assert dialogs.classify("WindowsForms10.Window.8.app.0.1", "Login", "") == "login"
    assert dialogs.classify("#32770", "Login", "") == "login"


def test_classify_other_dialogs_as_message():
    assert dialogs.classify("#32770", "Oracle Smart View for Office", "Empty response from provider") == "message"
    assert dialogs.classify("NUIDialog", "Microsoft Excel", "") == "message"


def test_classify_ignores_ordinary_windows():
    assert dialogs.classify("XLMAIN", "Book1 - Excel", "") is None
    assert dialogs.classify("Chrome_WidgetWin_1", "", "") is None


def test_describe_quotes_title_and_text():
    d = Dialog(2, "#32770", "Smart View", "Empty response from provider", "message")
    assert dialogs.describe(d) == "'Smart View: Empty response from provider'"
    assert dialogs.describe(Dialog(1, "#32770", "Login", "", "login")) == "'Login'"


def test_only_the_generic_finder_and_the_login_guard_are_left():
    """Task 6: the WebView2 / Internet Explorer recovery and Smart View's message-box handling are gone."""
    for gone in ("blocking", "ie_fallback", "sv_boxes", "close_sv_boxes", "press_ok", "end_stale_sv_webview2",
                 "webview2_holders", "not_signed_in", "close"):
        assert not hasattr(dialogs, gone), gone
