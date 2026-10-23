"""Smart View dialog classification and the restart decision (no Win32 calls)."""
from smartview_mcp import dialogs
from smartview_mcp.dialogs import Dialog


def test_classify_webview2_fallback_message_box():
    text = "WebView2 Runtime failed to initialize. Switching to Internet Explorer."
    assert dialogs.classify("#32770", "Oracle Smart View for Office", text) == "webview2_fallback"


def test_classify_ie_script_error_as_webview2_fallback():
    assert dialogs.classify("#32770", "Script Error", "Function expected") == "webview2_fallback"


def test_classify_login_window():
    assert dialogs.classify("WindowsForms10.Window.8.app.0.1", "Login", "") == "login"
    assert dialogs.classify("#32770", "Login", "") == "login"


def test_classify_other_dialogs_as_message():
    assert dialogs.classify("#32770", "Oracle Smart View for Office", "Empty response from provider") == "message"
    assert dialogs.classify("NUIDialog", "Microsoft Excel", "") == "message"


def test_classify_ignores_ordinary_windows():
    assert dialogs.classify("XLMAIN", "Book1 - Excel", "") is None
    assert dialogs.classify("Chrome_WidgetWin_1", "", "") is None


def _login(hwnd=1, ie=False):
    return Dialog(hwnd, "#32770", "Login", "", "login", ie)


def test_fresh_webview2_login_is_left_to_the_sso_watcher():
    first_seen: dict[int, float] = {}
    assert dialogs.blocking([_login()], first_seen, now=100.0, grace=45) is None
    assert dialogs.blocking([_login()], first_seen, now=144.0, grace=45) is None


def test_login_still_open_after_grace_blocks():
    first_seen: dict[int, float] = {}
    dialogs.blocking([_login()], first_seen, now=100.0, grace=45)
    assert dialogs.blocking([_login()], first_seen, now=146.0, grace=45) == _login()


def test_internet_explorer_login_blocks_at_once():
    assert dialogs.blocking([_login(ie=True)], {}, now=0.0, grace=45) == _login(ie=True)


def test_webview2_fallback_blocks_at_once_and_message_never_does():
    msg = Dialog(2, "#32770", "Smart View", "Empty response from provider", "message")
    fb = Dialog(3, "#32770", "Smart View", "WebView2 Runtime failed", "webview2_fallback")
    assert dialogs.blocking([msg], {}, now=0.0, grace=45) is None
    assert dialogs.blocking([msg, fb], {}, now=0.0, grace=45) == fb


def test_describe_quotes_title_and_text():
    d = Dialog(2, "#32770", "Smart View", "Empty response from provider", "message")
    assert dialogs.describe(d) == "'Smart View: Empty response from provider'"
    assert dialogs.describe(_login()) == "'Login'"


def test_webview2_holders_names_other_apps_on_smart_views_profile():
    sv = r"C:\Users\x\AppData\Roaming\Oracle\SmartView\WebView\EBWebView"
    rows = [
        (10, 96396, f'msedgewebview2.exe --webview-exe-name=POWERPNT.EXE --user-data-dir="{sv}" --noerrdialogs'),
        (11, 10, f'msedgewebview2.exe --type=renderer --user-data-dir="{sv}"'),          # child, not a browser
        (12, 555, f'msedgewebview2.exe --webview-exe-name=EXCEL.EXE --user-data-dir="{sv}"'),   # our own Excel
        (13, 7, r'msedgewebview2.exe --webview-exe-name=OneDrive.exe --user-data-dir="C:\x\OneDrive\EBWebView"'),
    ]
    assert dialogs.webview2_holders_from(rows, exclude_parent=555) == ["POWERPNT.EXE"]
