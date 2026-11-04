"""Smart View dialog classification and the restart decision (no Win32 calls)."""
from smartass.xl import dialogs
from smartass.xl.dialogs import Dialog


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


SV = r"--user-data-dir=C:\Users\x\AppData\Roaming\Oracle\SmartView\WebView\EBWebView"


def test_stale_smart_view_webview2_is_this_excels_own_or_orphaned():
    rows = [(10, 500, f"msedgewebview2.exe --embedded-browser-webview=1 {SV} --webview-exe-name=EXCEL.EXE"),
            (11, 10, f"msedgewebview2.exe --type=renderer {SV}"),                      # a child: ended with its tree
            (20, 777, f"msedgewebview2.exe {SV} --webview-exe-name=EXCEL.EXE"),        # its Excel is gone
            (30, 900, f"msedgewebview2.exe {SV} --webview-exe-name=POWERPNT.EXE"),     # a live other app: never
            (40, 500, r"msedgewebview2.exe --user-data-dir=C:\OneDrive\EBWebView")]    # not Smart View's folder
    alive = {500, 900}.__contains__
    assert dialogs.stale_sv_webview2_from(rows, 500, alive) == [10, 20]


def test_recover_only_when_the_login_already_fell_back_to_ie():
    web = Dialog(1, "#32770", "Login", "", "login", ie=False)
    ie = Dialog(2, "#32770", "Login", "", "login", ie=True)
    err = Dialog(3, "Internet Explorer_TridentDlgFrame", "Script Error", "", "webview2_fallback")
    assert dialogs.ie_fallback([web]) == []
    assert dialogs.ie_fallback([web, ie, err]) == [2, 3]


def test_smart_view_boxes_and_the_not_signed_in_verdict():
    box = Dialog(5, "#32770", "Oracle Smart View for Office", "Empty response from provider.", "message")
    other = Dialog(6, "#32770", "Microsoft Excel", "Something else", "message")
    assert dialogs.sv_boxes([box, other]) == [box]
    # per failed connect attempt: the box texts it raised
    assert not dialogs.not_signed_in([["Empty response from provider.", "Empty response from provider."]])
    assert not dialogs.not_signed_in([["Empty response from provider."], ["other"]])
    assert dialogs.not_signed_in([["Empty response from provider."], [], ["Empty response from provider."]])
    assert dialogs.ok_only(["OK"]) and not dialogs.ok_only(["OK", "Cancel"]) and not dialogs.ok_only(["Yes", "No"])

def test_a_hung_webview2_login_is_stuck_at_once():
    d = Dialog(21, "#32770", "Login", "", "login", False, hung=True)
    assert dialogs.blocking([d], {}, now=100.0, grace=30.0) is d


def test_a_responsive_young_webview2_login_still_gets_the_grace():
    d = Dialog(22, "#32770", "Login", "", "login", False)
    assert dialogs.blocking([d], {}, now=100.0, grace=30.0) is None
