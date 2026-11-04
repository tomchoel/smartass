"""Headless Microsoft Edge that reads Data Integration drill pages (proven 15.10.2026 with Chrome, see the sv_drill
spec; moved to Edge 22.10.2026 because an HKCU Chrome policy, UserDataDir, now forces every Chrome launch onto the
user's own profile, so a separate headless Chrome cannot start). Channels are tried in config.DRILL_CHANNELS order:
one blocked by a UserDataDir policy is skipped, one that fails to launch is retried once, then the next is tried.
A browser that died between or during drills is restarted (one retry of the drill).

Why this route: the drill link's sso_token alone does not pass the IDCS cloud gate, and Smart View's own execute
call opens the user's browser. Edge signs in silently with the Windows account after a click on 'Telenor SSO'
(Chrome needed the IT-policy 'Windows Accounts' extension; Playwright's bundled Chromium has neither and stops at
the Microsoft login). The drill page renders only after /HyperionPlanning/ was opened in the same tab.

Threading: all Playwright calls run on ONE worker thread (the sync API is thread-bound). The worker owns its
_Session; other threads never touch it. kill() kills processes (thread-safe), lets the worker reset its session,
and replaces the worker with a fresh thread if it stays hung.

Never visible: headless=True always. Our chrome.exe processes are recognised by the exact --user-data-dir of our
profile; the Playwright driver node.exe is recognised by pid + create time recorded when we started it. Only those
are ever killed - never the user's Chrome, Edge or other node processes.

Error messages never carry a URL query (the drill link holds an sso_token)."""
from __future__ import annotations

import concurrent.futures as cf
import itertools
import json
import logging
import os
import queue
import re
import threading
import time
from pathlib import Path
from typing import Callable

import psutil

from smartass.core import config
from .core import drilltime

try:
    from playwright.sync_api import Error as PwError
except ImportError:   # pragma: no cover - playwright is a hard dependency of the drill route
    PwError = Exception

log = logging.getLogger("smartview.drill")

SIGNIN = "/ui/v1/signin"
SIGNIN_FAILED = ("Sign-in to EPM did not complete - open EPM once in Edge and log in, then retry sv_drill. Nothing "
                 "was shown on screen.")
CRASH_RETRY_S = 45             # a browser restart + sign-in + capture after a crash needs this much of the deadline
MIN_RETRY_S = 15               # a second login attempt / page reload needs at least this much of the deadline left
LOGIN_NUDGE_S = 15             # a login page unchanged this long (not the sign-in form): reopen the EPM home once


def login_action(url: str, base: str, elapsed: float, unchanged: float, clicked: bool, nudged: bool,
                 left: float = float("inf")) -> str:
    """Next step of the headless SSO login: 'click' the Telenor SSO button, 'done' (on the EPM home), 'home'
    (reopen the home: a redirect such as /oauth2/callback has hung - seen for 60 s in the baseline) or 'wait'."""
    if SIGNIN in url:
        return "wait" if clicked else "click"
    if url.startswith(base + "/HyperionPlanning/"):
        return "done" if elapsed > 3 else "wait"
    return "home" if unchanged >= LOGIN_NUDGE_S and not nudged and left >= MIN_RETRY_S else "wait"


def load_timeout_ms(deadline: float, now: float) -> int:
    """Playwright timeout for a page load: at most 90 s and never past the drill's deadline (at least 1 s)."""
    return int(min(90.0, max(1.0, deadline - now)) * 1000)


def classify_stall(url: str, page_text: str) -> str:
    """Why a drill page never asked for its data: 'sso' (stuck in sign-in), 'blank' (nothing rendered), 'page'."""
    if any(k in url for k in ("login.microsoftonline", "/oauth2/", SIGNIN)):
        return "sso"
    return "blank" if not (page_text or "").strip() else "page"
KILL_WAIT_S = 15                 # kill(): how long the worker gets to return after its browser died
WANTED = ".wanted"               # in the profile folder: another smartass process asked the owner to let go
KEEPALIVE_WAIT_S = 10            # keep-alive: how long its GET of the EPM home may take
POLL_S = 0.01                    # how often the idle worker collects the answers of requests in flight

# In-tab requests that do not block the worker (request()): _FIRE_JS starts a fetch and returns at once; its answer
# lands in window.__smartass.done, which _TAKE_JS empties. The box's token changes when the tab navigates (the
# fetches in flight are then lost: their futures fail).
_FIRE_JS = """([id, u, o, ms]) => {
  const w = window;
  if (!w.__smartass) w.__smartass = {tok: Math.random().toString(36).slice(2), done: []};
  const box = w.__smartass;
  const opt = Object.assign({}, o, {credentials: 'include', signal: AbortSignal.timeout(ms)});
  fetch(u, opt).then(async r => { let t; try { t = await r.text(); } catch (e) { t = String(e); }
                                  box.done.push([id, r.status, t]); },
                     e => box.done.push([id, 0, String(e)]));
  return box.tok; }"""
_TAKE_JS = """() => { const box = window.__smartass; if (!box) return [null, []];
  const d = box.done; box.done = []; return [box.tok, d]; }"""
_ids = itertools.count(1)


def _fail(fut: cf.Future, exc: BaseException) -> None:
    if not fut.done():
        try:
            fut.set_exception(exc)
        except cf.InvalidStateError:   # cancelled meanwhile by its caller
            pass


def _relay(src: cf.Future, dst: cf.Future) -> None:
    """dst gets src's outcome (unless dst is already done, e.g. cancelled by its caller)."""
    if dst.done():
        return
    if src.cancelled():
        _fail(dst, cf.CancelledError())
    elif src.exception() is not None:
        _fail(dst, src.exception())
    else:
        try:
            dst.set_result(src.result())
        except cf.InvalidStateError:
            pass


class DrillError(Exception):
    pass


class DrillBusy(DrillError):
    """Another MCP server process owns the drill profile (its hidden browser is alive)."""


class TooManyLines(DrillError):
    def __init__(self, total: int):
        super().__init__(f"{total} lines")
        self.total = total


class CellTooWide(DrillError):
    """EPMFDM-140373: the cell spans more than 3 dimensions with over 1000 children each (summary drill refused).
    Narrow it (e.g. child entities) and drill again."""


# ------------------------------------------------------------------ pure helpers
def redact(text: str) -> str:
    """Drop every URL query string (and any stray sso_token) from text meant for messages or logs."""
    text = re.sub(r"sso_token=[^&\s'\"]*", "sso_token=...", str(text))
    return re.sub(r"\?[^\s'\"]*", "?...", text)


def _page_url(captured_url: str, off: int) -> str:
    """The captured getData URL with ONLY limit and offset replaced (every other query parameter kept)."""
    from urllib.parse import parse_qsl, quote, urlencode, urlsplit, urlunsplit
    u = urlsplit(captured_url)
    q = [(k, v) for k, v in parse_qsl(u.query, keep_blank_values=True) if k not in ("limit", "offset")]
    q[:0] = [("limit", str(config.DRILL_PAGE)), ("offset", str(off))]
    return urlunsplit((u.scheme, u.netloc, u.path, urlencode(q, quote_via=quote),
                       u.fragment))


def _payload(text: str) -> dict | None:
    try:
        pay = json.loads(text)["responsePayload"]
    except (ValueError, KeyError, TypeError):
        return None
    return pay if isinstance(pay, dict) else None


def _expired(status: int, pay: dict | None) -> bool:
    """A getData answer that means 'signed out': 401/403, a failed fetch (status 0: cross-origin redirect to the
    sign-in page), or a 200 without JSON (the sign-in page itself)."""
    return status in (0, 401, 403) or (status == 200 and pay is None)


def page_through(post: Callable[[int], tuple[int, str]], max_lines: int,
                 progress: Callable[[int, int | None], None] | None = None,
                 recover: Callable[[], None] | None = None) -> tuple[list[dict], int | None]:
    """Page getData by offset. post(offset) -> (http status, body text). Follows hasMore (totalResults may be one
    short, spec §9.2). Safety caps: stop at totalResults + one page, and TooManyLines past max_lines (also when
    totalResults is missing).
    On a signed-out answer recover() is called ONCE (re-login + recapture) and paging resumes at the same offset."""
    pg = Pager(max_lines, progress, recover is not None)
    while True:
        step = pg.take(*post(pg.off))
        if step == "recover":
            recover()
        elif step == "done":
            return pg.result()


class Pager:
    """page_through's bookkeeping, one getData answer at a time (also drill_direct's async loop). take() says what
    comes next: 'more' (post again at .off), 'recover' (sign in again, then post at the same .off) or 'done'."""

    def __init__(self, max_lines: int, progress: Callable[[int, int | None], None] | None = None,
                 can_recover: bool = True):
        self.max_lines, self.progress = max_lines, progress
        self.items: list[dict] = []
        self.off, self.total, self.recovered = 0, None, not can_recover

    def take(self, status: int, text: str) -> str:
        pay = _payload(text)
        if status != 200 or pay is None:
            if _expired(status, pay) and not self.recovered:
                log.info("drill session expired at offset %d (HTTP %s); signing in again", self.off, status)
                self.recovered = True
                return "recover"
            why = "without data" if pay is None else "with an error"
            raise DrillError(f"getData returned HTTP {status} {why} at offset {self.off}.")
        if self.total is None and pay.get("totalResults") is not None:
            self.total = int(pay["totalResults"])
            if self.total > self.max_lines:
                raise TooManyLines(self.total)
        got = pay.get("items") or []
        self.items += got
        if len(self.items) > self.max_lines:
            raise TooManyLines(max(len(self.items), self.total or 0))
        if self.progress:
            self.progress(len(self.items), self.total)
        if not got or not pay.get("hasMore"):
            return "done"
        if self.total is not None and len(self.items) >= self.total + config.DRILL_PAGE:
            log.warning("drill paging stopped at the safety cap (%d items, totalResults %d, hasMore still set)",
                        len(self.items), self.total)
            return "done"
        self.off += len(got)
        return "more"

    def result(self) -> tuple[list[dict], int | None]:
        if self.total is not None and len(self.items) != self.total:
            log.warning("drill paged %d items but totalResults was %d", len(self.items), self.total)
        return self.items, self.total


def _user_data_dir(cmdline: list[str] | None) -> str | None:
    for a in cmdline or []:
        if a.startswith("--user-data-dir="):
            return a.split("=", 1)[1].strip('"')
    return None


def _same_path(a: str, b) -> bool:
    return os.path.normcase(os.path.normpath(a)) == os.path.normcase(os.path.normpath(str(b)))


def is_our_chrome(name: str | None, cmdline: list[str] | None, profile) -> bool:
    """chrome.exe whose --user-data-dir IS our profile (exact path, not a prefix)."""
    if (name or "").lower() not in ("chrome.exe", "msedge.exe"):
        return False
    udd = _user_data_dir(cmdline)
    return bool(udd) and _same_path(udd, profile)


def kill_our_chrome(profile) -> list[int]:
    """Kill only OUR chrome.exe processes (see is_our_chrome)."""
    killed = []
    for pr in psutil.process_iter(["name", "cmdline"]):
        try:
            if is_our_chrome(pr.info["name"], pr.info["cmdline"], profile):
                pr.kill()
                killed.append(pr.pid)
        except (psutil.Error, OSError):
            pass
    return killed


_POLICY_KEYS = {"msedge": r"SOFTWARE\Policies\Microsoft\Edge", "chrome": r"SOFTWARE\Policies\Google\Chrome"}


def user_data_dir_policy(channel: str) -> str | None:
    """The UserDataDir browser policy for this channel (HKCU, then HKLM), or None. When set, the browser ignores
    --user-data-dir and hands off to the user's own running browser, so a separate headless one cannot start."""
    key = _POLICY_KEYS.get(channel)
    if not key:
        return None
    import winreg
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, key) as k:
                v = winreg.QueryValueEx(k, "UserDataDir")[0]
                if v:
                    return str(v)
        except OSError:
            pass
    return None


def browser_health() -> dict:
    """Which drill browsers can start (cheap: registry + file lookup; launches nothing)."""
    import shutil
    exe = {"msedge": [r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
                      r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"],
           "chrome": [r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
                      r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
                      r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"]}
    out = {}
    for ch in config.DRILL_CHANNELS:
        found = any(os.path.exists(os.path.expandvars(p)) for p in exe.get(ch, [])) or bool(shutil.which(ch))
        pol = user_data_dir_policy(ch)
        out[ch] = "ok" if found and not pol else ("blocked by UserDataDir policy" if pol else "not installed")
    return {"channels": out, "usable": [c for c, v in out.items() if v == "ok"]}


def _kill_driver(driver: list[tuple[int, float]]) -> list[int]:
    """Kill the Playwright driver node.exe we started (pid + create time recorded at start), if still alive."""
    killed = []
    for pid, ctime in driver or []:
        try:
            pr = psutil.Process(pid)
            if pr.name().lower() == "node.exe" and abs(pr.create_time() - ctime) < 0.01:
                pr.kill()
                killed.append(pid)
        except (psutil.Error, OSError):
            pass
    return killed


# ------------------------------------------------------------------ cross-process profile lock
class ProfileLock:
    """Exclusive lock on <profile>.lock (msvcrt byte lock; Windows frees it when the process dies). Held while OUR
    Chrome runs on the profile, so two MCP servers never share one Chrome profile and never kill each other's
    browser. Idempotent per instance; usable from any thread."""

    def __init__(self, profile):
        self.path = Path(str(profile) + ".lock")
        self._fd: int | None = None
        self._mu = threading.Lock()

    @property
    def held(self) -> bool:
        return self._fd is not None

    def acquire(self) -> bool:
        import msvcrt
        with self._mu:
            if self._fd is not None:
                return True
            self.path.parent.mkdir(parents=True, exist_ok=True)
            fd = os.open(self.path, os.O_RDWR | os.O_CREAT)
            try:
                os.lseek(fd, 0, os.SEEK_SET)
                msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
            except OSError:
                os.close(fd)
                return False
            self._fd = fd
            return True

    def release(self) -> None:
        import msvcrt
        with self._mu:
            fd, self._fd = self._fd, None
            if fd is None:
                return
            try:
                os.lseek(fd, 0, os.SEEK_SET)
                msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
            finally:
                os.close(fd)


# ------------------------------------------------------------------ session (owned by the worker thread)
class _Session:
    def __init__(self):
        self.pw = self.ctx = self.page = None
        self.ready = False
        self.driver: list[tuple[int, float]] = []
        self.plock: ProfileLock | None = None     # set while this session's browser runs on the profile
        self.channel: str | None = None           # browser that started (msedge / chrome)
        self.deadline = float("inf")               # sign-in + capture of the current drill end by then
        self.paging = False                        # the current drill got past its capture
        self.gen = 0                               # sign-ins so far: a failure seen before a newer one is stale
        self.inflight: dict[int, tuple] = {}       # request id -> (page, box token, future, deadline)
        self.profile: Path | None = None           # profile folder (holds the .wanted flag of another process)
        self.started_at = time.time()              # wall clock of the browser start (compared with .wanted's mtime)
        self.last_alive = time.monotonic()         # last keep-alive check (or browser start)
        self.last_used = time.monotonic()          # last job or request (the keep-alive does not count)
        self.warmup = False                        # the start-up warm-up is running: it never asks the owner to let go


class DrillBrowser:
    def __init__(self, base: str = config.EPM_BASE, profile=config.DRILL_PROFILE):
        self.base, self.profile = base.rstrip("/"), profile
        self.base_profile = profile                    # slot 1; _take_slot may move self.profile to '<name>_2' ...
        self.plock = ProfileLock(profile)
        self._lock = threading.Lock()          # guards self.q / self.thread / self.s against kill()'s replacement
        with self._lock:
            self._spawn()

    # ------------------------------------------------------------ worker thread
    def _spawn(self):
        """Start a fresh worker (caller holds self._lock)."""
        self.q: queue.Queue = queue.Queue()
        self.s = _Session()
        self.s.profile = Path(self.profile)
        self.thread = threading.Thread(target=self._loop,
                                       args=(self.q, self.s, config.DRILL_IDLE_S, self._keepalive),
                                       name="drill-chrome", daemon=True)
        self.thread.start()

    @staticmethod
    def _loop(q: queue.Queue, s: _Session, idle_s: float | None = None, keepalive=None):
        """Run jobs; while requests are in flight, collect their answers between jobs (every POLL_S when idle).
        Idle: keep-warm off closes the browser after idle_s (frees the profile lock for other sessions). Keep-warm
        on keeps it, runs keepalive(s) every DRILL_KEEPALIVE_S and closes it only when another smartass process
        asked for the profile (.wanted newer than this browser)."""
        while True:
            warm = config.DRILL_KEEP_WARM
            try:
                fn, fut = q.get(timeout=POLL_S if s.inflight else
                                (min(idle_s, 30) if warm and idle_s else idle_s))
            except queue.Empty:
                if s.inflight:
                    DrillBrowser._collect(s)
                elif warm:
                    DrillBrowser._idle_warm(s, keepalive)
                elif s.ctx is not None or s.plock is not None:
                    DrillBrowser._close(s)
                continue
            if fn is None:
                DrillBrowser._fail_inflight(s, "the hidden browser was shut down")
                break
            if fut.set_running_or_notify_cancel():
                s.last_used = time.monotonic()
                try:
                    fut.set_result(fn(s))
                except BaseException as e:   # noqa: BLE001 - forwarded to the caller
                    fut.set_exception(e)
                s.last_used = time.monotonic()
            if s.inflight:
                DrillBrowser._collect(s)

    @staticmethod
    def _idle_warm(s: _Session, keepalive) -> None:
        """Worker thread, idle with keep-warm on: yield the profile to a process that asked for it, else keep alive.
        Never raises."""
        try:
            if s.ctx is None and s.plock is None:
                return
            flag = s.profile / WANTED if s.profile is not None else None
            if (flag is not None and flag.exists() and flag.stat().st_mtime > s.started_at
                    and time.monotonic() - s.last_used >= config.DRILL_YIELD_IDLE_S):
                log.info("another smartass process wants the drill profile; closing the hidden browser")
                DrillBrowser._close(s)
                flag.unlink(missing_ok=True)
            elif (keepalive is not None and s.ctx is not None
                  and time.monotonic() - s.last_alive >= config.DRILL_KEEPALIVE_S):
                s.last_alive = time.monotonic()
                keepalive(s)
        except Exception as e:   # noqa: BLE001 - the keep-alive must never take the worker down
            log.warning("drill keep-alive: %s", redact(str(e).splitlines()[0] if str(e) else type(e).__name__)[:200])

    def _keepalive(self, s: _Session) -> None:
        """Worker thread: restart a dead browser / sign in again, then one GET of the EPM home through the tab; a
        failed or sign-in answer means the session expired: sign in again. Bounded by KEEPALIVE_WAIT_S."""
        s.deadline = time.monotonic() + config.DRILL_SETUP_S
        used = s.last_used
        self._ensure(s)
        fut = self._fire(s, s.page, self.base + "/HyperionPlanning/", {"method": "GET"}, KEEPALIVE_WAIT_S)
        end = time.monotonic() + KEEPALIVE_WAIT_S + 5
        while not fut.done() and time.monotonic() < end:
            self._collect(s)
            time.sleep(0.05)
        ok = False
        if fut.done() and not fut.cancelled() and fut.exception() is None:
            ok = fut.result()[0] == 200
        for rid in [r for r, e in s.inflight.items() if e[2] is fut]:
            del s.inflight[rid]
        if not ok:
            log.warning("drill keep-alive: the EPM session is gone; signing in again")
            s.ready = False
            s.deadline = time.monotonic() + config.DRILL_SETUP_S
            self._ensure(s)
        s.last_used = used

    def submit(self, fn: Callable[[_Session], object]) -> cf.Future:
        fut: cf.Future = cf.Future()
        with self._lock:
            self.q.put((fn, fut))
        return fut

    # ------------------------------------------------------------ requests in flight
    def request(self, prep: Callable[[_Session], tuple], timeout_s: float) -> cf.Future:
        """One fetch from a tab that does not hold the worker while EPM answers (many run at once). prep(s) runs
        on the worker and returns (page, url, fetch options); the future gives (http status, body text), or
        DrillError when the tab navigated, closed or no answer came in time (the fetch aborts itself after
        timeout_s). The returned future's .job is the worker job that starts it."""
        out: cf.Future = cf.Future()

        def start(s: _Session):
            if out.cancelled():
                return
            page, url, opts = prep(s)
            DrillBrowser._fire(s, page, url, opts, timeout_s).add_done_callback(lambda f: _relay(f, out))
        job = self.submit(start)
        job.add_done_callback(lambda f: (_fail(out, cf.CancelledError()) if f.cancelled()
                                         else _fail(out, f.exception()) if f.exception() is not None else None))
        out.job = job
        return out

    @staticmethod
    def _fire(s: _Session, page, url: str, opts: dict, timeout_s: float) -> cf.Future:
        """Worker thread: start the fetch in the tab; its answer is collected later (_collect)."""
        fut: cf.Future = cf.Future()
        rid = next(_ids)
        tok = page.evaluate(_FIRE_JS, [rid, url, opts, max(1, int(timeout_s * 1000))])
        s.inflight[rid] = (page, tok, fut, time.monotonic() + timeout_s + 5)
        s.last_used = time.monotonic()
        return fut

    @staticmethod
    def _collect(s: _Session) -> None:
        """Worker thread: hand finished answers to their futures (one evaluate per tab); requests on a tab that
        navigated or closed, or past their deadline, fail."""
        tabs: dict[int, tuple] = {}
        for rid, (page, _, _, _) in s.inflight.items():
            tabs.setdefault(id(page), (page, []))[1].append(rid)
        now = time.monotonic()
        for page, rids in tabs.values():
            try:
                tok, done = page.evaluate(_TAKE_JS)
                lost = "the browser tab navigated away before EPM answered"
            except Exception as e:   # noqa: BLE001 - closed tab or dead browser
                tok, done = None, []
                lost = f"the browser tab is gone ({redact(str(e).splitlines()[0] if str(e) else type(e).__name__)})"
            for rid, status, text in done:
                ent = s.inflight.pop(rid, None)
                if ent is not None and not ent[2].done():
                    try:
                        ent[2].set_result((int(status), text))
                    except cf.InvalidStateError:
                        pass
            for rid in rids:
                ent = s.inflight.get(rid)
                if ent is None:
                    continue
                if tok != ent[1]:
                    del s.inflight[rid]
                    _fail(ent[2], DrillError(lost[:200]))
                elif now > ent[3]:
                    del s.inflight[rid]
                    _fail(ent[2], DrillError("EPM did not answer in time"))

    @staticmethod
    def _fail_inflight(s: _Session, why: str) -> None:
        for _, _, fut, _ in s.inflight.values():
            _fail(fut, DrillError(why))
        s.inflight.clear()

    # ------------------------------------------------------------ session
    def _slots(self) -> list[Path]:
        """Profile folders this browser may use: the configured one, then '<name>_2', '<name>_3' ... so that every
        smartass server (one per Claude/Copilot session) has its own warm, signed-in browser."""
        base = Path(self.base_profile)
        return [base] + [base.with_name(f"{base.name}_{k}") for k in range(2, max(1, config.DRILL_PROFILE_SLOTS) + 1)]

    def _take_slot(self) -> bool:
        """Hold the lock of the first free profile slot (switching self.profile to it); False when all are taken."""
        if self.plock.acquire():
            return True
        for prof in self._slots():
            if Path(prof) == Path(self.profile):
                continue
            lock = ProfileLock(prof)
            if lock.acquire():
                log.info("drill browser: profile %s is in use by another smartass process; using %s",
                         Path(self.profile).name, prof.name)
                self.profile, self.plock = prof, lock
                return True
        return False

    def _start(self, s: _Session):
        from playwright.sync_api import sync_playwright
        if not self._take_slot():
            if not s.warmup:                               # a start-up warm-up never pushes the owner out
                try:                                       # ask the owner (idle worker) to free the profile
                    Path(self.profile).mkdir(parents=True, exist_ok=True)
                    (Path(self.profile) / WANTED).write_text(str(time.time()))
                except OSError:
                    pass
            if config.DRILL_KEEP_WARM:
                gone = f"the other session lets go after {config.DRILL_YIELD_IDLE_S // 60:.0f} idle minutes"
            else:
                gone = f"its hidden browser closes after {config.DRILL_IDLE_S // 60:.0f} idle minutes"
            raise DrillBusy(f"another sv_drill is running in a different session ({gone}); "
                            "try again in a minute.")
        s.plock = self.plock
        s.profile = Path(self.profile)
        try:
            self._kill_ours()                              # we own the profiles: anything still on them is an orphan
            me = psutil.Process()
            before = {c.pid for c in me.children(recursive=True)}
            s.pw = sync_playwright().start()
            s.driver = []
            for c in me.children(recursive=True):
                try:
                    if c.pid not in before and c.name().lower() == "node.exe":
                        s.driver.append((c.pid, c.create_time()))
                except psutil.Error:
                    pass
            s.ctx, s.channel = self._launch(s.pw)
            s.page = s.ctx.pages[0] if s.ctx.pages else s.ctx.new_page()
            s.started_at, s.last_alive = time.time(), time.monotonic()
        except BaseException:
            self._close(s)                             # no browser started: stop the driver, free the lock
            raise

    def _profile_for(self, channel: str) -> Path:
        p = Path(self.profile)
        return p if channel == "msedge" else p.with_name(f"{p.name}_{channel}")

    def _kill_ours(self) -> list[int]:
        """Kill our orphan browsers on every channel's profile (caller owns the profile lock)."""
        killed = []
        for ch in dict.fromkeys(config.DRILL_CHANNELS + ("msedge",)):
            killed += kill_our_chrome(self._profile_for(ch))
        if killed:
            log.info("killed %d orphan browser process(es) on the drill profiles", len(killed))
        return killed

    def _launch(self, pw):
        """First channel that starts: skip channels a UserDataDir policy hijacks, retry each once after killing
        orphans on its profile. Fails with one readable line per channel."""
        why = []
        for ch in config.DRILL_CHANNELS:
            pol = user_data_dir_policy(ch)
            if pol:
                why.append(f"{ch}: skipped, policy UserDataDir={pol} overrides our profile")
                continue
            prof = self._profile_for(ch)
            prof.mkdir(parents=True, exist_ok=True)
            for attempt in range(1, config.DRILL_LAUNCH_TRIES + 1):
                try:
                    ctx = pw.chromium.launch_persistent_context(
                        str(prof), channel=ch, headless=True,
                        ignore_default_args=["--disable-extensions"],         # the SSO extension must load
                        args=["--disable-blink-features=AutomationControlled"],
                        viewport={"width": 1600, "height": 1000})
                    if ch != config.DRILL_CHANNELS[0]:
                        log.warning("drill browser: fell back to %s (%s)", ch, "; ".join(why))
                    return ctx, ch
                except Exception as e:   # noqa: BLE001 - not installed, crashed, handed off: try again / next
                    msg = redact(str(e).splitlines()[0] if str(e) else type(e).__name__)[:200]
                    log.warning("drill browser %s launch %d failed: %s", ch, attempt, msg)
                    kill_our_chrome(prof)
                    if attempt == config.DRILL_LAUNCH_TRIES:
                        why.append(f"{ch}: {msg}")
                    else:
                        time.sleep(1)
        raise DrillError("No headless browser could start for the drill (" + " | ".join(why) + "). Check that "
                         "Microsoft Edge is installed; a browser policy may block a separate profile.")

    @staticmethod
    def _alive(s: _Session) -> bool:
        try:
            return s.page is not None and not s.page.is_closed() and s.page.evaluate("1") == 1
        except Exception:   # noqa: BLE001 - dead browser or driver
            return False

    @staticmethod
    def _goto(p, url: str, deadline: float = float("inf")):
        try:
            p.goto(url, wait_until="domcontentloaded", timeout=load_timeout_ms(deadline, time.monotonic()))
        except PwError as e:
            raise DrillError(f"Could not open {redact(url)}: {redact(str(e).splitlines()[0] if str(e) else e)}") \
                from None

    def _login_once(self, p, deadline: float, base: str | None = None):
        """Open the EPM home of `base` (default: this browser's pod; signs in via Telenor SSO if needed). The drill
        page needs its session state; the Smart View HTTP route signs in on other pods (FinPlan: stratmodel)."""
        base = (base or self.base).rstrip("/")
        home = base + "/HyperionPlanning/"
        self._goto(p, home, deadline)
        t0, clicked, nudged = time.monotonic(), False, False
        last, changed = None, t0
        end = min(t0 + config.DRILL_LOGIN_TIMEOUT_S, deadline)
        while time.monotonic() < end:
            url, now = p.url, time.monotonic()
            if url != last:
                last, changed = url, now
            act = login_action(url, base, now - t0, now - changed, clicked, nudged, deadline - now)
            if act == "click":
                b = p.get_by_role("button", name="Telenor SSO")
                if b.count():
                    b.first.click()
                    clicked = True
            elif act == "home":                     # e.g. stuck on /oauth2/callback: the session is usually set
                log.warning("drill SSO stalled on %s; reopening the EPM home", redact(url)[:120])
                nudged = True
                self._goto(p, home, deadline)
                last, changed = None, time.monotonic()
                continue
            elif act == "done":
                try:
                    p.wait_for_load_state("networkidle", timeout=min(15_000, load_timeout_ms(deadline, now)))
                except PwError:   # a busy home page is fine once we are on it
                    pass
                return
            p.wait_for_timeout(1000)
        raise DrillError(f"last page: {redact(p.url)[:120]}")

    def _login(self, s: _Session):
        """Sign in; on failure keep the profile and retry ONCE in a fresh tab (spec §6) if the drill's time allows."""
        s.ready = False
        if s.deadline - time.monotonic() < 1:
            raise DrillError("The drill ran out of time before it could sign in (the hidden browser was slow to "
                             "start or restarted) - retry sv_drill.")
        try:
            self._login_once(s.page, s.deadline)
        except DrillError as first:
            if s.deadline - time.monotonic() < MIN_RETRY_S:
                raise DrillError(f"{SIGNIN_FAILED} ({first})") from None
            log.warning("drill SSO attempt 1 failed (%s); retrying in a fresh tab", first)
            old, s.page = s.page, s.ctx.new_page()
            try:
                old.close()
            except PwError:
                pass
            try:
                self._login_once(s.page, s.deadline)
            except DrillError as second:
                raise DrillError(f"{SIGNIN_FAILED} (2 attempts; {second})") from None
        s.ready = True
        s.gen += 1

    def _ensure(self, s: _Session):
        if s.ctx is not None and not self._alive(s):   # crashed or killed since the last drill: start fresh
            log.warning("drill browser was dead; restarting it")
            self._close(s)
        if s.ctx is None:
            with drilltime.span("browser start"):
                self._start(s)
        if not s.ready:
            with drilltime.span("browser sign-in"):
                self._login(s)

    # ------------------------------------------------------------ one drill
    def _capture(self, s: _Session, url: str) -> dict:
        """Open the drill page and capture its own getData request and view definition. A sign-in redirect gets
        ONE re-login, after which the capture timer starts again."""
        got: dict = {}
        errs: list[str] = []

        def on_request(req):
            if "/aif/ui/model/drilldown/getData?limit=" in req.url and "req" not in got:
                got["req"] = {"url": req.url, "body": req.post_data, "headers": dict(req.headers)}

        def on_response(resp):
            if "/aif/ui/model/view?page=DRILL_THROUGH" in resp.url and "view" not in got:
                try:
                    got["view"] = resp.json()
                except Exception:   # noqa: BLE001
                    pass

        def on_console(msg):
            if msg.type == "error":
                errs.append(msg.text)

        def on_pageerror(err):
            errs.append(f"page error: {err}")
        handlers = {"request": on_request, "response": on_response, "console": on_console, "pageerror": on_pageerror}

        def attach(p):
            for ev, fn in handlers.items():
                p.on(ev, fn)

        def detach(p):
            for ev, fn in handlers.items():
                try:
                    p.remove_listener(ev, fn)
                except Exception:   # noqa: BLE001 - a closed tab
                    pass
        target = self.base + "/" + url.lstrip("./")
        p = s.page
        attach(p)
        try:
            self._goto(p, target, s.deadline)
            t0, relogged, reloaded, wait_s = time.monotonic(), False, False, config.DRILL_CAPTURE_TIMEOUT_S
            while True:
                if "req" in got and "view" in got:
                    break
                if time.monotonic() - t0 >= wait_s or time.monotonic() >= s.deadline:
                    if "req" in got:
                        break
                    kind = classify_stall(p.url, self._page_text(p))
                    if kind == "sso":
                        raise DrillError(f"{SIGNIN_FAILED} Page: {redact(p.url)[:120]}.")
                    if reloaded or s.deadline - time.monotonic() < MIN_RETRY_S:
                        break
                    log.warning("drill page never asked for its data (%s); reloading it once", kind)
                    reloaded, wait_s = True, config.DRILL_RELOAD_WAIT_S
                    try:
                        p.reload(wait_until="domcontentloaded",
                                 timeout=min(30_000, load_timeout_ms(s.deadline, time.monotonic())))
                    except PwError as e:            # fall through to the 'never requested data' error
                        log.warning("drill page reload failed: %s", redact(str(e).splitlines()[0] if str(e) else e))
                        break
                    t0 = time.monotonic()
                    continue
                if SIGNIN in p.url:                 # session expired: log in again (once), then reopen the drill
                    if relogged:
                        raise DrillError("EPM signed us out again right after a fresh SSO login. Nothing was shown "
                                         "on screen.")
                    relogged = True
                    self._login(s)
                    if s.page is not p:             # the login retried in a fresh tab
                        detach(p)
                        p = s.page
                        attach(p)
                    self._goto(p, target, s.deadline)
                    t0 = time.monotonic()
                    continue
                epm = self._epm_error(p)
                if epm:                             # the page reported a refusal (toast): no point waiting
                    if "EPMFDM-140373" in epm:
                        raise CellTooWide(epm)
                    raise DrillError(f"EPM refused the drill: {epm}")
                p.wait_for_timeout(500)
            if "req" not in got:
                tail = "; ".join(redact(e)[:160] for e in errs[-8:]) or "none"
                dump = self._dump_failure(p)
                raise DrillError("The drill page loaded but never requested data"
                                 + (" (also after one reload)" if reloaded else "") + " - retry "
                                 "sv_drill; if it repeats, the Oracle drill page may have changed. "
                                 f"Page: {redact(p.url)[:120]}. Console errors: {tail}. "
                                 f"Screenshot/HTML: {dump}")
        finally:
            detach(p)
        return got

    @staticmethod
    def _page_text(p) -> str:
        try:
            return str(p.evaluate("() => document.body ? document.body.innerText : ''") or "")
        except Exception:   # noqa: BLE001 - a navigating or closed tab: treat as blank
            return ""

    @staticmethod
    def _epm_error(p) -> str | None:
        """'EPMFDM-nnnnnn: text' when the drill page shows an EPM error toast, else None."""
        try:
            loc = p.get_by_text(re.compile(r"EPMFDM-\d+"))
            if not loc.count():
                return None
            return re.sub(r"\s+", " ", loc.first.inner_text(timeout=2000)).strip()[:400]
        except Exception:   # noqa: BLE001 - page navigating; check again next tick
            return None

    @staticmethod
    def _dump_failure(p) -> str:
        """Screenshot + HTML of a failed drill page to svkit\\logs (local only; sso_token redacted)."""
        try:
            d = config.SVKIT / "logs"
            d.mkdir(parents=True, exist_ok=True)
            stem = d / f"drill_fail_{time.strftime('%Y%m%d_%H%M%S')}"
            p.screenshot(path=str(stem) + ".png", full_page=True)
            html = re.sub(r"sso_token=[^&\s'\"]*", "sso_token=...", p.content())
            Path(str(stem) + ".html").write_text(html, encoding="utf-8")
            return str(stem) + ".png/.html"
        except Exception as e:   # noqa: BLE001
            return f"not saved ({type(e).__name__})"

    @staticmethod
    def _post(p, req: dict, off: int) -> tuple[int, str]:
        hdr = {k: v for k, v in req["headers"].items() if k.lower() not in ("cookie", "content-length", "host")}
        try:
            res = p.evaluate("""async ([u, b, h]) => { try { const r = await fetch(u, {method: 'POST', body: b,
                                    headers: h, credentials: 'include'}); return {s: r.status, t: await r.text()}; }
                                catch (e) { return {s: 0, t: String(e)}; } }""",
                             [_page_url(req["url"], off), req["body"], hdr])
        except PwError as e:                         # the tab navigated away (e.g. to the sign-in page)
            return 0, str(e)
        return int(res["s"]), res["t"]

    def _fetch(self, s: _Session, url: str, max_lines: int, progress) -> dict:
        """One drill; if the browser died under it, restart the browser and try ONCE more. Sign-in and capture
        share one deadline (DRILL_SETUP_S), so a stuck login fails with a next step instead of hanging."""
        s.deadline, s.paging = time.monotonic() + config.DRILL_SETUP_S, False
        try:
            return self._fetch_once(s, url, max_lines, progress)
        except (TooManyLines, CellTooWide, DrillBusy):
            raise
        except DrillError as e:
            if s.ctx is None or self._alive(s):     # a real answer from a live page: do not repeat it
                raise
            self._close(s)
            if s.paging:                            # long drills page past the setup deadline: start a new one
                s.deadline = time.monotonic() + config.DRILL_SETUP_S
            elif s.deadline - time.monotonic() < CRASH_RETRY_S:
                raise DrillError(f"The hidden drill browser crashed ({e}) with too little time left to start it "
                                 "again - retry sv_drill.") from None
            log.warning("drill browser died mid-drill (%s); restarting it and retrying once", e)
            return self._fetch_once(s, url, max_lines, progress)

    def _fetch_once(self, s: _Session, url: str, max_lines: int, progress) -> dict:
        try:
            self._ensure(s)
            got = self._capture(s, url)
            s.paging = True                         # captured: a crash from here on is mid-paging
            state = {"req": got["req"]}

            def recover():                          # spec §6: re-login, then resume from the last offset
                s.deadline = time.monotonic() + config.DRILL_SETUP_S
                self._login(s)
                state["req"] = self._capture(s, url)["req"]
            t1 = time.monotonic()
            items, total = page_through(lambda off: self._post(s.page, state["req"], off), max_lines, progress,
                                        recover)
        except PwError as e:                        # Playwright messages can quote the drill URL (sso_token)
            raise DrillError(redact(str(e).splitlines()[0] if str(e) else type(e).__name__)) from None
        return {"items": items, "total": total, "view": got.get("view"), "secs": round(time.monotonic() - t1, 1)}

    def fetch(self, url: str, max_lines: int = config.DRILL_MAX_LINES,
              progress: Callable[[int, int | None], None] | None = None) -> cf.Future:
        """Drill one link. progress(lines_so_far, total_or_None) is called on the WORKER thread, so it must be
        thread-safe and quick (e.g. put into a queue or set a value)."""
        return self.submit(lambda s: self._fetch(s, url, max_lines, progress))

    # ------------------------------------------------------------ cleanup
    @staticmethod
    def _close(s: _Session):
        """Worker thread: close the browser and ALWAYS stop the Playwright driver, even if closing fails."""
        DrillBrowser._fail_inflight(s, "the hidden browser closed")
        try:
            if s.ctx:
                s.ctx.close()
        except Exception as e:   # noqa: BLE001 - the browser may already be dead
            log.info("drill browser close: %s", redact(str(e))[:200])
        finally:
            try:
                if s.pw:
                    s.pw.stop()
            except Exception as e:   # noqa: BLE001
                log.info("playwright stop: %s", redact(str(e))[:200])
            finally:
                _kill_driver(s.driver)               # no-op when stop() ended it
                s.pw = s.ctx = s.page = None
                s.driver = []
                s.ready = False
                if s.plock is not None:              # our Chrome is gone: other sessions may use the profile
                    s.plock.release()
                    s.plock = None

    def _abandon_worker(self):
        """The worker is hung: cancel its queued jobs, tell it to exit when it wakes, and start a fresh one.
        Under the lock, so no submit() can land in the old queue after it was drained."""
        with self._lock:
            old_q, old_s = self.q, self.s
            while True:
                try:
                    fn, fut = old_q.get_nowait()
                except queue.Empty:
                    break
                if fut is not None:
                    fut.cancel()
            old_q.put((None, None))
            try:                                    # their answers would never be collected
                hung = list(old_s.inflight.values())
            except RuntimeError:                    # the hung worker changed it meanwhile: its own close fails them
                hung = []
            for _, _, fut, _ in hung:
                _fail(fut, DrillError("the hidden browser hung and was replaced"))
            _kill_driver(old_s.driver)              # read-only look at the hung worker's state
            self._spawn()

    def shutdown(self):
        try:
            self.submit(self._close).result(timeout=30)
        except Exception:   # noqa: BLE001
            if self.plock.held:                     # never another server's live browser
                self._kill_ours()
            _kill_driver(self.s.driver)
            self.plock.release()
        with self._lock:
            self.q.put((None, None))

    def kill(self) -> list[int]:
        """Hang recovery: kill only OUR chrome.exe (by profile) and our driver; the next fetch starts fresh."""
        killed = self._kill_ours() if self.plock.held else []   # only when the profile is OURS
        try:
            self.submit(self._close).result(timeout=KILL_WAIT_S)   # the dead browser lets the hung call return
        except Exception:   # noqa: BLE001 - still hung or failed: replace the worker
            self._abandon_worker()
        return killed


_shared: DrillBrowser | None = None
_shared_lock = threading.Lock()


def shared_browser() -> DrillBrowser:
    """The process's one drill browser: sv_drill and the Smart View HTTP route (svhttp) share it, its worker thread
    and its profile lock."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = DrillBrowser()
        return _shared


def shutdown_shared() -> None:
    """Shut the process's shared drill browser down (if one was made) and forget it, so a later shared_browser()
    builds a fresh one. Safe to call twice."""
    global _shared
    with _shared_lock:
        b, _shared = _shared, None
    if b is not None:
        b.shutdown()
