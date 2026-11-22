"""xl.hostproc._PidRegistry, an older release's record of the hidden Excels its engine started: saves are atomic and
serialised, so pids.json is never half-written; the sweep at server start (sweep_leftovers) kills only what the
registry proves ours and removes the file once no entry is left."""
import json
import threading
import time

from smartass.xl.hostproc import _PidRegistry


def test_concurrent_save_never_leaves_invalid_json(tmp_path):
    path = tmp_path / "pids.json"
    reg = _PidRegistry(path)
    reg.save([])
    stop = threading.Event()
    bad: list[str] = []

    def writer(n: int):
        items = [{"pid": n * 1000 + i, "ctime": 1.0, "owner": n, "owner_ctime": 2.0, "pad": "x" * 300}
                 for i in range(20)]
        for _ in range(60):
            reg.save(items)

    def reader():
        while not stop.is_set():
            try:
                raw = path.read_text()
            except OSError:
                continue                      # a reader may briefly lose the race with os.replace on Windows
            try:
                json.loads(raw)
            except ValueError:
                bad.append(raw[:40])
            time.sleep(0.001)

    r = threading.Thread(target=reader)
    r.start()
    ws = [threading.Thread(target=writer, args=(n,)) for n in (1, 2)]
    for w in ws:
        w.start()
    for w in ws:
        w.join()
    stop.set()
    r.join()
    assert not bad, f"{len(bad)} reads saw invalid JSON, e.g. {bad[0]!r}"
    assert len(json.loads(path.read_text())) == 20
    assert not list(tmp_path.glob("*.tmp"))


def _entry(pid, owner=None):
    return {"pid": pid, "ctime": 1.0, "owner": owner, "owner_ctime": 2.0}


def test_the_sweep_kills_only_proven_orphans_and_then_removes_the_file(tmp_path, monkeypatch):
    import smartass.xl.hostproc as hp
    path = tmp_path / "pids.json"
    path.write_text(json.dumps([_entry(11), _entry(12)]))
    killed = []
    monkeypatch.setattr(hp, "kill_ours", lambda pid, ctime: killed.append(pid) or pid == 11)   # 12: not ours (reused)
    assert hp.sweep_leftovers(path) == [11] and killed == [11, 12]
    assert not path.exists()                     # nothing left to exclude: the file goes


def test_an_entry_of_a_live_older_server_stays_and_its_excel_stays_excluded(tmp_path, monkeypatch):
    import smartass.xl.hostproc as hp
    from smartass.xl import instances
    path = tmp_path / "pids.json"
    path.write_text(json.dumps([_entry(21, owner=99), _entry(22)]))
    monkeypatch.setattr(hp._PidRegistry, "_owner_alive", staticmethod(lambda item: item["owner"] == 99))
    monkeypatch.setattr(hp, "kill_ours", lambda pid, ctime: True)
    monkeypatch.setattr(hp.config, "PID_FILE", path)
    assert hp.sweep_leftovers() == [22]
    assert path.exists() and [i["pid"] for i in json.loads(path.read_text())] == [21]
    assert instances.engine_excel_pids() == {21}


def test_without_a_registry_the_sweep_does_nothing(tmp_path, monkeypatch):
    import smartass.xl.hostproc as hp
    monkeypatch.setattr(hp, "kill_ours", lambda pid, ctime: (_ for _ in ()).throw(AssertionError("killed")))
    assert hp.sweep_leftovers(tmp_path / "pids.json") == [] and not (tmp_path / "pids.json").exists()


def test_the_server_sweeps_at_start(monkeypatch):
    import asyncio

    from smartass.core import serverkit
    from smartass.xl import hostproc
    swept = []
    monkeypatch.setattr(hostproc, "sweep_leftovers", lambda: swept.append(1) or [])
    monkeypatch.setattr(serverkit, "_keep_signed_in", lambda: None)

    async def run():
        async with serverkit.lifespan(None):
            assert swept == [1]
    asyncio.run(run())
