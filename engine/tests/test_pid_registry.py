"""xl.hostproc._PidRegistry: saves are atomic and serialised, so pids.json is never half-written."""
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


def test_concurrent_add_keeps_every_pid(tmp_path, monkeypatch):
    import smartass.xl.hostproc as hp
    monkeypatch.setattr(hp.psutil, "pid_exists", lambda pid: True)
    reg = _PidRegistry(tmp_path / "pids.json")

    def adder(base: int):
        for i in range(25):
            reg.add(base + i, 1.0)

    ts = [threading.Thread(target=adder, args=(b,)) for b in (100, 200)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert sorted(i["pid"] for i in reg.load()) == [*range(100, 125), *range(200, 225)]
