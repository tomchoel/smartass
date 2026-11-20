"""xl/fastcom.py: the late-bound wrapper that learns each interface once, against pywin32's standard wrapper on a COM
server every Windows has (scrrun.dll: Scripting.FileSystemObject and Scripting.Dictionary, in-process). No Excel."""
import threading
from types import SimpleNamespace

import pythoncom
import pytest
import winerror
from win32com.client import build, dynamic
from win32com.server import util as server_util

from smartass.core import config
from smartass.xl import fastcom

COM_TYPES = tuple(pythoncom.TypeIIDs.values())          # PyIDispatch, PyITypeInfo, PyITypeComp ...


@pytest.fixture(autouse=True)
def fresh(monkeypatch):
    """Every test learns its interfaces from nothing, with the wrapper switched on."""
    monkeypatch.setattr(fastcom, "_IFACE", {})
    monkeypatch.setattr(config, "FASTCOM", True)


@pytest.fixture
def tree(tmp_path):
    for sub in ("a", "b"):
        (tmp_path / sub).mkdir()
    (tmp_path / "x.txt").write_text("xx")
    (tmp_path / "y.txt").write_text("yyy")
    (tmp_path / "a" / "z.txt").write_text("z")
    return tmp_path


def raw(progid: str):
    return pythoncom.CoCreateInstance(progid, None, pythoncom.CLSCTX_INPROC_SERVER, pythoncom.IID_IDispatch)


def both(progid: str):
    """(standard, fast) wrappers of two new objects of progid."""
    return dynamic.Dispatch(raw(progid)), fastcom.wrap(raw(progid))


def folder_ops(fso, path) -> tuple:
    f = fso.GetFolder(str(path))
    return (f.Files.Count, f.Name, f.Path, f.SubFolders.Count, sorted(s.Name for s in f.SubFolders),
            sorted((x.Name, x.Size) for x in f.Files), f.Files.Item("x.txt").Size, f.SubFolders("a").Files.Count,
            f.ParentFolder.Path, f.IsRootFolder, f.DateCreated, fso.FolderExists(str(path)),
            fso.BuildPath(str(path), "q.txt"), fso.GetBaseName("q.txt"), len(f.Files))


def dictionary_ops(d) -> tuple:
    d.CompareMode = 1                                   # a property put (only on an empty Dictionary)
    d.Add("k", 1)
    d.Add("t", (1, "two", 3.5))
    d.Add("n", None)
    d.SetItem("k", 5)                                   # a put with an argument
    return (d.CompareMode, d.Count, d.Item("k"), d.Exists("K"), d("t"), d["k"], d.Keys(), d.Items(), d.Item("n"),
            d.Exists("missing"), len(d))


def test_results_equal_the_standard_wrapper(tree):
    std_fso, fast_fso = both("Scripting.FileSystemObject")
    expected = folder_ops(std_fso, tree)
    assert folder_ops(fast_fso, tree) == expected
    assert folder_ops(fastcom.wrap(raw("Scripting.FileSystemObject")), tree) == expected     # fresh objects, warm
    std_d, fast_d = both("Scripting.Dictionary")
    expected = dictionary_ops(std_d)
    assert dictionary_ops(fast_d) == expected
    assert dictionary_ops(fastcom.wrap(raw("Scripting.Dictionary"))) == expected


def test_objects_returned_by_gets_calls_and_iteration_are_fast(tree):
    fso = fastcom.wrap(raw("Scripting.FileSystemObject"))
    f = fso.GetFolder(str(tree))                                           # a method result
    got = {"call": f, "get": f.Files, "get of a get": f.ParentFolder, "default member": f.SubFolders("a"),
           "Item": f.SubFolders.Item("b"), "for": next(iter(f.Files)), "index": f.SubFolders[0],
           "__iter__": next(f.Files.__iter__()), "list": [s for s in f.SubFolders][1]}
    assert {k: type(v).__name__ for k, v in got.items()} == {k: "Fast" for k in got}
    assert got["for"].Name in ("x.txt", "y.txt")


def _use_folder(f) -> tuple:
    with pytest.raises(AttributeError):
        f.NotInTheTypeLibrary                                              # bound once, found nowhere
    return f.Name, f.Path, f.Files.Count, f.SubFolders.Count, f.IsRootFolder, [x.Name for x in f.Files]


def test_an_interface_is_learnt_once(tree, monkeypatch):
    fso = fastcom.wrap(raw("Scripting.FileSystemObject"))
    a = fso.GetFolder(str(tree))
    _use_folder(a)
    ifaces, methods = len(fastcom._IFACE), {i: dict(s.methods) for i, s in fastcom._IFACE.items()}
    binds, typecomps = [], []
    real_bind, real_typecomp = fastcom._bind, fastcom.Fast._typecomp_
    monkeypatch.setattr(fastcom, "_bind", lambda *args: binds.append(args[-1]) or real_bind(*args))
    monkeypatch.setattr(fastcom.Fast, "_typecomp_", lambda self: typecomps.append(self) or real_typecomp(self))
    b = fso.GetFolder(str(tree / "a"))
    assert _use_folder(b) == ("a", str(tree / "a"), 1, 0, False, ["z.txt"])
    assert b._iface_ is a._iface_
    assert len(fastcom._IFACE) == ifaces
    assert {i: dict(s.methods) for i, s in fastcom._IFACE.items()} == methods
    assert binds == [] and typecomps == []                                 # no ITypeComp.Bind for the second


def test_names_are_bound_as_pywin32_binds_them():
    names = {"Scripting.Dictionary": ["Item", "Add", "Key", "Keys", "Items", "Count", "CompareMode", "Exists",
                                      "Remove", "RemoveAll", "HashVal", "GetItem", "SetItem", "Nothing"],
             "Scripting.FileSystemObject": ["GetFolder", "Drives", "BuildPath", "FolderExists", "CreateTextFile",
                                            "GetSpecialFolder", "OpenTextFile", "Nothing"]}
    for progid, attrs in names.items():
        std, fast = both(progid)
        for attr in attrs:
            assert fast._LazyAddAttr_(attr) == std._LazyAddAttr_(attr), attr
        for m in fastcom._MAPS:
            mine, theirs = getattr(fast._iface_.olerepr, m), getattr(std._olerepr_, m)
            assert {k: (e.dispid, e.desc[4] if e.desc else None) for k, e in mine.items()} \
                == {k: (e.dispid, e.desc[4] if e.desc else None) for k, e in theirs.items()}, (progid, m)
        assert fast._iface_.olerepr.defaultDispatchName == std._olerepr_.defaultDispatchName
        assert fast._iface_.known["Nothing"] == 0


def test_a_property_put_and_a_method_with_arguments_work():
    d = fastcom.wrap(raw("Scripting.Dictionary"))
    d.CompareMode = 1
    d.Add("k", 1)
    d.Add("other", "v")
    d.Remove("other")
    assert (d.Item("k"), d.CompareMode, d.Count, d.Exists("K")) == (1, 1, 1, True)
    with pytest.raises(AttributeError):
        d.NoSuchProperty = 1                                               # as pywin32: cannot be set


def test_a_com_error_is_raised_unchanged(tree):
    errors = []
    for fso in both("Scripting.FileSystemObject"):
        with pytest.raises(pythoncom.com_error) as e:
            fso.GetFolder(str(tree / "missing"))
        errors.append((e.value.hresult, e.value.excepinfo[5] if e.value.excepinfo else None))
    assert errors[0] == errors[1] and errors[0][0] == winerror.DISP_E_EXCEPTION


def test_wrapping_twice_returns_the_same_object_and_the_switch_gives_the_standard_wrapper(monkeypatch):
    f = fastcom.wrap(raw("Scripting.Dictionary"))
    assert type(f) is fastcom.Fast and fastcom.wrap(f) is f
    std = dynamic.Dispatch(raw("Scripting.Dictionary"))
    assert type(fastcom.wrap(std)) is fastcom.Fast                       # a standard wrapper's _oleobj_ is used
    monkeypatch.setattr(config, "FASTCOM", False)
    off = fastcom.wrap(raw("Scripting.Dictionary"))
    assert type(off) is dynamic.CDispatch
    off.Add("k", 1)
    assert off.Item("k") == 1


def test_what_is_not_a_com_object_passes_through_unchanged():
    class Fake:                                         # like the fake Excel of the tests: anything goes
        def __getattr__(self, name):
            return lambda *a: None
    for ob in (Fake(), SimpleNamespace(_oleobj_="not COM"), None, 3, "Excel.Application"):
        assert fastcom.wrap(ob) is ob


def on_thread(fn):
    """fn() on a new STA thread of its own (CoInitialize there); every COM object it makes is dropped there too."""
    out = {}

    def run():
        pythoncom.CoInitialize()
        try:
            out["value"] = fn()
        except BaseException as e:   # noqa: BLE001 - reported to the test as text, no COM object leaves the thread
            out["error"] = repr(e)
        finally:
            pythoncom.CoUninitialize()
    t = threading.Thread(target=run)
    t.start()
    t.join(60)
    assert "error" not in out, out.get("error")
    return out["value"]


def _com_pointers(iface) -> list[str]:
    """Every COM pointer held by an interface's shared state (there must be none)."""
    found = []
    for where, ob in [("iface", iface), ("olerepr", iface.olerepr)]:
        for k, v in vars(ob).items():
            if isinstance(v, COM_TYPES):
                found.append(f"{where}.{k}")
    for m in fastcom._MAPS:
        for name, entry in getattr(iface.olerepr, m).items():
            found += [f"{m}[{name}].{k}" for k, v in vars(entry).items() if isinstance(v, COM_TYPES)]
    return found


def test_an_object_wrapped_on_one_thread_does_not_hand_its_type_information_to_another(tree, monkeypatch):
    def thread_a():
        return fastcom.wrap(raw("Scripting.FileSystemObject")).GetFolder(str(tree)).Name    # learns Folder.Name
    assert on_thread(thread_a) == tree.name
    folder = next(s for s in fastcom._IFACE.values() if "Name" in s.known)
    assert "Path" not in folder.known and _com_pointers(folder) == []
    seen = []
    real_typecomp = fastcom.Fast._typecomp_
    monkeypatch.setattr(fastcom.Fast, "_typecomp_",
                        lambda self: seen.append((threading.get_ident(), id(self))) or real_typecomp(self))

    def thread_b():
        f = fastcom.wrap(raw("Scripting.FileSystemObject")).GetFolder(str(tree / "a"))
        assert f._iface_ is folder
        path = f.Path                                                     # not bound yet: bound on this thread
        return path, (threading.get_ident(), id(f)), f._lazydata_[0] is f._typeinfo_
    path, me, own = on_thread(thread_b)
    assert path == str(tree / "a") and own
    assert me in seen                                   # thread B's object fetched its own ITypeComp, on thread B
    assert folder.known["Path"] == 1 and _com_pointers(folder) == []


class _Refuses:
    """An IDispatch that answers one property get with E_INVALIDARG, as Office does for a chart title that is not
    there; everything else goes to the real object."""

    def __init__(self, ob, dispid: int):
        self.ob, self.dispid = ob, dispid

    def Invoke(self, dispid, *args):
        if dispid == self.dispid:
            raise pythoncom.com_error(winerror.E_INVALIDARG, "The parameter is incorrect.", None, None)
        return self.ob.Invoke(dispid, *args)

    def __getattr__(self, name):
        return getattr(self.ob, name)


def test_what_one_object_learns_from_its_own_answers_stays_with_that_object(tree):
    for fso in both("Scripting.FileSystemObject"):
        a, b = fso.GetFolder(str(tree)), fso.GetFolder(str(tree / "a"))
        a.__dict__["_oleobj_"] = _Refuses(a._oleobj_, a._oleobj_.GetIDsOfNames("Name"))
        assert callable(a.Name)                         # pywin32's fallback: a property that refused is a method
        assert b.Name == "a"                            # ... for that object only
        assert fso.GetFolder(str(tree / "b")).Name == "b"
        b._FlagAsMethod("ShortName")
        assert callable(b.ShortName) and isinstance(fso.GetFolder(str(tree)).ShortName, str)
    shared = a._iface_
    assert "Name" not in shared.olerepr.mapFuncs and "Name" not in shared.methods


def test_an_object_without_type_information_falls_back_to_the_standard_wrapper():
    class Server:
        _public_methods_ = ["Hello"]
        _public_attrs_ = ["x"]
        x = 3

        def Hello(self, who):
            return "hi " + who
    ob = server_util.wrap(Server())                     # a real PyIDispatch whose GetTypeInfo raises com_error
    with pytest.raises(pythoncom.com_error):
        ob.GetTypeInfo()
    w = fastcom.wrap(ob)
    assert type(w) is dynamic.CDispatch
    assert (w.Hello("you"), w.x) == ("hi you", 3)
    assert fastcom._IFACE == {}


def test_the_users_excel_is_attached_through_the_fast_wrapper(monkeypatch):
    from smartass.xl import instances
    pointer, wrapped = object(), []
    monkeypatch.setattr(instances.win32gui, "SendMessageTimeout", lambda *a: (0, 7))
    monkeypatch.setattr(instances.pythoncom, "ObjectFromLresult", lambda lres, iid, wp: pointer)
    monkeypatch.setattr(instances.fastcom, "wrap",
                        lambda ob, name=None: wrapped.append(ob) or SimpleNamespace(Application="app"))
    assert instances._app_from_window(1) == "app" and wrapped == [pointer]


def test_the_hidden_excel_is_started_through_the_fast_wrapper(monkeypatch):
    from smartass.xl import host

    class Stop(Exception):
        pass
    started, wrapped = object(), []
    monkeypatch.setattr(host.sso, "ensure_global_watcher", lambda: None)
    monkeypatch.setattr(host.wc, "DispatchEx", lambda progid: started)

    def wrap(ob, name=None):
        wrapped.append(ob)
        raise Stop
    monkeypatch.setattr(host.fastcom, "wrap", wrap)
    with pytest.raises(Stop):
        host.ExcelHost._start(SimpleNamespace())
    assert wrapped == [started]


def test_the_pywin32_internals_this_module_relies_on_are_there():
    assert callable(dynamic.MakeOleRepr) and len(dynamic.ALL_INVOKE_TYPES) == 4
    for name in ("_LazyAddAttr_", "_make_method_", "_wrap_dispatch_", "_NewEnum", "_get_good_object_"):
        assert callable(getattr(dynamic.CDispatch, name)), name
    item = build.LazyDispatchItem((pythoncom.IID_NULL,), None)
    assert all(isinstance(getattr(item, m), dict) for m in fastcom._MAPS) and item.defaultDispatchName is None
