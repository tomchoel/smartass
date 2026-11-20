"""A late-bound COM wrapper that learns each interface once, for every Excel call the engine makes.

pywin32's standard late-bound wrapper (win32com.client.dynamic.CDispatch) starts from nothing for every object a call
returns: it fetches the type information again, binds each attribute name again on first use (four ITypeComp.Bind
calls per name) and compiles a new method stub per object. Measured in the owner's Excel that is about two thirds of
the cost of a call. `Fast` is that same wrapper with one change: what it learns about an interface (which names are
properties or methods, their ids, the compiled method stubs) is kept per interface IID and shared by every object of
that interface. After warm-up a new object costs the two cheap type-information calls in `wrap` and nothing more.

`wrap` is applied at the engine's COM roots (xl/instances.py: the user's Excel behind a workbook window; xl/host.py:
the hidden Excel); everything a Fast object returns (attribute gets, calls, method results, iteration) is Fast too.
SV_FASTCOM=0 (config.FASTCOM) gives the standard wrapper instead, to rule this module out when diagnosing.

Rules:
- Threading: a COM pointer belongs to the thread that got it (user-Excel jobs run on their own STA threads). The
  state shared per interface is pure Python: name maps of MapEntry / FUNCDESC data and compiled functions. A name
  that is not known yet is bound with the type information of the object at hand, fetched on its own thread. The lock
  guards only the merge into the shared maps, never a COM call: a hung Excel must not stall a job in another Excel.
- A name is bound as pywin32 binds it on an object that has used every name used so far on any object of the
  interface (into a copy of the shared description), and the changes are merged. It is remembered as known (found,
  or not in the type library) only when the binding gave a clean answer (no error beyond ITypeComp.Bind's "not this
  invoke kind") and a found name sits under its own key; otherwise the next object binds it again, as pywin32 would.
- What pywin32 learns from one object's behaviour rather than from the type library stays with that object, as in
  pywin32: the method fallback after a property get failed with a "bad context" error (CDispatch.__getattr__), a name
  found by GetIDsOfNames on a put (__setattr__), _FlagAsMethod, and the stubs built from those. Office answers some
  property gets with E_INVALIDARG when the object is in a certain state (a chart without a title); shared, that one
  answer would turn the property into a method for every object of the interface.
- The default member's name (defaultDispatchName, used by obj(...)) is shared once an object has bound it.

Coupling (accepted: pywin32 is pinned, pywin32==312 in engine/pyproject.toml): this module relies on pywin32's
internals - CDispatch's instance attributes (_oleobj_, _olerepr_, _builtMethods_, _lazydata_, _username_),
_LazyAddAttr_ (whose loop _bind repeats), _make_method_, _wrap_dispatch_, _NewEnum, the stubs' global "Dispatch",
build.DispatchItem / LazyDispatchItem with their four maps and _AddFunc_ / _AddVar_ / _propMap*Check_,
dynamic.MakeOleRepr and util.EnumVARIANT / Iterator. tests/test_fastcom.py exercises each of them against the
standard wrapper, so an upgrade that changes them fails there.
"""
from __future__ import annotations

import threading
import traceback
from collections.abc import MutableMapping
from types import MethodType

import pythoncom
import winerror
from win32com.client import build, dynamic, util

from ..core import config

_PyIDispatch = pythoncom.TypeIIDs[pythoncom.IID_IDispatch]
_MAPS = ("propMap", "propMapGet", "propMapPut", "mapFuncs")
NOT_SHARED = (pythoncom.IID_NULL, pythoncom.IID_IDispatch)      # no real interface behind these: never shared
_ENUM_INVOKE = pythoncom.DISPATCH_METHOD | pythoncom.DISPATCH_PROPERTYGET


class _Interface:
    """What is learnt about one interface. Pure Python, shared by its objects on every thread."""

    def __init__(self, olerepr: build.DispatchItem, iid):
        self.iid = iid                             # the interface's IID (inproc/runner.py keys its property check on it)
        self.olerepr = olerepr                     # build.LazyDispatchItem: the name maps, filled as names are used
        self.methods: dict = {}                    # compiled method stubs by name
        self.known: dict[str, int] = {}            # name -> 1 bound, 0 not in the type library
        self.lock = threading.Lock()               # merges into olerepr, methods and known

    def merge(self, base: build.DispatchItem, bound: build.DispatchItem, attr: str, found: int, clean: bool) -> None:
        """Apply what binding `attr` changed (base -> bound, both copies) to the shared description."""
        with self.lock:
            for m in _MAPS:
                before, after, shared = getattr(base, m), getattr(bound, m), getattr(self.olerepr, m)
                for key, entry in after.items():
                    if before.get(key) is not entry:
                        shared[key] = entry
                for key in before.keys() - after.keys():
                    shared.pop(key, None)
            if bound.defaultDispatchName != base.defaultDispatchName:
                self.olerepr.defaultDispatchName = bound.defaultDispatchName
            if clean and (not found or any(attr in getattr(self.olerepr, m) for m in _MAPS)):
                self.known[attr] = found           # last: a reader that sees the name known also sees its entries


_IFACE: dict = {}                                  # interface IID -> _Interface


class _Overlay(MutableMapping):
    """One object's view of a shared map: its own entries first, then the interface's. Writes stay with the object."""

    __slots__ = ("own", "shared")

    def __init__(self, shared: dict):
        self.own: dict = {}
        self.shared = shared

    def __contains__(self, key) -> bool:
        return key in self.own or key in self.shared

    def __getitem__(self, key):
        own = self.own
        return own[key] if key in own else self.shared[key]

    def get(self, key, default=None):
        own = self.own
        return own[key] if key in own else self.shared.get(key, default)

    def __setitem__(self, key, value) -> None:
        self.own[key] = value

    def __delitem__(self, key) -> None:
        del self.own[key]

    def __iter__(self):
        return iter({**self.shared, **self.own})

    def __len__(self) -> int:
        return len(self.shared.keys() | self.own.keys())


class _Repr:
    """One object's olerepr: the four maps as overlays on the interface's; anything else (defaultDispatchName,
    MakeFuncMethod, clsid ...) is the interface's own."""

    def __init__(self, shared: build.DispatchItem):
        d = self.__dict__
        d["_shared"] = shared
        for m in _MAPS:
            d[m] = _Overlay(getattr(shared, m))

    def __getattr__(self, name):
        return getattr(self.__dict__["_shared"], name)


def _copy(olerepr: build.DispatchItem) -> build.DispatchItem:
    """A description with copies of olerepr's maps (the same entries) and its default member."""
    item = build.DispatchItem()
    for m in _MAPS:
        setattr(item, m, dict(getattr(olerepr, m)))
    item.defaultDispatchName = olerepr.defaultDispatchName
    return item


def _bind(olerepr: build.DispatchItem, typeinfo, typecomp, attr: str) -> tuple[int, bool]:
    """pywin32's CDispatch._LazyAddAttr_ (312) on `olerepr`: bind `attr` for each invoke kind and add what is found.
    Returns (found, clean). Bind answers a name that exists, but not as this invoke kind, with TYPE_E_TYPEMISMATCH;
    any other error makes the answer unclean (pywin32 skips that invoke kind silently either way)."""
    found, clean = 0, True
    for kind in dynamic.ALL_INVOKE_TYPES:
        try:
            x, t = typecomp.Bind(attr, kind)
            if x == 0 and attr[:3] in ("Set", "Get"):          # 'Get' and 'Set' properties, as pywin32
                x, t = typecomp.Bind(attr[3:], kind)
            if x == pythoncom.DESCKIND_FUNCDESC:
                r = olerepr._AddFunc_(typeinfo, t, 0)
            elif x == pythoncom.DESCKIND_VARDESC:
                r = olerepr._AddVar_(typeinfo, t, 0)
            else:                                               # not found, or a TYPEDESC / IMPLICITAPPOBJ
                r = None
            if r is not None:
                key, where = r
                if where is olerepr.propMapPut:
                    olerepr._propMapPutCheck_(key, where[key])
                elif where is olerepr.propMapGet:
                    olerepr._propMapGetCheck_(key, where[key])
                found = 1
        except pythoncom.com_error as e:
            clean = clean and e.hresult == winerror.TYPE_E_TYPEMISMATCH
        except Exception:   # noqa: BLE001 - as pywin32: that invoke kind is skipped
            clean = False
    return found, clean


def _good(ob, userName=None):
    """win32com.client._get_good_object_ with wrap: what an enumerator hands out."""
    if isinstance(ob, tuple):
        return tuple(_good(o, userName) for o in ob)
    if isinstance(ob, _PyIDispatch):
        return wrap(ob, userName)
    return ob


class _Enum(util.EnumVARIANT):
    """The enumerator behind `for x in coll` and coll[i]: its items are Fast."""

    def _make_retval_(self, result):
        return _good(result)


class _Iterator(util.Iterator):
    """The iterator obj.__iter__() returns: its items are Fast."""

    def __next__(self):
        return _good(next(self._iter_))


def _from_stub(ob, userName=None, resultCLSID=None):
    """"Dispatch" as the compiled method stubs see it: an object a method returns is Fast."""
    return wrap(ob, userName)


_STUB_GLOBALS = {**vars(dynamic), "Dispatch": _from_stub}


class Fast(dynamic.CDispatch):
    """pywin32's own late-bound wrapper with one change: what it learns about an interface is shared by every object
    of that interface (see the module docstring for what stays with the object)."""

    def __init__(self, ob, iface: _Interface, userName=None, typeinfo=None, typecomp=None):
        lazy = (typeinfo, typecomp) if typecomp is not None else None
        dynamic.CDispatch.__init__(self, ob, _Repr(iface.olerepr), userName, lazydata=lazy)
        d = self.__dict__
        d["_builtMethods_"] = _Overlay(iface.methods)
        d["_iface_"] = iface
        d["_typeinfo_"] = typeinfo                 # this object's own, fetched on its thread in wrap

    def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
        return wrap(ob, userName)

    def _typecomp_(self):
        """This object's own type information and its ITypeComp, the first time it binds a name."""
        ti = self._typeinfo_
        lazy = (ti, ti.GetTypeComp())
        self.__dict__["_lazydata_"] = lazy
        return lazy

    def _LazyAddAttr_(self, attr):
        iface = self._iface_
        known = iface.known.get(attr)
        if known is not None:
            return known
        try:
            typeinfo, typecomp = self._lazydata_ or self._typecomp_()
        except pythoncom.com_error:
            return 0
        base, bound = _copy(iface.olerepr), _copy(iface.olerepr)
        found, clean = _bind(bound, typeinfo, typecomp, attr)
        iface.merge(base, bound, attr, found, clean)
        return found

    def _make_method_(self, name):
        method_name = build.MakePublicAttributeName(name)
        entry = self._olerepr_.mapFuncs[name]
        code = "\n".join(self._olerepr_.MakeFuncMethod(entry, method_name, 0))
        try:
            ns: dict = {}
            exec(compile(code, "<COMObject %s>" % self._username_, "exec"), _STUB_GLOBALS, ns)
            fn = ns[method_name]
        except Exception:   # noqa: BLE001 - as pywin32: print the stub and return None
            dynamic.debug_print("Error building OLE definition for code ", code)
            traceback.print_exc()
            return None
        iface = self._iface_
        if iface.olerepr.mapFuncs.get(name) is entry:          # from the type library: for every object
            with iface.lock:
                iface.methods[method_name] = fn
        else:                                                   # learnt from this object only
            self._builtMethods_.own[method_name] = fn
        return MethodType(fn, self)

    def _NewEnum(self):
        try:
            enum = self._oleobj_.InvokeTypes(pythoncom.DISPID_NEWENUM, dynamic.LCID, _ENUM_INVOKE, (13, 10), ())
        except pythoncom.com_error:
            return None
        if not isinstance(enum, pythoncom.TypeIIDs[pythoncom.IID_IEnumVARIANT]):
            enum = enum.QueryInterface(pythoncom.IID_IEnumVARIANT)
        return _Enum(enum, None)

    def __getattr__(self, attr):
        if attr == "__iter__":
            try:
                enum = self._oleobj_.InvokeTypes(pythoncom.DISPID_NEWENUM, dynamic.LCID, _ENUM_INVOKE, (13, 10), ())
            except pythoncom.com_error:
                raise AttributeError("This object can not function as an iterator") from None
            return lambda: _Iterator(enum)
        return dynamic.CDispatch.__getattr__(self, attr)


def wrap(obj, name: str | None = None):
    """`obj` (a PyIDispatch or any pywin32 dispatch object) wrapped Fast; a Fast object is returned unchanged, and so
    is anything that is not a COM object (the fakes of the tests). With SV_FASTCOM=0 the standard wrapper."""
    if isinstance(obj, Fast):
        return obj
    raw = getattr(obj, "_oleobj_", obj)
    if not isinstance(raw, _PyIDispatch):
        return obj
    if name is None and isinstance(obj, dynamic.CDispatch):
        name = obj._username_
    if not config.FASTCOM:
        return dynamic.Dispatch(raw, name)
    try:
        typeinfo = raw.GetTypeInfo()
        iid = typeinfo.GetTypeAttr()[0]
    except (pythoncom.com_error, AttributeError):          # no type information: pywin32 copes without it
        return dynamic.Dispatch(raw, name)
    if iid in NOT_SHARED:
        return dynamic.Dispatch(raw, name, typeinfo=typeinfo)
    iface = _IFACE.get(iid)
    typecomp = None
    if iface is None:                                      # first object of this interface
        try:
            typecomp = typeinfo.GetTypeComp()
        except pythoncom.com_error:
            return dynamic.Dispatch(raw, name, typeinfo=typeinfo)
        iface = _IFACE.setdefault(iid, _Interface(dynamic.MakeOleRepr(raw, typeinfo, typecomp), iid))
    return Fast(raw, iface, name, typeinfo, typecomp)
