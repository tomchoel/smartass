"""A batch of Excel operations recorded as plain data, to be run in one go by xl/inproc/runner.py.

The calling code reads like the COM code it replaces: b.workbook(name) and b.app are nodes, and attribute access,
calls and assignment on a node record operations and return nodes:

    ws = b.workbook("Book1.xlsx").Worksheets("Q3 Data")
    lo = ws.ListObjects.Add(1, ws.Range("A1:C9"), MISSING, 1)
    lo.Name = "Data_Q3"
    b.ret(ws.Range("A1").Value2, "a1")

What is recorded: a call (node.Name(args), also with no arguments) when it is made; an assignment (node.Name = v)
when it is made; a get with no arguments (node.Name) only when something is done with its node - an attribute of it
is used, it is assigned to, passed as an argument, asked back (ret) or marked (attempt). So a get is recorded where it
is used, not where it is written: write a read where it should happen (b.ret right away). A bare node.Method without
the call records nothing, as in pywin32. A call passes its arguments to the member named (Range.Offset(2, 3),
Address(False, False)), as VBA does; pywin32's own aliases (GetAddress) are not Excel names.

What a batch cannot express (it is data, not code): Python decisions on values read inside the same batch (a value
is known only when the batch has run: split the batch there), loops over a collection (the verb each reads named
properties of every item of one: the only loop a batch has), keyword arguments (late-bound
COM ignores them), truth tests, ==, iteration or len() of a node (all raise), calling what a call gave
(ws.Range("A1:C3")(2, 2) raises: name the member, .Item(2, 2)), names starting with "_", a COM object asked back
(refused here when the recorder knows the node is an object: b.app, a workbook, a node used as the target of an
operation; otherwise when it runs). Values: None, bool, int (64-bit), finite float, str, a grid (a tuple or list of
equal rows of those) and a list of scalars (as Worksheets([...]) takes). MISSING is an omitted optional argument.

How the two executors call Excel (runner.py): the add-in makes every operation an IDispatch::Invoke with LCID 0 -
gets and calls with DISPATCH_METHOD | DISPATCH_PROPERTYGET, sets with DISPATCH_PROPERTYPUT. The COM executor makes
calls the same raw way, but gets and sets through pywin32's attribute access (PROPERTYGET / PROPERTYPUT, LCID 0). They
agree on every idiom scripts/check_inproc.py runs; a get of a name pywin32's type library calls a method would differ
(pywin32 hands back the method instead of calling it): write such a member as a call, node.Name(). A failing
operation's message starts "operation <index> (<kind> <name>): " in both executors; the text after it is each
executor's own (pywin32's and the add-in's words for the same COM error can differ).

Plain data (to_dict, version 1): {"version", "quiet", "ops": [...], "finally": [...]}. Each operation is one of
    {"k": "get",  "on": id, "name": n, "as": id}
    {"k": "call", "on": id, "name": n, "args": [v, ...], "as": id}
    {"k": "set",  "on": id, "name": n, "value": v}
    {"k": "ret",  "on": id, "name": key}
    {"k": "mark", "name": key}
    {"k": "verb", "name": verb, "args": [v, ...], "as": id}
with "try": true on an operation whose failure is reported and does not stop the batch (attempt). Id 0 is the
Application. A value is a scalar, {"r": id}, {"grid": [[...], ...]}, {"list": [...]} or {"missing": 1}. Operations
are numbered in one sequence: ops first, then finally.

The verbs (VERBS; each is implemented in both executors, runner.py and Runner.cs):
    each(collection, paths, prefixes): for every item of the collection, in its order, the list of the values at each
        path - dotted property names ("Tab.Color"), where "*" maps the rest of the path over the items of the
        collection reached ("ListObjects.*.Name": the names of a sheet's tables). A value that cannot be read (Excel
        refuses it, or it is a COM object or a date) is None; a busy Excel is not swallowed. prefixes: None, or a list
        of texts; then only the items whose first path's value starts with one of them (case ignored) are given."""
from __future__ import annotations

import contextlib
import math

VERSION = 1
VERBS: dict[str, int] = {"each": 3}    # name -> number of arguments; each verb exists in both executors (runner.py)
INT64 = 2 ** 63
TEXT_MAX = 40                     # a string argument longer than this is shortened in an operation's text


class _Missing:
    def __repr__(self) -> str:
        return "MISSING"


MISSING = _Missing()


class Node:
    """A reference to what an operation gives inside Excel. Plain attributes start with "_"; everything else is
    recorded."""

    __slots__ = ("_b", "_id", "_parent", "_name", "_text", "_object", "_returned", "_op", "_final")

    def __init__(self, b: Batch, text: str, parent: Node | None = None, name: str | None = None):
        for k, v in (("_b", b), ("_id", None), ("_parent", parent), ("_name", name), ("_text", text),
                     ("_object", False), ("_returned", False), ("_op", None), ("_final", False)):
            object.__setattr__(self, k, v)

    def __getattr__(self, name: str) -> Node:
        if name.startswith("_"):
            raise AttributeError(name)
        self._b._receiver(self)
        return Node(self._b, f"{self._text}.{name}", parent=self, name=name)

    def __setattr__(self, name: str, value) -> None:
        if name.startswith("_"):
            raise AttributeError(f"{name}: a node has no attributes of its own to set")
        self._b._set(self, name, value)

    def __call__(self, *args, **kw) -> Node:
        if kw:
            raise TypeError("keyword arguments are not recorded (late-bound COM ignores them): pass them by position")
        if self._parent is None or (self._op is not None and self._op["k"] != "get"):
            raise TypeError(f"{self._text} is not a property or method to call: name the member (.Item(...))")
        return self._b._call(self._parent, self._name, args)

    def __eq__(self, other):
        raise TypeError(f"{self._text} has no value while the batch is recorded: compare after it ran")

    __ne__ = __eq__
    __hash__ = object.__hash__

    def __bool__(self):
        raise TypeError(f"{self._text} has no value while the batch is recorded: read it with ret and decide after")

    def __iter__(self):
        raise TypeError(f"{self._text} cannot be iterated while the batch is recorded")

    def __len__(self):
        raise TypeError(f"{self._text} has no length while the batch is recorded")

    def __repr__(self) -> str:
        return f"<Node {self._text}>"


class Batch:
    def __init__(self, quiet: bool = True):
        self.quiet = quiet
        self._ops: list[dict] = []
        self._finally: list[dict] = []
        self._texts: list[str] = []
        self._ftexts: list[str] = []
        self._in_finally = False
        self._trying = False
        self._next = 1
        self._rets: set[str] = set()
        self._books: dict[str, Node] = {}
        self.app = Node(self, "Application")
        object.__setattr__(self.app, "_id", 0)
        object.__setattr__(self.app, "_object", True)

    # ------------------------------------------------------------------ public
    def workbook(self, name: str) -> Node:
        """The open workbook of that name (Application.Workbooks(name)), one node per name."""
        if name not in self._books:
            node = self._call(self.app, "Workbooks", (name,))
            object.__setattr__(node, "_text", f"Workbooks({_short(name)})")
            object.__setattr__(node, "_object", True)
            self._books[name] = node
        return self._books[name]

    def ret(self, node: Node, name: str) -> None:
        """Ask for node's value back: Result.values[name]."""
        if not isinstance(node, Node):
            raise TypeError("ret takes a node")
        if node._object:
            raise TypeError(f"{node._text} is a COM object: only a value can be asked back")
        if name in self._rets:
            raise ValueError(f"a value named {name!r} is already asked back")
        self._record(node)
        self._rets.add(name)
        object.__setattr__(node, "_returned", True)
        self._add({"k": "ret", "on": self._ref(node), "name": name}, f"ret {name} = {node._text}")

    @property
    def position(self) -> int:
        """The index the next operation recorded outside always() gets (a failure's BatchError.index tells which
        step it was in; Result.skipped's indexes too)."""
        return len(self._ops)

    def mark(self, name: str) -> None:
        """Result.marks[name]: the milliseconds since the batch started, when this point is reached."""
        self._add({"k": "mark", "name": name}, f"mark {name}")

    def verb(self, name: str, *args) -> Node:
        """A named step the operation list cannot say (VERBS, see the module docstring)."""
        if VERBS.get(name) != len(args):
            raise ValueError(f"no verb {name!r} with {len(args)} arguments (verbs: {VERBS})")
        vals = [self._value(a) for a in args]
        node = Node(self, f"{name}({', '.join(_arg_text(a) for a in args)})")
        self._bind(node, {"k": "verb", "name": name, "args": vals})
        return node

    @contextlib.contextmanager
    def always(self):
        """Operations recorded inside run after the others, also when one of those failed (a finally); a failure
        there is reported in Result.skipped and the rest go on."""
        if self._in_finally:
            raise RuntimeError("always() is not nested")
        self._in_finally = True
        try:
            yield
        finally:
            self._in_finally = False

    def attempt(self, node: Node | None = None):
        """attempt(node): the operation that gives node may fail without stopping the batch (Result.skipped).
        attempt() with no node: a context in which every operation recorded is such an operation."""
        if node is None:
            return self._attempting()
        self._record(node)
        if node._op is None:
            raise TypeError(f"{node._text} is not an operation")
        node._op["try"] = True
        return node

    def to_dict(self) -> dict:
        return {"version": VERSION, "quiet": self.quiet, "ops": [dict(o) for o in self._ops],
                "finally": [dict(o) for o in self._finally]}

    def text(self, index: int) -> str:
        """The operation at index (ops, then finally, in one sequence) as it reads in code."""
        texts = self._texts + self._ftexts
        return texts[index] if 0 <= index < len(texts) else f"operation {index}"

    # ------------------------------------------------------------------ recording
    @contextlib.contextmanager
    def _attempting(self):
        before, self._trying = self._trying, True
        try:
            yield
        finally:
            self._trying = before

    def _add(self, op: dict, text: str) -> None:
        if self._trying:
            op["try"] = True
        (self._finally if self._in_finally else self._ops).append(op)
        (self._ftexts if self._in_finally else self._texts).append(text)

    def _bind(self, node: Node, op: dict) -> None:
        op["as"] = self._next
        object.__setattr__(node, "_id", self._next)
        object.__setattr__(node, "_op", op)
        object.__setattr__(node, "_final", self._in_finally)
        self._next += 1
        self._add(op, node._text)

    def _record(self, node: Node) -> None:
        """A pending get is recorded now (its parent was recorded when the attribute was taken)."""
        if node._b is not self:
            raise ValueError(f"{node._text} belongs to another batch")
        if node._id is None:
            self._bind(node, {"k": "get", "on": self._ref(node._parent), "name": node._name})

    def _ref(self, node: Node) -> int:
        if node._final and not self._in_finally:
            raise ValueError(f"{node._text} is made inside always(): it cannot be used outside it")
        return node._id

    def _receiver(self, node: Node) -> None:
        """node is the target of an operation: it is an object."""
        if node._returned:
            raise TypeError(f"{node._text} is asked back as a value: it cannot be used as a COM object")
        self._record(node)
        object.__setattr__(node, "_object", True)

    def _call(self, parent: Node, name: str, args: tuple) -> Node:
        self._receiver(parent)
        vals = [self._value(a) for a in args]
        node = Node(self, f"{parent._text}.{name}({', '.join(_arg_text(a) for a in args)})", parent=parent, name=name)
        self._bind(node, {"k": "call", "on": self._ref(parent), "name": name, "args": vals})
        return node

    def _set(self, node: Node, name: str, value) -> None:
        self._receiver(node)
        v = self._value(value)
        self._add({"k": "set", "on": self._ref(node), "name": name, "value": v},
                  f"{node._text}.{name} = {_arg_text(value)}")

    def _value(self, v):
        if v is MISSING:
            return {"missing": 1}
        if isinstance(v, Node):
            self._record(v)
            return {"r": self._ref(v)}
        if isinstance(v, (list, tuple)):
            if not v:
                raise ValueError("an empty list or grid cannot be sent")
            if all(isinstance(r, (list, tuple)) for r in v):
                width = len(v[0])
                if not width or any(len(r) != width for r in v):
                    raise ValueError("a grid must be rectangular, with at least one column")
                return {"grid": [[_scalar(c) for c in r] for r in v]}
            return {"list": [_scalar(c) for c in v]}
        return _scalar(v)


def _scalar(v):
    if v is None or isinstance(v, (bool, str)):
        return v
    if isinstance(v, int):
        if not -INT64 <= v < INT64:
            raise ValueError(f"{v} does not fit a 64-bit integer")
        return v
    if isinstance(v, float):
        if not math.isfinite(v):
            raise ValueError(f"{v} cannot be sent to Excel")
        return v
    raise TypeError(f"a {type(v).__name__} cannot be sent to Excel (None, bool, int, float, str, grids, lists)")


def _short(v) -> str:
    if isinstance(v, str):
        s = '"' + v.replace('"', '\\"') + '"'
        return s if len(v) <= TEXT_MAX else f'"{v[:TEXT_MAX]}..." ({len(v)} characters)'
    return repr(v)


def _arg_text(a) -> str:
    if isinstance(a, Node):
        return f"<{a._text}>"
    if isinstance(a, (list, tuple)) and a and all(isinstance(r, (list, tuple)) for r in a):
        return f"<grid {len(a)}x{len(a[0])}>"
    if isinstance(a, (list, tuple)):
        return "[" + ", ".join(_short(x) for x in a[:5]) + (", ..." if len(a) > 5 else "") + "]"
    return _short(a)
