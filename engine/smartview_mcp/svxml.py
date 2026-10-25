"""Smart View provider XML for EPM Cloud Planning (`/HyperionPlanning/SmartView`), learnt by probing on 2026-10-25
(docs/specs/2026-10-25-stage0-epm-access-probe.md). Pure functions: request bodies and response parsing.

Read-only by construction: check_read_only() admits only the requests built here, and ExecuteGrid only with the
'retrieve' action. svhttp.BrowserTransport passes every body through it before sending."""
from __future__ import annotations

import html
import re

from .grid import Layout

ALLOWED = frozenset({"ConnectToProvider", "GetProvisionedDataSources", "OpenApplication", "OpenCube", "ExecuteGrid"})
CLIENT_XML_VERSION = "4.2.5.7.3"
_SPECIAL = {"#noaccess": "#NoAccess", "#no access": "#NoAccess", "#missing": "#Missing", "#invalid": "#Invalid"}


class SvXmlError(Exception):
    """The provider answered with <exception>, or with something that is not provider XML."""

    def __init__(self, desc: str, code: int | None = None):
        super().__init__(desc)
        self.code = code


def check_read_only(body: str) -> None:
    """The transport's last guard: the body must be exactly one allowed request element, retrieve-only."""
    m = re.fullmatch(r"\s*<req_(\w+)>.*</req_\1>\s*", body, re.S)
    name = m.group(1) if m else body[:40]
    if not m or body.count("<req_") != 1 or name not in ALLOWED:
        raise ValueError(f"refused: {name!r} is not a single read-only Smart View request")
    if name == "ExecuteGrid" and (len(re.findall(r"<\s*action\b", body, re.I)) != 1
                                  or "<action>retrieve</action>" not in body):
        raise ValueError("refused: ExecuteGrid is only sent with the 'retrieve' action")


def _x(v) -> str:
    return html.escape(str(v), quote=False)


def connect() -> str:
    return (f'<req_ConnectToProvider><ClientXMLVersion>{CLIENT_XML_VERSION}</ClientXMLVersion><lngs enc="0">en_US'
            '</lngs><usr></usr><pwd></pwd><sso></sso></req_ConnectToProvider>')


def provisioned_data_sources() -> str:
    return '<req_GetProvisionedDataSources><usr></usr><pwd></pwd><filters></filters></req_GetProvisionedDataSources>'


def open_application(sid: str, app: str) -> str:
    return (f'<req_OpenApplication><sID>{_x(sid)}</sID><srv>{_x(app)}</srv><app>{_x(app)}</app><type></type>'
            '<url></url><sessionInfo></sessionInfo></req_OpenApplication>')


def open_cube(sid: str, app: str, cube: str) -> str:
    return (f'<req_OpenCube><sID>{_x(sid)}</sID><srv>{_x(app)}</srv><app>{_x(app)}</app><cube>{_x(cube)}</cube>'
            '<type></type><url></url></req_OpenCube>')


def grid_cells(lay: Layout) -> tuple[list[str], list[str]]:
    """Row-major cell values and types of the free-form layout: 7 = blank, 0 = member, 2 = data cell."""
    m = lay.matrix()
    data_rows = range(1 + lay.nc, len(m))
    data_cols = range(lay.nr, lay.nr + len(lay.col_tuples))
    vals, types = [], []
    for r, row in enumerate(m):
        for c, v in enumerate(row):
            text = "" if v is None else str(v)
            if "|" in text:
                raise ValueError(f"member {text!r} contains '|', which the provider uses as separator")
            vals.append(_x(text))
            types.append("2" if r in data_rows and c in data_cols else ("7" if v is None else "0"))
    return vals, types


def execute_grid(sid: str, cube: str, lay: Layout, alias: str = "Default") -> str:
    """Free-form ad hoc retrieve. Alias table 'Default' returns the labels Smart View shows in Excel."""
    m = lay.matrix()
    h, w = len(m), len(m[0])
    vals, types = grid_cells(lay)
    prefs = (f'<preferences><aliasTableName val="{html.escape(alias, quote=True)}"/><missingLabelText val="#Missing"/>'
             '<noAccessText val="#NoAccess"/><repeatMemberLabels val="1"/></preferences>')
    return (f'<req_ExecuteGrid><sID>{_x(sid)}</sID><action>retrieve</action><ords></ords>{prefs}<grid>'
            f'<cube>{_x(cube)}</cube><gridType>1</gridType><dims/><slices><slice rows="{h}" cols="{w}"><data>'
            f'<range start="0" end="{h * w - 1}"><vals enc="1">{"|".join(vals)}</vals>'
            f'<types enc="0">{"|".join(types)}</types></range></data></slice></slices></grid></req_ExecuteGrid>')


def raise_for_exception(text: str) -> None:
    m = re.search(r'<exception\b[^>]*\berrcode="(-?\d+)"[^>]*>.*?<desc>(.*?)</desc>', text, re.S)
    if m:
        raise SvXmlError(html.unescape(m.group(2)).strip(), int(m.group(1)))
    if "<res_" not in text:
        raise SvXmlError("not a Smart View provider answer: " + re.sub(r"\s+", " ", text)[:120])


def sid_of(text: str) -> str:
    raise_for_exception(text)
    m = re.search(r"<sID>(.*?)</sID>", text, re.S)
    if not m:
        raise SvXmlError("no session id (sID) in the provider answer")
    return html.unescape(m.group(1)).strip()


def sid_in(text: str) -> str | None:
    """The answer's sID if it carries one (the provider may hand out a new one with any answer), else None."""
    m = re.search(r"<sID>(.*?)</sID>", text, re.S) if isinstance(text, str) else None
    return (html.unescape(m.group(1)).strip() or None) if m else None


def data_sources(text: str) -> dict[str, str]:
    """{app: pod base URL} from res_GetProvisionedDataSources."""
    raise_for_exception(text)
    out = {}
    for tag in re.findall(r"<Server\b[^>]*>", text):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', tag))
        if attrs.get("context") and attrs.get("appName"):
            out[html.unescape(attrs["appName"])] = html.unescape(attrs["context"]).split("/HyperionPlanning/")[0]
    return out


def _cell(v: str, t: str):
    if v == "":
        return None
    special = _SPECIAL.get(v.strip().casefold())
    if special:
        return special
    if t not in ("0", "7"):
        try:
            return float(v)
        except ValueError:
            return v
    return v


def grid_values(text: str) -> list[list]:
    """The returned slice as a dense 2D list (like Excel's UsedRange.Value2): floats for data, None for blanks."""
    raise_for_exception(text)
    s = re.search(r'<slice\s+rows="(\d+)"\s+cols="(\d+)"', text)
    r = re.search(r'<range\s+start="(\d+)"\s+end="(\d+)"', text)
    v = re.search(r"<vals[^>]*>(.*?)</vals>", text, re.S)
    t = re.search(r"<types[^>]*>(.*?)</types>", text, re.S)
    if not (s and r and v and t):
        raise SvXmlError("no grid in the provider answer")
    if len(re.findall(r"<range\b", text)) > 1:
        raise SvXmlError("grid answer has several ranges; not supported")
    rows, cols, start, end = int(s.group(1)), int(s.group(2)), int(r.group(1)), int(r.group(2))
    vals = [html.unescape(x) for x in v.group(1).split("|")]
    types = t.group(1).split("|")
    if not (len(vals) == len(types) == end - start + 1 and cols > 0 and start + len(vals) <= rows * cols):
        raise SvXmlError(f"grid answer has {len(vals)} values and {len(types)} types for cells {start}..{end} "
                         f"of a {rows}x{cols} slice")
    out: list[list] = [[None] * cols for _ in range(rows)]
    for i, (x, ty) in enumerate(zip(vals, types)):
        k = start + i
        out[k // cols][k % cols] = _cell(x, ty)
    return out
