"""Provider answers for tests - shapes copied from the live probe of 2026-10-25; every figure is invented."""


def answer(matrix):
    """A res_ExecuteGrid like the provider's, for a dense matrix (float = data cell)."""
    h, w = len(matrix), len(matrix[0])
    cells = [v for row in matrix for v in row]
    vals = "|".join("" if v is None else (repr(v) if isinstance(v, float) else v) for v in cells)
    types = "|".join("2" if isinstance(v, float) else ("7" if v is None else "0") for v in cells)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><res_ExecuteGrid><grid><slices>'
            f'<slice rows="{h}" cols="{w}"><data><generations/><range start="0" end="{h * w - 1}">'
            f'<vals enc="1">{vals}</vals><types enc="0">{types}</types></range></data></slice></slices></grid>'
            '<stats><Total_Provider_Time>9</Total_Provider_Time></stats></res_ExecuteGrid>')


SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
           'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
           '</res_GetProvisionedDataSources>')


def _request_matrix(body):
    import html
    import re
    h, w = (int(x) for x in re.search(r'<slice rows="(\d+)" cols="(\d+)">', body).groups())
    vals = [html.unescape(x) for x in re.search(r"<vals[^>]*>(.*?)</vals>", body, re.S).group(1).split("|")]
    types = re.search(r"<types[^>]*>(.*?)</types>", body, re.S).group(1).split("|")
    return h, w, vals, types


class EchoTransport:
    """A provider that answers every retrieve by echoing the requested grid and filling each data cell with
    value(row_members, col_members, pov_members) - a float or None (blank). Labels come back as the member names.
    Figures are whatever the test's value function invents."""

    def __init__(self, value, closed="[CurrRepMonth].[Y26M09]"):
        self.value, self.closed, self.grids = value, closed, []

    def post(self, base, body):
        import re

        from smartview_mcp import svxml
        svxml.check_read_only(body)
        tag = re.match(r"\s*<req_(\w+)>", body).group(1)
        if tag == "GetProvisionedDataSources":
            return SOURCES
        if tag in ("ConnectToProvider", "OpenApplication"):
            return f"<res_{tag}><sID>s1</sID></res_{tag}>"
        if tag == "OpenCube":
            return "<res_OpenCube></res_OpenCube>"
        if tag == "EnumMembers":
            return f'<res_EnumMembers><mbrs enc="1">{self.closed}</mbrs></res_EnumMembers>'
        h, w, vals, types = _request_matrix(body)
        cells = [[None if vals[r * w + c] == "" else vals[r * w + c] for c in range(w)] for r in range(h)]
        data = [(r, c) for r in range(h) for c in range(w) if types[r * w + c] == "2"]
        r0, c0 = min(r for r, _ in data), min(c for _, c in data)          # 1 + column dims, row dims
        pov = tuple(v for v in cells[0][c0:] if v is not None)
        for r, c in data:
            cells[r][c] = self.value(tuple(cells[r][:c0]), tuple(cells[d][c] for d in range(1, r0)), pov)
        self.grids.append((base, len(data)))
        return answer(cells)


FINMAX_DIMS = ["All Accounts", "All Balance Types", "All Currencies", "All Entities", "All Scenarios",
               "All Time Periods"]
FINPLAN_DIMS = ["Account", "All Balance Types", "Entity", "Period", "Scenario", "Version", "Years"]


def atlas_with_dimensions(path):
    """A temporary atlas that knows the two cubes' dimensions (all svhttp needs to lay out a grid)."""
    from smartview_mcp import atlas
    db = atlas.connect(path)
    for conn, dims in (("FinMax", FINMAX_DIMS), ("FinPlan", FINPLAN_DIMS)):
        for d in dims:
            atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)
    db.close()
    return path


# A consistent synthetic P&L for EchoTransport (invented figures; the catalogue identities hold).
FM = {"ac.C30": 1000.0, "ac.C40S": -400.0, "ac.C48": 600.0, "ac.C50S": -200.0, "ac.C60": -100.0, "ac.C70": -300.0,
      "ac.EBITDA_BOI": 300.0, "ac.C73": -10.0, "ac.EBITDA": 290.0, "ac.C75": -50.0, "ac.STA20Y": 10.0}
FP = {"Revenue": 1000.0, "COGS": -400.0, "Gross_Profit": 600.0, "OPEX": -300.0, "EBITDA before Other Items": 300.0,
      "Total Other Items": -10.0, "EBITDA": 290.0}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def pnl_value(rows, cols, pov):
    """A consistent synthetic P&L: the same every month; an Accumulated flow = month number x the month (the sum of
    the months) in both cubes; FTE is a level. The account may be on rows, columns or in the POV."""
    cells = (*rows, *cols, *pov)
    account = next(x for x in cells if x in FM or x in FP)
    v = FM.get(account, FP.get(account))
    month = next((x[:3] for x in cells if x[:3] in MONTHS and (len(x) == 3 or x[3] == " ")), None)
    if "Accumulated" in cells and account != "ac.STA20Y" and month:
        return v * (MONTHS.index(month) + 1)
    return v
