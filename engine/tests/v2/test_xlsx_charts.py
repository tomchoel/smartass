# engine/tests/v2/test_xlsx_charts.py
"""Phase 3 Task 3: the answer's charts in the built file (render/xlsx_charts.py): one native chart per ChartSpec,
with render/charts.py's kinds, colours, labels, axes and legend rule. The chart is read back from its XML inside the
.xlsx zip (openpyxl reads charts only partly). Every colour is the map's literal hex
(docs/audit/2026-11-19-render-map.md section 5.3), never one run through the module's own helpers. Every figure is
invented."""
import re
import xml.etree.ElementTree as ET
import zipfile

import pytest
import xlsxwriter
from test_xlsx import build_with

from smartass.render import xlsx_charts
from smartass.render.layout import ChartSpec, SheetPlan
from smartass.render.xlsx_charts import add_charts

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
SHEET = "Q3 Answer"
MNOK = '#,##0.0,,;"−"#,##0.0,,;;'


def spec(kind, series, rows="5:9", **kw) -> ChartSpec:
    """A chart over the rows 'a:b' (or 'a:b,c' with a gap) of columns C (categories) and the series' columns."""
    def rng(col):
        return ",".join(f"{col}{p.split(':')[0]}:{col}{p.split(':')[1]}" if ":" in p else f"{col}{p}"
                        for p in rows.split(","))
    return ChartSpec(kind, kw.pop("title", f"{kind} chart"), cat_range=kw.pop("cat_range", rng("C")),
                     series=[(n, rng(c), r) for n, c, r in series], anchor=kw.pop("anchor", "B30"), **kw)


def draw(tmp_path, *specs, lint=()):
    """The specs drawn on an empty 'Q3 Answer' sheet: (titles drawn, lint, the chart XML roots in order)."""
    path = tmp_path / "charts.xlsx"
    with xlsxwriter.Workbook(str(path)) as wb:
        ws = wb.add_worksheet(SHEET)
        titles, out = add_charts(wb, ws, SHEET, SheetPlan(charts=list(specs), chart_lint=list(lint)))
    return titles, out, charts_of(path)


def charts_of(path) -> list:
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", n)),
                       key=lambda n: int(re.search(r"\d+", n.rsplit("/", 1)[1]).group()))
        return [ET.fromstring(z.read(n)) for n in names]


def drawing_of(path):
    with zipfile.ZipFile(path) as z:
        return ET.fromstring(z.read("xl/drawings/drawing1.xml"))


def sers(el) -> list:
    return list(el.iter(C + "ser"))


def val(el, tag):
    found = el.find(tag)
    return None if found is None else found.get("val")


def name(ser) -> str:
    return ser.find(f"{C}tx/{C}v").text


def ref(ser, part) -> str:
    return ser.find(f"{C}{part}//{C}f").text


def fill(ser) -> str | None:
    """The solid fill colour of a series' shape; 'none' for no fill."""
    sp = ser.find(C + "spPr")
    if sp.find(A + "noFill") is not None:
        return "none"
    solid = sp.find(A + "solidFill")
    return None if solid is None else solid.find(A + "srgbClr").get("val")


def line(ser) -> dict:
    """The shape's outline: colour ('none' for no line), width in EMU, dash."""
    ln = ser.find(f"{C}spPr/{A}ln")
    if ln is None:
        return {}
    if ln.find(A + "noFill") is not None:
        return {"colour": "none"}
    clr = ln.find(f"{A}solidFill/{A}srgbClr")
    return {"colour": None if clr is None else clr.get("val"), "w": ln.get("w"), "dash": val(ln, A + "prstDash")}


def labels(ser):
    return ser.find(C + "dLbls")


def label_format(ser) -> str:
    return labels(ser).find(C + "numFmt").get("formatCode")


def label_font(el) -> dict:
    """A label font: typeface, size, bold, colour (from c:txPr's defRPr)."""
    rpr = el.find(f"{C}txPr/{A}p/{A}pPr/{A}defRPr")
    clr = rpr.find(f"{A}solidFill/{A}srgbClr")
    return {"face": rpr.find(A + "latin").get("typeface"), "sz": rpr.get("sz"), "b": rpr.get("b"),
            "colour": None if clr is None else clr.get("val")}


def custom(ser) -> dict[int, ET.Element]:
    """The per-point labels: 0-based index -> its c:dLbl."""
    return {int(val(d, C + "idx")): d for d in labels(ser).findall(C + "dLbl")}


def value_axes(root) -> list:
    return root.findall(f"{C}chart/{C}plotArea/{C}valAx")


def category_axis(root):
    return root.find(f"{C}chart/{C}plotArea/{C}catAx")


def title(root) -> tuple[str, dict]:
    t = root.find(f"{C}chart/{C}title")
    run = t.find(f".//{A}r")
    rpr = run.find(A + "rPr")
    return run.find(A + "t").text, {"face": rpr.find(A + "latin").get("typeface"), "sz": rpr.get("sz"),
                                    "b": rpr.get("b"), "colour": rpr.find(f"{A}solidFill/{A}srgbClr").get("val")}


def legend(root) -> str | None:
    lg = root.find(f"{C}chart/{C}legend")
    return None if lg is None else val(lg, C + "legendPos")


def group(root, tag="barChart"):
    return root.find(f"{C}chart/{C}plotArea/{C}{tag}")


def assert_value_axis_hidden(ax):
    """No gridlines, no tick marks, no labels, no axis line: the data labels carry the numbers."""
    assert ax.find(C + "majorGridlines") is None and ax.find(C + "minorGridlines") is None
    assert val(ax, C + "majorTickMark") == "none" and val(ax, C + "tickLblPos") == "none"
    assert ax.find(f"{C}spPr/{A}ln/{A}noFill") is not None


def assert_category_axis(root, label_pos="low", reversed_=False):
    cat = category_axis(root)
    assert val(cat, C + "majorTickMark") == "none" and val(cat, C + "tickLblPos") == label_pos
    assert val(cat.find(C + "scaling"), C + "orientation") == ("maxMin" if reversed_ else "minMax")


def assert_refs(ser, values, categories):
    assert ref(ser, "val") == values and ref(ser, "cat") == categories


# ------------------------------------------------------------------ one test per kind
def test_a_column_chart_is_clustered_with_ac_solid_and_py_tinted_in_the_metric_hue(tmp_path):
    s = spec("column", [("AC YTD Oct 26", "D", "ac"), ("PY YTD Oct 25", "E", "py")], hues=[1, 1],
             title="Revenue vs last year")
    titles, lint, (root,) = draw(tmp_path, s)
    assert titles == ["Revenue vs last year"] and lint == []
    g = group(root)
    assert val(g, C + "barDir") == "col" and val(g, C + "grouping") == "clustered"
    assert val(g, C + "gapWidth") == "50" and g.find(C + "overlap") is None
    ac, py = sers(g)
    assert [name(ac), name(py)] == ["AC YTD Oct 26", "PY YTD Oct 25"]
    assert_refs(ac, "'Q3 Answer'!$D$5:$D$9", "'Q3 Answer'!$C$5:$C$9")
    assert_refs(py, "'Q3 Answer'!$E$5:$E$9", "'Q3 Answer'!$C$5:$C$9")
    assert (fill(ac), line(ac)) == ("008080", {"colour": "none"})          # teal, no border
    assert (fill(py), line(py)) == ("8CC6C6", {"colour": "none"})          # teal 55 % towards white
    assert label_format(ac) == label_format(py) == MNOK
    assert val(labels(ac), C + "showVal") == "1" and labels(ac).find(C + "dLblPos") is None   # Excel's outside end
    assert label_font(labels(ac))["face"] == "Arial"
    text, font = title(root)
    assert text == "Revenue vs last year" and font == {"face": "Arial", "sz": "1200", "b": "1", "colour": "1F3864"}
    assert legend(root) == "t"                                             # two visible series
    (ax,) = value_axes(root)
    assert_value_axis_hidden(ax)
    assert_category_axis(root)
    assert ac.find(C + "dPt") is None                                      # no highlight


def test_a_column_chart_over_rows_with_gaps_points_at_a_union(tmp_path):
    s = spec("column", [("AC", "D", "ac")], rows="5:7,9")
    _, _, (root,) = draw(tmp_path, s)
    (ac,) = sers(root)
    assert_refs(ac, "('Q3 Answer'!$D$5:$D$7,'Q3 Answer'!$D$9)", "('Q3 Answer'!$C$5:$C$7,'Q3 Answer'!$C$9)")
    assert val(ac.find(f"{C}cat//{C}strCache"), C + "ptCount") == "4"      # the labels are text: four of them
    assert legend(root) is None                                            # one AC series: the title names it


def test_a_bar_chart_keeps_the_first_row_on_top_and_outlines_bu(tmp_path):
    s = spec("bar", [("AC", "D", "ac"), ("BU", "E", "bu")], sort="desc")
    _, lint, (root,) = draw(tmp_path, s)
    assert lint == []
    g = group(root)
    assert val(g, C + "barDir") == "bar" and val(g, C + "grouping") == "clustered" and val(g, C + "gapWidth") == "50"
    ac, bu = sers(g)
    assert fill(ac) == "1F3864"
    assert fill(bu) == "FFFFFF" and line(bu) == {"colour": "1F3864", "w": "12700", "dash": None}   # 1 pt navy
    assert label_format(bu) == MNOK
    assert legend(root) == "t"
    (ax,) = value_axes(root)
    assert_value_axis_hidden(ax)
    assert_category_axis(root, reversed_=True)                             # ReversePlotOrder: first row on top


def test_a_bar_chart_sorted_ascending_keeps_the_category_order(tmp_path):
    _, _, (root,) = draw(tmp_path, spec("bar", [("AC", "D", "ac")], sort="asc"))
    assert_category_axis(root, reversed_=False)


def test_a_line_chart_labels_only_its_chosen_points_and_reports_the_positions_it_cannot_place(tmp_path):
    s = spec("line", [("AC", "D", "ac"), ("PY", "E", "py"), ("BU", "F", "bu"), ("FC", "G", "fc")],
             label_points=[[3, 5], [5], [], [1, 5]], label_below=[False, True, False, False],
             label_pos=[{3: "below"}, {}, {}, {}], whole=True, millions=False, title="Opex trend")
    titles, lint, (root,) = draw(tmp_path, s)
    assert titles == ["Opex trend"]
    g = group(root, "lineChart")
    assert g is not None and val(g, C + "grouping") == "standard" and g.find(C + "gapWidth") is None
    ac, py, bu, fc = sers(g)
    assert line(ac) == {"colour": "1F3864", "w": "28575", "dash": None}    # 2.25 pt
    assert line(py) == {"colour": "9AA5B9", "w": "28575", "dash": None}    # navy 55 % towards white
    assert line(bu) == {"colour": "1F3864", "w": "28575", "dash": "dash"}
    assert line(fc) == {"colour": "1F3864", "w": "28575", "dash": "sysDot"}   # round dot
    assert all(val(x.find(C + "marker"), C + "symbol") == "none" for x in (ac, py, bu, fc))
    assert label_format(ac) == '#,##0;"−"#,##0;;'                          # whole NOK
    assert val(labels(ac), C + "dLblPos") == "t" and val(labels(py), C + "dLblPos") == "b"   # a cost line: below
    assert {i: d.find(C + "delete") is not None for i, d in custom(ac).items()} == {0: True, 1: True, 3: True}
    assert set(custom(py)) == {0, 1, 2, 3} and set(custom(bu)) == {0, 1, 2, 3, 4} and set(custom(fc)) == {1, 2, 3}
    assert all(d.find(C + "delete") is not None for x in (py, bu, fc) for d in custom(x).values())
    assert legend(root) == "t"
    (ax,) = value_axes(root)
    assert_value_axis_hidden(ax)
    assert_category_axis(root)
    (msg,) = lint                                                          # one line for the chart, never an error
    assert "'Opex trend'" in msg and "position" in msg


def test_a_line_without_chosen_points_has_no_limit_line(tmp_path):
    _, lint, (root,) = draw(tmp_path, spec("line", [("BU", "D", "bu")]))
    (bu,) = sers(root)
    assert lint == [] and labels(bu).find(C + "dLbl") is None and labels(bu).find(C + "dLblPos") is None
    assert legend(root) == "t"                                             # a BU-only line must say BU


def test_a_stacked_chart_colours_its_parts_and_always_has_a_legend(tmp_path):
    s = spec("stacked", [("Mobile", "D", "part"), ("Fixed", "E", "part"), ("Other", "F", "part")])
    _, lint, (root,) = draw(tmp_path, s)
    g = group(root)
    assert lint == [] and val(g, C + "grouping") == "stacked" and val(g, C + "barDir") == "col"
    assert val(g, C + "gapWidth") == "50" and val(g, C + "overlap") == "100"
    parts = sers(g)
    assert [fill(x) for x in parts] == ["1F3864", "546E96", "8CA0BE"]
    assert all(val(labels(x), C + "dLblPos") == "ctr" for x in parts)
    assert [label_font(labels(x))["colour"] for x in parts] == ["FFFFFF", "FFFFFF", "262626"]   # white on the dark two
    assert legend(root) == "t"
    assert_value_axis_hidden(value_axes(root)[0])


def test_a_stacked_chart_over_a_header_row_points_at_that_row(tmp_path):
    s = ChartSpec("stacked", "Mix", cat_range="D10:F10", series=[("2025", "D12:F12", "part")], anchor="B30")
    _, _, (root,) = draw(tmp_path, s)
    (part,) = sers(root)
    assert_refs(part, "'Q3 Answer'!$D$12:$F$12", "'Q3 Answer'!$D$10:$F$10")
    assert val(part.find(f"{C}cat//{C}strCache"), C + "ptCount") == "3"


def test_a_stacked100_chart_is_percent_stacked_with_a_legend_for_one_part(tmp_path):
    _, _, (root,) = draw(tmp_path, spec("stacked100", [("Mobile", "D", "part")]))
    g = group(root)
    assert val(g, C + "grouping") == "percentStacked" and val(g, C + "gapWidth") == "50"
    assert fill(sers(g)[0]) == "1F3864" and legend(root) == "t"


def test_a_combo_chart_draws_pct_as_a_line_on_a_hidden_secondary_axis(tmp_path):
    s = spec("combo", [("Revenue", "D", "ac"), ("Margin %", "E", "pct")])
    _, lint, (root,) = draw(tmp_path, s)
    assert lint == []
    bars, lines = group(root), group(root, "lineChart")
    assert val(bars, C + "grouping") == "clustered" and val(bars, C + "gapWidth") == "50"
    (rev,), (pct,) = sers(bars), sers(lines)
    assert name(rev) == "Revenue" and name(pct) == "Margin %"
    assert_refs(pct, "'Q3 Answer'!$E$5:$E$9", "'Q3 Answer'!$C$5:$C$9")
    assert fill(rev) == "1F3864" and line(pct) == {"colour": "404040", "w": "28575", "dash": None}
    assert label_format(pct) == '0.0 %;"−"0.0 %;;' and val(labels(pct), C + "dLblPos") == "t"
    assert label_format(rev) == MNOK
    primary, secondary = value_axes(root)
    assert_value_axis_hidden(primary)
    assert_value_axis_hidden(secondary)
    assert val(secondary, C + "axId") in {x.get("val") for x in lines.findall(C + "axId")}   # the % line's axis
    assert legend(root) == "t"


def test_a_variance_chart_overlaps_better_and_worse_without_a_legend(tmp_path):
    s = spec("variance", [("Better", "P", "fav"), ("Worse", "Q", "unfav")])
    _, lint, (root,) = draw(tmp_path, s)
    g = group(root)
    assert lint == [] and val(g, C + "barDir") == "bar" and val(g, C + "grouping") == "clustered"
    assert val(g, C + "gapWidth") == "50" and val(g, C + "overlap") == "100"
    fav, unfav = sers(g)
    assert (fill(fav), fill(unfav)) == ("00703C", "C00000")
    assert label_format(fav) == label_format(unfav) == MNOK
    assert legend(root) is None
    (ax,) = value_axes(root)
    assert_value_axis_hidden(ax)
    assert val(ax.find(C + "scaling"), C + "orientation") == "minMax"     # reverse_values False: never reversed
    assert_category_axis(root, reversed_=True)


WATERFALL = [("Run", "P", "run"), ("Base", "Q", "base"), ("Base <0", "R", "base"), ("Up", "S", "up"),
             ("Up <0", "T", "up_neg"), ("Down", "U", "down"), ("Down <0", "V", "down_neg"), ("Total", "W", "total")]


def waterfall(**kw) -> ChartSpec:
    return spec("waterfall", [x for x in WATERFALL if x[2] != "run"], title=kw.pop("title", "EBITDA bridge"), **kw)


def test_a_waterfall_is_stacked_with_green_steps_up_red_steps_down_and_white_labels(tmp_path):
    _, lint, (root,) = draw(tmp_path, waterfall())
    assert lint == []
    g = group(root)
    assert val(g, C + "grouping") == "stacked" and val(g, C + "gapWidth") == "50" and val(g, C + "overlap") == "100"
    base, base_neg, up, up_neg, down, down_neg, total = sers(g)
    assert_refs(up, "'Q3 Answer'!$S$5:$S$9", "'Q3 Answer'!$C$5:$C$9")
    assert [fill(x) for x in (up, up_neg, down, down_neg, total)] == ["00703C", "00703C", "C00000", "C00000",
                                                                      "1F3864"]
    assert label_format(up) == label_format(up_neg) == "#,##0.0,,;#,##0.0,,;;"
    assert label_format(down) == label_format(down_neg) == '"−"#,##0.0,,;"−"#,##0.0,,;;'
    assert label_format(total) == MNOK
    assert val(labels(total), C + "dLblPos") == "inEnd"
    assert all(val(labels(x), C + "dLblPos") == "ctr" for x in (up, up_neg, down, down_neg))
    assert all(label_font(labels(x))["colour"] == "FFFFFF" and label_font(labels(x))["b"] == "1"
               for x in (up, up_neg, down, down_neg, total))
    assert legend(root) is None
    assert_value_axis_hidden(value_axes(root)[0])
    assert_category_axis(root)


def test_base_series_of_a_waterfall_are_invisible(tmp_path):
    _, _, (root,) = draw(tmp_path, waterfall())
    bases = [x for x in sers(root) if name(x).startswith("Base")]
    assert len(bases) == 2
    for b in bases:
        assert fill(b) == "none" and line(b) == {"colour": "none"} and labels(b) is None   # no fill, no line, no label


def test_thin_waterfall_steps_get_dark_ink_without_the_lift_and_one_lint_line(tmp_path):
    _, lint, (root,) = draw(tmp_path, waterfall(thin_steps=[3]))
    steps = [x for x in sers(root) if name(x) in ("Up", "Up <0", "Down", "Down <0")]
    for x in steps:
        (d,) = custom(x).values()
        assert val(d, C + "idx") == "2" and val(d, C + "dLblPos") == "ctr" and val(d, C + "showVal") == "1"
        assert label_font(d) == {"face": "Arial", "sz": "900", "b": "1", "colour": "262626"}
    total = next(x for x in sers(root) if name(x) == "Total")
    assert labels(total).find(C + "dLbl") is None                         # a total is never thin
    (msg,) = lint
    assert "'EBITDA bridge'" in msg and "10 pt" in msg


def test_a_reversed_value_axis_puts_the_category_labels_at_its_high_end(tmp_path):
    _, _, (root,) = draw(tmp_path, spec("column", [("AC", "D", "ac")], reverse_values=True))
    (ax,) = value_axes(root)
    assert val(ax.find(C + "scaling"), C + "orientation") == "maxMin"
    assert_category_axis(root, label_pos="high")


@pytest.mark.parametrize("axis_min, bound", [(120_000_000.0, "min"), (-50_000_000.0, "max")])
def test_a_truncated_waterfall_axis_starts_at_axis_min(tmp_path, axis_min, bound):
    _, _, (root,) = draw(tmp_path, waterfall(axis_min=axis_min))
    scaling = value_axes(root)[0].find(C + "scaling")
    assert float(val(scaling, C + bound)) == axis_min
    assert scaling.find(C + ("max" if bound == "min" else "min")) is None


def test_a_hatched_forecast_bar_and_the_legend_for_a_lone_comparison(tmp_path):
    _, _, (root,) = draw(tmp_path, spec("column", [("FC", "D", "fc")], hues=[2]))
    (fc,) = sers(root)
    patt = fc.find(f"{C}spPr/{A}pattFill")
    assert patt.get("prst") == "wdUpDiag"
    assert patt.find(f"{A}fgClr/{A}srgbClr").get("val") == "D68910"       # amber on white
    assert patt.find(f"{A}bgClr/{A}srgbClr").get("val") == "FFFFFF"
    assert line(fc)["colour"] == "D68910"
    assert legend(root) == "t"                                             # one series, but not AC


def test_highlighted_points_get_the_accent_and_the_rest_grey(tmp_path):
    s = spec("column", [("AC", "D", "ac"), ("PY", "E", "py")], highlight=[2, 4])
    _, _, (root,) = draw(tmp_path, s)
    ac, py = sers(root)
    assert fill(ac) == "A6A6A6"                                            # the legend key grey like its bars
    points = {int(val(p, C + "idx")): p.find(f"{C}spPr/{A}solidFill/{A}srgbClr").get("val")
              for p in ac.findall(C + "dPt")}
    assert points == {0: "A6A6A6", 1: "1C16C5", 2: "A6A6A6", 3: "1C16C5", 4: "A6A6A6"}
    assert fill(py) == "FFFFFF" and line(py) == {"colour": "A6A6A6", "w": "12700", "dash": None}
    assert py.find(C + "dPt") is None


def test_a_highlight_on_a_line_chart_is_not_drawn(tmp_path):
    _, _, (root,) = draw(tmp_path, spec("line", [("AC", "D", "ac")], highlight=[2]))
    (ac,) = sers(root)
    assert ac.find(C + "dPt") is None and line(ac)["colour"] == "1F3864"


# ------------------------------------------------------------------ placing, skipping, failing
def test_a_chart_is_placed_at_its_anchor_480_by_260_pt_and_named(tmp_path):
    path = tmp_path / "charts.xlsx"
    with xlsxwriter.Workbook(str(path)) as wb:
        ws = wb.add_worksheet(SHEET)
        add_charts(wb, ws, SHEET, SheetPlan(charts=[spec("column", [("AC", "D", "ac")], anchor="C40")]))
    (anchor,) = drawing_of(path).findall(XDR + "twoCellAnchor")
    assert (anchor.find(f"{XDR}from/{XDR}col").text, anchor.find(f"{XDR}from/{XDR}row").text) == ("2", "39")
    assert anchor.find(f".//{XDR}cNvPr").get("name") == "smartass chart"
    to = {t: int(anchor.find(f"{XDR}to/{XDR}{t}").text) for t in ("col", "colOff", "row", "rowOff")}
    width = (to["col"] - 2) * 64 + to["colOff"] / 9525                     # default columns: 64 px, rows: 20 px
    height = (to["row"] - 39) * 20 + to["rowOff"] / 9525
    assert (width, height) == (640, 347)                                   # 480 x 260 pt at 96 dpi


def test_a_spec_without_series_is_skipped_and_not_reported_as_drawn(tmp_path):
    empty = ChartSpec("column", "Plan only", cat_range="C5:C9", series=[], anchor="B30")
    titles, lint, roots = draw(tmp_path, empty, spec("bar", [("AC", "D", "ac")], title="Drawn"),
                               lint=["chart 'Plan only': no series has data; left out"])
    assert titles == ["Drawn"] and len(roots) == 1
    assert lint == ["chart 'Plan only': no series has data; left out"]    # the layout's lint, nothing added


def test_a_chart_that_cannot_be_built_is_one_lint_line_and_the_others_are_drawn(tmp_path):
    bad = spec("pie", [("AC", "D", "ac")], title="Pie")
    titles, lint, roots = draw(tmp_path, spec("column", [("AC", "D", "ac")], title="First"), bad,
                               spec("line", [("AC", "D", "ac")], title="Last"), lint=["layout note"])
    assert titles == ["First", "Last"]
    assert [name(sers(r)[0]) for r in roots] == ["AC", "AC"] and group(roots[1], "lineChart") is not None
    assert lint[0] == "layout note" and len(lint) == 2
    assert lint[1].startswith("chart 'Pie' failed: ") and "fix: " in lint[1]


def test_a_combo_without_a_column_series_fails_that_chart_only(tmp_path):
    bad = spec("combo", [("Margin %", "E", "pct")], title="Only a line")
    titles, lint, roots = draw(tmp_path, bad, spec("column", [("AC", "D", "ac")], title="Kept"))
    assert titles == ["Kept"] and len(roots) == 1
    (msg,) = lint
    assert msg.startswith("chart 'Only a line' failed: ")


def test_an_option_xlsxwriter_refuses_fails_that_chart_only(tmp_path, monkeypatch):
    real = xlsx_charts._labels

    def refused(spec, role, k, n):                 # a position a line chart has not: XlsxWriter warns, drops it
        return real(spec, role, k, n) | ({"position": "inside_base"} if spec.kind == "line" else {})
    monkeypatch.setattr(xlsx_charts, "_labels", refused)
    titles, lint, roots = draw(tmp_path, spec("line", [("AC", "D", "ac")], title="Refused"),
                               spec("column", [("AC", "D", "ac")], title="Kept"))
    assert titles == ["Kept"] and len(roots) == 1 and group(roots[0]) is not None
    (msg,) = lint
    assert msg.startswith("chart 'Refused' failed: ") and "inside_base" in msg


# ------------------------------------------------------------------ in the built file
def test_the_built_file_carries_the_answer_charts_and_reports_them(tmp_path):
    built = build_with(tmp_path)
    assert built.charts == ["P&L vs last year"] and built.chart_lint == []
    (root,) = charts_of(built.path)
    ac, py = sers(root)
    assert_refs(ac, "'Q3 Answer'!$I$26:$I$27", "'Q3 Answer'!$C$26:$C$27")
    assert ref(py, "val") == "'Q3 Answer'!$J$26:$J$27"
    with zipfile.ZipFile(built.path) as z:
        rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    assert "drawing1.xml" in rels                                          # on the Answer sheet, the first one
