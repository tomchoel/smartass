"""TN rapportstil: ren rapportformatering (årsrapport-/konsulentstil) for Excel-ark.

Pipeline:  roller (openpyxl, fra en arkspesifikasjon + automatisk gjenkjenning)
        -> stiler og formatering i ekte Excel (COM)
        -> kontroll (lint) av lagret fil
        -> PDF-eksport for visuell sjekk.

Alle celler får en navngitt cellestil "TN ...". Endrer man en stil, endres alle cellene med den.
Tallformater skrives i US-syntaks og oversettes til maskinens lokale syntaks ved kjøring.
"""
import os, re, sys, math, itertools
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

# ------------------------------------------------------------------ design tokens
FONT = "Arial"
NAVY, INK, GREY, RULE_DARK = (31, 56, 100), (38, 38, 38), (118, 118, 118), (64, 64, 64)
INPUT_BLUE, INPUT_FILL, KEY_FILL = (0, 0, 255), (255, 247, 221), (237, 242, 249)
BREACH_FONT, BREACH_FILL = (192, 0, 0), (253, 233, 233)

NUM = {  # US-syntaks; oversettes til lokal syntaks ved kjøring
    "belop":    '#,##0.00_);(#,##0.00);"–"_)',
    "prosent":  '0.00 %_);(0.00 %);"–"_)',
    "prosent1": '0.0 %_);(0.0 %);"–"_)',
    "tall1":    '0.0_);(0.0);"–"_)',
    "heltall":  '#,##0_);(#,##0);"–"_)',   # antall (abonnenter, stk.)
    "faktor":   '0.000_);(0.000);"–"_)',   # påslag og andre multiplikatorer
}
KIND_NAME = {"etikett": "Etikett", "belop": "Beløp", "prosent": "Prosent", "prosent1": "Prosent 1d", "tall1": "Tall 1d", "heltall": "Heltall", "faktor": "Faktor"}
EMPH = ["", "Input", "Nøkkel", "Sum", "Hovedsum"]

# COM-konstanter
L, T, B, R = 7, 8, 9, 10          # Range.Borders (xlEdge*)
S_L, S_T, S_B, S_R = -4131, -4160, -4107, -4152   # Style.Borders (xlLeft/Top/Bottom/Right)
CONT, DOUBLE, NONE = 1, -4119, -4142
THIN, MEDIUM, HAIR = 2, -4138, 1
H_LEFT, H_RIGHT, H_GEN = -4131, -4152, 1
V_BOTTOM, V_CENTER, V_TOP = -4107, -4108, -4160


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def style_defs():
    """Alle TN-stiler: navn -> egenskaper."""
    d = {
        "TN Tittel":      dict(size=16, bold=True, color=NAVY, h=H_LEFT, v=V_BOTTOM, bottom=(MEDIUM, NAVY)),
        "TN Undertittel": dict(size=9, italic=True, color=GREY, h=H_LEFT, v=V_CENTER),
        "TN Seksjon":     dict(size=11, bold=True, color=NAVY, h=H_LEFT, v=V_BOTTOM, bottom=(THIN, NAVY)),
        "TN Kolonne":     dict(size=9, bold=True, color=NAVY, h=H_LEFT, v=V_BOTTOM, wrap=True, bottom=(THIN, NAVY)),
        "TN Kolonne tall": dict(size=9, bold=True, color=NAVY, h=H_RIGHT, v=V_BOTTOM, wrap=True, bottom=(THIN, NAVY)),
        "TN Note":        dict(size=9, italic=True, color=GREY, h=H_LEFT, v=V_CENTER, indent=1),
        "TN Tekst":       dict(size=10, color=INK, h=H_LEFT, v=V_TOP, wrap=True),
        "TN Tekst tallkolonne": dict(size=10, color=INK, h=H_RIGHT, v=V_CENTER),   # f.eks. «n.a.» blant tall
        "TN ID":          dict(size=10, color=INK, h=H_RIGHT, v=V_CENTER, num="0"),  # løpenummer, ikke et beløp
        "TN Lenke":       dict(size=10, color=NAVY, h=H_LEFT, v=V_CENTER, underline=True),
        "TN Kolonne faktor": dict(size=9, bold=True, color=NAVY, h=H_RIGHT, v=V_BOTTOM, bottom=(THIN, NAVY), num='0.000'),
        "TN Stortall etikett": dict(size=9, color=GREY, h=H_LEFT, v=V_BOTTOM),
        "TN Stortall prosent": dict(size=22, bold=True, color=NAVY, h=H_LEFT, v=V_CENTER, num='0.0 %;(0.0 %);"–"'),
        "TN Stortall tall1":   dict(size=22, bold=True, color=NAVY, h=H_LEFT, v=V_CENTER, num='#,##0.0;(#,##0.0);"–"'),
        "TN Stortall tekst":   dict(size=16, bold=True, color=NAVY, h=H_LEFT, v=V_CENTER),
        "TN Tom":         dict(size=10, color=INK, h=H_GEN, v=V_CENTER),
        "TN Produkt privat":  dict(size=10, color=INK, h=H_LEFT, v=V_CENTER, num='@"  · privat"'),
        "TN Produkt bedrift": dict(size=10, color=INK, h=H_LEFT, v=V_CENTER, num='@"  · bedrift"'),
    }
    for kind, label in KIND_NAME.items():
        for emph in EMPH:
            name = f"TN {label}" + (f" · {emph}" if emph else "")
            p = dict(size=10, color=INK, h=H_LEFT if kind == "etikett" else H_RIGHT, v=V_CENTER)
            if kind != "etikett":
                p["num"] = NUM[kind]
                p["color"] = (0, 0, 0)
            if emph == "Input" and kind != "etikett":
                p.update(color=INPUT_BLUE, fill=INPUT_FILL)
            if emph == "Nøkkel":
                p.update(bold=True, fill=KEY_FILL)
            if emph in ("Sum", "Hovedsum"):
                p.update(bold=True, top=(THIN, RULE_DARK))
            if emph == "Hovedsum":
                p.update(bottom=(DOUBLE, RULE_DARK))
            d[name] = p
    return d


def style_name(kind, emph=""):
    return f"TN {KIND_NAME[kind]}" + (f" · {emph}" if emph else "")


# ------------------------------------------------------------------ locale
def to_local(fmt, dec, thou):
    out, q, esc = [], False, False
    for ch in fmt:
        if esc:
            out.append(ch); esc = False; continue
        if ch == "\\":
            out.append(ch); esc = True; continue
        if ch == '"':
            q = not q; out.append(ch); continue
        if not q and ch == ".":
            out.append("\x00"); continue
        if not q and ch == ",":
            out.append(thou); continue
        out.append(ch)
    return "".join(out).replace("\x00", dec)


# ------------------------------------------------------------------ tekstmål (kalibrert mot Excel-render, Arial)
# Arial tegnbredder (1/1000 em), vanlig og fet. Excel-kolonneenhet = bredden av "0" (556).
_AR = dict(zip("abcdefghijklmnopqrstuvwxyzæøå", [556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,889,611,556]))
_AR.update(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ", [667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,1000,778,667]))
_AR.update({c: 556 for c in "0123456789"}); _AR.update({" ": 278, "(": 333, ")": 333, "/": 278, "%": 889, ".": 278, ",": 278, "-": 333, "–": 556, ":": 278, "=": 584, "≥": 549, "·": 278})
_ARB = dict(zip("abcdefghijklmnopqrstuvwxyzæøå", [556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,889,611,556]))
_ARB.update(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ", [722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,1000,778,722]))
_ARB.update({c: 556 for c in "0123456789"}); _ARB.update({" ": 278, "(": 333, ")": 333, "/": 278, "%": 889, ".": 278, ",": 278, "-": 333, "–": 556, ":": 333, "=": 584, "≥": 549, "·": 278})


def text_units(text, size=10, bold=False):
    """Tekstbredde i Excel-kolonneenheter (Arial), skalert med skriftstørrelse."""
    tbl = _ARB if bold else _AR
    return sum(tbl.get(ch, 600) for ch in str(text)) / 556 * size / 10 * 1.08   # pikselavrunding i Excel


def wrap_lines(text, width, size=10, bold=False):
    """Antall linjer en ordbrutt celle trenger, eller None hvis et enkeltord er bredere enn cellen."""
    avail = width - 0.6   # cellemarg
    n, cur = 1, 0.0
    sp = text_units(" ", size, bold)
    for word in str(text).split():
        wl = text_units(word, size, bold)
        if wl > avail:
            return None
        if cur and cur + sp + wl > avail:
            n, cur = n + 1, wl
        else:
            cur = cur + (sp if cur else 0) + wl
    return n


def column_widths(ws_openpyxl, last_col):
    """Kolonnebredder, også når Excel har lagret dem som et område (f.eks. C:I)."""
    w = {c: 8.43 for c in range(1, last_col + 1)}
    for dim in ws_openpyxl.column_dimensions.values():
        if dim.width and dim.min and dim.max:
            for c in range(dim.min, min(dim.max, last_col) + 1):
                w[c] = dim.width
    return w


# ------------------------------------------------------------------ roles
def kind_from_format(fmt):
    f = (fmt or "").replace("\\", "")
    if "%" in f:
        return "prosent" if re.search(r"0[.,]00", f) else "prosent1"
    if re.search(r"0[.,]000", f):
        return "faktor"
    if "#" in f and not re.search(r"0[.,]0", f):
        return "heltall"
    if "#" in f and re.search(r"0[.,]00", f):
        return "belop"
    if re.search(r"0[.,]0(?!0)", f):
        return "tall1"
    return "belop"


def _cells(rng):
    c1, r1, c2, r2 = range_boundaries(rng)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield f"{get_column_letter(c)}{r}"


def _span(row, cols):
    a, b = cols.split(":")
    return f"{a}{row}:{b}{row}"


def build_roles(xlsx_path, sheet, spec):
    """Returnerer {celle: stilnavn} for hele det brukte området + regellinjer."""
    wf = openpyxl.load_workbook(xlsx_path)[sheet]
    wv = openpyxl.load_workbook(xlsx_path, data_only=True)[sheet]
    last_col = spec["last_col"]
    roles = {}
    fixed = {}

    def put(rng, name):
        for a in _cells(rng):
            fixed[a] = name

    put(spec["title"], "TN Tittel"); put(spec["subtitle"], "TN Undertittel")
    for a in _cells(_span(1, f"A:{last_col}")):
        fixed.setdefault(a, "TN Tittel")
    for r in spec["sections"]:
        put(_span(r, f"A:{last_col}"), "TN Seksjon")
    for r, cols in spec["headers"].items():
        c1, c2 = cols.split(":")
        put(f"{c1}{r}", "TN Kolonne")
        nxt = get_column_letter(column_index_from_string(c1) + 1)
        if column_index_from_string(nxt) <= column_index_from_string(c2):
            put(f"{nxt}{r}:{c2}{r}", "TN Kolonne tall")
    for rng in spec.get("header_left", []):
        put(rng, "TN Kolonne")
    for rng in spec.get("header_right", []):
        put(rng, "TN Kolonne tall")
    for rng in spec.get("notes", []):
        put(rng, "TN Note")
    for seg, rng in spec.get("products", {}).items():
        put(rng, f"TN Produkt {seg}")
    for rng, name in spec.get("styles", {}).items():   # eksplisitt stil der rolle ikke kan utledes
        put(rng, name)
    for rng in spec.get("header_num", []):
        put(rng, "TN Kolonne faktor")
    for rng in spec.get("wrap", []):      # ombrutt brødtekst (lange beskrivelser i tabeller)
        put(rng, "TN Tekst")
    for rng in spec.get("ids", []):
        put(rng, "TN ID")
    for rng in spec.get("links", {}):
        put(rng, "TN Lenke")
    ints = {a for rng in spec.get("integers", []) for a in _cells(rng)}
    numcols = {a for rng in spec.get("num_cols", []) for a in _cells(rng)}

    emph = {}
    for key, e in (("inputs", "Input"), ("key", "Nøkkel"), ("totals", "Sum"), ("grand_totals", "Hovedsum")):
        for rng in spec.get(key, []):
            for a in _cells(rng):
                emph[a] = e

    c_max = column_index_from_string(last_col)
    for row in wf.iter_rows(min_row=1, max_row=max(wf.max_row, spec.get("min_rows", 0)), max_col=c_max):
        for c in row:
            a = c.coordinate
            if a in fixed:
                roles[a] = fixed[a]; continue
            v = wv[a].value
            e = emph.get(a, "")
            if v is None:
                if e in ("Sum", "Hovedsum", "Nøkkel"):
                    roles[a] = style_name("etikett", e)
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                roles[a] = style_name("heltall" if a in ints else kind_from_format(c.number_format), e)
            else:
                roles[a] = "TN Tekst tallkolonne" if a in numcols and not e else style_name("etikett", e if e != "Input" else "")
    return roles


# ------------------------------------------------------------------ Excel (COM)
def _chunks(addrs, sep, limit=240):
    cur, n = [], 0
    for a in addrs:
        if n + len(a) + 1 > limit and cur:
            yield sep.join(cur); cur, n = [], 0
        cur.append(a); n += len(a) + 1
    if cur:
        yield sep.join(cur)


def ensure_styles(xl, wb):
    tup = xl.International
    dec, thou = tup[2], tup[3]
    for name, p in style_defs().items():
        try:
            s = wb.Styles(name)
        except Exception:
            s = wb.Styles.Add(name)
        for flag in ("IncludeFont", "IncludeAlignment", "IncludeBorder", "IncludePatterns"):
            setattr(s, flag, True)
        s.IncludeNumber = bool(p.get("num"))
        s.IncludeProtection = False
        s.Font.Name = FONT
        s.Font.Size = p.get("size", 10)
        s.Font.Bold = p.get("bold", False)
        s.Font.Italic = p.get("italic", False)
        s.Font.Underline = 2 if p.get("underline") else NONE   # xlUnderlineStyleSingle
        s.Font.Color = rgb(p.get("color", INK))
        s.HorizontalAlignment = p.get("h", H_GEN)
        s.VerticalAlignment = p.get("v", V_CENTER)
        s.WrapText = p.get("wrap", False)
        s.IndentLevel = p.get("indent", 0)
        if p.get("num"):
            s.NumberFormatLocal = to_local(p["num"], dec, thou)
        if p.get("fill"):
            s.Interior.Pattern = 1
            s.Interior.Color = rgb(p["fill"])
        else:
            s.Interior.Pattern = NONE
        for edge in (S_L, S_T, S_B, S_R):
            s.Borders(edge).LineStyle = NONE
        for edge, key in ((S_T, "top"), (S_B, "bottom")):
            if p.get(key):
                w, col = p[key]
                bd = s.Borders(edge)
                if w == DOUBLE:
                    bd.LineStyle = DOUBLE
                else:
                    bd.LineStyle = CONT; bd.Weight = w
                bd.Color = rgb(col)


def apply(xl, wb, sheet, spec, roles):
    ws = wb.Worksheets(sheet)
    last_col = spec["last_col"]
    used = f"A1:{last_col}{max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, spec.get('min_rows', 0))}"
    ws.Range(used).ClearFormats()          # fjerner direkte formatering og betinget formatering
    # lenker til cellen det vises til: arknavn i en annen kolonne på samme rad, adressen er celleverdien
    for rng, sheet_col in spec.get("links", {}).items():
        for a in _cells(rng):
            addr = ws.Range(a).Value
            if not addr:
                continue
            r = int(re.findall(r"\d+", a)[0])
            target = ws.Range(f"{sheet_col}{r}").Value
            ws.Range(a).Hyperlinks.Delete()
            ws.Hyperlinks.Add(Anchor=ws.Range(a), Address="", SubAddress=f"'{target}'!{addr}",
                              ScreenTip=f"Gå til {target}!{addr}", TextToDisplay=str(addr))
    ws.Range(used).Style = "TN Tom"
    by_style = {}
    for a, s in roles.items():
        by_style.setdefault(s, []).append(a)
    sep = xl.International[4]  # listeskilletegn (';' på norsk Excel)
    for s, addrs in by_style.items():
        for chunk in _chunks(addrs, sep):
            ws.Range(chunk).Style = s
    # regellinje mellom segmenter (fin, grå) – brukes på hele tabellbredden
    for r, cols in spec.get("segment_breaks", {}).items():
        bd = ws.Range(_span(r, cols)).Borders(T)
        bd.LineStyle = CONT; bd.Weight = HAIR; bd.Color = rgb((166, 166, 166))
    # kolonnebredder og radhøyder
    for cols, wdt in spec["col_widths"].items():
        ws.Columns(cols if ":" in cols else f"{cols}:{cols}").ColumnWidth = wdt
    last_row = int(re.findall(r"\d+", used.split(":")[1])[0])
    ws.Range(f"A3:A{last_row}").EntireRow.RowHeight = 15
    ws.Range("A1").EntireRow.RowHeight = 28
    ws.Range("A2").EntireRow.RowHeight = 16
    for r in spec["sections"]:
        ws.Range(f"A{r}").EntireRow.RowHeight = 22
    for r, cols in spec["headers"].items():
        c1, c2 = (column_index_from_string(x) for x in cols.split(":"))
        extra = [column_index_from_string(re.match(r"[A-Z]+", a).group()) for a in spec.get("header_left", []) if int(re.findall(r"\d+", a)[0]) == r]
        lines = 1
        for ci in list(range(c1, c2 + 1)) + extra:
            txt = ws.Cells(r, ci).Value
            if txt:
                lines = max(lines, wrap_lines(str(txt), ws.Cells(r, ci).ColumnWidth, 9, True) or 4)
        ws.Range(f"A{r}").EntireRow.RowHeight = max(26, lines * 12 + 3)
    # rader med ombrutt brødtekst: høyde etter lengste celle
    wrap_rows = {}
    for rng in spec.get("wrap", []):
        for a in _cells(rng):
            txt = ws.Range(a).Value
            if txt:
                n = wrap_lines(str(txt), ws.Range(a).ColumnWidth) or 4
                r = int(re.findall(r"\d+", a)[0])
                wrap_rows[r] = max(wrap_rows.get(r, 1), n)
    for r, n in wrap_rows.items():
        ws.Range(f"A{r}").EntireRow.RowHeight = max(15, n * 13 + 2)
    for r, h in spec.get("row_heights", {}).items():
        ws.Range(f"A{r}").EntireRow.RowHeight = h
    for rng in spec.get("merges", []):
        ws.Range(rng).Merge()
    # tabeller med flerlinjers rader: toppjuster hele raden så tall og tekst står på samme linje
    for rows in spec.get("valign_top", []):
        ws.Range(f"A{rows.split(':')[0]}:{last_col}{rows.split(':')[1]}").VerticalAlignment = V_TOP
    # betinget formatering: kun brudd (verdi < 0) i rødt
    for rng in spec.get("breach", []):
        fc = ws.Range(rng).FormatConditions.Add(Type=1, Operator=6, Formula1="=0")
        fc.Font.Color = rgb(BREACH_FONT); fc.Font.Bold = True; fc.Interior.Color = rgb(BREACH_FILL)
    # visning og utskrift
    ws.Activate()
    win = wb.Windows(1)
    win.DisplayGridlines = False
    try:
        win.FreezePanes = False; win.SplitColumn = 0; win.SplitRow = spec.get("freeze_rows", 2); win.FreezePanes = True
    except Exception:
        pass
    win.Zoom = 100
    ws.Tab.Color = rgb(NAVY)
    ps = ws.PageSetup
    ps.PrintArea = spec.get("print_area", used)     # flere områder (skilt med komma) gir én side per område
    ps.PrintTitleRows = spec.get("print_titles", "$1:$2")
    ps.Orientation = 2
    ws.ResetAllPageBreaks()
    if spec.get("page_breaks"):          # faste sideskift krever fast zoom (tilpass-til-side overstyrer skiftene)
        ps.Zoom = spec.get("print_zoom", 65)
        for r in spec["page_breaks"]:
            ws.HPageBreaks.Add(ws.Range(f"A{r}"))
    else:
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = spec.get("pages_tall", 2)
    ps.CenterHorizontally = True
    ps.LeftMargin = ps.RightMargin = xl.CentimetersToPoints(1.5)
    ps.TopMargin = ps.BottomMargin = xl.CentimetersToPoints(1.8)
    ps.CenterHeader = '&"Arial"&8' + spec.get("print_header", "Intern – inneholder grossistpriser, abonnementstall og marginer")
    ps.RightFooter = '&"Arial"&8Side &P av &N'
    ps.LeftFooter = '&"Arial"&8&F'
    ws.Range("A1").Select() if xl.Visible else None


def export_pdf(xl, wb, sheet, pdf_path):
    wb.Worksheets(sheet).ExportAsFixedFormat(0, pdf_path)


# ------------------------------------------------------------------ lint
def lint(xlsx_path, sheet, spec):
    """Returnerer en liste med brudd. Tom liste = bestått."""
    wf = openpyxl.load_workbook(xlsx_path)[sheet]
    wv = openpyxl.load_workbook(xlsx_path, data_only=True)[sheet]
    last_col = column_index_from_string(spec["last_col"])
    widths = column_widths(wf, last_col)
    allowed_fmt = set()
    errs = []
    fonts = set()
    for row in wf.iter_rows():
        for c in row:
            v = wv[c.coordinate].value
            if v is None:
                continue
            if c.column > last_col:
                errs.append(f"{c.coordinate}: innhold utenfor tabellbredden (til {spec['last_col']})")
            if not str(c.style).startswith("TN "):
                errs.append(f"{c.coordinate}: mangler TN-stil (har '{c.style}')")
            fonts.add(c.font.name)
            if isinstance(v, (int, float)) and c.number_format in ("General", "Standard"):
                errs.append(f"{c.coordinate}: tall uten tallformat")
            if isinstance(v, (int, float)) and "%" in c.number_format:
                allowed_fmt.add(c.number_format)
            if c.border.left.style or c.border.right.style:
                errs.append(f"{c.coordinate}: loddrett strek (ikke tillatt i rapportstil)")
            if isinstance(v, str) and v.startswith("="):
                errs.append(f"{c.coordinate}: formel lagret som tekst ('{v[:30]}')")
            if isinstance(v, str):
                size = c.font.sz or 10
                need = text_units(v, size, bool(c.font.b))
                if c.alignment.wrap_text:
                    h = wf.row_dimensions[c.row].height or 15
                    lines = wrap_lines(v, widths.get(c.column, 8.43), size, bool(c.font.b))
                    if lines is None or lines * (size + 3) > h + 1:
                        errs.append(f"{c.coordinate}: overskriften får ikke plass ({lines or 'ord for langt'} linjer, {h:.0f} pt)")
                elif c.alignment.horizontal != "right":
                    avail = 0
                    for k in range(c.column, last_col + 1):
                        if k > c.column and wv.cell(c.row, k).value is not None:
                            break
                        avail += widths[k]
                    if need > avail:
                        errs.append(f"{c.coordinate}: teksten kuttes eller renner ut (~{need:.0f} > {avail:.0f} tegn)")
            st = str(c.style)
            if st in ("TN Seksjon", "TN Kolonne", "TN Kolonne tall", "TN Tittel") and not c.border.bottom.style:
                errs.append(f"{c.coordinate}: {st} mangler strek under")
            if (" · Sum" in st or " · Hovedsum" in st) and not c.border.top.style:
                errs.append(f"{c.coordinate}: {st} mangler strek over")
            if " · Hovedsum" in st and c.border.bottom.style != "double":
                errs.append(f"{c.coordinate}: hovedsum mangler dobbel strek under")
    if len(fonts) > 1:
        errs.append(f"flere fonter i bruk: {sorted(fonts)}")
    if len(allowed_fmt) > 2:
        errs.append(f"for mange prosentformater: {sorted(allowed_fmt)}")
    return errs
