"""TN rapportstil, trimmed: the design tokens, the TN cell styles, local number-format syntax and the text-width
measures that styling.py uses. The skill's roles, apply, lint and PDF export are not vendored.
"""
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

