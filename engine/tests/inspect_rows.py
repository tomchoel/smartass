"""Show labels / formulas / values of given rows in source vs refreshed workbook (read-only)."""
import sys
import openpyxl

src, out, sheet, r0, r1 = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5])
for tag, path in (("SRC", src), ("OUT", out)):
    f = openpyxl.load_workbook(path, data_only=False)[sheet]
    v = openpyxl.load_workbook(path, data_only=True)[sheet]
    print(f"===== {tag}")
    for r in [1, 2, 3] + list(range(r0, r1 + 1)):
        cells = []
        for c in range(1, 9):
            fv, vv = f.cell(r, c).value, v.cell(r, c).value
            if fv is None and vv is None:
                continue
            s = f"{f.cell(r, c).coordinate}={vv!r}"
            if isinstance(fv, str) and fv.startswith("="):
                s += f" [{fv[:60]}]"
            cells.append(s)
        print(f"r{r}: " + " | ".join(cells))
