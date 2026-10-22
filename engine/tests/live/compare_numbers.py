"""Compare numeric cells of an archive workbook with its refreshed copy; group differences by sheet and
show the row/column labels so a difference can be explained (reorg, open period, restatement).
Usage: python compare_numbers.py <archive> <refreshed> [tolerance]"""
import collections
import sys

import openpyxl

src, out = sys.argv[1], sys.argv[2]
tol = float(sys.argv[3]) if len(sys.argv) > 3 else 0.5


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    cells, text = {}, {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, bool) or v is None:
                    continue
                if isinstance(v, (int, float)):
                    cells[(ws.title, c.row, c.column)] = float(v)
                elif isinstance(v, str) and c.column <= 4:
                    text[(ws.title, c.row, c.column)] = v.strip()
    wb.close()
    return cells, text


a, at = load(src)
b, _ = load(out)
common = set(a) & set(b)
diff = [k for k in common if abs(a[k] - b[k]) > tol]
by_sheet = collections.Counter(k[0] for k in common)
d_sheet = collections.Counter(k[0] for k in diff)
print(f"numeric cells: archive={len(a)} refreshed={len(b)} common={len(common)} "
      f"identical={len(common) - len(diff)} differing={len(diff)}")
for s in by_sheet:
    print(f"  sheet {s!r}: {by_sheet[s]} cells, {d_sheet.get(s, 0)} differ")
rows = collections.defaultdict(list)
for k in diff:
    rows[(k[0], k[1])].append(k)
print("differing rows (labels from columns A-D):")
for (s, r), ks in sorted(rows.items())[:25]:
    lab = " / ".join(at.get((s, r, c), "") for c in range(1, 5) if at.get((s, r, c)))
    tot_a = sum(a[k] for k in ks)
    tot_b = sum(b[k] for k in ks)
    print(f"  {s} r{r} [{lab[:90]}] cells={len(ks)} archive_sum={tot_a:,.0f} refreshed_sum={tot_b:,.0f}")
if len(rows) > 25:
    print(f"  ... {len(rows) - 25} more rows")
