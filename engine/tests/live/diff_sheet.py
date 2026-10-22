"""For one sheet: list differing rows with the row labels (any text in the row) and per-row sums."""
import collections
import sys

import openpyxl

src, out, sheet = sys.argv[1], sys.argv[2], sys.argv[3]
lim = int(sys.argv[4]) if len(sys.argv) > 4 else 30


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True)[sheet]
    nums, labels = {}, collections.defaultdict(list)
    hdr = {}
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                nums[(c.row, c.column)] = float(v)
            elif isinstance(v, str) and v.strip():
                labels[c.row].append(v.strip())
                if c.row <= 8:
                    hdr[(c.row, c.column)] = v.strip()
    return nums, labels, hdr


a, la, ha = load(src)
b, _, _ = load(out)
print("header:", {k: v for k, v in sorted(ha.items())[:30]})
rows = collections.defaultdict(list)
for k in set(a) & set(b):
    if abs(a[k] - b[k]) > 0.5:
        rows[k[0]].append(k)
same_rows = {k[0] for k in set(a) & set(b)} - set(rows)
print(f"rows with differences: {len(rows)}; rows fully identical: {len(same_rows)}")
for r in sorted(rows)[:lim]:
    ks = rows[r]
    print(f" r{r} {' / '.join(la[r])[:110]} | cells={len(ks)} archive={sum(a[k] for k in ks):,.0f} refreshed={sum(b[k] for k in ks):,.0f}")
