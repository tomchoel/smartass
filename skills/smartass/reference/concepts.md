# Products and product concepts

## Products

A query takes `products` (FinMax/FinLite `pd.SY.<code>`, FinPlan e.g. `Mobile`; find them with
`fin_lookup("<code or name>", kind="member")`). Without products a query reads the cube's product total, as
always. With products the grid adds the dimension `All Products` (all three cubes) as a third row dimension, the
Data table gains a `Product` column and every number of the query sums the listed products. Unknown products stop
the answer.

`table` and `ranking` take `"rows": "products"`: one line per product of the query (top N, Other, and the total
over the listed products).

```json
{"queries": [{"id": "q1", "cube": "FinMax", "metrics": ["Revenue"], "entity": "un.SY.2WHS",
              "periods": "closed-ytd", "compare": ["ly"], "products": ["pd.SY.P502", "pd.SY.P542", "pd.SY.P540"]}],
 "blocks": [{"type": "ranking", "query": "q1", "metric": "Revenue", "top": 10, "rows": "products"}]}
```

## Concepts

A product concept is an approved definition in `knowledge/semantic/concepts/<id>.toml`: `splitt` (handsets on
instalments, hand-made with allocation keys), `streamix`, and 22 product families `fam_*` (booked figures only).
Find one with `fin_lookup("splitt", kind="concept")` (blocks, items, notes).

```json
{"queries": [{"id": "q1", "cube": "FinMax", "metrics": [], "entity": "un.SY", "periods": "closed-ytd",
              "compare": ["ly"]}],
 "blocks": [{"type": "concept", "query": "q1", "concept": "splitt"}]}
```

- The query gives the cube (FinMax or FinLite), the entity, the periods and the compare; the concept gives the
  accounts and products. A query that serves concept blocks alone may have `metrics: []`.
- One section per concept block, a heading per concept group, one line per item:
  - **booked**: tagged on the products in the cube (account x products).
  - **attributed**: not tagged, belongs to the concept in full (Key share 100 %, a typed input).
  - **estimated**: a shared base x a key share (num / den, or another item's key).
  - **missing**: not in the cube: no value; its note comes as a caveat.
  - sums, the total and the checks.
- Base and Key share sit in their own helper columns (and again for last year): the line is `=Base*Key share`, so
  no share is ever a constant inside a formula. Change a key in the sheet and the line follows.
- A check that does not come to zero is a **finding** (warning), never a stop: a concept is an estimate.
- The concept's notes, its missing items and how each item is estimated come as `info` caveats: pass on those
  that matter.
