# smartass v2 · Plan 2C: products and product concepts

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Product profitability in v2.
- **Product dimension.** It is a first-class query dimension, in FinMax and FinLite ("All Products", members
  `pd.SY.*`) and in FinPlan ("All Products"; check the member names in the atlas).
- **Concepts.** The approved concepts from `knowledge/concepts.json` become semantic **concept** files: `splitt`
  (hand-made with keys), `streamix`, and the 22 `fam_*` families (booked only).
- **Concept block.** A concept answers through a new `concept` block. It shows every block of the concept (booked /
  attributed / estimated / missing items) with live SUMIFS formulas, a total and the concept's checks.

**Also in this plan:** `knowledge.db` is rebuilt automatically when the semantic files changed (hash differs), not
only when it is missing.

**Conventions:** as Plan 2B.
- Branch `v2/2c` from `main`. TDD.
- Gates after every task: pytest, ruff, and the test set at 8/8 with goldens unchanged.
- Never run python through a stdin heredoc.
- Stage files explicitly.
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

### Task 0: Branch

### Task 1: Rebuild knowledge.db when the semantic files changed

The resolver factory used by `run_prepare`, the test set and `fin_lookup` compares the semantic hash stored in
`knowledge.db` meta with the current files. When they differ, it rebuilds first and adds a warning ("knowledge.db
rebuilt: semantic files changed"). A build error stops the answer with a GuardError naming it.

**Test:** change a semantic file in a temp dir; the next resolver build rebuilds.

### Task 2: Product dimension end to end

- **Spec.** `Query.products: list[str] = []`. Empty means the cube's product total, as today.
- **Requests.** `GridRequest.products`. With products, the grid adds Products as a third row dimension
  (Accounts × Entities × Products). FinPlan uses its product dimension name (check `epm` grids and the atlas
  dimension list).
- **Facts.** `Fact.product` (default `""`). The Data table gains a `Product` column. `Pick` filters on Product when
  the query has products.
- **Constant columns.** Product joins the constant-column omission from Plan 1E. An empty product is left out of
  the formula.
- **Validation.** Product members must exist in the atlas (the guard).
- **Blocks.** Table, ranking and matrix can take `rows: "products"` (the query's products) as well as entities. Keep
  this to `table` and `ranking` in this plan (YAGNI for the others).
- **Tests:**
  - provider grids with products (fake transport);
  - pick and formulas with Product;
  - a ranking by product;
  - unchanged formulas when there are no products, so the goldens are untouched.

### Task 3: Concept files from concepts.json

- **Schema.** A semantic `Concept` model in `knowledge/semantic/concepts/<id>.toml`:
  - `id`, `title`, `names`/aliases, `kind` (product_offer / product_family), `description`, `entity` (default
    un.SY);
  - `[[block]]` with title and `[[block.item]]` (key, label, kind booked|attributed|estimated|missing, account,
    products[], `key_share` for attributed or estimated);
  - `checks` (label, expr over item keys);
  - `total` (label, of[]);
  - `notes`, `status`, `approved`.
- **Seed.** `seed_concepts()` converts `knowledge/concepts.json` (24 concepts). Its status is approved: the user
  approved the business map on 15.10, and the `fam_*` families are auto but booked-only. Read how
  `attributed`/`estimated` items carry their keys in concepts.json and map them exactly. If an item's key is not
  machine-readable, keep it as `missing` with the text in `notes`, and report it.
- **Build.** The concepts and their aliases go into `knowledge.db`. Concept aliases must not clash with metric
  aliases.
- **Lookup.** `fin_lookup(kind="concept")` and `kind="any"` include concepts.
- **Tests:** the seed produces 24 concepts that load; `splitt` has its blocks and keys; aliases are unique.

### Task 4: The concept block

- **Spec.** `ConceptBlock {type "concept", concept: id, query: id}`. The query supplies the cube (FinMax or
  FinLite), the entity, the periods and the compare. The block adds the concept's accounts and products to that
  query's requests automatically: `requests()` unions the concept items' accounts and products into the grid.
- **Builder.** `model/blocks/concept.py` produces one section per concept block and one line per item:
  - booked = `Pick(account, products)`;
  - attributed / estimated = the booked base × the key share. Write the key share in a labelled helper column, not
    as a constant inside the formula: tn-formelstil BLOCKs constants inside formulas;
  - missing = an empty value with its note;
  - block subtotals, the concept total, and the concept's checks as check lines. A check failure is a finding, not
    a stop, because a concept is an estimate.
  - The columns follow the query's compare (value, ly, var, var_pct).
- **Caveats.** The concept's notes become info caveats in the reply.
- **Tests** use a tiny concept with booked, attributed and missing items over fake facts:
  - values and formulas;
  - the key share in the helper column;
  - the total;
  - a check line;
  - the requests contain the concept's accounts and products.

### Task 5: Skill, docs, test-set case

- **Skill.** Mention the concept block and `rows: products` in one or two lines. `SKILL.md` is close to its 7000
  character limit, so move detail into `reference/concepts.md` and link it.
- **Test set.** Add case `b09` ("What do we earn on mobile subscriptions?", the concept nearest to v1's B09). Check
  `concepts.json` names for the mobile family and use that concept. Record it live in Task 6.
- **Docs.** CHANGELOG and DESIGN.

### Task 6: Live and finish (controller)

1. `fin_status(restart=True)`.
2. `fin_lookup("splitt", kind="concept")`.
3. `fin_analyze` with a concept block for `splitt`, FinMax un.SY closed-ytd with ly. Expect Sheet = model OK, the
   checks reported, and the caveats from the notes.
4. A product ranking: top 10 products by revenue in Wholesale YTD.
5. `testset record b09`, then `approve b09`, then `run` (9/9).
6. Merge and push.
