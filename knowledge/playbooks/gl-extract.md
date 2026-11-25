# Playbook: GL extract (ledger lines behind a number)

Use when the user asks which suppliers, projects, invoices or journals sit behind a cube number, or needs the
ledger lines per account and cost centre for a review, an audit question or a model input.

## Method

1. **Find the cell.** `fin_lookup` gives the account (`ac.<code>`, e.g. `ac.C60` for other opex) and the entity
   (`un.SY.<code>`). A metric with several accounts (Opex = `ac.C50S` + `ac.C60`) is drilled per account: one
   cell is one account × one entity × one period.
2. **Choose the period.** One month: `Sep 26`. A year to date: the year parent `Yr 26` returns every line of the
   year in one drill; filter the months afterwards with `fin_lines(drill_id, where={"period_to": "Oct 26"})` or a
   filtered `lines` block (the lines carry `Period` and `period_ym`). Do not drill month by month for a YTD.
3. **Explore & Drill.**
   - Investigate first (zero Excel): `fin_drill(cell)` (without `target`) caches all 45 columns in SQLite and returns
     `drill_id` + multi-facet `profile`. Then use `fin_lines(drill_id, ...)` to filter (`where`, `exclude_reversals`),
     multi-group (e.g. `["ACCOUNT", "Entity"]`), or run read-only `SELECT ... FROM lines` SQL.
   - Inside an Excel deliverable: pass `target="session"` to `fin_drill`, or add a `lines` block (`cell`, optional
     `drill_id`, `where`, `exclude_reversals`, `group_by`) next to the table it explains (`basis: "data"`).
   - Several accounts or cost centres: multiple `lines` blocks run in parallel.
4. **Reconcile.** The unfiltered drill total must equal the cube value: the gap line must be zero and the reconciliation
   must say `exact`. A gap is a finding: lines posted after the drill, a different balance type, or the wrong
   cell. Never report filtered subtotals as full cell balances without showing the unfiltered cube reconciliation.
5. **Read the lines.** Labels mislead: an account group called travel can hold a VAT-interest journal, and an
   energy line can be an allocation journal. Inspect `GL_DESCRIPTION_LINE`, `JOURNAL_NAME`, and `JE_SOURCE` via
   `fin_lines` — name what the lines actually are, not just the account label. Lines without a supplier are journals.

## Report

- Table: the top groups in full NOK, negatives in parentheses, share %, Other, lines without a group, the lines
  total, the cube value and the gap (zero).
- Text: 2–3 sentences on what the money is, the largest items with the minus sign, and the reconciliation.
- Caveats on the account or entity (from the reply) go into the text.
- The lines stay on the `Q<n> Lines` sheet as a table; a user who needs a file filters there. Never write the
  lines into the Produktregnskapet archive.

## Judgement

- An extract that feeds a filing or a model is taken for the full closed period and stated with its as-of
  date (the reply's `closed`).
- Per-person data (names, employee ids) from journal texts is not copied into a summary; aggregate.
