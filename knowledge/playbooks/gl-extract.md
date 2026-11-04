# Playbook: GL extract (ledger lines behind a number)

Use when the user asks which suppliers, projects, invoices or journals sit behind a cube number, or needs the
ledger lines per account and cost centre for a review, an audit question or a model input.

## Method

1. **Find the cell.** `fin_lookup` gives the account (`ac.<code>`, e.g. `ac.C60` for other opex) and the entity
   (`un.SY.<code>`). A metric with several accounts (Opex = `ac.C50S` + `ac.C60`) is drilled per account: one
   cell is one account × one entity × one period.
2. **Choose the period.** One month: `Sep 26`. A year to date: the year parent `Yr 26` returns every line of the
   year in one drill; filter the months afterwards (the lines carry their period) and add them up to the YTD.
   Do not drill month by month for a YTD. The drill is Periodic only.
3. **Drill.**
   - A quick look: `fin_drill(cell, group_by, top)` with
     `{"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Yr 26"}`.
   - Inside an analysis: a `lines` block with the same cell, next to the table it explains, so a commentary
     driver can cite the group as evidence (`basis: "data"`).
   - `group_by`: `supplier` for purchases, `project` for capex-like or programme costs, `account` when the cell
     is a parent account. `top` 10–20, the rest in Other.
   - Several accounts or cost centres: one `lines` block per cell in one spec; they are drilled in parallel.
4. **Reconcile.** The lines total must equal the cube value: the gap line must be zero and the reconciliation
   must say `exact`. A gap is a finding: lines posted after the drill, a different balance type, or the wrong
   cell. Never report lines that do not tie without saying so.
5. **Read the lines.** Labels mislead: an account group called travel can hold a VAT-interest journal, and an
   energy line can be an allocation journal. Name what the lines are (supplier, journal text), not the account
   label. Lines without a supplier or project are journals: say so.

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
