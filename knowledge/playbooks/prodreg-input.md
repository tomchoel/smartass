# Playbook: Produktregnskapet input (interview support)

Use when preparing a division's input to the product cost accounting (Produktregnskapet): the support workbook
for the interview where the division's costs are split over activities and products. The full workflow, file
layout and scripts are in the produktregnskapet skill, `reference/intervjuforberedelse.md`; this playbook is the
smartass part and the judgement rules.

## Method

1. **Nodes.** Take the division's entity nodes from the interview file (only nodes with someone who can answer
   for FTE). The entity is `un.SY.<node>` in FinMax and FinLite.
2. **Drill per division node.** One `fin_drill` per node on the year:
   `{"All Accounts": "ac.C60", "All Entities": "un.SY.<node>", "All Time Periods": "Yr 26"}`, `group_by`
   `supplier` (or `account` for the split per account group). One drill returns every opex line of the year.
   The reconciliation must say `exact`; a gap stops the work until it is explained.
3. **Sum to YTD.** Keep the closed months only (Jan to the latest closed month) and sum the lines to YTD. The
   YTD total must equal FinMax `ac.C60` YTD for the node (`fin_analyze`, `periods: "closed-ytd"`): report the
   check, it must be zero.
4. **Activities from the lines.** Draft the special-cost activities from what the lines are (suppliers, journal
   texts), not from last year's labels or the account-group name. Flag one-off items separately.
5. **Last year.** The same drill on `Yr 25` gives the comparison and last year's keys; a split that moved more
   than the cost did needs a reason in the interview.

## Interview support workbook rule

- One support workbook per division in the division's interview folder (`Forberedelse`), never loose files, and
  never inside the Produktregnskapet archive (read only).
- Shares per account group sum to exactly 100 % across activities; an account never carries both FTE and
  special-cost rows; the FTE on every sheet sums exactly to the average FTE.
- Numbers are a snapshot with their as-of month (the reply's `closed`), not live links into the interview file.
  The user transfers the numbers to the interview file.
- Salaries (C50) sit outside C60: reconcile them separately when the interview covers FTE.

## Report

Per node: the YTD total that ties to FinMax, the top suppliers or journals in full NOK (negatives in
parentheses), the drafted activities with their amounts, and what to clarify in the interview. Caveats on the
node (reorgs, moved units) are named, because they change which costs the interviewee can answer for.
