# Handoff 2026-11-02 — advisor plans done; next: drill speed

Branch `feat/advisor-plans` in `C:\dev\smartass` (checked out there, so the user's MCP runs it). Not merged, no PR
yet; the owner is testing it first. Plan: `docs/plans/2026-11-02-advisor-plans.md`; live bench:
`docs/audit/2026-11-02-advisor-plans*.md`.

## State
- Batches A, B, C of the advisor-plans plan are done, plus the performance work (parallel browser tabs, one getData
  per piece) and the fixes up to `db33ff9`.
- After `db33ff9`:
  - `7bfccc5` + chain of 3+ lines written as SUM (tn-formelstil PLUS_CHAIN block on lines_ranking)
  - `751c323` lines_ranking: drivers only for Total (cube) + the 3 largest supplier changes, rest one Info line
  - `4e0f9f8` lines_ranking: short English title, entity in the subtitle
  - `d6bd7fc` first frame drawn right after the workbook opens (before closed month / drill plan); stage timers
    for the frame and the drill (`core/drilltime.py`)
  - drill cache REMOVED (owner: no caching in this workflow); every drill is live
- Tests: `cd engine; uv run pytest -q -p no:cacheprovider` (904); `uv run ruff check smartview_mcp tests`.

## Owner decisions this session
- **No caching of drills at all** ("does not make any sense to cache any of this kind of workflow"). The old
  closed-period cache (ba63829) was removed; last year is drilled live again (~55–60 s).
- Owner expected nothing to go through Excel. True for cube data (HTTP, ~0.3 s); **not true for the drill**:
  `ops.drill_link` (engine/smartview_mcp/ops.py ~402) puts each cell on a hidden grid sheet in Excel and calls
  Smart View `B_DrillReports` to get the drill-through URL. That is the one place Excel is still primary.

## Where the time goes (live run, consultant question, lines_ranking, un.SY, C60080+C60090, YTD Oct 26) ~100 s
| Stage | Time | Why |
|---|---:|---|
| Before first frame | 24 s | cold hidden Excel start (~8 s), new workbook (~13 s), EPM "Empty response" retries |
| Drill 2026 | 65 s | un.SY too wide (EPMFDM-140373) → ~22 child entities × 2 accounts ≈ 44 pieces |
| – drill links via Excel | ~30 s | ~0.8 s per piece, sequential (Smart View COM) |
| – browser pages | ~30 s | ~5 s per page in parallel tabs; opening each tab blocks ~2.2 s |
| Drill 2025 | ~55–60 s | live again (cache removed); runs in parallel with 2026 |
| Excel build | 11–15 s | styling ~4 s (page setup 2.5 s), Lines sheets 2.4–3.3 s, findings 0.7 s, HTTP 0.3 s, save |

Measured cold run after `d6bd7fc` (fresh server, no Excel open, 105 s total, 2025 still cached then):
workbook open 31.6 s (hidden Excel start 6.9, file create ~1.5, visible Excel open ~9.5, fixed 12 s Smart View login
wait), frame 0.7 s after open, drills 57.4 s elapsed (15 Excel drill links sequential = 20.2 s; browser start 3.0 +
sign-in 5.7; 13 browser pages summed 344.8 s in parallel tabs, i.e. ~26 s per page — **the browser pages are the
biggest cost, not the Excel links**), Excel write 15.3 s. Not done: open the new workbook directly in the visible
Excel from the TN template (~8 s), the 12 s login wait; `tools_core.py` is 406 lines.

## Next (owner to confirm)
1. **Drill links without Excel.** Find out whether the Smart View HTTP provider (`svhttp.py`) can list drill-through
   reports/URLs for a cell; if not, check whether the FDMEE/Data Integration drill URL is a fixed pattern that can
   be built from the cell (compare URLs `drill_link` returns for a few cells). Short live spike first, then TDD.
   Removes ~30 s, the cold Excel start and the Smart View "Empty response" failures from drills.
2. Browser pages: ~26 s each on average in parallel tabs (find why — EPM page load, tab contention, the ~2.2 s
   tab-open block) — this is now the biggest drill cost.
3. Then: PR for `feat/advisor-plans` once the owner has tested.
4. Cleanup later: `git worktree remove --force C:\dev\smartass.worktrees\walking-skeleton-answer-sheet-harness`,
   `git branch -D walking-skeleton-answer-sheet-harness`, delete remote `rework/pipeline`.

## Gotchas
- Windows PowerShell 5.1: no `&&`; never edit with Get-Content|Set-Content (mojibake). `git push` exits 1 (stderr).
- The chat cancels a running MCP tool call when the user sends a message.
- Fresh hidden Excel sometimes answers "Empty response from provider"; retry works.
- Live helper used for measurements: `files/live_b4.py` in the old Copilot session folder
  `C:\Users\t992115\.copilot\session-state\09673b0c-ab80-492c-a857-e1daf56b6ab1\` (gate → sv_answer for the
  consultant question; `$env:PYTHONIOENCODING="utf-8"`).
- Logs: `%LOCALAPPDATA%\smartass\logs\server.log` (stages), `runs.jsonl` (per-answer timings of the Excel job only).
- Commit trailer: `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`.
