# Parallel spike (2026-11-03)

Question: can the drills run in parallel, and where is the limit? All live, read-only, FinLite, 20 cells
(un.SY.2TEC/2BUS/2SSM/2EI/2WHS x C60090/C60080 x Sep 26/Aug 26). Times only, no figures. Scripts: session scratchpad
`spike_par.py`, `spike_par2.py`, `spike_par3.py`.

## Results

| Request (one signed-in session, fired from one tab with `Promise.all`) | One after another | All at once |
|---|---|---|
| Smart View `GetDrillThroughReports` (drill link) | 8 in 0.8 s | 8 in 0.3 s, 16 in 0.5 s; all answered with a URL |
| Smart View `ExecuteGrid` retrieve | 8 in 0.5 s | 8 in 0.24 s, 16 in 0.38 s |
| Drill query `FdmeeMain.jsp ... FORMAT=SUI` | 6 in 8.2 s (~1.4 s each) | 20 in 14.7 s: the answers arrive ~0.7 s apart |
| `getData` (lines) from the warm tab | 4 in 1.6 s | 4 in 1.2 s, 10 in 1.2 s; all JSON, no sign-out |

No "concurrent requests on the same session" error (a known Smart View message, see Oracle Doc 2014943.1) on either
Smart View request. 20 drill links one after another, with an empty queue: 11-22 s (grid + reports each).

**A second EPM session does not help the drill query.** Context 2 = the signed-in state without the pod's cookies
(own JSESSIONID, signed in again through SSO in 3.3 s). 10 SUI from session 1 alone: 7.2 s; 10 from session 2 alone:
7.8 s; both at once: 18.0 s. EPM handles drill queries one at a time per user (or for the whole pod), ~0.7 s each.

## What this means
- The worker-thread queue (the bench's ~22 s wait per drill link) is our own and can go: links, grids and getData
  scale with concurrency.
- The floor is EPM's drill query, ~0.7 s per non-empty piece, one at a time, whatever we do on our side. The
  consultant bench has 26 non-empty pieces: about 18 s of drill queries.
- Below that floor only fewer drill queries help: one per split (siblings differ in one member) with the other
  siblings' getData bodies built from it. That builds request content ourselves (a change to the "only send what EPM
  produced" rule); the tie of every piece to its cube value would catch a wrong body.
