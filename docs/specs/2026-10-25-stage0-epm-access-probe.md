# Stage 0: EPM access probe (2026-10-25)

**Verdict: GO, via the Smart View provider over HTTP. No admin, no OAuth client, no Excel.**

## What was tested

| Route | Result |
|---|---|
| Planning REST (`/HyperionPlanning/rest/...`, `/interop/rest/`) without credentials | Reachable from the laptop; `401 Bearer error="invalid_session"`. |
| Planning REST with the drill's SSO browser session (Playwright request context and in-page `fetch`) | `401`. The REST gate accepts only OAuth bearer tokens, not the web session. The EPM web UI itself makes no `/rest` calls and sends no bearer header (237 requests checked). |
| **Smart View provider `/HyperionPlanning/SmartView` with the SSO browser session** | **Works.** `ConnectToProvider` → `sID`; `GetProvisionedDataSources` lists FinMax (`plan1`), FinPlan (`stratmodel`), FinMaxDK and others; `OpenApplication` → `OpenCube` → `ExecuteGrid` (action `retrieve`) returns data. |

## Proof

EBITDA, `un.SY.2WHS`, Actual, NOK, Periodic, Aug 26 + Sep 26 from FinMax:
- **HTTP:** identical to the øre with **the COM route** (`ops.query` through the hidden Excel).
- **Timing:** the HTTP chain (connect + open + retrieve) took about 0.2 s after sign-in; Oracle reports 19 ms. The COM
  route took 49.8 s cold (1.3 s for the retrieve itself).
- The figures themselves are not stored in the repo.

Reproduce: `scripts\probe_sv_http.py` (read-only; refuses any request containing a write operation).

## Protocol notes (learnt by probing; undocumented by Oracle)

- Request names come from the installed Smart View 25.100 DLLs (`HsAddin.dll`, `SmartQuery.dll`): 176 `req_*`
  names, including `ExecuteGrid`, `GetSubVar`, `EnumMembers`, `QueryMembers`, `ZoomIn`, `GetDrillThroughReports`.
- `req_Refresh` is the **forms** call ("you are trying to open the form ..."); ad hoc is `req_ExecuteGrid` with
  `<action>retrieve</action><ords/>`. `freeformsubmit` is a write action: never use it.
- **Free-form grid layout:**
  - every POV member must be a cell in the grid (row 0), not a `<dims pov=...>` attribute;
  - cells are row-major from index 0 (the corner included);
  - `vals` take `enc="1"` (pipe-separated) and `types` take `enc="0"`;
  - type `7` = blank, `0` = member, `2` = data.
- The response returns member *names* (`Y26M08`, `Per`), so the atlas gives the mapping to labels.
- Each response carries a new `sID`; use the latest one.
- Every response has `<stats>` with provider/Essbase timings.

## Consequences for the design

- **New primary source `SvHttpSource`**: drill-browser sign-in plus provider XML. It covers the warehouse, planning
  and exploration. Expect seconds, not minutes, and no Excel, COM, dialogs or box guard.
- **SmartViewSource (COM) stays** for the live, refreshable source sheets ("kilder") in the delivered workbook, and as
  a fallback. Every delivered workbook is reconciled: source-sheet values = HTTP values.
- OAuth REST (`exportdataslice`) is no longer needed. It stays an option if an admin ever registers a client.

## Open risks

- The protocol is undocumented and could change with an Oracle monthly update. Mitigations: a nightly live probe in
  the live matrix, and the COM fallback.
- Session lifetime and re-login under load; large-grid limits (Essbase QUERYRESULTLIMIT); concurrency with an active
  `sv_drill` (they share the browser profile lock).
- FinPlan lives on another pod (`stratmodel`): sign-in there is not yet tested.
