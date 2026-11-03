# Drill speed spike (2026-11-02)

Live, read-only spikes on FinLite (plan1). No figures are recorded here; every drilled sum tied to the cube.
Scripts: session scratchpad (`spike_dt.py`, `spike_com.py`, `spike_getdata.py`, `spike_direct.py`, `spike_hybrid.py`).
Research (no documented alternative): `2026-11-02-drill-api-research.md`.

## 1. Drill link without Excel: works
- Smart View in Excel asks the provider `req_GetDrillThroughReports` (captured with Smart View's own diagnostics,
  `%APPDATA%\Oracle\SmartView\Diagnostics\<run>\xml\NN_reqGetDrillThroughReports.xml`).
- Body: `<sID/><usr/><pwd/><rowGov/><timeGov/><cube>FinLite</cube><ords>{cell index in the retrieved grid}</ords>
  <alsTbl>Default</alsTbl><pov/><cells>` + one `<dim name=".." row|col="0"|pov="pov" enc="0">{label}</dim>` per
  dimension `</cells>`. Labels are the Default-alias labels the ExecuteGrid answer returns.
- Answer: `<drillThroughURL name="Drill Through to source" type="post" ...>./aif/...FdmeeMain.jsp?...` - the same
  URL `ops.drill_link` gets through Excel. ~0.05 s over HTTP (Excel: ~0.8 s per cell, sequential, plus a cold Excel).
- The provider returns a URL for any cell (also `ac.EBITDA`, also empty cells); "no data / not Data Integration"
  is reported by the drill query (section 2), not here.
- ExecuteGrid cell status: 2097153 (0x200001) on cells with data, 8193 (0x2001) on empty ones (not needed).

## 2. Drill lines without loading a page per cell: works (with one warm page)
What the drill page does (recorded): loads ~84 requests of UI; at ~5 s it GETs the drill URL's JSON form
(`FdmeeMain.jsp?module=aif.drilldown&attribute=...&FORMAT=SUI`), which returns EPM's own drill query:
`QUERY`, `PARAMSARRAY`, `PARAMSTYPEARRAY` (+ `PARTITIONKEY`, `APPLICATION_ID`, `DEFAULT_HEADER`, `POV_CONTEXT`).
It then POSTs `getData?limit=500&offset=N` with body `{"QUERY": QUERY, "PARAMS": PARAMSARRAY,
"PARAMSTYPE": PARAMSTYPEARRAY}` - field for field EPM's answer, nothing built by the client.

| Try | Result |
|---|---|
| SUI GET from a plain signed-in tab | Works, 1.1-1.5 s. Empty / non-Data-Integration cell: at once `status 1`, "Drill query does not return any data. Data was not loaded to this intersection using Data Integration." Too-wide cell: at once `EPMFDM-140373`. |
| getData POST from a plain tab (with or without `page_id` / `X-AppBuilder-Client-Id` headers) | **EPM signs the session out** (epmexit.jsp; the Smart View session dies too). Never do this. |
| One real drill page loaded in a tab (4.2 s), then SUI + getData for OTHER cells from that tab, reusing the page's client id | **Works**: 1.4-2.3 s per cell (one 8 s outlier), lines = cube on every cell. No new cookies: the state is the page's client registration. |

Today a drill page takes ~5 s alone and ~26 s each with 13 in parallel; an empty cell waits for the 60 s capture
timeout; a too-wide cell costs a ~20 s round trip before the split.

## Open
- Parallel SUI/getData from one warm tab (Promise.all) - not measured yet.
- How long a warm tab stays usable (session lifetime, idle) and what a sign-out mid-drill looks like.
- Whether the warm page must be a drill of the same application/cube (only FinLite tested).
