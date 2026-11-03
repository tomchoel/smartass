# Benchmarks — route `core` (2026-11-02 21:03)

| ID | Result | First visible s | Total s | Oracle | Notes |
|---|---|---|---|---|---|
| B02 | PASS | 1.03 | 15.88 | 3/3 | bench pins recipe free-form (gate named monthly_trend); routes {'Q1 FinMax': 'http', 'Q1 FinMax ytd': 'http'} |
| B03 | PASS | 4.89 | 9.7 | 4/4 | reused ['m', 'ytd'] |
| B04 | PASS | 1.38 | 4.61 | 3/3 | reused ['m', 'ytd', 'fp'] |
| B12 | ERROR | — | — | — | ToolFailed: sv_drill: Error executing tool sv_drill: Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to every connect, and Smart View shows no Login to click. Ask the user to sign in once in that Excel (Smart View ribbon -> Panel -> connect FinMax, or close that Excel so the next call starts a fresh one), then ask again. If every Excel does this and  |
| B16 | FAIL | 3.23 | 71.81 | 1/1 | first frame 2.94 s after the workbook opened > 2.5 s (seen in Excel at 3.23 s) |
