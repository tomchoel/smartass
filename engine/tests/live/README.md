# Live checks (run before every release)

These need desktop Excel, Smart View, an EPM login and the network, so CI cannot run them. From `engine`:

```
uv run python tests\live\matrix.py
```

| Option | Default | Meaning |
|---|---|---|
| `--only TAGS` | all | Comma-separated tags or scenario IDs |
| `--repeat N` | 1 | Run each selected scenario N times |
| `--fresh` | off | Run only the fresh-user scenarios |
| `--keep` | off | Leave the scratch workbooks and the `_matrix` folder |
| `--out NAME` | `matrix` | Report name: `docs\audit\<date>-<NAME>.json/.md` |
| `--before NAME` | latest run of the same kind | Compare with `docs\audit\<NAME>.json` (the before column) |

## Scenarios

| IDs | Tags | Covers |
|---|---|---|
| S0 | `smoke` | `sv_status` |
| S1 | `lookup`, `input` | find member, `sv_query` on FinMax/FinLite/FinPlan, bad input |
| S2 | `build` | `sv_ws_build`, patch, plan |
| S3 | `template` | the `sv_analyse` templates |
| S4 | `concept` | `sv_scope`, `sv_concept` |
| S5 | `drill` | `sv_drill` cold, warm, quarter, report |
| S6 | `xl` | the `xl_*` tools |
| S7 | `refresh` | `sv_refresh_workbook` |
| S8 | `inject` | SSO expiry, dialogs, edit mode, parallel calls, killed Excel/Edge, cold start, sleep |
| S9 | `fresh` | install and first lookup as a new user (`--fresh`) |
| S10 | `feature`, some `fresh` | bundle, setup, relay reload, update, text fixes, pack, style, cold Excel, options, HTTP |
| S11 | `fin` | `fin_pnl` Wholesale: workbook checks incl. warehouse = sheet, facts = `sv_query`, warm call refills nothing |
| S12 | `fin` | `fin_pnl` for the unit `un.SY.5OK_W` (no FinPlan member): read live, "compared with last year only" warning, live read = sheet and every check OK, EBITDA YTD = `sv_query` |

The doctor lines checked by S9 include `warehouse` (OK on a fresh install: "not filled yet").

`test_styling_live.py` is separate: it uses its own Excel (`SV_STYLING_LIVE=1` under pytest).

## Benchmarks (pipeline rework)

`bench.py` runs the acceptance set B01–B15 (`docs/specs/2026-11-02-benchmarks.md`):

```
uv run python tests\live\bench.py [--only B01,B02] [--route baseline|core] [--keep] [--out NAME]
```

- `--route baseline` runs today's tools; `--route core` runs the new pipeline (`sv_answer`, answer spec v1). Cases
  without a core route yet report `UNSUPPORTED` ("no core route yet"). `--out` defaults to `bench-<route>`.
- A case passes when every oracle value appears on the answer sheet, every check is OK, and there are no formula
  errors.
  - The oracle is independent single-cell `sv_query` calls, run after the route.
  - Matching allows 1 NOK, or the value rounded to kNOK/MNOK; sign is ignored.
- Each case records time to first visible output (the first new sheet in a visible Excel, polled through its
  windows) and total time.
- Some cases are not scored by the script:
  - B08 is `UNSUPPORTED` on the baseline (no tool for it today).
  - B10, B11 and the traps T01–T10 are `AGENT`, because gate quality needs an agent run.
- Reports go to `docs\audit\<date>-<out>.json/.md`.

## Safety

- It drives the real Excel and needs SSO. **Warn the owner of the PC first.**
- Injection scenarios (S8) kill or suspend only engine-owned Excel and Edge processes, never the owner's.
- `--fresh` uses a gitignored `.matrix-fresh` profile in the repo, deleted after the run.
- Scratch workbooks go in `<analyser>\_matrix`.
- Reports go to `docs\audit`.
