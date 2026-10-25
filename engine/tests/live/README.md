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

The doctor lines checked by S9 include `warehouse` (OK on a fresh install: "not filled yet").

`test_styling_live.py` is separate: it uses its own Excel (`SV_STYLING_LIVE=1` under pytest).

## Safety

- It drives the real Excel and needs SSO. **Warn the owner of the PC first.**
- Injection scenarios (S8) kill or suspend only engine-owned Excel and Edge processes, never the owner's.
- `--fresh` uses a gitignored `.matrix-fresh` profile in the repo, deleted after the run.
- Scratch workbooks go in `<analyser>\_matrix`.
- Reports go to `docs\audit`.
