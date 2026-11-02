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
| S1 | `lookup` | `sv_find_member` |
| S5 | `drill` | `sv_drill` cold, warm, quarter |
| S6 | `xl` | the `xl_*` tools (on a tiny `sv_answer` workbook) |
| S8 | `inject` | SSO expiry, dialogs, edit mode, parallel calls, killed Excel/Edge, cold start, sleep (each checked with a tiny `sv_answer`) |
| S9 | `fresh` | install and first lookup as a new user (`--fresh`) |
| S10 | `feature`, some `fresh` | bundle, setup, relay reload, update, text fixes, pack, style, cold Excel, options, HTTP |

The bench (below) covers answers; the matrix covers reliability.

`test_styling_live.py` is separate: it uses its own Excel (`SV_STYLING_LIVE=1` under pytest).

## Benchmarks (pipeline rework)

`bench.py` runs the acceptance set B01–B15 (`docs/specs/2026-11-02-benchmarks.md`) on the pipeline:

```
uv run python tests\live\bench.py [--only B01,B02] [--keep] [--out NAME]
```

- `--out` defaults to `bench-core`.
- A case passes when every oracle value appears on the answer sheet, every check is OK, and there are no formula
  errors.
  - The oracle is an independent single-cell HTTP fetch per value, in the bench process, after the MCP server has
    exited (so it neither warms the route nor shares its browser).
  - Matching allows 1 NOK, or the value rounded to kNOK/MNOK; sign is ignored.
- Each case records time to first visible output (the first new sheet in a visible Excel, polled through its
  windows) and total time.
- B10, B11 and the traps T01–T10 are scored by `sv_gate` (suggested outcome = expected).
- Reports go to `docs\audit\<date>-<out>.json/.md`.

## Safety

- It drives the real Excel and needs SSO. **Warn the owner of the PC first.**
- Injection scenarios (S8) kill or suspend only engine-owned Excel and Edge processes, never the owner's.
- `--fresh` uses a gitignored `.matrix-fresh` profile in the repo, deleted after the run.
- Scratch workbooks go in `<analyser>\_matrix`.
- Reports go to `docs\audit`.
