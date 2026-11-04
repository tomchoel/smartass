# Vendored style helpers

Pinned copies of two skill modules (`tn_rapportstil.py` is trimmed to the parts `styling.py` uses: tokens,
cell styles, number-format syntax, text widths; the skill's roles, apply, lint and PDF export are left out):

| File | Source skill |
|---|---|
| `tn_rapportstil.py` | `tn-rapportstil` (`~\.claude\skills\tn-rapportstil\tn_rapportstil.py`) |
| `formelsjekk.py` | `tn-formelstil` (`~\.claude\skills\tn-formelstil\formelsjekk.py`) |

The engine always uses these vendored copies (`styling.py` and `formulas.py` import them from
`smartass.xl.vendor`). It never looks in the user profile, so behaviour is the same on every machine and on CI.
Both modules need only openpyxl and the standard library.

## Refresh

1. Copy the files from the skill folders (then trim `tn_rapportstil.py` again), for example from `engine`:

   ```powershell
   Copy-Item "$env:USERPROFILE\.claude\skills\tn-rapportstil\tn_rapportstil.py" smartass\xl\vendor\
   Copy-Item "$env:USERPROFILE\.claude\skills\tn-formelstil\formelsjekk.py" smartass\xl\vendor\
   ```

2. Run the tests: `uv run pytest -q` and `uv run python ..\scripts\check_staged.py --all`.
3. Commit the new copies together with any engine changes they need.

Do not edit the copies here beyond the trim; change the skill and copy again. Ruff skips this folder (`pyproject.toml`).
