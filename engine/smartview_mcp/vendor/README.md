# Vendored style helpers

Pinned, unmodified copies of two skill modules:

| File | Source skill |
|---|---|
| `tn_rapportstil.py` | `tn-rapportstil` (`~\.claude\skills\tn-rapportstil\tn_rapportstil.py`) |
| `formelsjekk.py` | `tn-formelstil` (`~\.claude\skills\tn-formelstil\formelsjekk.py`) |

The engine always uses these vendored copies (`styling.py` and `formulas.py` import them from
`smartview_mcp.vendor`). It never looks in the user profile, so behaviour is the same on every machine and on CI.
Both modules need only openpyxl and the standard library.

## Refresh

1. Copy the files byte-for-byte from the skill folders, for example from `engine`:

   ```powershell
   Copy-Item "$env:USERPROFILE\.claude\skills\tn-rapportstil\tn_rapportstil.py" smartview_mcp\vendor\
   Copy-Item "$env:USERPROFILE\.claude\skills\tn-formelstil\formelsjekk.py" smartview_mcp\vendor\
   ```

2. Run the tests: `uv run pytest -q` and `uv run python ..\scripts\check_staged.py --all`.
3. Commit the new copies together with any engine changes they need.

Do not edit the copies here; change the skill and copy again. Ruff skips this folder (`pyproject.toml`).
