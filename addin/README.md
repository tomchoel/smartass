# addin

`tn-claude/`: the local Excel task pane "Claude" (Vite + React 19 + Tailwind 4 + shadcn) and its server. Just
for this PC: nothing is deployed anywhere.

## What runs

- `tn-claude/server.mjs` (Node) serves `dist/` on `https://localhost:3000` (127.0.0.1 and ::1) with the
  `office-addin-dev-certs` certificate in `%USERPROFILE%\.office-addin-dev-certs` (never in git).
- `POST /api/chat` streams a headless `claude -p --output-format stream-json` run in the Produktregnskapet folder,
  limited to `mcp__smartass` (v2: fin_*, xl_snapshot/xl_revert/xl_tidy), `mcp__excel` (Excel MCP) and read-only
  tools. The pane prompt (`PANE_PROMPT`) forces `target="active"` and snapshot-before-edit.
- `GET /api/changes`, `POST /api/revert`, `POST /api/check` call the engine CLI:
  `engine\.venv\Scripts\python.exe -m smartass.ops.cli xl changes|revert|check` (one JSON line,
  `{ok, result}` / `{ok: false, error}`). Override with `SMARTASS_PY` / `SMARTASS_ENGINE`.
- `GET /api/health` -> `{ok, claude, runs}`. Log: `%LOCALAPPDATA%\svkit\logs\addin.log`.

## Build

```
cd addin\tn-claude
npm ci
npm run build          # tsc -b && vite build -> dist\
```

`npm run dev` runs Vite on :3000 with `/api` proxied to :3001 (start `server.mjs` with `PORT=3001`).

## Run (hidden)

`wscript addin\tn-claude\start-claude-pane-hidden.vbs` starts `node server.mjs` in its own folder without a
window (`start-claude-pane.cmd` does the same with a window, for debugging). To restart: stop only the `node`
process whose command line is `node server.mjs` and listens on :3000 (`netstat -ano | findstr :3000`), then run
the vbs again. A new `dist\` needs no restart: the pane reloads it when it is reopened.

## Registration (already done - do not redo)

- Classic host: Excel-DNA XLL in `%LOCALAPPDATA%\svkit\addin\dna` (`TNClaude64.xll`, ribbon group "Claude" on
  Home, docked task pane with WebView2 -> `https://localhost:3000/?host=excel`).
- Web add-in: `tn-claude/manifest.xml` via the shared-folder catalog on H: (`TNClaudeCatalog`, trusted catalog in
  `HKCU\...\WEF\TrustedCatalogs`), label "Claude web".
- Both point at `https://localhost:3000`, so switching the server folder needs no re-registration.
