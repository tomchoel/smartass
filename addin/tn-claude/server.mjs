// TN Claude task pane - local server (just for this PC).
//   https://localhost:3000            the built pane (dist/), over the office-addin-dev-certs localhost certificate
//   POST /api/chat     {prompt, context, sessionId?}  -> Server-Sent Events from a headless Claude Code run
//                                                        (one run at a time: 409 while one is going)
//   POST /api/stop     {runId}                        -> stops that run
//   GET  /api/changes?workbook=                       -> the smartass snapshot/change log (revertable)
//   POST /api/revert   {id?, workbook?, force?}       -> undo a change from its snapshot
//   POST /api/check    {workbook?, sheet?}            -> tn-formelstil check of the live formulas
// Claude runs as `claude -p` in the Produktregnskapet folder (CLAUDE.md, memory, skills, MCP servers as in the
// terminal) and may ONLY use smartass v2 (fin_* EPM tools, xl_snapshot/xl_revert/xl_tidy), the Excel MCP server
// (mcp__excel__*: generic edits to the open workbook) and read-only file tools.
// /api/changes|revert|check run `python -m smartass.ops.cli xl ...` from the smartass engine venv (one JSON line).
import http from "node:http"
import https from "node:https"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { spawn } from "node:child_process"
import { randomUUID } from "node:crypto"
import { fileURLToPath } from "node:url"

const HERE = path.dirname(fileURLToPath(import.meta.url))
const PORT = Number(process.env.PORT || 3000)
const CLAUDE = process.env.CLAUDE_EXE || path.join(os.homedir(), ".local", "bin", "claude.exe")
const CLAUDE_JS = process.env.CLAUDE_JS     // tests only: a node script that stands in for claude.exe
const WORKDIR = process.env.CLAUDE_CWD ||
  path.join(os.homedir(), "OneDrive - Telenor", "Analyse (Telenor Norge AS) - Produktregnskapet")
const SVKIT = path.join(process.env.LOCALAPPDATA, "svkit")
const ENGINE = process.env.SMARTASS_ENGINE || "C:\\dev\\smartass\\engine"
const PY = process.env.SMARTASS_PY || path.join(ENGINE, ".venv", "Scripts", "python.exe")
const CERTS = path.join(os.homedir(), ".office-addin-dev-certs")
const DIST = path.join(HERE, "dist")
const LOG = path.join(SVKIT, "logs", "addin.log")
fs.mkdirSync(path.dirname(LOG), { recursive: true })
const log = (...a) => fs.appendFileSync(LOG, `${new Date().toISOString()} ${a.join(" ")}\n`)

const ALLOWED_TOOLS = ["mcp__smartass", "mcp__excel", "Read", "Grep", "Glob", "Skill", "ToolSearch"]

const PANE_PROMPT = `You are Claude working INSIDE the user's Excel, in a narrow task pane (about 350 px wide).
- The user's live context (workbook, sheet, selection and its values) comes with every message. "This", "here",
  "the selection", "this cell" mean that context. Load tool schemas with ToolSearch first if needed.
- Finance numbers (FinMax/FinLite/FinPlan): smartass v2 (skill smartass). Pass target="active" to EVERY
  fin_analyze (new question AND follow-up patch), fin_critic, fin_refresh and fin_drill call - never a new file
  unless the user asks. Metrics: use metric NAMES from fin_lookup (e.g. "Gross margin"), not account codes or ids.
  Quote numbers only from key_values, pass on caveats, warnings and findings; standard/deep answers: run
  fin_critic(q, target="active"), fix by patch below 4, store the score. Ledger lines: fin_drill.
- If a fin_* call fails: tell the user the error message and stop. never rebuild an analysis by hand with Excel MCP
  (no own data sheets, lookup formulas or copied numbers instead of fin_analyze).
- Excel MCP is only for edits the user explicitly asks for in the open workbook (file_read discover -> file attach
  with the context workbook, then range, range_format, worksheet ...). ALWAYS xl_snapshot(workbook, sheet, range)
  BEFORE an edit; xl_revert undoes it; finish with xl_tidy (zero BLOCK findings, formulas per tn-formelstil: SUM
  over ranges, label-keyed XLOOKUP, variance % empty on a zero base). Never save or close the user's workbook (no
  file save/close, no save:true).
- Finance format: 1 000 000, negatives (1 000 000) in tables, a minus sign in text; FinMax signs (costs negative);
  variance = current - comparison. Sheet text English.
- Answer short for a narrow pane: a few sentences or bullets, no wide tables (max 3 columns). Reply in the user's
  language.`

// ------------------------------------------------------------------ helpers
const runs = new Map()

function json(res, code, obj) {
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8" })
  res.end(JSON.stringify(obj))
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let b = ""
    req.on("data", (c) => (b += c))
    req.on("end", () => { try { resolve(b ? JSON.parse(b) : {}) } catch (e) { reject(e) } })
    req.on("error", reject)
  })
}

function xlcli(args) {
  return new Promise((resolve) => {
    const p = spawn(PY, ["-B", "-m", "smartass.ops.cli", "xl", ...args], {
      cwd: ENGINE, env: { ...process.env, PYTHONIOENCODING: "utf-8", SV_XL_JUMP: "1" },
      windowsHide: true, stdio: ["ignore", "pipe", "pipe"],
    })
    let out = "", err = ""
    p.stdout.on("data", (d) => (out += d))
    p.stderr.on("data", (d) => (err += d))
    p.on("close", () => {
      try { resolve(JSON.parse(out.trim().split("\n").pop())) } catch { resolve({ ok: false, error: (err || out).slice(-500) }) }
    })
  })
}

function contextBlock(ctx) {
  if (!ctx || !ctx.workbook) return "Context: (the pane could not read the selection this time). Work in the user's ACTIVE workbook: fin_analyze with target=\"active\", xl_* tools without workbook=, Excel MCP attach to the active workbook."
  const lines = [`Context: workbook "${ctx.workbook}", sheet "${ctx.sheet}", selection ${ctx.selection}.`]
  if (ctx.values?.length) {
    const preview = ctx.values.slice(0, 12).map((r) => r.slice(0, 8).map((v) => (v === "" ? "·" : String(v))).join(" | "))
    lines.push("Selection values (top-left):", ...preview)
  }
  if (ctx.formulas?.length) {
    const f = ctx.formulas.flat().filter((x) => typeof x === "string" && x.startsWith("=")).slice(0, 6)
    if (f.length) lines.push("Formulas in the selection: " + f.join("  ;  "))
  }
  return lines.join("\n")
}

function summarizeInput(name, input) {
  if (!input) return ""
  const keys = ["action", "sheet_name", "range_address", "range", "sheet", "workbook", "q", "query", "kind", "target", "change_id", "cell"]
  const parts = keys.filter((k) => input[k] != null)
    .map((k) => `${k}=${(typeof input[k] === "object" ? JSON.stringify(input[k]) : String(input[k])).slice(0, 40)}`)
  return parts.join(" ")
}

// ------------------------------------------------------------------ chat (SSE)
const BUSY_MESSAGE = "Claude arbeider allerede med en forespørsel. Vent til den er ferdig, eller stopp den først."

function chat(req, res, body) {
  if (runs.size > 0) {          // one run at a time: two runs in the same Excel step on each other's workbook
    log("chat refused (busy)", JSON.stringify(body.prompt).slice(0, 200))
    return json(res, 409, { ok: false, error: BUSY_MESSAGE })
  }
  const runId = randomUUID()
  res.writeHead(200, { "Content-Type": "text/event-stream", "Cache-Control": "no-cache", Connection: "keep-alive" })
  const send = (ev) => res.write(`data: ${JSON.stringify(ev)}\n\n`)
  const prompt = `${contextBlock(body.context)}\n\nUser: ${body.prompt}`
  const args = ["-p", prompt, "--output-format", "stream-json", "--verbose", "--include-partial-messages",
    "--append-system-prompt", PANE_PROMPT, "--allowedTools", ...ALLOWED_TOOLS]
  if (body.sessionId) args.push("--resume", body.sessionId)
  log("chat", runId, body.sessionId || "new", JSON.stringify(body.prompt).slice(0, 200))
  const [exe, argv] = CLAUDE_JS ? [process.execPath, [CLAUDE_JS, ...args]] : [CLAUDE, args]
  const child = spawn(exe, argv, { cwd: WORKDIR, windowsHide: true, env: { ...process.env } })
  runs.set(runId, child)
  send({ type: "run", runId })
  const tools = new Map()       // tool_use id -> name
  let buf = ""
  child.stdout.on("data", (chunk) => {
    buf += chunk
    let i
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim()
      buf = buf.slice(i + 1)
      if (!line) continue
      let ev
      try { ev = JSON.parse(line) } catch { continue }
      if (ev.type === "system" && ev.subtype === "init") send({ type: "session", sessionId: ev.session_id })
      else if (ev.type === "stream_event") {
        const e = ev.event
        if (e?.type === "content_block_delta" && e.delta?.type === "text_delta") send({ type: "text", text: e.delta.text })
        else if (e?.type === "message_start") send({ type: "turn" })
      } else if (ev.type === "assistant") {
        for (const c of ev.message?.content || []) {
          if (c.type === "tool_use") {
            tools.set(c.id, c.name)
            send({ type: "tool_start", id: c.id, name: c.name, summary: summarizeInput(c.name, c.input) })
          }
        }
      } else if (ev.type === "user") {
        for (const c of ev.message?.content || []) {
          if (c.type === "tool_result") {
            const text = Array.isArray(c.content) ? c.content.map((x) => x.text || "").join(" ") : String(c.content || "")
            const changeId = (text.match(/"change_id":\s*"([^"]+)"/) || [])[1]
            send({ type: "tool_end", id: c.tool_use_id, name: tools.get(c.tool_use_id), isError: !!c.is_error,
                   preview: text.slice(0, 300), changeId })
          }
        }
      } else if (ev.type === "result") {
        send({ type: "done", sessionId: ev.session_id, isError: ev.is_error, durationMs: ev.duration_ms,
               costUsd: ev.total_cost_usd, turns: ev.num_turns, result: ev.result })
      }
    }
  })
  let errText = ""
  child.stderr.on("data", (d) => { errText += d })
  child.on("close", (code) => {
    runs.delete(runId)
    if (code && code !== 0) send({ type: "error", message: errText.slice(-600) || `claude exited with ${code}` })
    send({ type: "closed", code })
    res.end()
    log("chat end", runId, code)
  })
  req.on("close", () => { if (runs.has(runId)) { /* pane closed: let the run finish its tool calls safely */ } })
}

// ------------------------------------------------------------------ static
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".svg": "image/svg+xml",
  ".woff2": "font/woff2", ".json": "application/json", ".ico": "image/x-icon", ".xml": "application/xml" }

function serveStatic(req, res) {
  const u = new URL(req.url, "https://localhost")
  let p = path.normalize(path.join(DIST, decodeURIComponent(u.pathname)))
  if (!p.startsWith(DIST)) return json(res, 403, { error: "forbidden" })
  if (fs.existsSync(p) && fs.statSync(p).isDirectory()) p = path.join(p, "index.html")
  if (!fs.existsSync(p)) p = path.join(DIST, "index.html")
  res.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream", "Cache-Control": "no-cache" })
  fs.createReadStream(p).pipe(res)
}

// ------------------------------------------------------------------ router
async function handler(req, res) {
  try {
    const u = new URL(req.url, "https://localhost")
    if (u.pathname === "/api/health") return json(res, 200, { ok: true, claude: fs.existsSync(CLAUDE), runs: runs.size })
    if (u.pathname === "/api/chat" && req.method === "POST") return chat(req, res, await readBody(req))
    if (u.pathname === "/api/stop" && req.method === "POST") {
      const { runId } = await readBody(req)
      const c = runs.get(runId)
      if (c) spawn("taskkill", ["/PID", String(c.pid), "/T", "/F"], { windowsHide: true })
      return json(res, 200, { ok: !!c })
    }
    if (u.pathname === "/api/changes") {
      const wb = u.searchParams.get("workbook")
      return json(res, 200, await xlcli(["changes", "--limit", "30", ...(wb ? ["--workbook", wb] : [])]))
    }
    if (u.pathname === "/api/revert" && req.method === "POST") {
      const b = await readBody(req)
      return json(res, 200, await xlcli(["revert", ...(b.id ? ["--id", b.id] : []), ...(b.workbook ? ["--workbook", b.workbook] : []),
        ...(b.force ? ["--force"] : [])]))
    }
    if (u.pathname === "/api/check" && req.method === "POST") {
      const b = await readBody(req)
      return json(res, 200, await xlcli(["check", ...(b.workbook ? ["--workbook", b.workbook] : []), ...(b.sheet ? ["--sheet", b.sheet] : [])]))
    }
    return serveStatic(req, res)
  } catch (e) {
    log("error", e.stack || e)
    json(res, 500, { ok: false, error: String(e) })
  }
}

const hasCert = fs.existsSync(path.join(CERTS, "localhost.crt"))
const tls = hasCert ? { cert: fs.readFileSync(path.join(CERTS, "localhost.crt")), key: fs.readFileSync(path.join(CERTS, "localhost.key")) } : null
// "localhost" may resolve to ::1 or 127.0.0.1 in Office's WebView2: listen on both loopbacks (never on the network)
for (const host of ["127.0.0.1", "::1"]) {
  const server = tls ? https.createServer(tls, handler) : http.createServer(handler)
  server.on("error", (e) => log("listen", host, e.message))
  server.listen(PORT, host, () => {
    const msg = `TN Claude pane on ${tls ? "https" : "http (NO CERT - Excel will refuse)"}://localhost:${PORT} (${host})`
    console.log(msg)
    log(msg)
  })
}
