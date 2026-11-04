// node --test: the pane server allows one /api/chat run at a time, and PANE_PROMPT carries the routing rules.
// The real server.mjs runs on a spare port; "claude" is fake-claude.mjs (CLAUDE_JS: a slow run, no AI).
import { test } from "node:test"
import assert from "node:assert/strict"
import { spawn } from "node:child_process"
import fs from "node:fs"
import https from "node:https"
import http from "node:http"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const HERE = path.dirname(fileURLToPath(import.meta.url))
const SERVER = path.join(HERE, "..", "server.mjs")
const PORT = 3917

function request(method, p, body) {
  const lib = fs.existsSync(path.join(os.homedir(), ".office-addin-dev-certs", "localhost.crt")) ? https : http
  return new Promise((resolve, reject) => {
    const req = lib.request({ host: "127.0.0.1", port: PORT, path: p, method, rejectUnauthorized: false,
      headers: { "Content-Type": "application/json" } }, (res) => {
      let data = ""
      res.on("data", (c) => (data += c))
      res.on("end", () => resolve({ status: res.statusCode, body: data }))
    })
    req.on("error", reject)
    if (body) req.write(JSON.stringify(body))
    req.end()
  })
}

async function started() {
  for (let i = 0; i < 100; i++) {
    try { if ((await request("GET", "/api/health")).status === 200) return } catch { /* not listening yet */ }
    await new Promise((r) => setTimeout(r, 100))
  }
  throw new Error("pane server did not start")
}

test("a second /api/chat while one runs is refused with 409", async (t) => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "pane-test-"))
  const srv = spawn(process.execPath, [SERVER], {
    env: { ...process.env, PORT: String(PORT), CLAUDE_JS: path.join(HERE, "fake-claude.mjs"), CLAUDE_CWD: tmp,
           LOCALAPPDATA: tmp, FAKE_CLAUDE_MS: "2500" },
    stdio: "ignore", windowsHide: true })
  t.after(() => srv.kill())
  await started()
  const first = request("POST", "/api/chat", { prompt: "one", context: {} })
  await new Promise((r) => setTimeout(r, 500))                 // the first run is going
  const second = await request("POST", "/api/chat", { prompt: "two", context: {} })
  assert.equal(second.status, 409)
  assert.match(JSON.parse(second.body).error, /arbeider allerede/)
  const one = await first
  assert.equal(one.status, 200)
  assert.match(one.body, /"type":"closed"/)
  const third = request("POST", "/api/chat", { prompt: "three", context: {} })   // free again after the run
  assert.equal((await third).status, 200)
})

test("PANE_PROMPT requires target=active on every fin_* call and forbids hand-rebuilt analyses", () => {
  const src = fs.readFileSync(SERVER, "utf8")
  const prompt = src.slice(src.indexOf("const PANE_PROMPT"), src.indexOf("// ---", src.indexOf("const PANE_PROMPT")))
  for (const call of ["fin_analyze", "fin_critic", "fin_refresh", "fin_drill"]) assert.ok(prompt.includes(call), call)
  assert.match(prompt, /target="active" to EVERY/)
  assert.match(prompt, /never rebuild an analysis by hand with Excel MCP/)
  assert.match(prompt, /metric NAMES from fin_lookup/)
})
