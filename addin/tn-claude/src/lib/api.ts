import type { ExcelContext } from "./excel"

export type ChatEvent =
  | { type: "run"; runId: string }
  | { type: "session"; sessionId: string }
  | { type: "turn" }
  | { type: "text"; text: string }
  | { type: "tool_start"; id: string; name: string; summary: string }
  | { type: "tool_end"; id: string; name?: string; isError: boolean; preview: string; changeId?: string }
  | { type: "done"; sessionId: string; isError: boolean; durationMs: number; costUsd: number; turns: number; result: string }
  | { type: "error"; message: string }
  | { type: "closed"; code: number }

/** POST /api/chat and yield the Server-Sent Events as they arrive. */
export async function* chat(prompt: string, context: ExcelContext, sessionId?: string, signal?: AbortSignal) {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, context, sessionId }),
    signal,
  })
  if (!res.ok || !res.body) throw new Error(`server answered ${res.status}`)
  const reader = res.body.getReader()
  const dec = new TextDecoder()
  let buf = ""
  for (;;) {
    const { value, done } = await reader.read()
    if (done) break
    buf += dec.decode(value, { stream: true })
    let i: number
    while ((i = buf.indexOf("\n\n")) >= 0) {
      const chunk = buf.slice(0, i)
      buf = buf.slice(i + 2)
      const line = chunk.split("\n").find((l) => l.startsWith("data: "))
      if (line) yield JSON.parse(line.slice(6)) as ChatEvent
    }
  }
}

export async function stop(runId: string) {
  await fetch("/api/stop", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ runId }) })
}

export type Change = { id: string; ts: string; workbook: string; sheet: string; op: string; address: string; label?: string; reverted: string | false }

async function call<T>(url: string, body?: unknown): Promise<T> {
  const res = await fetch(url, body === undefined ? undefined : {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  })
  const j = await res.json()
  if (!j.ok) throw new Error(j.error || "failed")
  return j.result as T
}

export const changes = (workbook?: string) =>
  call<Change[]>(`/api/changes${workbook ? `?workbook=${encodeURIComponent(workbook)}` : ""}`)
export const revert = (id: string, workbook?: string, force = false) => call<Record<string, unknown>>("/api/revert", { id, workbook, force })

export type Finding = { cell: string; rule: string; message: string; formula: string; suggestion?: string }
export type CheckResult = { ok: boolean; checked: number; counts: Record<string, number>; block: Finding[]; warn: Finding[]; block_cells: string[] }
export const check = (workbook?: string, sheet?: string) => call<CheckResult>("/api/check", { workbook, sheet })

export async function health(): Promise<boolean> {
  try { return (await (await fetch("/api/health")).json()).ok === true } catch { return false }
}
