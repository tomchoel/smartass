// Live Excel context through Office.js (the pane runs inside the user's Excel). Outside Excel (browser preview)
// everything degrades to "not in Excel" so the UI can still be developed and screenshotted.

export type ExcelContext = {
  inExcel: boolean
  workbook: string
  sheet: string
  selection: string
  rows: number
  cols: number
  values: unknown[][]
  formulas: unknown[][]
}

export const EMPTY: ExcelContext = {
  inExcel: false, workbook: "", sheet: "", selection: "", rows: 0, cols: 0, values: [], formulas: [],
}

let officeReady: Promise<boolean> | null = null

// ---- classic add-in host (Excel-DNA + WebView2): Telenor's tenant blocks Office web add-ins, so the pane can also
// run inside a docked Custom Task Pane. The host posts {type:"context", ...} and accepts {cmd:"context"|"select"}.
type HostBridge = {
  postMessage: (m: unknown) => void
  addEventListener: (t: "message", cb: (e: MessageEvent) => void) => void
}
// Office's own web add-ins ALSO run in WebView2 (window.chrome.webview exists there too): only the classic host
// identifies itself, by loading the pane with ?host=excel.
const host: HostBridge | undefined = new URLSearchParams(location.search).get("host") === "excel"
  ? (window as unknown as { chrome?: { webview?: HostBridge } }).chrome?.webview
  : undefined
let hostCtx: ExcelContext = EMPTY
const hostListeners = new Set<() => void>()
let hostWaiters: (() => void)[] = []
if (host) {
  host.addEventListener("message", (e) => {
    const d = e.data as Partial<ExcelContext> & { type?: string; message?: string }
    if (d?.type === "context") {
      hostCtx = { ...EMPTY, ...d, inExcel: true } as ExcelContext
      if (hostWaiters.length) { hostWaiters.forEach((w) => w()); hostWaiters = [] }   // answer to readContext()
      else hostListeners.forEach((f) => f())                                         // the user moved the selection
    } else if (d?.type === "host_error") console.warn("host:", d.message)
  })
}
export const inClassicHost = !!host

async function hostContext(): Promise<ExcelContext> {
  const fresh = new Promise<void>((resolve) => { hostWaiters.push(resolve); setTimeout(resolve, 1200) })
  host!.postMessage({ cmd: "context" })
  await fresh
  return hostCtx
}

export function whenOffice(): Promise<boolean> {
  if (host) return Promise.resolve(false)
  if (officeReady) return officeReady
  const p: Promise<boolean> = new Promise((resolve) => {
    const w = window as unknown as { Office?: typeof Office }
    if (!w.Office) return resolve(false)
    const t = setTimeout(() => resolve(false), 15000)         // plain browser: Office.js loaded but no host
    w.Office.onReady((info) => {
      clearTimeout(t)
      resolve(info.host === w.Office!.HostType.Excel)
    })
  })
  officeReady = p
  p.then((ok) => { if (!ok) officeReady = null })            // a slow start must not disable Excel for good
  return p
}

function nameFromUrl(url: string): string {
  const last = decodeURIComponent(url.split(/[\\/]/).pop() || "")
  return last.split("?")[0] || url
}

const MAX_CELLS = 400

let lastGood: ExcelContext = EMPTY

/** Live context; on a transient failure (Excel busy, dialog open) the last good one, never silently empty. */
export async function readContext(): Promise<ExcelContext> {
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const c = host ? await hostContext() : await readOffice()
      if (c.inExcel) { lastGood = c; return c }
    } catch (e) {
      console.warn("readContext", e)
    }
    await new Promise((r) => setTimeout(r, 400))
  }
  return lastGood
}

async function readOffice(): Promise<ExcelContext> {
  if (!(await whenOffice())) return EMPTY
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet()
    const rng = ctx.workbook.getSelectedRange()
    const wb = ctx.workbook
    sheet.load("name")
    rng.load(["address", "rowCount", "columnCount"])
    wb.load("name")
    await ctx.sync()
    let values: unknown[][] = []
    let formulas: unknown[][] = []
    if (rng.rowCount * rng.columnCount <= MAX_CELLS) {
      rng.load(["values", "formulas"])
      await ctx.sync()
      values = rng.values
      formulas = rng.formulas
    } else {
      const top = rng.getCell(0, 0).getResizedRange(Math.min(rng.rowCount, 15) - 1, Math.min(rng.columnCount, 10) - 1)
      top.load(["values", "formulas"])
      await ctx.sync()
      values = top.values
      formulas = top.formulas
    }
    let workbook = wb.name
    if (!workbook) workbook = nameFromUrl(Office.context.document.url || "")
    return {
      inExcel: true, workbook, sheet: sheet.name, selection: rng.address.split("!").pop() || rng.address,
      rows: rng.rowCount, cols: rng.columnCount, values, formulas,
    }
  })
}

export function onSelectionChanged(cb: () => void): () => void {
  if (host) {
    hostListeners.add(cb)
    return () => { hostListeners.delete(cb) }
  }
  let off = () => {}
  whenOffice().then((ok) => {
    if (!ok) return
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, cb)
    off = () => Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: cb })
  })
  return () => off()
}

/** Show the user a range (after Claude changed it). */
export async function selectRange(sheet: string, address: string) {
  if (host) { host.postMessage({ cmd: "select", sheet, address }); return }
  if (!(await whenOffice())) return
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet)
    ws.activate()
    ws.getRange(address).select()
    await ctx.sync()
  })
}
