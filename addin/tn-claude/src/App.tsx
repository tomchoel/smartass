import { useCallback, useEffect, useRef, useState } from "react"
import ReactMarkdown from "react-markdown"
import remarkGfm from "remark-gfm"
import {
  ArrowUp, Check, CircleAlert, FileSpreadsheet, History, Loader2, MousePointerClick, Plus, RefreshCw,
  ShieldCheck, Sparkles, Square, Undo2, Wrench,
} from "lucide-react"
import { toast } from "sonner"

import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible"
import { ScrollArea } from "@/components/ui/scroll-area"
import { Separator } from "@/components/ui/separator"
import { Skeleton } from "@/components/ui/skeleton"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Textarea } from "@/components/ui/textarea"
import { Toaster } from "@/components/ui/sonner"
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip"
import * as api from "@/lib/api"
import { EMPTY, type ExcelContext, onSelectionChanged, readContext, selectRange } from "@/lib/excel"
import { cn } from "@/lib/utils"

// ------------------------------------------------------------------ types
type ToolPart = { kind: "tool"; id: string; name: string; summary: string; status: "running" | "ok" | "error"; preview?: string; changeId?: string }
type TextPart = { kind: "text"; text: string }
type Msg = { id: string; role: "user" | "assistant"; parts: (TextPart | ToolPart)[]; meta?: string }

const TOOL_LABEL: Record<string, string> = {
  xl_workbooks: "Ser på åpne arbeidsbøker", xl_overview: "Kartlegger arbeidsboka", xl_read: "Leser celler",
  xl_write: "Skriver", xl_rows: "Rader/kolonner", xl_add_sheet: "Nytt ark", xl_format: "Formaterer",
  xl_format_sheet: "TN-formaterer arket", xl_trace: "Sporer formel", xl_check: "Sjekker formler",
  xl_changes: "Endringslogg", xl_revert: "Angrer", sv_query: "Henter fra EPM", sv_ws_build: "Bygger analyse",
  sv_ws_answer: "Oppdaterer svar", sv_find_member: "Finner medlem", sv_members: "Slår opp medlemmer",
  sv_scope: "Avgrenser spørsmålet", sv_concept: "Henter begrep", ToolSearch: "Laster verktøy", Read: "Leser fil",
  Grep: "Søker", Glob: "Finner filer", Skill: "Laster ferdighet",
}
const shortName = (n: string) => n.replace(/^mcp__[^_]+(?:_[^_]+)*?__/, "")
const toolLabel = (n: string) => TOOL_LABEL[shortName(n)] ?? shortName(n)

const SUGGESTIONS = [
  { icon: MousePointerClick, text: "Forklar hvordan tallet i markert celle er regnet ut" },
  { icon: ShieldCheck, text: "Sjekk formlene på dette arket og rett det som bryter TN-formelstil" },
  { icon: Sparkles, text: "Gi meg en kort oppsummering av hva dette arket viser" },
  { icon: Wrench, text: "Lag en sumrad under markeringen med SUM" },
]

const uid = () => Math.random().toString(36).slice(2, 10)

// ------------------------------------------------------------------ app
export default function App() {
  const [ctx, setCtx] = useState<ExcelContext>(EMPTY)
  const [online, setOnline] = useState<boolean | null>(null)
  const [tab, setTab] = useState("chat")

  const refreshCtx = useCallback(async () => {
    try { setCtx(await readContext()) } catch (e) { console.warn(e) }
  }, [])

  useEffect(() => {
    refreshCtx()
    api.health().then(setOnline)
    return onSelectionChanged(() => { refreshCtx() })
  }, [refreshCtx])

  return (
    <TooltipProvider delayDuration={300}>
      <div className="flex h-screen flex-col bg-background text-foreground">
        <Header ctx={ctx} online={online} onRefresh={refreshCtx} />
        <Tabs value={tab} onValueChange={setTab} className="flex min-h-0 flex-1 flex-col gap-0">
          <div className="px-3 pt-2">
            <TabsList className="w-full">
              <TabsTrigger value="chat" className="flex-1">Chat</TabsTrigger>
              <TabsTrigger value="changes" className="flex-1">Endringer</TabsTrigger>
              <TabsTrigger value="check" className="flex-1">Formelsjekk</TabsTrigger>
            </TabsList>
          </div>
          <TabsContent value="chat" className="flex min-h-0 flex-1 flex-col"><ChatPane ctx={ctx} /></TabsContent>
          <TabsContent value="changes" className="min-h-0 flex-1"><ChangesPane ctx={ctx} active={tab === "changes"} /></TabsContent>
          <TabsContent value="check" className="min-h-0 flex-1"><CheckPane ctx={ctx} /></TabsContent>
        </Tabs>
        <Toaster position="top-center" richColors />
      </div>
    </TooltipProvider>
  )
}

// ------------------------------------------------------------------ header
function Header({ ctx, online, onRefresh }: { ctx: ExcelContext; online: boolean | null; onRefresh: () => void }) {
  return (
    <div className="border-b px-3 py-2.5">
      <div className="flex items-center gap-2">
        <div className="flex size-7 items-center justify-center rounded-md bg-[#D97757] text-white">
          <Sparkles className="size-4" />
        </div>
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-1.5 text-sm font-semibold leading-none">
            Claude
            <span className={cn("size-1.5 rounded-full", online ? "bg-emerald-500" : online === false ? "bg-red-500" : "bg-muted-foreground/40")} />
          </div>
          <div className="mt-1 text-[11px] text-muted-foreground">Telenor · Produktregnskapet</div>
        </div>
        <Tooltip>
          <TooltipTrigger asChild>
            <Button variant="ghost" size="icon" className="size-7" onClick={onRefresh}><RefreshCw className="size-3.5" /></Button>
          </TooltipTrigger>
          <TooltipContent>Les markering på nytt</TooltipContent>
        </Tooltip>
      </div>
      <div className="mt-2 flex flex-wrap items-center gap-1">
        {ctx.inExcel ? (
          <>
            <Badge variant="secondary" className="max-w-full gap-1 font-normal">
              <FileSpreadsheet className="size-3" /><span className="truncate">{ctx.workbook}</span>
            </Badge>
            <Badge variant="outline" className="font-normal">{ctx.sheet}</Badge>
            <Badge variant="outline" className="font-mono text-[11px] font-normal">{ctx.selection}</Badge>
          </>
        ) : (
          <Badge variant="outline" className="font-normal text-muted-foreground">Ikke i Excel – forhåndsvisning</Badge>
        )}
      </div>
    </div>
  )
}

// ------------------------------------------------------------------ chat
function ChatPane({ ctx }: { ctx: ExcelContext }) {
  const [msgs, setMsgs] = useState<Msg[]>([])
  const [input, setInput] = useState("")
  const [busy, setBusy] = useState(false)
  const [sessionId, setSessionId] = useState<string>()
  const [startedAt, setStartedAt] = useState<number>(0)
  const [now, setNow] = useState(Date.now())
  const runId = useRef<string>("")
  const bottom = useRef<HTMLDivElement>(null)

  useEffect(() => { bottom.current?.scrollIntoView({ behavior: "smooth" }) }, [msgs])
  useEffect(() => {
    if (!busy) return
    const t = setInterval(() => setNow(Date.now()), 500)
    return () => clearInterval(t)
  }, [busy])

  const patchLast = (fn: (m: Msg) => Msg) => setMsgs((ms) => ms.map((m, i) => (i === ms.length - 1 ? fn(m) : m)))

  async function send(text: string) {
    const prompt = text.trim()
    if (!prompt || busy) return
    setInput("")
    setBusy(true)
    setStartedAt(Date.now())
    const fresh = await readContext().catch(() => ctx)
    setMsgs((ms) => [...ms, { id: uid(), role: "user", parts: [{ kind: "text", text: prompt }] },
      { id: uid(), role: "assistant", parts: [] }])
    try {
      for await (const ev of api.chat(prompt, fresh, sessionId)) {
        if (ev.type === "run") runId.current = ev.runId
        else if (ev.type === "session") setSessionId(ev.sessionId)
        else if (ev.type === "text") {
          patchLast((m) => {
            const parts = [...m.parts]
            const last = parts[parts.length - 1]
            if (last?.kind === "text") parts[parts.length - 1] = { ...last, text: last.text + ev.text }
            else parts.push({ kind: "text", text: ev.text })
            return { ...m, parts }
          })
        } else if (ev.type === "tool_start") {
          patchLast((m) => ({ ...m, parts: [...m.parts, { kind: "tool", id: ev.id, name: ev.name, summary: ev.summary, status: "running" }] }))
        } else if (ev.type === "tool_end") {
          patchLast((m) => ({ ...m, parts: m.parts.map((p) => p.kind === "tool" && p.id === ev.id
            ? { ...p, status: ev.isError ? "error" : "ok", preview: ev.preview, changeId: ev.changeId } : p) }))
        } else if (ev.type === "done") {
          setSessionId(ev.sessionId)
          patchLast((m) => ({ ...m, meta: `${(ev.durationMs / 1000).toFixed(0)} s · ${ev.turns} steg` }))
          if (ev.isError) toast.error("Claude stoppet med en feil")
        } else if (ev.type === "error") toast.error(ev.message.slice(0, 200))
      }
    } catch (e) {
      toast.error(`Fikk ikke kontakt med den lokale serveren: ${e}`)
    } finally {
      setBusy(false)
      runId.current = ""
    }
  }

  function newChat() {
    if (busy) return
    setMsgs([]); setSessionId(undefined)
  }

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <ScrollArea className="min-h-0 flex-1">
        <div className="space-y-4 px-3 py-3">
          {msgs.length === 0 && <EmptyState onPick={send} ctx={ctx} />}
          {msgs.map((m) => <Message key={m.id} m={m} busy={busy && m === msgs[msgs.length - 1]} />)}
          <div ref={bottom} />
        </div>
      </ScrollArea>
      <div className="border-t p-2">
        <div className="rounded-lg border bg-background focus-within:ring-2 focus-within:ring-ring/30">
          <Textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(input) } }}
            placeholder={ctx.inExcel ? `Spør om ${ctx.sheet}!${ctx.selection} …` : "Spør Claude …"}
            className="max-h-40 min-h-[60px] resize-none border-0 shadow-none focus-visible:ring-0"
          />
          <div className="flex items-center justify-between px-2 pb-2">
            <div className="flex items-center gap-1">
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button variant="ghost" size="icon" className="size-7" onClick={newChat} disabled={busy}><Plus className="size-4" /></Button>
                </TooltipTrigger>
                <TooltipContent>Ny samtale</TooltipContent>
              </Tooltip>
              <span className="text-[11px] text-muted-foreground">
                {busy ? `Arbeider … ${Math.max(0, Math.round((now - startedAt) / 1000))} s` : sessionId ? "Fortsetter samtalen" : "Lagrer aldri – du lagrer selv"}
              </span>
            </div>
            {busy ? (
              <Button size="icon" variant="secondary" className="size-7" onClick={() => runId.current && api.stop(runId.current)}>
                <Square className="size-3 fill-current" />
              </Button>
            ) : (
              <Button size="icon" className="size-7" onClick={() => send(input)} disabled={!input.trim()}>
                <ArrowUp className="size-4" />
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

function EmptyState({ onPick, ctx }: { onPick: (t: string) => void; ctx: ExcelContext }) {
  return (
    <div className="space-y-3 pt-2">
      <Card className="gap-2 py-4 shadow-none">
        <CardHeader className="px-4">
          <CardTitle className="text-sm">Jobb rett i arbeidsboka</CardTitle>
          <CardDescription className="text-xs">
            Claude leser og endrer {ctx.inExcel ? <b>{ctx.workbook}</b> : "den åpne arbeidsboka"} direkte. Hver endring kan angres
            under <i>Endringer</i>, og ingenting lagres automatisk.
          </CardDescription>
        </CardHeader>
      </Card>
      <div className="grid gap-1.5">
        {SUGGESTIONS.map(({ icon: Icon, text }) => (
          <Button key={text} variant="outline" className="h-auto justify-start gap-2 whitespace-normal py-2 text-left text-xs font-normal"
            onClick={() => onPick(text)}>
            <Icon className="size-3.5 shrink-0 text-muted-foreground" />{text}
          </Button>
        ))}
      </div>
    </div>
  )
}

function Message({ m, busy }: { m: Msg; busy: boolean }) {
  if (m.role === "user") {
    const t = m.parts[0] as TextPart
    return (
      <div className="flex justify-end">
        <div className="max-w-[85%] rounded-2xl rounded-br-sm bg-muted px-3 py-2 text-sm">{t.text}</div>
      </div>
    )
  }
  return (
    <div className="space-y-2 text-sm">
      {m.parts.length === 0 && busy && (
        <div className="space-y-2"><Skeleton className="h-3 w-3/4" /><Skeleton className="h-3 w-1/2" /></div>
      )}
      {m.parts.map((p, i) => p.kind === "text"
        ? <Markdown key={i} text={p.text} />
        : <ToolChip key={p.id} p={p} />)}
      {m.meta && <div className="text-[11px] text-muted-foreground">{m.meta}</div>}
    </div>
  )
}

function Markdown({ text }: { text: string }) {
  return (
    <div className="prose-pane leading-relaxed [&_code]:rounded [&_code]:bg-muted [&_code]:px-1 [&_code]:text-[12px] [&_li]:my-0.5 [&_ol]:list-decimal [&_ol]:pl-5 [&_p]:my-1.5 [&_strong]:font-semibold [&_table]:my-2 [&_table]:w-full [&_table]:text-xs [&_td]:border-b [&_td]:px-1 [&_td]:py-0.5 [&_th]:border-b [&_th]:px-1 [&_th]:text-left [&_ul]:list-disc [&_ul]:pl-5">
      <ReactMarkdown remarkPlugins={[remarkGfm]}>{text}</ReactMarkdown>
    </div>
  )
}

function ToolChip({ p }: { p: ToolPart }) {
  return (
    <Collapsible>
      <CollapsibleTrigger asChild>
        <button className="flex w-full items-center gap-2 rounded-md border bg-muted/40 px-2 py-1.5 text-left text-xs hover:bg-muted">
          {p.status === "running" ? <Loader2 className="size-3.5 animate-spin text-muted-foreground" />
            : p.status === "ok" ? <Check className="size-3.5 text-emerald-600" /> : <CircleAlert className="size-3.5 text-red-600" />}
          <span className="font-medium">{toolLabel(p.name)}</span>
          <span className="truncate text-muted-foreground">{p.summary}</span>
          {p.changeId && <Badge variant="secondary" className="ml-auto h-4 px-1 text-[10px]">kan angres</Badge>}
        </button>
      </CollapsibleTrigger>
      <CollapsibleContent>
        <pre className="mt-1 max-h-40 overflow-auto whitespace-pre-wrap rounded-md bg-muted/60 p-2 text-[10px] text-muted-foreground">
          {shortName(p.name)}{p.summary ? `  ${p.summary}` : ""}{"\n"}{p.preview ?? "…"}
        </pre>
      </CollapsibleContent>
    </Collapsible>
  )
}

// ------------------------------------------------------------------ changes
function ChangesPane({ ctx, active }: { ctx: ExcelContext; active: boolean }) {
  const [items, setItems] = useState<api.Change[] | null>(null)
  const [busyId, setBusyId] = useState<string>()

  const load = useCallback(async () => {
    setItems(null)
    try { setItems(await api.changes(ctx.workbook || undefined)) } catch (e) { toast.error(String(e)); setItems([]) }
  }, [ctx.workbook])

  useEffect(() => { if (active) load() }, [active, load])

  async function undo(c: api.Change, force = false) {
    setBusyId(c.id)
    try {
      await api.revert(c.id, c.workbook, force)
      toast.success(`Angret ${c.op} på ${c.sheet}!${c.address}`)
      load()
    } catch (e) {
      const msg = String(e).replace(/^Error: /, "")
      toast.error(msg, msg.includes("force=true")
        ? { action: { label: "Angre likevel", onClick: () => undo(c, true) } } : undefined)
    } finally { setBusyId(undefined) }
  }

  return (
    <ScrollArea className="h-full">
      <div className="space-y-2 px-3 py-3">
        <div className="flex items-center justify-between">
          <div className="text-xs text-muted-foreground">
            {ctx.workbook ? <>Endringer i <b>{ctx.workbook}</b> (7 dager)</> : "Alle endringer (7 dager)"}
          </div>
          <Button variant="ghost" size="icon" className="size-7" onClick={load}><RefreshCw className="size-3.5" /></Button>
        </div>
        {items === null && [0, 1, 2].map((i) => <Skeleton key={i} className="h-14 w-full" />)}
        {items?.length === 0 && (
          <div className="flex flex-col items-center gap-2 py-10 text-center text-xs text-muted-foreground">
            <History className="size-6" />Ingen endringer ennå.
          </div>
        )}
        {items?.map((c) => (
          <Card key={c.id} className={cn("gap-0 py-2.5 shadow-none", c.reverted && "opacity-60")}>
            <CardContent className="flex items-center gap-2 px-3">
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-1.5">
                  <Badge variant={c.reverted ? "outline" : "secondary"} className="h-4 px-1.5 text-[10px]">{c.op}</Badge>
                  <span className="truncate font-mono text-xs">{c.sheet}!{c.address}</span>
                </div>
                <div className="mt-1 text-[11px] text-muted-foreground">
                  {c.ts.slice(11, 19)}{c.reverted ? ` · angret ${String(c.reverted).slice(11, 16) || ""}` : ""}
                </div>
              </div>
              {!c.reverted && (
                <div className="flex gap-1">
                  {c.op !== "add_sheet" && (
                    <Button variant="ghost" size="icon" className="size-7" onClick={() => selectRange(c.sheet, c.address.includes(":") || /\d/.test(c.address) ? c.address : `${c.address}1`).catch(() => {})}>
                      <MousePointerClick className="size-3.5" />
                    </Button>
                  )}
                  <Button variant="outline" size="sm" className="h-7 gap-1 px-2 text-xs" disabled={busyId === c.id} onClick={() => undo(c)}>
                    {busyId === c.id ? <Loader2 className="size-3 animate-spin" /> : <Undo2 className="size-3" />}Angre
                  </Button>
                </div>
              )}
            </CardContent>
          </Card>
        ))}
      </div>
    </ScrollArea>
  )
}

// ------------------------------------------------------------------ formula check
function CheckPane({ ctx }: { ctx: ExcelContext }) {
  const [res, setRes] = useState<api.CheckResult | null>(null)
  const [busy, setBusy] = useState(false)

  async function run(all: boolean) {
    setBusy(true)
    try { setRes(await api.check(ctx.workbook || undefined, all ? undefined : ctx.sheet || undefined)) }
    catch (e) { toast.error(String(e)) } finally { setBusy(false) }
  }

  const goto = (cell: string) => {
    const [sheet, addr] = cell.split("!")
    selectRange(sheet, addr).catch(() => {})
  }

  return (
    <ScrollArea className="h-full">
      <div className="space-y-3 px-3 py-3">
        <div className="flex gap-1.5">
          <Button size="sm" className="h-8 flex-1 gap-1.5 text-xs" onClick={() => run(false)} disabled={busy}>
            {busy ? <Loader2 className="size-3.5 animate-spin" /> : <ShieldCheck className="size-3.5" />}Sjekk {ctx.sheet || "arket"}
          </Button>
          <Button size="sm" variant="outline" className="h-8 text-xs" onClick={() => run(true)} disabled={busy}>Hele boka</Button>
        </div>
        {res && (
          <>
            <Card className="gap-2 py-3 shadow-none">
              <CardContent className="space-y-2 px-3">
                <div className="flex items-center gap-2 text-sm font-medium">
                  {res.ok ? <Check className="size-4 text-emerald-600" /> : <CircleAlert className="size-4 text-red-600" />}
                  {res.ok ? "Følger TN-formelstil" : "Bryter TN-formelstil"}
                  <span className="ml-auto text-xs font-normal text-muted-foreground">{res.checked} formler</span>
                </div>
                <div className="flex flex-wrap gap-1">
                  {Object.entries(res.counts).map(([k, n]) => (
                    <Badge key={k} variant={k.includes("block") ? "destructive" : "secondary"} className="text-[10px] font-normal">
                      {k.replace(" (block)", "").replace(" (warn)", "")} · {n}
                    </Badge>
                  ))}
                </div>
              </CardContent>
            </Card>
            {[...res.block, ...res.warn].map((f, i) => (
              <button key={i} onClick={() => goto(f.cell)} className="w-full rounded-md border p-2 text-left text-xs hover:bg-muted/50">
                <div className="flex items-center gap-1.5">
                  <span className="font-mono">{f.cell}</span>
                  <Badge variant="outline" className="h-4 px-1 text-[10px]">{f.rule}</Badge>
                </div>
                <div className="mt-1 text-muted-foreground">{f.message}</div>
                {f.suggestion && (
                  <>
                    <Separator className="my-1.5" />
                    <div className="font-mono text-[10px]"><span className="text-muted-foreground">Forslag: </span>{f.suggestion}</div>
                  </>
                )}
              </button>
            ))}
          </>
        )}
        {!res && !busy && (
          <div className="py-8 text-center text-xs text-muted-foreground">
            Sjekker de levende formlene mot TN-formelstil: summer, oppslag, faste områder og avvik-prosent.
          </div>
        )}
      </div>
    </ScrollArea>
  )
}
