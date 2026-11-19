import { useState, useEffect, useMemo, useRef } from 'react'
import * as XLSX from 'xlsx'
import { LayoutDashboard, CalendarDays, ListChecks, Bell, Database, Plus, Search, X, ChevronLeft, ChevronRight, Copy, Trash2, Pencil, Menu, Download, Upload, Check, MapPin, Clock, AlertTriangle } from 'lucide-react'

/* ---------- constants & helpers ---------- */
const TZ = 'Asia/Phnom_Penh'
const KEY = 'psm.productions.v1'
const STATUSES = ['Upcoming', 'Confirmed', 'In Progress', 'Completed', 'Post Production', 'On Hold', 'Cancelled']
const TYPES = ['Video', 'Photo', 'Graphic Design', 'Event', 'Commercial', 'Other']
const SC = { Upcoming: 'bg-blue-600', Confirmed: 'bg-teal-600', 'In Progress': 'bg-amber-600', Completed: 'bg-green-700', 'Post Production': 'bg-purple-600', 'On Hold': 'bg-slate-500', Cancelled: 'bg-neutral-600 line-through' }
const DONE = ['Completed', 'Cancelled']
const inp = 'w-full mt-1 rounded-lg bg-[#1c1e21] border border-white/10 px-3 py-2 text-sm focus:outline-none focus:border-red-600'
const btn = 'inline-flex items-center gap-2 rounded-lg px-3 py-2 text-sm font-medium transition '
const pri = btn + 'bg-red-700 hover:bg-red-600'
const sec = btn + 'bg-white/10 hover:bg-white/20'

const nowStr = () => {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }).formatToParts(new Date()).map(x => [x.type, x.value]))
  return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}`
}
const addDays = (s, n) => { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10) }
const addMonths = (s, n) => { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() + n); return d.toISOString().slice(0, 10) }
const dow = s => new Date(s + 'T00:00:00Z').getUTCDay()
const fmtD = s => s ? new Date(s + 'T00:00:00Z').toLocaleDateString('en-GB', { timeZone: 'UTC', day: '2-digit', month: 'short', year: 'numeric' }) : '—'
const fmtDT = s => s ? fmtD(s.slice(0, 10)) + ' ' + s.slice(11, 16) : '—'
const monthTitle = s => new Date(s + 'T00:00:00Z').toLocaleDateString('en-US', { timeZone: 'UTC', month: 'long', year: 'numeric' })
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7)
const isOverdue = (p, t) => !DONE.includes(p.status) && ((p.deadline && p.deadline < t) || (p.date < t && ['Upcoming', 'Confirmed'].includes(p.status)))

const buildRem = (p, old = []) => {
  const out = []
  const add = (key, label, at) => out.push(old.find(r => r.key === key && r.at === at) || { key, label, at, done: false })
  if (p.r1 && p.date) add('d1', '1 day before production', addDays(p.date, -1) + 'T09:00')
  if (p.r3 && p.date) add('d3', '3 days before production', addDays(p.date, -3) + 'T09:00')
  if (p.r7 && p.deadline) add('dl7', '7 days before deadline', addDays(p.deadline, -7) + 'T09:00')
  if (p.custom) add('custom', 'Custom reminder', p.custom)
  return out
}
const dl = (name, text, type) => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000) }
const applyF = (ps, f) => ps.filter(p => {
  const q = f.q.toLowerCase()
  return (!q || [p.name, p.client, p.location, p.pid].join(' ').toLowerCase().includes(q)) && (!f.type || p.type === f.type) && (!f.status || p.status === f.status) &&
    (!f.date || p.date === f.date) && (!f.deadline || p.deadline === f.deadline) && (!f.loc || (p.location || '').toLowerCase().includes(f.loc.toLowerCase()))
})
const load = () => { try { const d = JSON.parse(localStorage.getItem(KEY)); return Array.isArray(d) ? d : [] } catch { return [] } }

/* ---------- Excel helpers ---------- */
const XCOLS = ['pid', 'name', 'client', 'type', 'status', 'date', 'start', 'end', 'location', 'deadline', 'team', 'description', 'notes', 'r1', 'r3', 'r7', 'custom', 'createdAt', 'updatedAt']
const XHEAD = [...XCOLS, 'id', 'reminders_json', 'history_json']
const xlsxBytes = ps => {
  const rows = ps.map(p => ({ ...Object.fromEntries(XCOLS.map(c => [c, typeof p[c] === 'boolean' ? (p[c] ? 'yes' : 'no') : (p[c] ?? '')])), id: p.id, reminders_json: JSON.stringify(p.reminders || []), history_json: JSON.stringify(p.history || []) }))
  const ws = XLSX.utils.json_to_sheet(rows, { header: XHEAD })
  ws['!cols'] = XHEAD.map(h => ({ wch: ['description', 'notes', 'reminders_json', 'history_json'].includes(h) ? 30 : 16 }))
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Productions')
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
}
const nDate = v => typeof v === 'number' ? new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10) : String(v ?? '').slice(0, 10)
const nTime = v => { if (typeof v === 'number') { const m = Math.round(v * 1440) % 1440; return String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0') } return String(v ?? '').slice(0, 5) }
const rowsFromXlsx = buf => {
  const wb = XLSX.read(buf, { type: 'array' })
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
  const used = new Set()
  return rows.filter(r => r.name && r.date).map((r, i) => {
    const s = k => r[k] == null ? '' : String(r[k])
    const b = v => v === true || ['yes', 'true', '1'].includes(String(v).toLowerCase())
    const j = (v, d) => { try { const x = JSON.parse(v); return Array.isArray(x) ? x : d } catch { return d } }
    let id = s('id') || uid(); if (used.has(id)) id = uid(); used.add(id)
    const n = nowStr()
    return { id, pid: s('pid') || 'PRJ-' + String(i + 1).padStart(4, '0'), name: s('name'), client: s('client'), type: TYPES.includes(s('type')) ? s('type') : 'Other', status: STATUSES.includes(s('status')) ? s('status') : 'Upcoming',
      date: nDate(r.date), start: nTime(r.start), end: nTime(r.end), location: s('location'), deadline: r.deadline ? nDate(r.deadline) : '', team: s('team'), description: s('description'), notes: s('notes'),
      r1: b(r.r1), r3: b(r.r3), r7: b(r.r7), custom: s('custom'), reminders: j(r.reminders_json, []), history: j(r.history_json, []), createdAt: s('createdAt') || n, updatedAt: s('updatedAt') || n }
  })
}
const b64ToBuf = b => Uint8Array.from(atob(b), c => c.charCodeAt(0)).buffer

/* ---------- remote storage: local Excel file (dev) or cloud (Vercel) ---------- */
const CLOUD = !import.meta.env.DEV
const api = async (method, body, qs = '') => {
  const r = await fetch('/api/data' + qs, { method, headers: { 'x-app-password': localStorage.getItem('psm.pw') || '', 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined })
  if (r.status === 401) { localStorage.removeItem('psm.pw'); throw Object.assign(new Error('Wrong password'), { auth: true }) }
  let j = {}; try { j = await r.json() } catch { /* not json */ }
  if (!r.ok) throw new Error(j.error || 'Server error')
  return j
}
const remoteGet = async meta => {
  if (CLOUD) { const j = await api('GET', null, meta ? '?meta=1' : ''); return j.rev ? { exists: true, mtime: j.rev, rows: j.data || [] } : { exists: false } }
  const j = await (await fetch('/api/excel/' + (meta ? '?meta=1' : ''))).json()
  return j.exists ? { exists: true, mtime: j.mtime, rows: meta ? null : rowsFromXlsx(b64ToBuf(j.data)) } : { exists: false }
}
const remotePut = async list => {
  if (CLOUD) return (await api('PUT', { data: list })).rev
  const r = await fetch('/api/excel/', { method: 'POST', body: xlsxBytes(list) }); const j = await r.json()
  if (!r.ok) throw new Error(j.error); return j.mtime
}
const Login = ({ onDone }) => {
  const [pw, setPw] = useState(''), [err, setErr] = useState(''), [busy, setBusy] = useState(false)
  const go = async e => {
    e.preventDefault(); setBusy(true); setErr(''); localStorage.setItem('psm.pw', pw)
    try { await api('GET', null, '?meta=1'); onDone() } catch (x) { localStorage.removeItem('psm.pw'); setErr(x.message || 'Could not connect') }
    setBusy(false)
  }
  return (
    <div className="min-h-screen bg-[#1c1e21] text-white flex items-center justify-center p-4">
      <form onSubmit={go} className="w-full max-w-sm rounded-2xl bg-[#26292d] border border-white/10 p-6 space-y-4">
        <div className="flex items-center gap-2"><div className="w-8 h-8 rounded-lg bg-red-700 flex items-center justify-center font-bold">P</div><h1 className="font-semibold">Production Schedule Manager</h1></div>
        <Field l="Password" e={err}><input type="password" autoFocus className={inp} value={pw} onChange={e => setPw(e.target.value)} /></Field>
        <button className={pri + ' w-full justify-center'} disabled={busy || !pw}>{busy ? 'Checking…' : 'Log in'}</button>
      </form>
    </div>
  )
}

/* ---------- small components ---------- */
const Badge = ({ s }) => <span className={`${SC[s]} text-xs px-2 py-0.5 rounded-full whitespace-nowrap`}>{s}</span>
const Card = ({ title, children, right }) => (
  <section className="rounded-xl bg-[#26292d] border border-white/10 p-4">
    {title && <div className="flex items-center justify-between mb-3"><h3 className="font-semibold">{title}</h3>{right}</div>}
    {children}
  </section>
)
const Empty = ({ text }) => <p className="text-sm text-white/40 py-6 text-center">{text}</p>
const Modal = ({ onClose, children }) => (
  <div className="fixed inset-0 z-50 bg-black/70 overflow-y-auto p-3 sm:p-6" onMouseDown={e => e.target === e.currentTarget && onClose()}>
    <div className="mx-auto max-w-2xl rounded-2xl bg-[#26292d] border border-white/10 p-5 relative">
      <button onClick={onClose} className="absolute right-4 top-4 text-white/50 hover:text-white"><X size={20} /></button>
      {children}
    </div>
  </div>
)
const Field = ({ l, e, children }) => <label className="block text-sm"><span className="text-white/60">{l}</span>{children}{e && <span className="text-red-400 text-xs">{e}</span>}</label>
const Row = ({ p, onOpen }) => (
  <button onClick={() => onOpen(p.id)} className="w-full text-left flex items-center justify-between gap-3 rounded-lg px-3 py-2 hover:bg-white/5">
    <div className="min-w-0"><div className="text-sm font-medium truncate">{p.name}</div><div className="text-xs text-white/50 truncate">{p.client} · {fmtD(p.date)} {p.start}</div></div>
    <Badge s={p.status} />
  </button>
)
const Table = ({ ps, onOpen }) => !ps.length ? <Empty text="No productions match. Add one or change the filters." /> : (
  <div className="overflow-x-auto"><table className="w-full text-sm min-w-[720px]">
    <thead className="text-left text-white/50"><tr>{['ID', 'Project', 'Client', 'Type', 'Date', 'Deadline', 'Status'].map(h => <th key={h} className="py-2 px-2 font-medium">{h}</th>)}</tr></thead>
    <tbody>{[...ps].sort((a, b) => a.date.localeCompare(b.date)).map(p => (
      <tr key={p.id} onClick={() => onOpen(p.id)} className="border-t border-white/5 hover:bg-white/5 cursor-pointer">
        <td className="py-2 px-2 text-white/50">{p.pid}</td><td className="px-2 font-medium">{p.name}</td><td className="px-2">{p.client}</td><td className="px-2">{p.type}</td>
        <td className="px-2 whitespace-nowrap">{fmtD(p.date)}</td><td className="px-2 whitespace-nowrap">{fmtD(p.deadline)}</td><td className="px-2"><Badge s={p.status} /></td>
      </tr>))}</tbody></table></div>
)
const emptyF = { q: '', type: '', status: '', date: '', deadline: '', loc: '' }
const FilterBar = ({ f, setF }) => {
  const set = k => e => setF({ ...f, [k]: e.target.value })
  return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-4">
      <div className="col-span-2 relative"><Search size={16} className="absolute left-3 top-3.5 text-white/40" /><input className={inp + ' pl-9'} placeholder="Search project, client, location…" value={f.q} onChange={set('q')} /></div>
      <select className={inp} value={f.type} onChange={set('type')}><option value="">All types</option>{TYPES.map(t => <option key={t}>{t}</option>)}</select>
      <select className={inp} value={f.status} onChange={set('status')}><option value="">All statuses</option>{STATUSES.map(t => <option key={t}>{t}</option>)}</select>
      <input type="date" title="Production date" className={inp} value={f.date} onChange={set('date')} />
      <input type="date" title="Deadline" className={inp} value={f.deadline} onChange={set('deadline')} />
      <input className={inp} placeholder="Location" value={f.loc} onChange={set('loc')} />
      <button className={sec + ' justify-center'} onClick={() => setF(emptyF)}>Clear filters</button>
    </div>
  )
}

/* ---------- production form ---------- */
const ProductionForm = ({ item, date, onSave, onClose, nextPid }) => {
  const [f, setF] = useState(item || { name: '', client: '', type: 'Video', description: '', date: date || '', start: '09:00', end: '17:00', location: '', deadline: '', team: '', notes: '', status: 'Upcoming', r1: true, r3: false, r7: false, custom: '' })
  const [err, setErr] = useState({})
  const set = k => e => setF({ ...f, [k]: e.target.type === 'checkbox' ? e.target.checked : e.target.value })
  const submit = e => {
    e.preventDefault()
    const x = {}
    if (!f.name.trim()) x.name = 'Project name is required'
    if (!f.client.trim()) x.client = 'Client name is required'
    if (!f.date) x.date = 'Production date is required'
    if (f.start && f.end && f.end <= f.start) x.end = 'End time must be after start time'
    if (f.r7 && !f.deadline) x.r7 = 'Set a deadline to use this reminder'
    setErr(x)
    if (Object.keys(x).length === 0) onSave({ ...f, name: f.name.trim(), client: f.client.trim() }, item?.id)
  }
  return (
    <Modal onClose={onClose}>
      <h2 className="text-lg font-semibold mb-4">{item ? 'Edit production' : 'New production'}</h2>
      <form onSubmit={submit} className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field l="Project ID"><input className={inp + ' opacity-60'} readOnly value={item ? item.pid : nextPid} /></Field>
        <Field l="Status"><select className={inp} value={f.status} onChange={set('status')}>{STATUSES.map(s => <option key={s}>{s}</option>)}</select></Field>
        <Field l="Project name *" e={err.name}><input className={inp} value={f.name} onChange={set('name')} /></Field>
        <Field l="Client name *" e={err.client}><input className={inp} value={f.client} onChange={set('client')} /></Field>
        <Field l="Production type"><select className={inp} value={f.type} onChange={set('type')}>{TYPES.map(s => <option key={s}>{s}</option>)}</select></Field>
        <Field l="Location"><input className={inp} value={f.location} onChange={set('location')} /></Field>
        <div className="sm:col-span-2"><Field l="Description"><textarea rows={2} className={inp} value={f.description} onChange={set('description')} /></Field></div>
        <Field l="Production date *" e={err.date}><input type="date" className={inp} value={f.date} onChange={set('date')} /></Field>
        <Field l="Deadline"><input type="date" className={inp} value={f.deadline} onChange={set('deadline')} /></Field>
        <Field l="Start time"><input type="time" className={inp} value={f.start} onChange={set('start')} /></Field>
        <Field l="End time" e={err.end}><input type="time" className={inp} value={f.end} onChange={set('end')} /></Field>
        <div className="sm:col-span-2"><Field l="Team members (comma separated)"><input className={inp} value={f.team} onChange={set('team')} /></Field></div>
        <div className="sm:col-span-2"><Field l="Notes"><textarea rows={2} className={inp} value={f.notes} onChange={set('notes')} /></Field></div>
        <div className="sm:col-span-2 rounded-lg border border-white/10 p-3 text-sm space-y-2">
          <div className="text-white/60">Reminders (fire at 09:00 Phnom Penh time)</div>
          <label className="flex gap-2"><input type="checkbox" checked={!!f.r1} onChange={set('r1')} /> 1 day before production</label>
          <label className="flex gap-2"><input type="checkbox" checked={!!f.r3} onChange={set('r3')} /> 3 days before production</label>
          <label className="flex gap-2"><input type="checkbox" checked={!!f.r7} onChange={set('r7')} /> 7 days before deadline</label>
          {err.r7 && <div className="text-red-400 text-xs">{err.r7}</div>}
          <Field l="Custom reminder date & time"><input type="datetime-local" className={inp} value={f.custom} onChange={set('custom')} /></Field>
        </div>
        <div className="sm:col-span-2 flex justify-end gap-2 pt-2"><button type="button" className={sec} onClick={onClose}>Cancel</button><button className={pri}>Save production</button></div>
      </form>
    </Modal>
  )
}

/* ---------- details ---------- */
const Details = ({ p, onClose, onEdit, onDup, onDel, onStatus }) => {
  const I = ({ l, v }) => <div><div className="text-xs text-white/50">{l}</div><div className="text-sm">{v || '—'}</div></div>
  return (
    <Modal onClose={onClose}>
      <div className="pr-8"><div className="text-xs text-white/50">{p.pid}</div><h2 className="text-xl font-semibold">{p.name}</h2></div>
      <div className="flex flex-wrap gap-2 my-3">
        <select className={inp + ' !w-auto !mt-0'} value={p.status} onChange={e => onStatus(p.id, e.target.value)}>{STATUSES.map(s => <option key={s}>{s}</option>)}</select>
        <button className={sec} onClick={() => onEdit(p)}><Pencil size={14} />Edit</button>
        <button className={sec} onClick={() => onDup(p)}><Copy size={14} />Duplicate</button>
        <button className={btn + 'bg-red-900/60 hover:bg-red-800'} onClick={() => onDel(p)}><Trash2 size={14} />Delete</button>
      </div>
      <div className="grid grid-cols-2 gap-4 my-4">
        <I l="Client" v={p.client} /><I l="Type" v={p.type} />
        <I l="Production date" v={fmtD(p.date)} /><I l="Time" v={`${p.start || '—'} – ${p.end || '—'}`} />
        <I l="Location" v={p.location} /><I l="Deadline" v={fmtD(p.deadline)} />
        <I l="Team" v={p.team} /><I l="Status" v={p.status} />
        <div className="col-span-2"><I l="Description" v={p.description} /></div>
        <div className="col-span-2"><I l="Notes" v={p.notes} /></div>
      </div>
      <h3 className="font-semibold text-sm mb-2">Reminders & history</h3>
      <div className="text-sm space-y-1 mb-4">
        {p.reminders.length === 0 && p.history.length === 0 && <div className="text-white/40">No reminders set.</div>}
        {p.reminders.map(r => <div key={r.key} className="flex justify-between"><span>{r.label} · {fmtDT(r.at)}</span><span className={r.done ? 'text-green-400' : 'text-white/50'}>{r.done ? 'Completed' : 'Pending'}</span></div>)}
        {p.history.map((h, i) => <div key={i} className="text-white/50">{fmtDT(h.at)} — {h.text}</div>)}
      </div>
      <div className="text-xs text-white/40 flex gap-4"><span>Created {fmtDT(p.createdAt)}</span><span>Updated {fmtDT(p.updatedAt)}</span></div>
    </Modal>
  )
}

/* ---------- calendar ---------- */
const Cal = ({ ps, today, onNew, onOpen, onMove }) => {
  const [mode, setMode] = useState('month')
  const [cur, setCur] = useState(today)
  const nav = n => setCur(c => mode === 'month' || mode === 'list' ? addMonths(c, n) : addDays(c, mode === 'week' ? 7 * n : n))
  let days = []
  if (mode === 'month') { const first = cur.slice(0, 8) + '01'; const s = addDays(first, -dow(first)); days = Array.from({ length: 42 }, (_, i) => addDays(s, i)) }
  if (mode === 'week') { const s = addDays(cur, -dow(cur)); days = Array.from({ length: 7 }, (_, i) => addDays(s, i)) }
  if (mode === 'day') days = [cur]
  const title = mode === 'month' || mode === 'list' ? monthTitle(cur) : mode === 'week' ? `${fmtD(days[0])} – ${fmtD(days[6])}` : fmtD(cur)
  const h = mode === 'month' ? 'min-h-[84px]' : mode === 'week' ? 'min-h-[280px]' : 'min-h-[360px]'
  return (
    <Card>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <button className={sec} onClick={() => nav(-1)}><ChevronLeft size={16} /></button>
        <button className={sec} onClick={() => setCur(today)}>Today</button>
        <button className={sec} onClick={() => nav(1)}><ChevronRight size={16} /></button>
        <h2 className="font-semibold mx-2">{title}</h2>
        <div className="ml-auto flex rounded-lg overflow-hidden border border-white/10">
          {['month', 'week', 'day', 'list'].map(m => <button key={m} onClick={() => setMode(m)} className={`px-3 py-2 text-sm capitalize ${mode === m ? 'bg-red-700' : 'hover:bg-white/10'}`}>{m}</button>)}
        </div>
      </div>
      {mode === 'list' ? <Table ps={ps.filter(p => p.date.startsWith(cur.slice(0, 7)))} onOpen={onOpen} /> : (
        <div className="overflow-x-auto"><div className={mode === 'day' ? '' : 'min-w-[640px]'}>
          {mode !== 'day' && <div className="grid grid-cols-7 text-xs text-white/50 mb-1">{['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(d => <div key={d} className="px-2">{d}</div>)}</div>}
          <div className={`grid gap-px bg-white/10 rounded-lg overflow-hidden ${mode === 'day' ? 'grid-cols-1' : 'grid-cols-7'}`}>
            {days.map(d => (
              <div key={d} className={`${h} bg-[#1f2124] p-1.5 cursor-pointer hover:bg-[#25282c] ${mode === 'month' && d.slice(0, 7) !== cur.slice(0, 7) ? 'opacity-40' : ''}`}
                onClick={() => onNew(d)} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); const id = e.dataTransfer.getData('text/plain'); if (id) onMove(id, d) }}>
                <div className={`text-xs mb-1 w-6 h-6 flex items-center justify-center rounded-full ${d === today ? 'bg-red-700' : 'text-white/60'}`}>{d.slice(8)}</div>
                <div className="space-y-1">{ps.filter(p => p.date === d).sort((a, b) => (a.start || '').localeCompare(b.start || '')).map(p => (
                  <div key={p.id} draggable onDragStart={e => e.dataTransfer.setData('text/plain', p.id)} onClick={e => { e.stopPropagation(); onOpen(p.id) }}
                    className={`${SC[p.status]} text-xs rounded px-1.5 py-1 truncate cursor-grab`} title={`${p.name} (${p.status})`}>{p.start} {p.name}</div>))}</div>
              </div>))}
          </div>
        </div></div>
      )}
      <div className="flex flex-wrap gap-3 mt-4 text-xs text-white/60">{STATUSES.map(s => <span key={s} className="flex items-center gap-1"><i className={`${SC[s].split(' ')[0]} w-2.5 h-2.5 rounded-sm inline-block`} />{s}</span>)}</div>
      <p className="text-xs text-white/40 mt-2">Click a date to add, click an event to open, drag an event to another day to reschedule.</p>
    </Card>
  )
}

/* ---------- app ---------- */
const NAV = [['dashboard', 'Dashboard', LayoutDashboard], ['calendar', 'Calendar', CalendarDays], ['productions', 'Productions', ListChecks], ['reminders', 'Reminders', Bell], ['data', 'Data & Backup', Database]]

export default function App() {
  const [ps, setPs] = useState(load)
  const [view, setView] = useState('dashboard')
  const [modal, setModal] = useState(null)
  const [menu, setMenu] = useState(false)
  const [filters, setFilters] = useState(emptyF)
  const [now, setNow] = useState(nowStr())
  const today = now.slice(0, 10)

  useEffect(() => { try { localStorage.setItem(KEY, JSON.stringify(ps)) } catch { alert('Could not save to LocalStorage (storage full or blocked).') } }, [ps])

  useEffect(() => {
    const tick = () => {
      const n = nowStr(); setNow(n)
      setPs(list => {
        let changed = false
        const out = list.map(p => {
          const due = p.reminders.filter(r => !r.done && !r.notified && r.at <= n)
          if (!due.length) return p
          changed = true
          due.forEach(r => { try { if ('Notification' in window && Notification.permission === 'granted') new Notification(p.name, { body: r.label, tag: p.id + r.key }) } catch { /* ignore */ } })
          return { ...p, reminders: p.reminders.map(r => due.includes(r) ? { ...r, notified: true } : r), history: [...p.history, ...due.map(r => ({ at: n, text: 'Reminder triggered: ' + r.label }))] }
        })
        return changed ? out : list
      })
    }
    tick(); const t = setInterval(tick, 30000); return () => clearInterval(t)
  }, [])

  const patch = (id, ch) => setPs(l => l.map(x => x.id === id ? { ...x, ...ch, reminders: buildRem({ ...x, ...ch }, x.reminders), updatedAt: nowStr() } : x))
  const save = (f, id) => {
    if (id) patch(id, f)
    else setPs(l => {
      const pid = 'PRJ-' + String(Math.max(0, ...l.map(x => parseInt(x.pid.slice(4)) || 0)) + 1).padStart(4, '0'), n = nowStr()
      return [...l, { ...f, id: uid(), pid, reminders: buildRem(f), history: [], createdAt: n, updatedAt: n }]
    })
    setModal(null)
  }
  const nextPid = 'PRJ-' + String(Math.max(0, ...ps.map(x => parseInt(x.pid.slice(4)) || 0)) + 1).padStart(4, '0')
  const del = p => { if (confirm(`Delete "${p.name}"? This cannot be undone.`)) { setPs(l => l.filter(x => x.id !== p.id)); setModal(null) } }
  const dup = p => { save({ ...p, name: p.name + ' (Copy)', status: 'Upcoming' }); }
  const doneRem = (pid, key) => setPs(l => l.map(p => p.id === pid ? { ...p, reminders: p.reminders.map(r => r.key === key ? { ...r, done: true } : r), history: [...p.history, { at: nowStr(), text: 'Reminder completed: ' + p.reminders.find(r => r.key === key).label }] } : p))
  const open = id => setModal({ t: 'detail', id })
  const newOn = date => setModal({ t: 'form', date })

  const rems = useMemo(() => ps.flatMap(p => p.reminders.map(r => ({ ...r, p }))).sort((a, b) => a.at.localeCompare(b.at)), [ps])
  const dueRems = rems.filter(r => !r.done && r.at <= now)
  const upRems = rems.filter(r => !r.done && r.at > now)
  const overdue = ps.filter(p => isOverdue(p, today))
  const filtered = applyF(ps, filters)

  const Dashboard = () => {
    const live = ps.filter(p => !DONE.includes(p.status))
    const stats = [
      ['Today', live.filter(p => p.date === today).length, CalendarDays],
      ['Upcoming', live.filter(p => p.date > today).length, Clock],
      ['In progress', ps.filter(p => p.status === 'In Progress').length, ListChecks],
      ['Completed', ps.filter(p => p.status === 'Completed').length, Check],
      ['Overdue', overdue.length, AlertTriangle],
    ]
    const deadlines = live.filter(p => p.deadline && p.deadline >= today && p.deadline <= addDays(today, 14)).sort((a, b) => a.deadline.localeCompare(b.deadline))
    return (
      <div className="space-y-4">
        {dueRems.length > 0 && <div className="rounded-xl bg-red-900/40 border border-red-700 p-3 text-sm flex items-center gap-2"><Bell size={16} />{dueRems.length} reminder(s) due. <button className="underline" onClick={() => setView('reminders')}>Review</button></div>}
        <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">{stats.map(([l, n, Ic]) => (
          <div key={l} className={`rounded-xl border p-4 ${l === 'Overdue' && n ? 'bg-red-900/30 border-red-700' : 'bg-[#26292d] border-white/10'}`}>
            <Ic size={18} className="text-red-500" /><div className="text-3xl font-semibold mt-2">{n}</div><div className="text-sm text-white/60">{l}</div></div>))}</div>
        <div className="grid lg:grid-cols-2 gap-4">
          <Card title="Today's productions">{live.filter(p => p.date === today).length ? live.filter(p => p.date === today).map(p => <Row key={p.id} p={p} onOpen={open} />) : <Empty text="Nothing scheduled today." />}</Card>
          <Card title="Upcoming productions">{live.filter(p => p.date > today).length ? live.filter(p => p.date > today).sort((a, b) => a.date.localeCompare(b.date)).slice(0, 6).map(p => <Row key={p.id} p={p} onOpen={open} />) : <Empty text="No upcoming productions. Add your first one." />}</Card>
          <Card title="Upcoming deadlines (14 days)">{deadlines.length ? deadlines.map(p => <button key={p.id} onClick={() => open(p.id)} className="w-full flex justify-between px-3 py-2 text-sm hover:bg-white/5 rounded-lg"><span className="truncate">{p.name}</span><span className="text-white/60">{fmtD(p.deadline)}</span></button>) : <Empty text="No deadlines coming up." />}</Card>
          <Card title="Overdue">{overdue.length ? overdue.map(p => <Row key={p.id} p={p} onOpen={open} />) : <Empty text="Nothing overdue." />}</Card>
        </div>
      </div>
    )
  }

  const RemList = ({ list, action }) => !list.length ? <Empty text="None." /> : list.map(r => (
    <div key={r.p.id + r.key} className="flex items-center justify-between gap-3 px-3 py-2 rounded-lg hover:bg-white/5">
      <button onClick={() => open(r.p.id)} className="text-left min-w-0"><div className="text-sm truncate">{r.p.name} — {r.label}</div><div className="text-xs text-white/50">{fmtDT(r.at)}</div></button>
      {action && <button className={sec} onClick={() => doneRem(r.p.id, r.key)}><Check size={14} />Done</button>}
    </div>))

  const importJSON = e => {
    const file = e.target.files[0]; if (!file) return
    const rd = new FileReader()
    rd.onload = () => {
      try {
        const d = JSON.parse(rd.result)
        if (!Array.isArray(d) || d.some(x => !x || !x.name || !x.date)) throw new Error()
        if (!confirm(`Replace current data with ${d.length} imported production(s)?`)) return
        setPs(d.map((x, i) => ({ type: 'Other', status: 'Upcoming', ...x, id: x.id || uid(), pid: x.pid || 'PRJ-' + String(i + 1).padStart(4, '0'), reminders: x.reminders || [], history: x.history || [], createdAt: x.createdAt || nowStr(), updatedAt: x.updatedAt || nowStr() })))
      } catch { alert('Invalid backup file. Please choose a JSON file exported from this app.') }
    }
    rd.readAsText(file); e.target.value = ''
  }
  const csv = () => {
    const cols = ['pid', 'name', 'client', 'type', 'status', 'date', 'start', 'end', 'location', 'deadline', 'team', 'description', 'notes']
    const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`
    dl('productions.csv', [cols.join(','), ...ps.map(p => cols.map(c => q(p[c])).join(','))].join('\n'), 'text/csv')
  }

  /* ----- Excel sync with data/production-data.xlsx (via the dev server) ----- */
  const [auth, setAuth] = useState(!CLOUD || localStorage.getItem('psm.pw') ? 'ok' : 'login')
  const [sync, setSync] = useState({ state: 'connecting', at: '' })
  const psRef = useRef(ps); psRef.current = ps
  const ready = useRef(false), skip = useRef(false), dirty = useRef(false), busy = useRef(false), lastM = useRef(0)
  const flush = async () => {
    if (busy.current || !ready.current) return
    busy.current = true
    try {
      while (dirty.current) {
        dirty.current = false
        lastM.current = await remotePut(psRef.current); setSync({ state: 'ok', at: nowStr() })
      }
    } catch (e) { dirty.current = true; if (e.auth) { ready.current = false; setAuth('login') } else setSync({ state: 'locked', at: '' }) }
    busy.current = false
  }
  const pull = async () => {
    const j = await remoteGet()
    if (!j.exists) return false
    lastM.current = j.mtime
    if (!j.rows.length && !CLOUD) return false
    skip.current = true; setPs(j.rows); return true
  }
  useEffect(() => {
    if (auth !== 'ok') return
    const init = async () => {
      try {
        const had = await pull()
        ready.current = true
        if (!had) dirty.current = true
        setSync({ state: 'ok', at: nowStr() }); flush()
      } catch (e) { if (e.auth) setAuth('login'); else setSync({ state: 'off', at: '' }) }
    }
    const check = async () => {
      if (busy.current) return
      if (!ready.current) return init()
      if (dirty.current) return flush()
      if (document.visibilityState === 'hidden') return
      try { const j = await remoteGet(true); if (j.exists && j.mtime !== lastM.current) { await pull(); setSync({ state: 'ok', at: nowStr() }) } } catch (e) { if (e.auth) { ready.current = false; setAuth('login') } }
    }
    init()
    const t = setInterval(check, CLOUD ? 15000 : 4000)
    document.addEventListener('visibilitychange', check)
    return () => { clearInterval(t); document.removeEventListener('visibilitychange', check) }
  }, [auth])
  useEffect(() => {
    if (!ready.current) return
    if (skip.current) { skip.current = false; return }
    dirty.current = true; flush()
  }, [ps])
  const replaceFromBuf = buf => {
    try {
      const d = rowsFromXlsx(buf)
      if (!d.length) throw new Error()
      if (ps.length && !confirm(`Replace current data with ${d.length} production(s) from Excel?`)) return
      setPs(d)
    } catch { alert('Could not read this Excel file. It needs a sheet with columns like name and date (use Export Excel to see the layout).') }
  }
  const importXlsx = e => { const f = e.target.files[0]; if (!f) return; f.arrayBuffer().then(replaceFromBuf); e.target.value = '' }
  const W = CLOUD ? 'Cloud' : 'Excel'
  const syncMsg = { connecting: ['text-white/50', W + ': connecting…'], ok: ['text-green-400', W + ': synced ' + (sync.at ? sync.at.slice(11) : '')], locked: CLOUD ? ['text-amber-400', 'Cloud: cannot save (offline?). Retrying…'] : ['text-amber-400', 'Excel file is open or locked. Close it in Excel; retrying…'], off: CLOUD ? ['text-amber-400', 'Cloud: offline. Retrying…'] : ['text-white/40', 'Excel: not connected (start with npm run dev)'] }[sync.state]

  const title = NAV.find(n => n[0] === view)[1]
  if (auth === 'login') return <Login onDone={() => setAuth('ok')} />
  return (
    <div className="min-h-screen bg-[#1c1e21] text-white flex" style={{ fontFamily: 'Inter, system-ui, sans-serif' }}>
      {menu && <div className="fixed inset-0 bg-black/60 z-20 md:hidden" onClick={() => setMenu(false)} />}
      <aside className={`fixed inset-y-0 left-0 z-30 w-60 bg-[#141517] border-r border-white/10 p-4 flex flex-col transition-transform ${menu ? 'translate-x-0' : '-translate-x-full'} md:translate-x-0`}>
        <div className="flex items-center gap-2 mb-8 px-2"><div className="w-8 h-8 rounded-lg bg-red-700 flex items-center justify-center font-bold">P</div><div className="font-semibold leading-tight">Production<br /><span className="text-white/50 text-xs font-normal">Schedule Manager</span></div></div>
        <nav className="space-y-1">{NAV.map(([k, l, Ic]) => (
          <button key={k} onClick={() => { setView(k); setMenu(false) }} className={`w-full flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm ${view === k ? 'bg-red-700' : 'text-white/70 hover:bg-white/10'}`}>
            <Ic size={18} />{l}{k === 'reminders' && dueRems.length > 0 && <span className="ml-auto bg-white text-red-700 text-xs rounded-full px-1.5">{dueRems.length}</span>}</button>))}</nav>
        <div className="mt-auto text-xs text-white/40 px-2 space-y-1"><div className={syncMsg[0]}>{syncMsg[1]}</div>{CLOUD && <button className="underline" onClick={() => { localStorage.removeItem('psm.pw'); ready.current = false; setAuth('login') }}>Log out</button>}<div>Time zone: {TZ}<br />{fmtDT(now)}</div></div>
      </aside>
      <main className="flex-1 md:ml-60 min-w-0 p-4 sm:p-6">
        <header className="flex items-center gap-3 mb-5">
          <button className="md:hidden" onClick={() => setMenu(true)}><Menu /></button>
          <h1 className="text-2xl font-semibold">{title}</h1>
          <button className={pri + ' ml-auto'} onClick={() => newOn(today)}><Plus size={16} />New production</button>
        </header>
        {view === 'dashboard' && <Dashboard />}
        {view === 'calendar' && <Cal ps={ps} today={today} onNew={newOn} onOpen={open} onMove={(id, d) => patch(id, { date: d })} />}
        {view === 'productions' && <Card><FilterBar f={filters} setF={setFilters} /><Table ps={filtered} onOpen={open} /></Card>}
        {view === 'reminders' && (
          <div className="space-y-4">
            {'Notification' in window && Notification.permission !== 'granted' && <button className={sec} onClick={() => Notification.requestPermission().then(() => setNow(nowStr()))}><Bell size={14} />Enable browser notifications</button>}
            <Card title="Due now"><RemList list={dueRems} action /></Card>
            <Card title="Upcoming"><RemList list={upRems} action /></Card>
            <Card title="Completed"><RemList list={rems.filter(r => r.done)} /></Card>
            <Card title="Overdue deadlines">{overdue.length ? overdue.map(p => <Row key={p.id} p={p} onOpen={open} />) : <Empty text="No overdue deadlines." />}</Card>
          </div>)}
        {view === 'data' && (
          <Card title="Backup & export">
            <p className="text-sm text-white/60 mb-4">{ps.length} production(s) saved in this browser. Data stays on this device.</p>
            <div className="flex flex-wrap gap-2">
              <button className={pri} onClick={() => dl(`production-backup-${today}.json`, JSON.stringify(ps, null, 2), 'application/json')}><Download size={14} />Export JSON</button>
              <label className={sec + ' cursor-pointer'}><Upload size={14} />Import JSON<input type="file" accept="application/json,.json" className="hidden" onChange={importJSON} /></label>
              <button className={sec} onClick={csv}><Download size={14} />Export CSV</button>
              <button className={sec} onClick={() => dl(`productions-${today}.xlsx`, xlsxBytes(ps), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}><Download size={14} />Export Excel</button>
              <label className={sec + ' cursor-pointer'}><Upload size={14} />Import Excel<input type="file" accept=".xlsx" className="hidden" onChange={importXlsx} /></label>
              <button className={btn + 'bg-red-900/60 hover:bg-red-800'} onClick={() => confirm('Delete ALL productions? Export a backup first.') && setPs([])}><Trash2 size={14} />Clear all data</button>
            </div>
          </Card>)}
        {view === 'data' && (
          <div className="mt-4"><Card title={CLOUD ? 'Online storage (shared across devices)' : 'Excel file (always connected)'}>
            {CLOUD
              ? <p className="text-sm text-white/60 mb-3">Your data is stored online, so every device you log in on shows the same data. Changes save automatically and other devices update within about 15 seconds. Use Export Excel above to download a copy any time.</p>
              : <p className="text-sm text-white/60 mb-3">Every change is saved automatically to <b className="text-white">data/production-data.xlsx</b> inside the project folder, and the app loads that file whenever it starts. If you edit and save the file in Excel, the app picks up the change within a few seconds.</p>}
            <p className={`text-sm mb-3 ${syncMsg[0]}`}>{syncMsg[1]}</p>
            <button className={sec} onClick={() => { dirty.current = true; flush() }}>Save now</button>
            {!CLOUD && <p className="text-xs text-white/40 mt-3">On Windows, Excel locks a file while it is open, so the app cannot save until you close it. Edit in Excel, save, then close Excel.</p>}
          </Card></div>)}
      </main>
      {modal?.t === 'form' && <ProductionForm item={modal.item} date={modal.date} nextPid={nextPid} onSave={save} onClose={() => setModal(null)} />}
      {modal?.t === 'detail' && ps.find(p => p.id === modal.id) && <Details p={ps.find(p => p.id === modal.id)} onClose={() => setModal(null)} onEdit={p => setModal({ t: 'form', item: p })} onDup={dup} onDel={del} onStatus={(id, s) => patch(id, { status: s })} />}
    </div>
  )
}
