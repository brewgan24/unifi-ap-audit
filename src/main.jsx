import { StrictMode, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import './styles.css'

const DEFAULT_APS = [
  ['ap1', 'LT1 - Cafe Bar'], ['ap2', 'LT1 - Casual Eat'], ['ap3', 'LT1 - Casual Work'],
  ['ap4', 'LT2 - Cafe Library'], ['ap5', 'LT2 - Concentration'], ['ap6', 'LT2 - Kiva'],
  ['ap7', 'LT2 - MultiRoom'], ['ap8', 'LT3 - Meeting Room'], ['ap9', 'Outdoor Area'],
].map(([id, name]) => ({ id, name }))

const FIELDS = [
  ['name', 'Client name'], ['mac', 'MAC address'], ['accessPoint', 'Access point'], ['network', 'Network'],
  ['quality', 'Connection quality'], ['signal', 'Signal'], ['wifiMode', 'Wi-Fi mode'], ['channel', 'Channel'],
  ['radioWidth', 'Radio / channel width'], ['ip', 'IP address'], ['uploadRate', 'Upload rate'],
  ['downloadRate', 'Download rate'], ['activity', 'Activity'], ['usage', 'Usage'], ['connectedTime', 'Connected time'],
  ['transmitRate', 'Transmit rate'], ['receiveRate', 'Receive rate'],
]
const EMPTY_CLIENT = Object.fromEntries(FIELDS.map(([key]) => [key, '']))
const emptyData = () => ({ auditDate: todayISO(), myDeviceRaw: '', otherClientsRaw: '', myDevice: null, clients: [], speedtestImg: null, speedtest: { download: '', upload: '', ping: '', server: '', testedAt: '' }, notes: '' })

function todayISO() { return new Date().toISOString().slice(0, 10) }
function loadJson(key, fallback) { try { const value = localStorage.getItem(key); return value ? JSON.parse(value) : fallback } catch { return fallback } }
function titleCase(value) { return value ? value.charAt(0).toUpperCase() + value.slice(1) : '' }

function parseRawRows(raw) {
  if (!raw?.trim()) return []
  const lines = raw.trim().split(/\r?\n/).map((line) => line.trim()).filter(Boolean)
  if (!lines.length) return []
  if (lines.some((line) => line.includes('\t'))) return lines.map((line) => line.split('\t').map((cell) => cell.trim()))

  const macIndexes = lines.reduce((indexes, line, index) => {
    if (/^(?:[0-9a-f]{2}:){5}[0-9a-f]{2}$/i.test(line)) indexes.push(index)
    return indexes
  }, [])
  if (!macIndexes.length) return [lines]
  return macIndexes.map((macIndex, index) => {
    const start = Math.max(0, macIndex - 1)
    const end = macIndexes[index + 1] ? Math.max(start + 1, macIndexes[index + 1] - 1) : lines.length
    return lines.slice(start, end)
  }).filter(Boolean)
}

function rowToClient(row) {
  const client = { ...EMPTY_CLIENT }
  FIELDS.forEach(([key], index) => { client[key] = row[index] ?? '' })
  return client
}
function rawToClients(raw) { return parseRawRows(raw).map(rowToClient).filter((client) => Object.values(client).some(Boolean)) }
function clientsToRaw(clients) { return clients.map((client) => FIELDS.map(([key]) => client[key] || '').join('\t')).join('\n') }
function getData(data) {
  const safe = { ...emptyData(), ...data, speedtest: { ...emptyData().speedtest, ...(data?.speedtest ?? {}) } }
  const myDevice = safe.myDevice && Object.values(safe.myDevice).some(Boolean) ? safe.myDevice : rawToClients(safe.myDeviceRaw)[0] ?? null
  const clients = Array.isArray(safe.clients) && safe.clients.length ? safe.clients : rawToClients(safe.otherClientsRaw)
  return { ...safe, myDevice, clients }
}
function hasData(data) { const safe = getData(data); return Boolean(safe.myDevice || safe.clients.length || safe.speedtestImg || safe.notes?.trim()) }
function statusOf(client) { return (client?.quality || '').toLowerCase() }
function tone(value) { const lower = String(value ?? '').toLowerCase(); if (lower === 'excellent' || lower === 'good') return 'good'; if (lower === 'fair') return 'warn'; if (lower === 'poor') return 'bad'; if (lower.includes('mbps')) return 'speed'; if (String(value).includes('dBm')) { const signal = Number.parseInt(value, 10); return signal >= -60 ? 'good' : signal >= -75 ? 'warn' : 'bad' } return '' }
function formatDate(date) { return date ? new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric', year: 'numeric' }).format(new Date(`${date}T00:00:00`)) : '—' }
function Icon({ children }) { return <span className="icon" aria-hidden="true">{children}</span> }

function StatCard({ icon, label, value, detail, color }) { return <div className={`stat-card ${color}`}><div className="stat-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></div> }
function StatusBadge({ value }) { const status = value || 'Not recorded'; return <span className={`status-badge ${tone(status)}`}>{status}</span> }

function DeviceFields({ client, compact = false }) {
  const visible = compact ? FIELDS.filter(([key]) => ['mac', 'network', 'quality', 'signal', 'wifiMode', 'channel', 'ip', 'connectedTime'].includes(key)) : FIELDS
  return <div className={`device-fields ${compact ? 'compact' : ''}`}>{visible.map(([key, label]) => <div className="device-field" key={key}><span>{label}</span><strong className={tone(client?.[key])}>{client?.[key] || '—'}</strong></div>)}</div>
}

function ClientTable({ clients, onEdit, onDelete }) {
  if (!clients.length) return <div className="table-empty"><span>♧</span><strong>No other connected devices</strong><small>Paste UniFi client data or add a device manually.</small></div>
  return <div className="client-table-wrap"><table className="client-table"><thead><tr><th>Client</th><th>MAC address</th><th>IP address</th><th>Network</th><th>Signal</th><th>Quality</th><th>Rates</th><th /></tr></thead><tbody>{clients.map((client, index) => <tr key={`${client.mac || client.name || 'client'}-${index}`}><td><strong>{client.name || 'Unnamed client'}</strong><small>{client.wifiMode || '—'}</small></td><td>{client.mac || '—'}</td><td>{client.ip || '—'}</td><td>{client.network || '—'}</td><td className={tone(client.signal)}>{client.signal || '—'}</td><td><StatusBadge value={client.quality} /></td><td><span className="rate-text">↓ {client.downloadRate || '—'}</span><span className="rate-text">↑ {client.uploadRate || '—'}</span></td><td className="row-actions"><button onClick={() => onEdit(client, index)} title="Edit client">✎</button><button onClick={() => onDelete(index)} title="Delete client">×</button></td></tr>)}</tbody></table></div>
}

function ImageDropzone({ image, onImage, onRemove, fileInputRef }) {
  const readImage = (file) => { if (!file || !file.type.startsWith('image/')) return; const reader = new FileReader(); reader.onload = (event) => onImage(event.target.result); reader.readAsDataURL(file) }
  return <div className={`dropzone ${image ? 'has-image' : ''}`} onClick={() => fileInputRef.current?.click()}>{image ? <><img src={image} alt="Ookla speed test" /><span className="image-overlay">↻ Replace screenshot</span><button className="remove-image" onClick={(event) => { event.stopPropagation(); onRemove() }}>×</button></> : <><span className="upload-icon">▧</span><strong>Drop screenshot here</strong><span>or click to browse · Ctrl + V</span></>}<input ref={fileInputRef} type="file" accept="image/*" hidden onChange={(event) => { readImage(event.target.files?.[0]); event.target.value = '' }} /></div>
}

function ClientEditor({ client, title, onSave, onClose }) {
  const [draft, setDraft] = useState({ ...EMPTY_CLIENT, ...client })
  const update = (key, value) => setDraft((current) => ({ ...current, [key]: value }))
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="modal large-modal"><div className="modal-title"><div><span className="eyebrow">DEVICE CRUD</span><h3>{title}</h3></div><button className="close-button" onClick={onClose}>×</button></div><div className="editor-grid">{FIELDS.map(([key, label]) => <label key={key}>{label}<input value={draft[key]} onChange={(event) => update(key, event.target.value)} placeholder={label} /></label>)}</div><div className="modal-actions"><button className="secondary-button" onClick={onClose}>Cancel</button><button className="primary-button" onClick={() => onSave(draft)}>Save device</button></div></section></div>
}

function LocationModal({ aps, onAdd, onEdit, onDelete, onClose }) {
  const [name, setName] = useState('')
  const [editing, setEditing] = useState(null)
  const submit = (event) => { event.preventDefault(); if (!name.trim()) return; editing ? onEdit(editing.id, name.trim()) : onAdd(name.trim()); setName(''); setEditing(null) }
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="modal"><div className="modal-title"><div><span className="eyebrow">WORKSPACE CRUD</span><h3>Manage locations</h3></div><button className="close-button" onClick={onClose}>×</button></div><form className="add-form" onSubmit={submit}><input value={name} onChange={(event) => setName(event.target.value)} placeholder={editing ? 'Rename location' : 'e.g. LT1 - Lobby'} autoFocus /><button className="primary-button" type="submit">{editing ? 'Save' : 'Add'}</button></form><div className="modal-list">{aps.map((ap) => <div className="modal-item" key={ap.id}><span>{ap.name}</span><div><button onClick={() => { setEditing(ap); setName(ap.name) }}>Edit</button><button className="danger-text" onClick={() => onDelete(ap.id)}>Delete</button></div></div>)}{!aps.length && <p className="modal-empty">No locations found.</p>}</div><button className="modal-done" onClick={onClose}>Done</button></section></div>
}

function exportExcel(aps, allData, dateLabel) {
  const summary = [['UniFi Wi-Fi Audit Dashboard'], ['Exported', dateLabel], [], ['Location', 'Audit date', 'Status', 'My device', 'Other devices', 'Total devices', 'Alerts']]
  const devices = [['Location', 'Record type', ...FIELDS.map(([, label]) => label)]]
  aps.forEach((ap) => { const data = getData(allData[ap.id]); const all = data.myDevice ? [data.myDevice, ...data.clients] : data.clients; const alerts = all.filter((client) => ['fair', 'poor'].includes(statusOf(client))).length; summary.push([ap.name, data.auditDate, hasData(data) ? 'In progress' : 'Not started', data.myDevice?.name || '', data.clients.length, all.length, alerts]); if (data.myDevice) devices.push([ap.name, 'My device', ...FIELDS.map(([key]) => data.myDevice[key] || '')]); data.clients.forEach((client) => devices.push([ap.name, 'Other device', ...FIELDS.map(([key]) => client[key] || '')])) })
  const workbook = XLSX.utils.book_new(); const addSheet = (name, rows, widths) => { const sheet = XLSX.utils.aoa_to_sheet(rows); sheet['!cols'] = widths.map((wch) => ({ wch })); XLSX.utils.book_append_sheet(workbook, sheet, name) }; addSheet('Dashboard', summary, [25, 15, 15, 24, 15, 15, 10]); addSheet('Devices', devices, [24, 16, ...FIELDS.map(() => 20)]); XLSX.writeFile(workbook, `unifi-audit-${todayISO()}.xlsx`)
}

function App() {
  const [aps, setAps] = useState(() => loadJson('unifi-audit-aps', DEFAULT_APS))
  const [allData, setAllData] = useState(() => loadJson('unifi-audit-data', {}))
  const [selectedApId, setSelectedApId] = useState(() => loadJson('unifi-audit-aps', DEFAULT_APS)[0]?.id ?? null)
  const [showLocations, setShowLocations] = useState(false); const [editor, setEditor] = useState(null); const [notice, setNotice] = useState(''); const fileInputRef = useRef(null)
  useEffect(() => localStorage.setItem('unifi-audit-aps', JSON.stringify(aps)), [aps])
  useEffect(() => { try { localStorage.setItem('unifi-audit-data', JSON.stringify(allData)) } catch { setNotice('Browser storage is full. Remove an old screenshot.') } }, [allData])
  useEffect(() => { setAllData((current) => { const next = { ...current }; aps.forEach((ap) => { if (!next[ap.id]) next[ap.id] = emptyData() }); return next }) }, [aps])
  useEffect(() => { const onPaste = (event) => { const item = [...(event.clipboardData?.items ?? [])].find((entry) => entry.kind === 'file' && entry.type.startsWith('image/')); if (!item || !selectedApId) return; event.preventDefault(); const file = item.getAsFile(); if (!file) return; const reader = new FileReader(); reader.onload = (loadEvent) => updateData({ speedtestImg: loadEvent.target.result }); reader.readAsDataURL(file) }; window.addEventListener('paste', onPaste); return () => window.removeEventListener('paste', onPaste) })
  useEffect(() => { if (!notice) return undefined; const timer = window.setTimeout(() => setNotice(''), 3600); return () => window.clearTimeout(timer) }, [notice])

  const currentAp = aps.find((ap) => ap.id === selectedApId); const currentData = getData(allData[selectedApId]); const allCurrentDevices = currentData.myDevice ? [currentData.myDevice, ...currentData.clients] : currentData.clients; const dateLabel = new Intl.DateTimeFormat('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).format(new Date()); const auditedCount = aps.filter((ap) => hasData(allData[ap.id])).length; const totalDevices = aps.reduce((total, ap) => { const data = getData(allData[ap.id]); return total + data.clients.length + (data.myDevice ? 1 : 0) }, 0); const alertCount = aps.reduce((total, ap) => total + getData(allData[ap.id]).clients.filter((client) => ['fair', 'poor'].includes(statusOf(client))).length, 0)

  function updateData(patch) { if (!selectedApId) return; setAllData((current) => ({ ...current, [selectedApId]: { ...emptyData(), ...(current[selectedApId] ?? {}), ...patch } })) }
  function importRaw(type, raw) { const parsed = rawToClients(raw); if (type === 'myDeviceRaw') updateData({ myDeviceRaw: raw, myDevice: parsed[0] ?? null }); else updateData({ otherClientsRaw: raw, clients: parsed }) }
  function saveDevice(device, index) { if (editor?.type === 'my') updateData({ myDevice: device, myDeviceRaw: clientsToRaw([device]) }); else { const clients = [...currentData.clients]; if (index === null) clients.push(device); else clients[index] = device; updateData({ clients, otherClientsRaw: clientsToRaw(clients) }) }; setEditor(null); setNotice('Device saved') }
  function deleteClient(index) { const clients = currentData.clients.filter((_, clientIndex) => clientIndex !== index); updateData({ clients, otherClientsRaw: clientsToRaw(clients) }); setNotice('Device removed') }
  function clearAudit() { if (!currentAp || !window.confirm(`Start a new audit for ${currentAp.name}? Existing data will be cleared.`)) return; updateData(emptyData()); setNotice('New audit started') }
  function addLocation(name) { const id = `ap_${Date.now()}`; setAps((items) => [...items, { id, name }]); setAllData((data) => ({ ...data, [id]: emptyData() })); setSelectedApId(id); setNotice('Location added') }
  function editLocation(id, name) { setAps((items) => items.map((ap) => ap.id === id ? { ...ap, name } : ap)); setNotice('Location renamed') }
  function deleteLocation(id) { const ap = aps.find((item) => item.id === id); if (!ap || !window.confirm(`Delete ${ap.name} and its audit data?`)) return; const remaining = aps.filter((item) => item.id !== id); setAps(remaining); setAllData((data) => { const next = { ...data }; delete next[id]; return next }); if (selectedApId === id) setSelectedApId(remaining[0]?.id ?? null) }

  return <div className="app-shell">
    <aside className="sidebar no-print"><div className="brand"><div className="brand-mark">⌁</div><div><h1>AP Audit</h1><span>UNIFI DASHBOARD</span></div></div><div className="sidebar-date"><span className="status-dot" /> {dateLabel}</div><div className="location-heading"><span>LOCATIONS</span><button className="icon-button" onClick={() => setShowLocations(true)} title="Manage locations">⚙</button></div><nav className="locations">{aps.map((ap) => <button key={ap.id} className={`location ${selectedApId === ap.id ? 'selected' : ''}`} onClick={() => setSelectedApId(ap.id)}><span className={`ap-icon ${hasData(allData[ap.id]) ? 'complete' : ''}`}>⌁</span><span className="location-name">{ap.name}</span>{hasData(allData[ap.id]) && <span className="check">✓</span>}</button>)}{!aps.length && <div className="sidebar-empty">No locations yet.</div>}</nav><div className="sidebar-footer"><div className="progress-label"><span>DAILY PROGRESS</span><strong>{auditedCount}/{aps.length}</strong></div><div className="progress-track"><div style={{ width: aps.length ? `${auditedCount / aps.length * 100}%` : '0%' }} /></div><button className="export-button" onClick={() => exportExcel(aps, allData, dateLabel)}><Icon>⇩</Icon> Export Excel</button><button className="print-button" onClick={() => window.print()}>▣ Print / PDF</button><p className="local-note">Saved locally in this browser</p></div></aside>
    <main className="workspace"><header className="topbar"><div><span className="eyebrow">WI-FI AUDIT DASHBOARD / {dateLabel.toUpperCase()}</span><h2>{currentAp?.name ?? 'No location selected'}</h2></div>{currentAp && <div className="top-actions"><button className="secondary-button" onClick={() => setShowLocations(true)}>⚙ Locations</button><button className="clear-button" onClick={clearAudit}>＋ New audit</button></div>}</header>
      <div className="dashboard-content"><section className="dashboard-intro"><div><span className="eyebrow">OVERVIEW</span><h3>Network health at a glance</h3><p>Manage daily UniFi audits and review connected devices by location.</p></div><span className="storage-pill">● LOCAL WORKSPACE</span></section><div className="stats-grid"><StatCard icon="⌁" label="Locations" value={aps.length} detail={`${auditedCount} audited today`} color="blue" /><StatCard icon="◈" label="Total devices" value={totalDevices} detail="Across all locations" color="violet" /><StatCard icon="✓" label="Audit progress" value={`${auditedCount}/${aps.length}`} detail="Locations completed" color="green" /><StatCard icon="!" label="Needs attention" value={alertCount} detail="Fair or poor signal" color="orange" /></div>
        <div className="location-tabs"><div className="tabs-label">AUDIT LOCATIONS</div><div className="tabs-scroll">{aps.map((ap) => <button key={ap.id} className={selectedApId === ap.id ? 'active' : ''} onClick={() => setSelectedApId(ap.id)}><span className={hasData(allData[ap.id]) ? 'tab-dot done' : 'tab-dot'} />{ap.name}</button>)}<button className="add-tab" onClick={() => setShowLocations(true)}>＋ Add location</button></div></div>
        {currentAp ? <><section className="location-heading-row"><div><span className="eyebrow">LOCATION AUDIT</span><h3>{currentAp.name}</h3><p>Audit date: <strong>{formatDate(currentData.auditDate)}</strong> · Last saved locally</p></div><div className="audit-actions"><label className="date-control">Audit date<input type="date" value={currentData.auditDate || todayISO()} onChange={(event) => updateData({ auditDate: event.target.value })} /></label><StatusBadge value={hasData(currentData) ? 'In progress' : 'Not started'} /></div></section><div className="dashboard-grid"><section className="dashboard-card device-card"><div className="card-top"><div><span className="eyebrow">PRIMARY DEVICE</span><h4>My device</h4></div><button className="outline-button" onClick={() => setEditor({ type: 'my', client: currentData.myDevice ?? EMPTY_CLIENT })}>✎ {currentData.myDevice ? 'Edit' : 'Add device'}</button></div>{currentData.myDevice ? <><div className="device-identity"><div className="device-avatar">▣</div><div><strong>{currentData.myDevice.name || 'Unnamed device'}</strong><span>{currentData.myDevice.ip || 'No IP'} · {currentData.myDevice.mac || 'No MAC'}</span></div><StatusBadge value={currentData.myDevice.quality} /></div><DeviceFields client={currentData.myDevice} compact /></> : <div className="card-empty"><span>▣</span><strong>No primary device recorded</strong><small>Paste UniFi data below or add it manually.</small></div>}<div className="raw-import"><label>Import My Device data <span>TSV or vertical UniFi copy</span></label><textarea value={currentData.myDeviceRaw} onChange={(event) => importRaw('myDeviceRaw', event.target.value)} placeholder={'LAPTOP-CINMP20D\n24:b2:b9:60:ed:0d\nLT1 - Casual East\nDBI_Office\nExcellent\n-49 dBm\n…'} spellCheck="false" /></div></section>
          <section className="dashboard-card speed-card"><div className="card-top"><div><span className="eyebrow">EVIDENCE</span><h4>Ookla speed test</h4></div>{currentData.speedtestImg && <button className="danger-button" onClick={() => updateData({ speedtestImg: null })}>Delete</button>}</div><ImageDropzone image={currentData.speedtestImg} onImage={(speedtestImg) => updateData({ speedtestImg })} onRemove={() => updateData({ speedtestImg: null })} fileInputRef={fileInputRef} /><div className="speed-fields"><label>Download<input value={currentData.speedtest.download} onChange={(event) => updateData({ speedtest: { ...currentData.speedtest, download: event.target.value } })} placeholder="192 Mbps" /></label><label>Upload<input value={currentData.speedtest.upload} onChange={(event) => updateData({ speedtest: { ...currentData.speedtest, upload: event.target.value } })} placeholder="78 Mbps" /></label><label>Ping<input value={currentData.speedtest.ping} onChange={(event) => updateData({ speedtest: { ...currentData.speedtest, ping: event.target.value } })} placeholder="12 ms" /></label><label>Server<input value={currentData.speedtest.server} onChange={(event) => updateData({ speedtest: { ...currentData.speedtest, server: event.target.value } })} placeholder="Ookla server" /></label></div></section>
        </div><section className="dashboard-card clients-card"><div className="card-top"><div><span className="eyebrow">CONNECTED DEVICES</span><h4>Other connected devices <span className="count-badge">{currentData.clients.length}</span></h4><p>Devices connected to this access point during the audit.</p></div><div className="card-buttons"><button className="outline-button" onClick={() => setEditor({ type: 'client', index: null, client: EMPTY_CLIENT })}>＋ Add device</button><button className="outline-button" onClick={() => document.getElementById('client-import')?.focus()}>⇧ Import data</button></div></div><ClientTable clients={currentData.clients} onEdit={(client, index) => setEditor({ type: 'client', index, client })} onDelete={deleteClient} /><div className="raw-import clients-import"><label htmlFor="client-import">Import other connected devices <span>Multiple TSV rows or vertical UniFi copies</span></label><textarea id="client-import" value={currentData.otherClientsRaw} onChange={(event) => importRaw('otherClientsRaw', event.target.value)} placeholder="Paste one or more client records here…" spellCheck="false" /></div><div className="client-footer"><span>Total connected devices</span><strong>{allCurrentDevices.length}</strong><span className="footer-separator" /> <span>Other devices</span><strong>{currentData.clients.length}</strong></div></section><section className="dashboard-card notes-card"><div className="card-top"><div><span className="eyebrow">AUDIT NOTES</span><h4>Notes for this location</h4></div></div><textarea value={currentData.notes} onChange={(event) => updateData({ notes: event.target.value })} placeholder="Add observations, maintenance notes, or follow-up actions…" /></section></> : <div className="no-ap"><span>⌁</span><h3>No locations available</h3><p>Add a location to start your first audit.</p><button className="primary-button" onClick={() => setShowLocations(true)}>Manage locations</button></div>}
      </div></main>
    {showLocations && <LocationModal aps={aps} onAdd={addLocation} onEdit={editLocation} onDelete={deleteLocation} onClose={() => setShowLocations(false)} />}{editor && <ClientEditor client={editor.client} title={editor.type === 'my' ? 'Edit my device' : editor.index === null ? 'Add connected device' : 'Edit connected device'} onSave={(client) => saveDevice(client, editor.index)} onClose={() => setEditor(null)} />}{notice && <div className="toast">✓ {notice}</div>}<PrintReport aps={aps} allData={allData} dateLabel={dateLabel} />
  </div>
}

function PrintReport({ aps, allData, dateLabel }) { return <div className="print-report"><h1>Daily Wi-Fi AP Audit Report</h1><p>{dateLabel}</p>{aps.map((ap) => { const data = getData(allData[ap.id]); const devices = data.myDevice ? [data.myDevice, ...data.clients] : data.clients; return <section key={ap.id}><h2>{ap.name} · {formatDate(data.auditDate)}</h2>{!devices.length && !data.speedtestImg ? <em>No data recorded for this location.</em> : <><table><tbody>{devices.map((client, index) => <tr key={index}>{FIELDS.slice(0, 10).map(([key]) => <td key={key}>{client[key] || '—'}</td>)}</tr>)}</tbody></table>{data.speedtestImg && <img className="print-speed" src={data.speedtestImg} alt="Speed test" />}</>}</section> })}</div> }

createRoot(document.getElementById('root')).render(<StrictMode><App /></StrictMode>)
