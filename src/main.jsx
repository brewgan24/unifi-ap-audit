import { StrictMode, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import './styles.css'

const DEFAULT_APS = [
  ['ap1', 'LT1 - Cafe Bar'],
  ['ap2', 'LT1 - Casual Eat'],
  ['ap3', 'LT1 - Casual Work'],
  ['ap4', 'LT2 - Cafe Library'],
  ['ap5', 'LT2 - Concentration'],
  ['ap6', 'LT2 - Kiva'],
  ['ap7', 'LT2 - MultiRoom'],
  ['ap8', 'LT3 - Meeting Room'],
  ['ap9', 'Outdoor Area'],
].map(([id, name]) => ({ id, name }))

const emptyData = () => ({ myDeviceRaw: '', otherClientsRaw: '', speedtestImg: null })

function loadJson(key, fallback) {
  try {
    const saved = localStorage.getItem(key)
    return saved ? JSON.parse(saved) : fallback
  } catch {
    return fallback
  }
}

function parseRows(raw) {
  if (!raw?.trim()) return []
  return raw.trim().split(/\r?\n/).map((row) => row.split('\t'))
}

function getCellTone(value) {
  const text = String(value ?? '').trim()
  const lower = text.toLowerCase()
  if (!text) return 'muted'
  if (lower === 'excellent' || lower === 'good') return 'good'
  if (lower === 'fair') return 'warn'
  if (lower === 'poor') return 'bad'
  if (lower.includes('mbps')) return 'speed'
  if (text.includes('dBm')) {
    const signal = Number.parseInt(text, 10)
    if (!Number.isNaN(signal)) return signal >= -60 ? 'good' : signal >= -75 ? 'warn' : 'bad'
  }
  return ''
}

function Icon({ children, className = '' }) {
  return <span className={`icon ${className}`} aria-hidden="true">{children}</span>
}

function AuditTable({ rows, emptyLabel }) {
  if (!rows.length) return <div className="empty-table">{emptyLabel}</div>
  return (
    <table className="audit-table">
      <tbody>
        {rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, cellIndex) => (
              <td className={getCellTone(cell)} title={cell} key={`${rowIndex}-${cellIndex}`}>
                {cell || '—'}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function ImageDropzone({ image, onImage, fileInputRef }) {
  const readImage = (file) => {
    if (!file || !file.type.startsWith('image/')) return
    const reader = new FileReader()
    reader.onload = (event) => onImage(event.target.result)
    reader.readAsDataURL(file)
  }

  return (
    <div className={`dropzone ${image ? 'has-image' : ''}`} onClick={() => fileInputRef.current?.click()}>
      {image ? (
        <>
          <img src={image} alt="Ookla speed test" />
          <span className="image-overlay">↻ Replace screenshot</span>
        </>
      ) : (
        <>
          <span className="upload-icon">▧</span>
          <strong>Drop screenshot here</strong>
          <span>or click to browse · Ctrl + V</span>
        </>
      )}
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        hidden
        onChange={(event) => {
          readImage(event.target.files?.[0])
          event.target.value = ''
        }}
      />
    </div>
  )
}

function App() {
  const [aps, setAps] = useState(() => loadJson('unifi-audit-aps', DEFAULT_APS))
  const [allData, setAllData] = useState(() => loadJson('unifi-audit-data', {}))
  const [selectedApId, setSelectedApId] = useState(() => loadJson('unifi-audit-aps', DEFAULT_APS)[0]?.id ?? null)
  const [showSettings, setShowSettings] = useState(false)
  const [newApName, setNewApName] = useState('')
  const [notice, setNotice] = useState('')
  const fileInputRef = useRef(null)

  useEffect(() => localStorage.setItem('unifi-audit-aps', JSON.stringify(aps)), [aps])
  useEffect(() => {
    try {
      localStorage.setItem('unifi-audit-data', JSON.stringify(allData))
    } catch {
      setNotice('Storage is full. Try removing an older screenshot.')
    }
  }, [allData])
  useEffect(() => {
    setAllData((current) => {
      const next = { ...current }
      aps.forEach((ap) => { if (!next[ap.id]) next[ap.id] = emptyData() })
      return next
    })
  }, [aps])
  useEffect(() => {
    const onPaste = (event) => {
      const item = [...(event.clipboardData?.items ?? [])].find((entry) => entry.kind === 'file' && entry.type.startsWith('image/'))
      if (!item || !selectedApId) return
      event.preventDefault()
      const file = item.getAsFile()
      if (!file) return
      const reader = new FileReader()
      reader.onload = (loadEvent) => updateCurrent({ speedtestImg: loadEvent.target.result })
      reader.readAsDataURL(file)
    }
    window.addEventListener('paste', onPaste)
    return () => window.removeEventListener('paste', onPaste)
  })
  useEffect(() => {
    if (!notice) return undefined
    const timer = window.setTimeout(() => setNotice(''), 3600)
    return () => window.clearTimeout(timer)
  }, [notice])

  const currentAp = aps.find((ap) => ap.id === selectedApId)
  const currentData = allData[selectedApId] ?? emptyData()
  const myDeviceRows = useMemo(() => parseRows(currentData.myDeviceRaw), [currentData.myDeviceRaw])
  const clientRows = useMemo(() => parseRows(currentData.otherClientsRaw), [currentData.otherClientsRaw])
  const completedCount = aps.filter((ap) => hasData(ap.id)).length
  const dateLabel = new Intl.DateTimeFormat('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).format(new Date())

  function hasData(id) {
    const data = allData[id]
    return Boolean(data?.myDeviceRaw?.trim() || data?.otherClientsRaw?.trim() || data?.speedtestImg)
  }

  function updateCurrent(patch) {
    if (!selectedApId) return
    setAllData((current) => ({ ...current, [selectedApId]: { ...(current[selectedApId] ?? emptyData()), ...patch } }))
  }

  function addAp(event) {
    event?.preventDefault()
    const name = newApName.trim()
    if (!name) return
    const id = `ap_${Date.now()}`
    setAps((current) => [...current, { id, name }])
    setAllData((current) => ({ ...current, [id]: emptyData() }))
    setSelectedApId(id)
    setNewApName('')
    setNotice(`${name} added`)
  }

  function removeAp(id) {
    const ap = aps.find((item) => item.id === id)
    if (!ap || !window.confirm(`Remove ${ap.name} and all its audit data?`)) return
    const remaining = aps.filter((item) => item.id !== id)
    setAps(remaining)
    setAllData((current) => { const next = { ...current }; delete next[id]; return next })
    if (selectedApId === id) setSelectedApId(remaining[0]?.id ?? null)
  }

  function clearCurrent() {
    if (!currentAp || !window.confirm(`Clear all data for ${currentAp.name}?`)) return
    updateCurrent(emptyData())
    setNotice('Current audit cleared')
  }

  return (
    <div className="app-shell">
      <aside className="sidebar no-print">
        <div className="brand">
          <div className="brand-mark"><Icon>⌁</Icon></div>
          <div><h1>AP Audit</h1><span>UNIFI WORKSPACE</span></div>
        </div>
        <div className="sidebar-date"><span className="status-dot" /> {dateLabel}</div>
        <div className="location-heading"><span>ACCESS POINTS</span><button className="icon-button" onClick={() => setShowSettings(true)} title="Manage access points">⚙</button></div>
        <nav className="locations" aria-label="Access points">
          {aps.map((ap) => (
            <button key={ap.id} className={`location ${selectedApId === ap.id ? 'selected' : ''}`} onClick={() => setSelectedApId(ap.id)}>
              <span className={`ap-icon ${hasData(ap.id) ? 'complete' : ''}`}>⌁</span>
              <span className="location-name">{ap.name}</span>
              {hasData(ap.id) && <span className="check">✓</span>}
            </button>
          ))}
          {!aps.length && <div className="sidebar-empty">No access points yet.</div>}
        </nav>
        <div className="sidebar-footer">
          <div className="progress-label"><span>DAILY PROGRESS</span><strong>{completedCount}/{aps.length}</strong></div>
          <div className="progress-track"><div style={{ width: aps.length ? `${(completedCount / aps.length) * 100}%` : '0%' }} /></div>
          <button className="export-button" onClick={() => window.print()}><Icon>⇩</Icon> Export report</button>
          <p className="local-note">Saved locally in this browser</p>
        </div>
      </aside>

      <main className="workspace">
        <header className="topbar">
          <div><span className="eyebrow">DAILY AUDIT / {dateLabel.toUpperCase()}</span><h2>{currentAp?.name ?? 'No access point selected'}</h2></div>
          {currentAp && <button className="clear-button" onClick={clearCurrent}>⌫ <span>Clear data</span></button>}
        </header>

        {currentAp ? (
          <div className="content">
            <section className="intro-row"><div><h3>Audit input</h3><p>Paste today’s UniFi client data and attach a speed test screenshot.</p></div><span className={`completion-pill ${hasData(currentAp.id) ? 'done' : ''}`}>{hasData(currentAp.id) ? '✓ IN PROGRESS' : '○ NOT STARTED'}</span></section>
            <div className="input-grid">
              <div className="input-column">
                <section className="card data-card">
                  <div className="card-heading"><div className="card-icon blue">▣</div><div><h4>My device data</h4><p>Raw row from UniFi · tab separated</p></div></div>
                  <textarea value={currentData.myDeviceRaw} onChange={(event) => updateCurrent({ myDeviceRaw: event.target.value })} placeholder="DESKTOP-2OIAP1F…    192.168.1.20    Excellent…" spellCheck="false" />
                  <div className="field-hint">One row only · data stays in your browser</div>
                </section>
                <section className="card data-card clients-card">
                  <div className="card-heading"><div className="card-icon violet">♧</div><div><h4>Other connected clients</h4><p>Multiple rows from UniFi · tab separated</p></div></div>
                  <textarea value={currentData.otherClientsRaw} onChange={(event) => updateCurrent({ otherClientsRaw: event.target.value })} placeholder={'Client-01…    192.168.1.21    Good…\nClient-02…    192.168.1.22    Fair…'} spellCheck="false" />
                  <div className="field-hint">Paste as many rows as needed</div>
                </section>
              </div>
              <section className="card screenshot-card">
                <div className="card-heading"><div className="card-icon green">◉</div><div><h4>Ookla speed test</h4><p>Screenshot evidence</p></div></div>
                <ImageDropzone image={currentData.speedtestImg} onImage={(speedtestImg) => updateCurrent({ speedtestImg })} fileInputRef={fileInputRef} />
                <div className="field-hint"><kbd>CTRL</kbd> + <kbd>V</kbd> to paste an image from clipboard</div>
              </section>
            </div>

            <section className="preview-section">
              <div className="section-heading"><div><span className="eyebrow">LIVE OUTPUT</span><h3>Spreadsheet preview</h3></div><span className="format-label">TSV / UNIFI EXPORT FORMAT</span></div>
              <div className="preview-card">
                <div className="preview-head"><span>DEVICE & CLIENT HEALTH</span><span>{myDeviceRows.length + clientRows.length} row{myDeviceRows.length + clientRows.length === 1 ? '' : 's'}</span></div>
                <div className="preview-body">
                  <AuditTable rows={myDeviceRows.slice(0, 1)} emptyLabel="No device data entered yet" />
                  <div className="preview-lower"><div className="preview-image">{currentData.speedtestImg ? <img src={currentData.speedtestImg} alt="Speed test preview" /> : <><span>▧</span><small>No screenshot</small></>}</div><div className="client-table"><AuditTable rows={clientRows} emptyLabel="No other client data entered yet" /></div></div>
                </div>
              </div>
              <p className="preview-note"><span>i</span> Values are color-coded automatically: signal strength, connection quality, and throughput.</p>
            </section>
          </div>
        ) : (
          <div className="no-ap"><span>⌁</span><h3>No access points available</h3><p>Use the settings button to add your first location.</p><button className="primary-button" onClick={() => setShowSettings(true)}>Manage access points</button></div>
        )}
      </main>

      {showSettings && <div className="modal-backdrop no-print" onMouseDown={(event) => event.target === event.currentTarget && setShowSettings(false)}><section className="modal" role="dialog" aria-modal="true" aria-labelledby="settings-title"><div className="modal-title"><div><span className="eyebrow">WORKSPACE SETTINGS</span><h3 id="settings-title">Manage access points</h3></div><button className="close-button" onClick={() => setShowSettings(false)}>×</button></div><form className="add-form" onSubmit={addAp}><input value={newApName} onChange={(event) => setNewApName(event.target.value)} placeholder="e.g. LT1 - Lobby" autoFocus /><button className="primary-button" type="submit">Add AP</button></form><div className="modal-list">{aps.map((ap) => <div className="modal-item" key={ap.id}><span><i className={hasData(ap.id) ? 'green-dot' : ''} />{ap.name}</span><button onClick={() => removeAp(ap.id)} aria-label={`Remove ${ap.name}`}>Remove</button></div>)}{!aps.length && <p className="modal-empty">No locations found.</p>}</div><button className="modal-done" onClick={() => setShowSettings(false)}>Done</button></section></div>}
      {notice && <div className="toast">✓ {notice}</div>}

      <PrintReport aps={aps} allData={allData} dateLabel={dateLabel} />
    </div>
  )
}

function PrintReport({ aps, allData, dateLabel }) {
  return <div className="print-report"><h1>Daily Wi-Fi AP Audit Report</h1><p>{dateLabel}</p>{aps.map((ap) => { const data = allData[ap.id] ?? emptyData(); return <section key={ap.id}><h2>{ap.name}</h2>{!data.myDeviceRaw?.trim() && !data.otherClientsRaw?.trim() && !data.speedtestImg ? <em>No data recorded for this location.</em> : <><AuditTable rows={parseRows(data.myDeviceRaw).slice(0, 1)} emptyLabel="No device data" /><div className="print-lower"><div className="print-image">{data.speedtestImg ? <img src={data.speedtestImg} alt="Speed test" /> : 'No speed test image'}</div><AuditTable rows={parseRows(data.otherClientsRaw)} emptyLabel="No other clients" /></div></>}</section>})}</div>
}

createRoot(document.getElementById('root')).render(<StrictMode><App /></StrictMode>)
