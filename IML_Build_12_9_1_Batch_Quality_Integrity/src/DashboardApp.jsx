import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2,
  AlertTriangle, Clock3, X, BarChart3, Download, Trash2, Save, Target,
  Gauge, CalendarCheck, BellRing, TrendingUp, FileSpreadsheet, ShieldCheck, RefreshCw, ClipboardList, Activity, LogOut, UserCircle, Cloud, WifiOff, ArrowLeft, HeartPulse, Printer, PanelRightClose, PanelRightOpen, Maximize2, Minimize2, Home, ChevronLeft, Settings2
} from 'lucide-react'
import { loadCloudDatasetOnce, getCloudDatasetMeta, uploadCloudDataset, uploadCloudDatasetIncremental, deleteAllCloudDatasets, getCloudHealth } from './cloudData'
import { supabase } from './supabase'
import { buildResourceRows } from './resourceEngine'
import { productionMappingKey, stationFamily } from './mappingEngine'
import './styles.css'

const LEGACY_DAILY_TARGETS = {
  '1519': 80000, '1521': 60000, '1523': 40000, '1524': 6000,
  '1525': 130000, '1528': 80000, '1540': 18000, '1541': 210000,
  '1542': 60000, '1543': 60000,
}
const FACILITY_ALIASES = {
  '1519': ['1519', '19', '19-F-01', '19-F-02'],
  '1521': ['1521', '21'],
  '1523': ['1523', '23'],
  '1524': ['1524', '24'],
  '1525': ['1525', '25'],
  '1528': ['1528', '28'],
  '1540': ['1540', '40'],
  '1541': ['1541', '41'],
  '1542': ['1542', '42-P-01', 'T42A'],
  '1543': ['1543', '43', '43-P-A', '43-P-B'],
  '1142': ['1142'],
  '1123': ['1123'],
}
const PRIMARY_FACILITIES = ['1519', '1541', '1540', '1525', '1523', '1521', '1528', '1524', '1542', '1543', '1142', '1123']
const DEFAULT_FACILITIES = PRIMARY_FACILITIES
const RESOURCE_LABELS = {
  '1542|LQ-P-1': { station: 'P-02', line: '1 ליטר' },
  '1542|LQ-P-5': { station: 'P-03', line: '5 ליטר' },
  '1542|LQ-P-10': { station: 'P-04', line: '10/20 ליטר' },
}
const STORAGE_KEY = 'iml-control-center-sprint7'
const DB_NAME = 'iml-control-center-db'
const DB_STORE = 'dashboard-state'
const DB_KEY = 'sprint1182-build2-batch-material'
const TARGET_FILE_KEY = 'latest-monthly-target-workbook'
const BUILD_LABEL = 'Sprint 11.9.0 Trial 12 — Unified Excel Upload'
const isoDate = value => {
  if (!value) return ''

  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return ''

  const pad = number => String(number).padStart(2, '0')

  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}
const MAPPING_STORAGE_KEY = 'iml-product-mappings-sprint1181'
const MAPPING_TIMELINE_KEY = 'iml-product-mapping-timeline-sprint1181'
const readLocalJson = (key, fallback = []) => {
  try { return JSON.parse(localStorage.getItem(key) || '') || fallback } catch { return fallback }
}

const initialToDate = () => isoDate(new Date())
const initialFromDate = () => { const date = new Date(); date.setDate(date.getDate() - 6); return isoDate(date) }

const openDashboardDb = () => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, 1)
  request.onupgradeneeded = () => {
    const db = request.result
    if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE)
  }
  request.onsuccess = () => resolve(request.result)
  request.onerror = () => reject(request.error)
})
const idbGet = async () => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly')
    const req = tx.objectStore(DB_STORE).get(DB_KEY)
    req.onsuccess = () => resolve(req.result || null)
    req.onerror = () => reject(req.error)
    tx.oncomplete = () => db.close()
  })
}
const idbSet = async (value) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).put(value, DB_KEY)
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}
const idbGetKey = async (key) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly')
    const req = tx.objectStore(DB_STORE).get(key)
    req.onsuccess = () => resolve(req.result || null)
    req.onerror = () => reject(req.error)
    tx.oncomplete = () => db.close()
  })
}
const idbSetKey = async (key, value) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).put(value, key)
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}
const idbClear = async () => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).delete(DB_KEY)
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}

const normalize = (v) => String(v ?? '').trim()
const batchMaterialKey = (batch, material) => {
  const b = normalize(batch)
  const m = normalize(material)
  return b && m ? `${b}|${m}` : ''
}
const normalizeRouting = (v) => normalize(v).toUpperCase()
const resourceMeta = (facility, routingGroup) => RESOURCE_LABELS[`${facility}|${normalizeRouting(routingGroup)}`] || {}
const planningName = (row) => row.resource || (row.routingGroup ? `מתקן ${row.facility} · ${row.station || row.routingGroup}${row.lineName ? ` · ${row.lineName}` : ''}` : `מתקן ${row.facility}`)
const normKey = (v) => normalize(v).toLowerCase().replace(/[\s_\-./()]+/g, '')
const num = (v) => {
  const n = Number(String(v ?? '').replace(/,/g, '').replace(/\s/g, ''))
  return Number.isFinite(n) ? n : 0
}
const excelDate = (v) => {
  if (v === '' || v === null || v === undefined) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.getFullYear() <= 1900 ? null : new Date(v)
  if (typeof v === 'number') {
    if (v < 1) return null
    const d = XLSX.SSF.parse_date_code(v)
    return d ? new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0) : null
  }
  const text = normalize(v)
  const match = text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/)
  if (match) {
    let year = Number(match[3]); if (year < 100) year += 2000
    return new Date(year, Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0))
  }
  const d = new Date(text)
  return Number.isNaN(d.getTime()) || d.getFullYear() <= 1900 ? null : d
}
const localDateTimeString = date => {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return ''

  const pad = value => String(value).padStart(2, '0')

  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
  ].join('-') + 'T' + [
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
  ].join(':')
}
const localDateOnlyString = value => {
  const date = excelDate(value)
  if (!date) return ''
  const pad = number => String(number).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}

const productionDateFromDay = day => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(day || ''))) return null
  const [year, month, date] = day.split('-').map(Number)
  return new Date(year, month - 1, date, 12, 0, 0, 0)
}

const excelTime = (v) => {
  if (v === '' || v === null || v === undefined) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return { h: v.getHours(), m: v.getMinutes(), s: v.getSeconds() }
  if (typeof v === 'number') {
    const seconds = Math.round((v % 1) * 86400) % 86400
    return { h: Math.floor(seconds / 3600), m: Math.floor((seconds % 3600) / 60), s: seconds % 60 }
  }
  const match = normalize(v).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/)
  return match ? { h: Number(match[1]), m: Number(match[2]), s: Number(match[3] || 0) } : null
}
const combineExcelDateTime = (dateValue, timeValue, fallbackValue = '') => {
  const date = excelDate(dateValue) || excelDate(fallbackValue)
  if (!date) return null
  const time = excelTime(timeValue)
  const result = new Date(date)
  if (time) result.setHours(time.h, time.m, time.s, 0)
  return result
}

const shiftInfo = (date) => {
  if (!date) return { key: 'unknown', label: 'ללא שעה', order: 9, workDate: '' }
  const d = new Date(date)
  const hour = d.getHours()
  let key = 'morning', label = 'בוקר', order = 0
  if (hour >= 15 && hour < 23) { key = 'evening'; label = 'ערב'; order = 1 }
  else if (hour >= 23 || hour < 7) { key = 'night'; label = 'לילה'; order = 2 }
  const work = new Date(d)
  if (hour < 7) work.setDate(work.getDate() - 1)
  return { key, label, order, workDate: `${work.getFullYear()}-${String(work.getMonth()+1).padStart(2,'0')}-${String(work.getDate()).padStart(2,'0')}` }
}
const iso = (d) => {
  if (!d) return ''
  const value = new Date(d)
  if (Number.isNaN(value.getTime()) || value.getFullYear() <= 1900) return ''
  return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`
}
const monthKey = (d) => iso(d).slice(0, 7)
const fmt = (n) => Math.round(n || 0).toLocaleString('he-IL')
const pctFmt = (n) => `${Math.round(n || 0)}%`
const getField = (row, names) => {
  const map = new Map(Object.keys(row || {}).map(k => [normKey(k), row[k]]))
  for (const name of names) {
    const value = map.get(normKey(name))
    if (value !== undefined) return value
  }
  return ''
}
const canonicalFacility = (value) => {
  const clean = normalize(value).toUpperCase()
  for (const [id, aliases] of Object.entries(FACILITY_ALIASES)) {
    if (aliases.some(alias => normalize(alias).toUpperCase() === clean)) return id
  }
  const digits = clean.match(/15\d{2}/)?.[0]
  return digits || clean
}
const matchesDateRange = (date, from, to) => {
  const value = iso(date)
  if (!from && !to) return true
  return Boolean(value) && (!from || value >= from) && (!to || value <= to)
}
const isWorkday = (date) => ![5, 6].includes(date.getDay()) // Israel default: Sunday–Thursday
const daysInMonth = (key) => {
  const [year, month] = key.split('-').map(Number)
  return new Date(year, month, 0).getDate()
}
const workdayCount = (key, startDay = 1, endDay = daysInMonth(key)) => {
  const [year, month] = key.split('-').map(Number)
  let count = 0
  for (let day = startDay; day <= endDay; day++) if (isWorkday(new Date(year, month - 1, day))) count++
  return count
}
const parseMonth = (value, fallbackDate = null) => {
  if (value instanceof Date) return monthKey(value)
  if (typeof value === 'number') {
    const d = excelDate(value); if (d) return monthKey(d)
  }
  const text = normalize(value)
  let m = text.match(/(20\d{2})[.\/-](\d{1,2})/)
  if (m) return `${m[1]}-${String(Number(m[2])).padStart(2, '0')}`
  m = text.match(/(\d{1,2})[.\/-](20\d{2})/)
  if (m) return `${m[2]}-${String(Number(m[1])).padStart(2, '0')}`
  return fallbackDate ? monthKey(fallbackDate) : ''
}

const TARGET_MONTHS = { jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4, may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8, sep:9, sept:9, september:9, oct:10, october:10, nov:11, november:11, dec:12, december:12 }
const targetMonthFromTitle = (value, fallbackDate = new Date()) => {
  const text = normalize(value)
  const english = text.match(/(?:production\s+report\s+)?([a-z]{3,9})\s+(20\d{2})/i)
  if (english) { const month = TARGET_MONTHS[english[1].toLowerCase()]; if (month) return `${english[2]}-${String(month).padStart(2,'0')}` }
  return parseMonth(text) || monthKey(fallbackDate)
}
const targetFacilityIds = (resource) => {
  const text = normalize(resource)
  const upperText = text.toUpperCase()
  // Approved business mappings. EC(23) is intentionally mapped to station 1523.
  if (/^EC\s*\(23\)/i.test(text)) return ['1523']
  if (/^EC\s*\(25\)/i.test(text)) return ['1525']
  // Bromacil is produced under station 1540. Product membership comes from the DATA sheet.
  if (/^BROMACIL\s*\(25\s*,\s*40\)/i.test(text)) return ['1540']
  if (/^SHAKED\s+ISO\s+42$/i.test(text)) return ['1142']
  if (/^SHAKED\s+ISO\s+23$/i.test(text)) return ['1123']
  if (/^LQ\s*43\b/i.test(text)) return ['1543']
  if (/^SC\s*\(28\)/i.test(text)) return ['1528']
  if (/WG\s*SMALL\s+PACKS?\s*\(19\)/i.test(text) || /^WG\s*\(19\)/i.test(text)) return ['1519']
  if (/LQ\s*(1|5|10\s*\/\s*20)\s*(LT|L)/i.test(text)) return ['1542']

  const raw = (text.match(/\(([^)]+)\)/)?.[1] || (text.match(/\b(19|23|24|25|28|40|41|42|43)\b/g) || []).join(','))
  const plantIds = [...new Set(String(raw).match(/19|23|24|25|28|40|41|42|43/g) || [])]
  const map = { '19':'1519','23':'1523','24':'1524','25':'1525','28':'1528','40':'1540','41':'1541','42':'1542','43':'1543' }
  return plantIds.map(id => map[id]).filter(Boolean)
}
const targetDescriptionTokens = (resource) => {
  const text = normalize(resource).toUpperCase()
  if (/LQ\s*1\s*(LT|L)\b/.test(text)) return ['1L','1 L','1LT','1 LT']
  if (/LQ\s*5\s*(LT|L)\b/.test(text)) return ['5L','5 L','5LT','5 LT']
  if (/LQ\s*10\s*\/\s*20/.test(text)) return ['10L','10 L','20L','20 L','10/20']
  if (/LQ\s*43/.test(text)) return []
  const clean = text.replace(/\([^)]*\)/g,'').replace(/\bSMALL\s+PACKS?\b/g,'SMALL PACK').trim()
  const generic = new Set(['EC','SC','WG','CS','LQ','24F'])
  return !clean || generic.has(clean) ? [] : [clean]
}
const APPROVED_TARGET_RESOURCES = new Set([
  'EC (23)','SHAKED ISO 42','SHAKED ISO 23','LQ 1LT (42)','LQ 5 LT (42)','LQ 10/20 LT (42)','LQ 43','SC (28)','WG (19)','WG SMALL PACKS (19)',
  '24F128','24F','EC (25)','DIURON (40)','TOLUREX (40)','CS (25,40)','BROMACIL (25,40)','GALIGAN (25,40)',
  'PROPA PREMIX (25,40)','FLUOROCHLORIDON (25,40)','SAFLUFENACIL TECH (25,40)','METAZACHLOR (41)',
  'ATRALONE (41)','NANA (41)','D. DAMASCONE (41)'
])
const isApprovedTargetResource = value => APPROVED_TARGET_RESOURCES.has(normalize(value).toUpperCase())

const parseTargetNumber = (value) => {
  if (value === null || value === undefined || value === '' || /^\s*-+\s*$/.test(String(value)) || /DIV\/0/i.test(String(value))) return 0
  const text = String(value).trim(); const negative = /^\(.*\)$/.test(text)
  const n = Number(text.replace(/[(),%\s]/g,'').replace(/,/g,''))
  return Number.isFinite(n) ? (negative ? -n : n) : 0
}

// Sprint 11.9.0 Trial 4 — normalize legacy monthly targets.
// SUM targets are expressed in thousands (t / m³), while production rows are L / kg.
// New uploads are already multiplied by 1000; old cloud/cache rows are normalized here once in memory.
const normalizeStoredTargetRow = row => {
  const scaleLegacy = value => {
    const n = Number(value) || 0
    return n > 0 && n < 10000 ? n * 1000 : n
  }
  return { ...row, target:scaleLegacy(row?.target), capacity:scaleLegacy(row?.capacity) }
}
const isLegacyCombinedTarget = value => {
  const text = normalize(value).toUpperCase().replace(/\s+/g,' ')
  return /EC\s*\(\s*(23\s*,\s*25|25\s*,\s*23)\s*\)/.test(text) ||
    /SHAKED\s+ISO\s*\(\s*(42\s*\+\s*23|42\s*\+\s*43)\s*\)/.test(text)
}
const normalizeStoredTargets = rows => (rows || [])
  .map(normalizeStoredTargetRow)
  .filter(row => !isLegacyCombinedTarget(row?.resource || row?.lineName || ''))


// A quality record is uniquely identified by the SAP triplet
// Inspection Lot + Sample # + Operation Activity + Characteristic.
// Sample # and Operation Activity used to be missing from this key, so two
// genuine measurements of the same characteristic that happened to return the
// same value collapsed into one row and the second was discarded at ingest.
const qualityRowKey = (row) => [
  normalize(row?.inspectionLot), normalize(row?.sample), normalize(row?.activity),
  normalize(row?.batch), normalize(row?.material),
  normalize(row?.characteristic), row?.date ? new Date(row.date).toISOString() : '',
  normalize(row?.value), normalize(row?.qualitative), normalize(row?.status)
].join('|')

// SAP inspection lots: 3xxxxxxxxx = in-process check during production,
// 4xxxxxxxxx = final / goods-receipt inspection. Only final lots ever receive a
// usage decision, so the batch verdict must be read off those lots alone.
const LOT_KIND_FINAL = 'final'
const LOT_KIND_INPROCESS = 'inprocess'
const lotKind = (lot) => {
  const value = normalize(lot)
  if (!value) return ''
  if (value.startsWith('4')) return LOT_KIND_FINAL
  if (value.startsWith('3') || value.startsWith('1')) return LOT_KIND_INPROCESS
  return ''
}
const lotKindLabel = (kind) => kind === LOT_KIND_FINAL ? 'סופי'
  : kind === LOT_KIND_INPROCESS ? 'תוך-תהליכי' : '—'

// A usage decision is "open" only while no UD code has been issued.
const isOpenDecision = (row) => {
  const code = normalize(row?.udCode)
  return !code || code === 'N/A' || code === '-'
}

async function filterNewQualityRows(existingRows, incomingRows, onProgress) {
  const keys = new Set()
  const total = existingRows.length + incomingRows.length
  let done = 0
  for (let i = 0; i < existingRows.length; i += 4000) {
    existingRows.slice(i, i + 4000).forEach(row => keys.add(qualityRowKey(row)))
    done += Math.min(4000, existingRows.length - i)
    onProgress?.(done, total)
    await new Promise(resolve => setTimeout(resolve, 0))
  }
  const fresh = []
  for (let i = 0; i < incomingRows.length; i += 4000) {
    incomingRows.slice(i, i + 4000).forEach(row => {
      const key = qualityRowKey(row)
      if (!keys.has(key)) { keys.add(key); fresh.push(row) }
    })
    done += Math.min(4000, incomingRows.length - i)
    onProgress?.(done, total)
    await new Promise(resolve => setTimeout(resolve, 0))
  }
  return fresh
}

async function readTargetWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type:'array', cellDates:true, dense:true })
  const output = []

  // The DATA sheet is the product master. Build Family -> Item Code dynamically.
  // Packaging facilities 1542 and 1519 intentionally keep their existing routing logic.
  const materialsByFamily = new Map()
  for (const sheetName of wb.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:true })
    if (!matrix.length) continue
    const headerIndex = matrix.findIndex(row => {
      const keys = row.map(normKey)
      return keys.includes(normKey('Item Code')) && keys.includes(normKey('Family'))
    })
    if (headerIndex < 0) continue
    const header = matrix[headerIndex].map(normalize)
    const itemCodeIndex = header.findIndex(value => normKey(value) === normKey('Item Code'))
    const familyIndex = header.findIndex(value => normKey(value) === normKey('Family'))
    if (itemCodeIndex < 0 || familyIndex < 0) continue
    matrix.slice(headerIndex + 1).forEach(row => {
      const material = normalize(row[itemCodeIndex])
      const family = normalize(row[familyIndex]).toUpperCase()
      if (!material || !family || family === '#N/A') return
      const list = materialsByFamily.get(family) || new Set()
      list.add(material)
      materialsByFamily.set(family, list)
    })
  }

  for (const sheetName of wb.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:true })
    if (!matrix.length) continue
    const titleText = matrix.slice(0,3).flat().map(normalize).join(' ')
    const fallbackMonth = targetMonthFromTitle(`${titleText} ${file.name}`, new Date())
    const headerIndex = matrix.findIndex(row => { const keys=row.map(normKey); return keys.some(k=>k==='capacity') && keys.some(k=>k==='plan') })
    if (headerIndex < 0) continue
    const header = matrix[headerIndex].map((v,i)=>normalize(v)||(i===0?'Resource':`Column_${i+1}`))
    for (const row of matrix.slice(headerIndex+1)) {
      const resource=normalize(row[0]); if(!resource||/^total$/i.test(resource)) continue
      const familyName = resource.replace(/\([^)]*\)/g, '').trim().toUpperCase()
      const protectedPackaging = /LQ\s*(1|5|10\s*\/\s*20)\s*(LT|L)/i.test(resource) || /WG/i.test(resource)
      const familyMaterials = protectedPackaging ? [] : [...(materialsByFamily.get(familyName) || [])]
      const item={__sheet:sheetName,Resource:resource,Month:fallbackMonth,__family:familyName,__materials:familyMaterials}
      header.forEach((key,i)=>{item[key]=row[i]??''}); output.push(item)
    }
  }
  return output
}

// Big SAP exports (280K+ rows) sit right at the browser's JS heap ceiling.
// Parsing this file costs ~1.1GB on its own, so every avoidable copy matters:
// the old version built the row objects, then cloned every one of them again to
// stamp __sheet, then loadFiles built a third array on top — roughly 1.7GB peak,
// which is where the tab died. Here each worksheet is released as soon as it has
// been converted, and __sheet is written in place instead of via a spread.
async function readWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: false, dense: true })
  const rows = []
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName]
    if (!sheet) continue
    const sheetRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true })
    delete wb.Sheets[sheetName]
    for (let i = 0; i < sheetRows.length; i++) {
      const row = sheetRows[i]
      row.__sheet = sheetName
      rows.push(row)
      sheetRows[i] = null
    }
  }
  wb.SheetNames = []
  return rows
}

// Maps rows into their compact shape while releasing each source row, so the raw
// array and the compact array are never both fully resident.
function mapRowsInPlace(rows, mapFn) {
  const out = []
  for (let i = 0; i < rows.length; i++) {
    const mapped = mapFn(rows[i], i)
    rows[i] = null
    if (mapped) out.push(mapped)
  }
  rows.length = 0
  return out
}
function classifyFile(rows) {
  // Some workbooks start with a title/summary sheet. Inspect headers across
  // a sample of rows instead of relying only on the first row.
  const keys = [...new Set(rows.slice(0, 250).flatMap(row => Object.keys(row || {}).map(normKey)))]
  const hasAny = (...terms) => terms.some(term => keys.some(k => k.includes(normKey(term))))
  // Whole-header match. 'plan' used to match by substring inside
  // "Production Plant ID", which every SAP export carries — that alone sent the
  // usage-decision file down the targets branch and made it fail to load.
  const hasHeader = (...terms) => terms.some(term => {
    const t = normKey(term)
    return keys.some(k => k === t || k.startsWith(t))
  })

  const productionLike = hasAny('actual finish time', 'delivered quantity', 'confirmed yield quantity') &&
    hasAny('storage location', 'order', 'batch')
  if (productionLike) return 'production'

  const deviationLike = hasAny('rejected characteristics', 'qa status', 'ud remarks', 'restricted - recycling')
  if (deviationLike) return 'deviations'

  const qualityLike = hasAny(
    'inspection lot', 'inspection lot #', 'inspection lot storage location',
    'master insp characteristic', 'master insp charactristic',
    'result status', 'qa approval', 'start date of inspection', 'end date of inspection'
  )
  if (qualityLike) return 'quality'

  // Targets is checked last: it is the loosest signature, so it must not be able
  // to claim a file that any of the SAP exports above already matched.
  const targetLike = hasHeader('monthly target', 'monthly plan', 'יעד חודשי', 'תוכנית חודשית', 'target', 'plan', 'capacity') &&
    hasHeader('facility', 'מתקן', 'resource', 'משאב')
  if (targetLike) return 'targets'

  return 'unknown'
}


const displayDatasetName = (kind, month = '') => {
  if (kind === 'production') return 'קובץ כמויות יצור ואריזה'
  if (kind === 'quality') return 'קובץ איכות תוצאות'
  if (kind === 'deviations') return 'קובץ איכות החלטת שימוש'
  if (kind === 'targets') {
    const mm = String(month || '').match(/-(\d{2})$/)?.[1] || String(new Date().getMonth()+1).padStart(2,'0')
    return `יעדים לחודש ${mm}`
  }
  return 'קובץ נתונים'
}

const normalizeDatasetMeta = (kind, meta, targetMonth = '') => {
  if (!meta) return null
  const month = kind === 'targets'
    ? (String(targetMonth || meta.month || '').match(/\d{4}-\d{2}/)?.[0] || '')
    : ''
  return {
    ...meta,
    fileName: displayDatasetName(kind, month),
    originalFileName: meta.originalFileName || meta.fileName || '',
  }
}



// Global table tools — generic search/filter/sort controls for every data table.
// Implemented at the DOM layer so existing calculation logic and row renderers stay untouched.
const useUniversalTableTools = () => {
  useEffect(() => {
    const normalizeTableText = value => String(value ?? '').trim().toLowerCase()

    const cellSortValue = cell => {
      const raw = String(cell?.innerText || '').trim()
      const numeric = Number(raw.replace(/,/g, '').replace(/[^0-9.\-]/g, ''))
      if (raw && Number.isFinite(numeric) && /\d/.test(raw)) return { numeric:true, value:numeric }
      const dateMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
      if (dateMatch) return { numeric:true, value:new Date(raw).getTime() }
      return { numeric:false, value:raw.toLocaleLowerCase('he-IL') }
    }

    const applyView = table => {
      if (!table?.tBodies?.[0]) return
      const toolbar = table.parentElement?.querySelector(':scope > .smart-table-toolbar')
      if (!toolbar) return
      const globalQuery = normalizeTableText(toolbar.querySelector('[data-smart-global]')?.value)
      const columnIndex = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
      const columnQuery = normalizeTableText(toolbar.querySelector('[data-smart-value]')?.value)
      let visible = 0
      const rows = [...table.tBodies[0].rows]
      rows.forEach(row => {
        if (row.classList.contains('smart-empty-row')) return
        const cells = [...row.cells]
        const allText = normalizeTableText(cells.map(cell => cell.innerText).join(' '))
        const selectedText = columnIndex >= 0 ? normalizeTableText(cells[columnIndex]?.innerText) : allText
        const matchesGlobal = !globalQuery || allText.includes(globalQuery)
        const matchesColumn = !columnQuery || selectedText.includes(columnQuery)
        const show = matchesGlobal && matchesColumn
        row.style.display = show ? '' : 'none'
        if (show) visible += 1
      })
      const count = toolbar.querySelector('[data-smart-count]')
      if (count) count.textContent = `${visible} רשומות`
      const sumBadge = toolbar.querySelector('[data-smart-sum]')
      const sumColumn = Number(table.dataset.smartSumColumn ?? -1)
      if (sumBadge && sumColumn >= 0) {
        const total = rows.reduce((sum,row) => {
          if (row.style.display === 'none' || row.classList.contains('smart-empty-row')) return sum
          const raw = String(row.cells[sumColumn]?.innerText || '').replace(/,/g,'').replace(/[^0-9.\-]/g,'')
          const value = Number(raw)
          return sum + (Number.isFinite(value) ? value : 0)
        }, 0)
        sumBadge.textContent = `סה״כ כמות: ${new Intl.NumberFormat('he-IL',{maximumFractionDigits:2}).format(total)}`
      }
    }

    const sortTable = (table, columnIndex, direction) => {
      const body = table?.tBodies?.[0]
      if (!body || columnIndex < 0 || !direction) return
      const rows = [...body.rows].filter(row => row.cells.length > columnIndex)
      rows.sort((a,b) => {
        const av = cellSortValue(a.cells[columnIndex])
        const bv = cellSortValue(b.cells[columnIndex])
        let result
        if (av.numeric && bv.numeric) result = av.value - bv.value
        else result = String(av.value).localeCompare(String(bv.value), 'he', { numeric:true, sensitivity:'base' })
        return direction === 'desc' ? -result : result
      })
      rows.forEach(row => body.appendChild(row))
      applyView(table)
    }

    const enhanceTable = table => {
      if (!table || table.dataset.smartTableReady === '1' || table.dataset.noSmartTable === '1') return
      const headerCells = [...(table.tHead?.rows?.[0]?.cells || [])]
      if (!headerCells.length || !table.tBodies?.length) return
      table.dataset.smartTableReady = '1'
      table.classList.add('smart-table-enabled')
      const wrap = table.parentElement
      if (!wrap) return

      const toolbar = document.createElement('div')
      toolbar.className = 'smart-table-toolbar'
      const options = headerCells.map((cell,index) => `<option value="${index}">${String(cell.innerText || `עמודה ${index+1}`).trim()}</option>`).join('')
      toolbar.innerHTML = `
        <div class="smart-table-search"><span>⌕</span><input data-smart-global type="search" placeholder="חיפוש בכל הטבלה..." autocomplete="off" /></div>
        <label class="smart-table-column"><span>עמודה</span><select data-smart-column><option value="-1">כל העמודות</option>${options}</select></label>
        <div class="smart-table-search smart-table-value"><span>≡</span><input data-smart-value type="search" placeholder="ערך לסינון, למשל A1" autocomplete="off" /></div>
        <label class="smart-table-sort"><span>מיון</span><select data-smart-sort><option value="">ללא</option><option value="asc">עולה ↑</option><option value="desc">יורד ↓</option></select></label>
        <button type="button" class="smart-table-clear">נקה סינון</button>
        ${table.dataset.smartSumColumn != null ? '<b data-smart-sum>סה״כ כמות: 0</b>' : ''}
        <b data-smart-count>${table.tBodies[0].rows.length} רשומות</b>
      `
      wrap.insertBefore(toolbar, table)

      const onFilter = () => applyView(table)
      toolbar.querySelector('[data-smart-global]')?.addEventListener('input', onFilter)
      toolbar.querySelector('[data-smart-column]')?.addEventListener('change', () => {
        applyView(table)
        const direction = toolbar.querySelector('[data-smart-sort]')?.value
        const index = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
        if (direction && index >= 0) sortTable(table,index,direction)
      })
      toolbar.querySelector('[data-smart-value]')?.addEventListener('input', onFilter)
      toolbar.querySelector('[data-smart-sort]')?.addEventListener('change', event => {
        const index = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
        if (index >= 0 && event.target.value) sortTable(table,index,event.target.value)
      })
      toolbar.querySelector('.smart-table-clear')?.addEventListener('click', () => {
        const global = toolbar.querySelector('[data-smart-global]')
        const value = toolbar.querySelector('[data-smart-value]')
        const column = toolbar.querySelector('[data-smart-column]')
        const sort = toolbar.querySelector('[data-smart-sort]')
        if (global) global.value = ''
        if (value) value.value = ''
        if (column) column.value = '-1'
        if (sort) sort.value = ''
        applyView(table)
      })

      headerCells.forEach((cell,index) => {
        if (cell.querySelector('button')) return
        cell.classList.add('smart-sortable-header')
        cell.title = 'לחץ למיון עולה / יורד'
        cell.addEventListener('click', event => {
          if (event.target.closest('input,select,button,a')) return
          const next = cell.dataset.sortDirection === 'asc' ? 'desc' : 'asc'
          headerCells.forEach(other => { if (other !== cell) delete other.dataset.sortDirection })
          cell.dataset.sortDirection = next
          sortTable(table,index,next)
        })
      })
      applyView(table)
    }

    const scan = () => document.querySelectorAll('.dashboard table').forEach(enhanceTable)
    let scheduled = false
    const scheduleScan = () => {
      if (scheduled) return
      scheduled = true
      requestAnimationFrame(() => {
        scheduled = false
        scan()
        document.querySelectorAll('.dashboard table[data-smart-table-ready="1"]').forEach(applyView)
      })
    }
    scan()
    const observer = new MutationObserver(scheduleScan)
    observer.observe(document.body, { childList:true, subtree:true })
    return () => observer.disconnect()
  }, [])
}

export default function DashboardApp({ currentUser, userRole = 'viewer', isGuest = false, onSignOut }) {
  useUniversalTableTools()

  const canManageData = ['admin', 'manager'].includes(userRole)
  const canDeleteData = userRole === 'admin'

  const [production, setProduction] = useState([])
  const [quality, setQuality] = useState([])
  const [deviations, setDeviations] = useState([])
  const [targets, setTargets] = useState([])
  const [status, setStatus] = useState('ממתין לטעינת קבצים')
  const [busy, setBusy] = useState(false)
  const [query, setQuery] = useState('')
  const [from, setFrom] = useState(initialFromDate)
  const [to, setTo] = useState(initialToDate)
  const [selectedFacilities, setSelectedFacilities] = useState([])
  const [activeTab, setActiveTab] = useState('production')
  const [productionSort, setProductionSort] = useState({ key:'date', direction:'desc' })
  const [facilityViewMode, setFacilityViewMode] = useState('relevant')
  const [showFacilityOverview, setShowFacilityOverview] = useState(true)
  const [showMonthlyForecast, setShowMonthlyForecast] = useState(true)
  const [simulatorOnlyIssues, setSimulatorOnlyIssues] = useState(true)
  const [manualMappings, setManualMappings] = useState(() => readLocalJson(MAPPING_STORAGE_KEY, []))
  const [mappingTimeline, setMappingTimeline] = useState(() => readLocalJson(MAPPING_TIMELINE_KEY, []))
  const [mappingDialog, setMappingDialog] = useState(null)
  const [mappingTargetKey, setMappingTargetKey] = useState('')
  const [mappingMessage, setMappingMessage] = useState('')
  const [planningMonth, setPlanningMonth] = useState('')
  const [additionalFacilities, setAdditionalFacilities] = useState([])
  const [facilityToAdd, setFacilityToAdd] = useState('')
  const [periodYear, setPeriodYear] = useState('')
  const [periodQuarter, setPeriodQuarter] = useState('')
  const [dataMeta, setDataMeta] = useState({ production:null, quality:null, deviations:null, targets:null })
  const [selectedBatch, setSelectedBatch] = useState('')
  const [selectedBatchMaterial, setSelectedBatchMaterial] = useState('')
  const [selectedResource, setSelectedResource] = useState(null)
  const [cloudState, setCloudState] = useState({ mode:'connecting', lastSync:null, message:'מתחבר למסד המשותף...', latencyMs:null, live:false })
  const [uploadProgress, setUploadProgress] = useState(null)
  const [perfStats, setPerformance] = useState({ startedAt:0, cache:'MISS', queries:0, rows:0, loadMs:0, phase:'אתחול' })
  const [sidebarCollapsed, setSidebarCollapsed] = useState(() => localStorage.getItem('iml-ui-sidebar-collapsed') === '1')
  const [managementMode, setManagementMode] = useState(() => localStorage.getItem('iml-ui-management-mode') === '1')
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(false)
  const [showHome, setShowHome] = useState(true)
  const [facilityPickerOpen, setFacilityPickerOpen] = useState(false)

  useEffect(() => { localStorage.setItem('iml-ui-sidebar-collapsed', sidebarCollapsed ? '1' : '0') }, [sidebarCollapsed])
  useEffect(() => { localStorage.setItem('iml-ui-management-mode', managementMode ? '1' : '0') }, [managementMode])

  useEffect(() => { localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify(manualMappings)) }, [manualMappings])
  useEffect(() => { localStorage.setItem(MAPPING_TIMELINE_KEY, JSON.stringify(mappingTimeline.slice(0,500))) }, [mappingTimeline])

  useEffect(() => {
    if (!supabase || !currentUser?.id) return
    let active = true
    supabase.from('iml_product_mappings').select('*').order('created_at', { ascending:false })
      .then(({ data, error }) => {
        if (!active || error || !data?.length) return
        setManualMappings(data.map(row => ({
          id:row.id, rowKey:row.row_key, material:row.material || '', description:row.description || '',
          family:row.facility_family, targetKey:row.target_key, targetResource:row.target_resource,
          status:row.status, active:row.active !== false, createdBy:row.created_by_email || '',
          createdAt:row.created_at, approvedAt:row.approved_at || null, cloud:true,
        })))
      })
    return () => { active = false }
  }, [currentUser?.id])

  useEffect(() => {
    let active = true
    const startedAt = performance.now()
    const kinds = ['production', 'quality', 'deviations', 'targets']

    const applyDataset = (kind, dataset) => {
      const rows = dataset?.rows || []
      if (kind === 'production') {
  setProduction(
    rows.map(row => {
      const sourceDate = row.finishDate || row.date

      return {
        ...row,
        date: sourceDate
          ? new Date(sourceDate)
          : null,
      }
    })
  )
}
      if (kind === 'quality') setQuality(rows.map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row))
      if (kind === 'deviations') setDeviations(rows)
      if (kind === 'targets') setTargets(normalizeStoredTargets(rows))
      setDataMeta(current => ({ ...current, [kind]:normalizeDatasetMeta(kind, dataset?.meta || null, kind === 'targets' ? (rows?.[0]?.month || planningMonth) : '') }))
      return rows.length
    }

    ;(async () => {
      let cached = null
      try {
        cached = await idbGet()
        if (active && cached) {
          setProduction(
  (cached.production || []).map(row => {
    const sourceDate = row.finishDate || row.date

    return {
      ...row,
      date: sourceDate
        ? sourceDate instanceof Date
          ? sourceDate
          : new Date(sourceDate)
        : null,
    }
  })
)
          setQuality(cached.quality || [])
          setDeviations(cached.deviations || [])
          setTargets(normalizeStoredTargets(cached.targets || []))
          setDataMeta({
            production: normalizeDatasetMeta('production', cached.dataMeta?.production),
            quality: normalizeDatasetMeta('quality', cached.dataMeta?.quality),
            deviations: normalizeDatasetMeta('deviations', cached.dataMeta?.deviations),
            targets: normalizeDatasetMeta('targets', cached.dataMeta?.targets, (cached.targets || [])[0]?.month || planningMonth),
          })
          setStatus('מוצג מטמון מקומי — בודק עדכונים מהשרת...')
          setCloudState({ mode:'connecting', lastSync:cached.savedAt || null, message:'הדשבורד זמין; בודק גרסאות חדשות ברקע...', latencyMs:null, live:false })
          setPerformance(current => ({ ...current, cache:'HIT', rows:(cached.production?.length||0)+(cached.quality?.length||0)+(cached.deviations?.length||0)+(cached.targets?.length||0), phase:'בדיקת גרסאות' }))
        }
      } catch (error) { console.warn('Cache restore failed', error) }

      try {
        const remoteMeta = {}
        await Promise.all(kinds.map(async kind => { remoteMeta[kind] = await getCloudDatasetMeta(kind) }))
        if (!active) return
        setPerformance(current => ({ ...current, queries:current.queries + 4 }))

        const changed = kind => {
          const remote = remoteMeta[kind]
          const local = cached?.dataMeta?.[kind]
          if (!remote) return !local
          const remoteId = remote.active_version_id || remote.updated_at || remote.loaded_at
          const localId = local?.versionId || local?.loadedAt
          return !cached || !local || String(remoteId || '') !== String(localId || '')
        }

        let loadedRows = 0
        if (changed('production')) {
          setStatus('טוען נתוני ייצור מעודכנים...')
          loadedRows += applyDataset('production', await loadCloudDatasetOnce('production'))
          setPerformance(current => ({ ...current, queries:current.queries + 1, phase:'הדשבורד זמין' }))
          await new Promise(resolve => setTimeout(resolve, 0))
        }

        for (const kind of ['targets', 'quality', 'deviations']) {
          if (!active) return
          if (!changed(kind)) continue
          setStatus(`טוען ${kind} ברקע...`)
          loadedRows += applyDataset(kind, await loadCloudDatasetOnce(kind))
          setPerformance(current => ({ ...current, queries:current.queries + 1, phase:`נטען ${kind}` }))
          await new Promise(resolve => setTimeout(resolve, 0))
        }

        const health = await getCloudHealth().catch(() => null)
        if (!active) return
        const elapsed = Math.round(performance.now() - startedAt)
        const lastSync = Object.values(remoteMeta).map(x => x?.loaded_at || x?.updated_at).filter(Boolean).sort().at(-1) || new Date().toISOString()
        setCloudState({ mode:'cloud', lastSync, message:'מחובר לענן — טעינה חכמה של 7 ימים כברירת מחדל', latencyMs:health?.latencyMs ?? null, live:true })
        setStatus(cached && loadedRows === 0 ? 'הנתונים במטמון מעודכנים — לא נדרשה הורדה מחדש' : 'הנתונים המעודכנים נטענו בהצלחה')
        setPerformance(current => ({ ...current, loadMs:elapsed, rows:current.rows + loadedRows, phase:'הושלם' }))
      } catch (cloudError) {
        console.warn('Smart cloud restore failed', cloudError)
        if (!active) return
        setCloudState({ mode:cached ? 'offline' : 'error', lastSync:cached?.savedAt || null, message:cached ? 'השרת אינו זמין — מוצג מטמון מקומי' : (cloudError?.message || 'טעינת הנתונים נכשלה'), latencyMs:null, live:false })
        setStatus(cached ? 'מוצג גיבוי מקומי; הסנכרון לענן נכשל' : 'לא נמצאו נתונים זמינים')
      }
    })()
    return () => { active = false }
  }, [])

  useEffect(() => {
    if (!supabase) return
    let refreshTimer
    const channel = supabase.channel('iml-data-sources-live')
      .on('postgres_changes', { event:'*', schema:'public', table:'iml_data_sources' }, payload => {
        clearTimeout(refreshTimer)
        refreshTimer = setTimeout(async () => {
          const kind = payload?.new?.kind || payload?.old?.kind
          if (!['production','quality','deviations','targets'].includes(kind)) return
          try {
            setStatus(`התקבל עדכון ${kind} — מסנכרן רק את הקובץ שהשתנה...`)
            const dataset = await loadCloudDatasetOnce(kind)
            const rows = dataset?.rows || []
            if (kind === 'production') {
  setProduction(
    rows.map(row => {
      const sourceDate = row.finishDate || row.date

      return {
        ...row,
        date: sourceDate
          ? new Date(sourceDate)
          : null,
      }
    })
  )
}
            if (kind === 'quality') setQuality(rows.map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row))
            if (kind === 'deviations') setDeviations(rows)
            if (kind === 'targets') setTargets(normalizeStoredTargets(rows))
            setDataMeta(current => ({ ...current, [kind]:dataset?.meta || null }))
            setCloudState(current => ({ ...current, mode:'cloud', live:true, lastSync:new Date().toISOString(), message:`עודכן ${kind} בלבד` }))
            setPerformance(current => ({ ...current, queries:current.queries + 1, rows:rows.length, phase:'עדכון חי' }))
            setStatus(`העדכון של ${kind} הושלם`)
          } catch (error) {
            console.warn('Realtime refresh failed', error)
            setCloudState(current => ({ ...current, live:false, message:'מחובר לענן, אך העדכון החי נכשל' }))
          }
        }, 700)
      })
      .subscribe(status => {
        if (status === 'SUBSCRIBED') setCloudState(current => ({ ...current, live:true }))
        if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') setCloudState(current => ({ ...current, live:false }))
      })
    return () => { clearTimeout(refreshTimer); supabase.removeChannel(channel) }
  }, [])

  useEffect(() => {
    if (!production.length && !quality.length && !deviations.length && !targets.length) return
    const timer = setTimeout(() => {
      idbSet({ production, quality, deviations, targets, dataMeta, savedAt: new Date().toISOString() })
        .catch(e => { console.error(e); setStatus('הנתונים בענן, אך יצירת גיבוי מקומי נכשלה') })
    }, 800)
    return () => clearTimeout(timer)
  }, [production, quality, deviations, targets, dataMeta])

  const validateRows = (kind, rows) => {
    const sample = rows.slice(0, 250)
    const present = (...names) => sample.some(r => names.some(n => getField(r, [n]) !== ''))
    const checks = {
      production: [
        ['מתקן / Storage Location', present('Storage Location', 'Storage location')],
        ['כמות', present('Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity')],
        ['Order או Batch', present('Order', 'Process Order', 'Batch', 'Batch Number')],
      ],
      quality: [
        ['Inspection Lot או Batch', present('Inspection Lot', 'Inspection Lot #', 'Batch', 'Batch Number')],
        ['מאפיין או סטטוס איכות', present('Master Insp Charactristic', 'Master Inspection Characteristic', 'Result Status', 'QA Approval')],
      ],
      deviations: [
        ['Batch', present('Batch', 'Batch Number')],
        ['סטטוס / מאפיין חריג', present('QA Status', 'Rejected characteristics', 'UD Remarks')],
      ],
      targets: [
        ['משאב', present('Resource', 'משאב', 'Storage Location', 'Facility', 'מתקן')],
        ['תוכנית חודשית', present('Plan', 'Monthly Target', 'Monthly Plan', 'יעד חודשי', 'תוכנית חודשית', 'Target')],
      ],
    }
    return (checks[kind] || []).filter(([, ok]) => !ok).map(([label]) => label)
  }

  const loadFiles = async (files, forcedKind = '') => {
    if (!files.length) return
    setBusy(true)
    try {
      const loaded = []
      for (const file of files) {
        const initialDisplayName = forcedKind ? displayDatasetName(forcedKind, forcedKind === 'targets' ? planningMonth : '') : file.name
        setStatus(`קורא את ${initialDisplayName}...`)
        if (forcedKind === 'targets') {
          const originalBuffer = await file.arrayBuffer()
          await idbSetKey(TARGET_FILE_KEY, {
            name: file.name,
            type: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            lastModified: file.lastModified || Date.now(),
            savedAt: new Date().toISOString(),
            bytes: originalBuffer,
          })
        }
        const sizeMB = file.size ? file.size / 1048576 : 0
        let rows
        try {
          rows = forcedKind === 'targets' ? await readTargetWorkbook(file) : await readWorkbook(file)
        } catch (readError) {
          // A 40MB / 280K-row SAP export needs well over 1GB of JS heap just to
          // parse. When the tab runs out, the failure used to surface further
          // down as the misleading "סוג הקובץ לא זוהה".
          throw new Error(sizeMB >= 15
            ? `${file.name}: הקובץ (${sizeMB.toFixed(0)}MB) גדול מכדי להיקרא בדפדפן — נגמר הזיכרון של הלשונית. פצל את הייצוא מ-SAP (לפי חודש או מתקן) וטען את החלקים בזה אחר זה.`
            : `${file.name}: קריאת הקובץ נכשלה — ${readError?.message || 'שגיאה לא ידועה'}`)
        }
        if (!rows.length) throw new Error(`${file.name}: הקובץ נקרא אך לא נמצאו בו שורות נתונים. ודא ששורת הכותרות היא הראשונה בגיליון.`)
        const detected = forcedKind === 'targets' ? 'targets' : classifyFile(rows)
        const kind = forcedKind || detected
        const missing = validateRows(kind, rows)
        if (missing.length) throw new Error(`${file.name}: חסרות עמודות חובה — ${missing.join(', ')}`)
        const rawRowCount = rows.length
        // Sampled before the rows are consumed by mapRowsInPlace below.
        const facilitiesFound = new Set(rows.slice(0, 5000).map(r => canonicalFacility(getField(r, ['Storage Location','Inspection Lot Storage Location','Process Order Storage Location','Facility','Production Line','מתקן']))).filter(Boolean)).size
        let storedCount = rows.length
        let rowsForCloud = rows
        if (kind === 'production') {
          const compact = mapRowsInPlace(rows, r => ({
            __compactProduction: true,
            facility: canonicalFacility(getField(r, ['Storage Location', 'Storage location'])),
            productionDay: localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date'])),
            finishDate: localDateTimeString(
  combineExcelDateTime(
    getField(r, ['Actual finish date', 'Actual Finish Date']),
    getField(r, ['Actual Finish Time', 'Actual finish time']),
    getField(r, ['Release date (actual)', 'Time Stamp'])
  )
),
            qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
            order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])),
            material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
            desc: normalize(getField(r, ['Material description', 'Material Description'])),
            orderType: normalize(getField(r, ['Order Type'])),
            routingGroup: normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup'])),
            routingDescription: normalize(getField(r, ['Description', 'Routing Description'])),
          })).filter(r => r.facility && (r.qty || r.order || r.batch))
          storedCount = compact.length
          rowsForCloud = compact
        }
        else if (kind === 'quality') {
          const compact = mapRowsInPlace(rows, r => ({
            __compactQuality: true,
            facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
            date: combineExcelDateTime(
              getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
              getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
              getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
            ),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
            order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])), status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
            approval: normalize(getField(r, ['QA Approval'])), inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
            sample: normalize(getField(r, ['Sample #', 'Sample Number', 'Sample No.', 'Sample'])),
            activity: normalize(getField(r, ['Operation Activity', 'Activity', 'Operation'])),
            operationText: normalize(getField(r, ['Operation short text'])),
            inspectionType: normalize(getField(r, ['Inspection Type'])),
            characteristic: normalize(getField(r, ['Master Insp Charactristic', 'Master Inspection Characteristic'])),
            value: normalize(getField(r, ['Arithmetic Mean of Valid Measured Values'])), lower: normalize(getField(r, ['Lower Specif Limit', 'Lower Spec Limit'])),
            upper: normalize(getField(r, ['Upper Specif Limit', 'Upper Spec Limit'])), unit: normalize(getField(r, ['Unit of Measurement'])),
            line: normalize(getField(r, ['Production Line'])), remarks: normalize(getField(r, ['Charactristic Remarks', 'Characteristic Remarks', 'Batch Remarks'])),
            qualitative: normalize(getField(r, ['Qualitative'])),
            udCode: normalize(getField(r, ['UD Code', 'Usage Decision', 'Usage decision', 'החלטת שימוש'])),
          })).filter(r => r.batch || r.inspectionLot)
          setStatus(`בודק אילו רשומות איכות חדשות קיימות ב-${displayDatasetName('quality')}...`)
          setUploadProgress({ fileName:displayDatasetName('quality'), kind, phase:'dedupe', percent:0, message:'משווה מול נתוני האיכות הקיימים' })
          const fresh = await filterNewQualityRows(quality, compact, (completed,total) => {
            const percent = total ? Math.round(completed / total * 100) : 100
            setUploadProgress({ fileName:displayDatasetName('quality'), kind, phase:'dedupe', percent, message:'מסנן רשומות שכבר קיימות' })
          })
          storedCount = fresh.length; rowsForCloud = fresh
          if (!fresh.length) {
            loaded.push(`${displayDatasetName('quality')}: לא נמצאו רשומות איכות חדשות`)
            setStatus(`${displayDatasetName('quality')}: כל ${fmt(compact.length)} הרשומות כבר קיימות — לא בוצעה העלאה`)
            continue
          }
        } else if (kind === 'deviations') {}
        else if (kind === 'targets') {
          const fallbackMonth = targetMonthFromTitle(file.name, new Date())
          const parsed = rows.flatMap(r => {
            const sourceResource = normalize(getField(r, ['Resource','משאב','1','Product','Line']))
            const makeTarget = (resource, overrides = {}) => {
              const facilities = targetFacilityIds(resource)
              const facility = facilities[0] || canonicalFacility(getField(r, ['Storage Location','Facility','מתקן']))
              return {
                resource, facility, facilities: facilities.length ? facilities : (facility ? [facility] : []),
                facilityLabel:(resource.match(/\(([^)]+)\)/)?.[1]||'').trim(), descriptionTokens:targetDescriptionTokens(resource),
                productFamily:normalize(r.__family), materials:Array.isArray(r.__materials) ? r.__materials.map(normalize).filter(Boolean) : [],
                routingGroup:normalizeRouting(getField(r,['Routing group','Routing Group','RoutingGroup','קבוצת ניתוב'])),
                station:normalize(getField(r,['Station','Work Center','תחנה']))||facility, lineName:resource,
                month:parseMonth(getField(r,['Month','חודש','Target Month','Plan Month']),fallbackMonth)||fallbackMonth,
                activity:normalize(getField(r,['Activity','Type','סוג פעילות','Production/Packaging']))||'ייצור / אריזה',
                capacity:parseTargetNumber(getField(r,['Capacity','קיבולת','Monthly Capacity','קיבולת חודשית'])) * 1000,
                target:parseTargetNumber(getField(r,['Plan','Monthly Target','Monthly Plan','יעד חודשי','תוכנית חודשית','Target'])) * 1000,
                fileProduction:parseTargetNumber(getField(r,['Production','ייצור'])), fileAchievement:parseTargetNumber(getField(r,['% Achievement','Achievement'])),
                requiredPerDay:parseTargetNumber(getField(r,['Req. t/d','Required t/d'])), lastDay:parseTargetNumber(getField(r,['Last day'])),
                adjustedRequiredPerDay:parseTargetNumber(getField(r,['Adjusted Req. t/d'])), actualPerDay:parseTargetNumber(getField(r,['Actual t/d'])),
                ratePercent:parseTargetNumber(getField(r,['%Rate','Rate'])), adherence:parseTargetNumber(getField(r,['New Adherence'])),
                recyclingPlan:parseTargetNumber(getField(r,['Recycling Plan'])), recycled:parseTargetNumber(getField(r,['Recycled'])),
                forPacking:parseTargetNumber(getField(r,['For Packing'])), restrictedRecycling:parseTargetNumber(getField(r,['Restricted - recycling'])),
                restrictedDisposal:parseTargetNumber(getField(r,['Restricted – risk for disposal','Restricted - risk for disposal'])),
                average:parseTargetNumber(getField(r,['AVERAGE','Average'])), notes:normalize(getField(r,['Notes','Remarks','הערות'])),
                ...overrides
              }
            }
            if (/^SHAKED\s+ISO\s*\(42\+23\)$/i.test(sourceResource)) {
              return [
                makeTarget('Shaked iso 42', { facility:'1142', facilities:['1142'], station:'1142', capacity:0, target:0 }),
                makeTarget('Shaked iso 23', { facility:'1123', facilities:['1123'], station:'1123', capacity:0, target:0 })
              ]
            }
            return [makeTarget(sourceResource)]
          }).filter(r=>r.resource&&r.facilities.length&&isApprovedTargetResource(r.resource))
          if(!parsed.length) throw new Error(`${file.name}: לא נמצאו שורות יעד תקינות. ודא שקיימות עמודות Resource, Capacity ו-Plan.`)
          storedCount=parsed.length; rowsForCloud=parsed; if(parsed[0]?.month) setPlanningMonth(parsed[0].month)
        } else throw new Error(`${file.name}: סוג הקובץ לא זוהה לפי הכותרות שנמצאו (${Object.keys(rows[0] || {}).slice(0, 6).join(', ')}...). טען אותו מאזור הטעינה הייעודי במרכז הנתונים.`)
        const displayName = displayDatasetName(kind, kind === 'targets' ? rowsForCloud?.[0]?.month : '')
        const nextMeta = { fileName:displayName, originalFileName:file.name, rows:storedCount, rawRows:rawRowCount, loadedAt:new Date().toISOString(), facilities:facilitiesFound, valid:true, source:'cloud' }
        setStatus(`מעלה את ${displayName} למסד המשותף...`)
        setUploadProgress({ fileName:displayName, kind, phase:'prepare', percent:0, message:'מכין את הנתונים' })
        const progressHandler = progress => {
          setUploadProgress({ fileName:displayName, kind, ...progress })
          setStatus(`${displayName}: ${progress.message} (${progress.percent}%)`)
        }
        const savedMeta = kind === 'quality'
          ? await uploadCloudDatasetIncremental(kind, rowsForCloud, { ...nextMeta, existingRows:quality.length }, currentUser, progressHandler)
          : await uploadCloudDataset(kind, rowsForCloud, nextMeta, currentUser, progressHandler)
        if (kind === 'production') setProduction(rowsForCloud)
        else if (kind === 'quality') setQuality(current => [...current, ...rowsForCloud])
        else if (kind === 'deviations') setDeviations(rowsForCloud)
        else if (kind === 'targets') setTargets(normalizeStoredTargets(rowsForCloud))
        setDataMeta(current => ({ ...current, [kind]: savedMeta }))
        setCloudState({ mode:'cloud', lastSync:savedMeta.loadedAt, message:'מחובר ומסונכרן עם Supabase', latencyMs:cloudState.latencyMs, live:true })
        loaded.push(`${displayName}: ${fmt(storedCount)} רשומות בענן`)
      }
      setStatus(`הטעינה לענן הושלמה — ${loaded.join(' | ')}`)
    } catch (e) {
      console.error('Cloud upload failed', e)
      const technical = [e?.message, e?.details, e?.hint, e?.code].filter(Boolean).join(' | ')
      setStatus(`שגיאה בטעינה לענן: ${technical || 'שגיאה לא ידועה'}`)
      setCloudState(current => ({
        ...current,
        mode: current.lastSync ? 'cloud' : current.mode,
        live: current.live,
        message: `החיבור לענן פעיל, אך העלאת הקובץ נכשלה: ${technical || 'שגיאה לא ידועה'}`,
      }))
    }
    finally { setBusy(false); setTimeout(() => setUploadProgress(null), 1800) }
  }

  const handleFiles = (files) => loadFiles(files)


  const prod = useMemo(() => production.map(r => {
    if (r?.__compactProduction) {
      return {
        facility: canonicalFacility(r.facility),
        productionDay: normalize(r.productionDay) || iso(r.finishDate),
        date: productionDateFromDay(r.productionDay) || (r.finishDate ? new Date(r.finishDate) : null),
        qty: num(r.qty),
        order: normalize(r.order),
        batch: normalize(r.batch),
        material: normalize(r.material),
        desc: normalize(r.desc),
        orderType: normalize(r.orderType),
        routingGroup: normalizeRouting(r.routingGroup),
        routingDescription: normalize(r.routingDescription),
        hour: r.finishDate ? new Date(r.finishDate).getHours() : null,
        shift: shiftInfo(r.finishDate ? new Date(r.finishDate) : null),
      }
    }
    const finish = combineExcelDateTime(
      getField(r, ['Actual finish date', 'Actual Finish Date']),
      getField(r, ['Actual Finish Time', 'Actual finish time']),
      getField(r, ['Release date (actual)', 'Time Stamp'])
    )
    return {
      facility: canonicalFacility(getField(r, ['Storage Location', 'Storage location'])),
      productionDay: localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date'])),
      date: productionDateFromDay(localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date']))) || finish,
      qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
      order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
      batch: normalize(getField(r, ['Batch', 'Batch Number'])),
      material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
      desc: normalize(getField(r, ['Material description', 'Material Description'])),
      orderType: normalize(getField(r, ['Order Type'])),
      routingGroup: normalizeRouting(getField(r, ['Routing group', 'Routing Group'])),
      routingDescription: normalize(getField(r, ['Description', 'Routing Description'])),
      hour: finish ? finish.getHours() : null,
      shift: shiftInfo(finish),
    }
  }).filter(r => r.facility), [production])
const materialByBatchDescription = useMemo(() => {
  const map = new Map()

  prod.forEach(row => {
    const batch = normalize(row.batch)
    const material = normalize(row.material)
    const desc = normalize(row.desc).toLowerCase()

    if (!batch || !material || !desc) return

    map.set(`${batch}|${desc}`, material)
  })

  return map
}, [prod])
  const qualityRows = useMemo(() => quality.map(r => {
  if (r.__compactQuality) {
    const repairedMaterial =
      materialByBatchDescription.get(
        `${normalize(r.batch)}|${normalize(r.material).toLowerCase()}`
      )

    return {
      ...r,
      material: repairedMaterial || r.material
    }
  }

  return ({
    facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
    date: combineExcelDateTime(
      getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
      getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
      getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
    ),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
 order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])),
status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
approval: normalize(getField(r, ['QA Approval'])),
inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
sample: normalize(getField(r, ['Sample #', 'Sample Number', 'Sample No.', 'Sample'])),
activity: normalize(getField(r, ['Operation Activity', 'Activity', 'Operation'])),
operationText: normalize(getField(r, ['Operation short text'])),
inspectionType: normalize(getField(r, ['Inspection Type'])),
characteristic: normalize(getField(r, ['Master Insp Charactristic', 'Master Inspection Characteristic'])),
value: normalize(getField(r, ['Arithmetic Mean of Valid Measured Values'])),
lower: normalize(getField(r, ['Lower Specif Limit', 'Lower Spec Limit'])),
upper: normalize(getField(r, ['Upper Specif Limit', 'Upper Spec Limit'])),
unit: normalize(getField(r, ['Unit of Measurement'])),
line: normalize(getField(r, ['Production Line'])),
remarks: normalize(getField(r, ['Charactristic Remarks', 'Characteristic Remarks', 'Batch Remarks'])),
qualitative: normalize(getField(r, ['Qualitative'])),
udCode: normalize(getField(r, ['UD Code', 'Usage Decision', 'Usage decision', 'החלטת שימוש']))
  })
}), [quality, materialByBatchDescription])

  const deviationRows = useMemo(() => deviations.map(r => ({
    facility: canonicalFacility(getField(r, ['Facility', 'Production Line', 'Storage Location'])),
    date: excelDate(getField(r, ['Date of Lot Creation', 'Inspection Lot UD Date', 'Process Order Delivered Date', 'Start Date of Inspection'])),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])),
material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
    inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
    status: normalize(getField(r, ['QA Status', 'Status'])),
    rejectedCount: num(getField(r, ['Rejected characteristics', 'Rejected characteristics '])),
    approvedCount: num(getField(r, ['Approved characteristics', 'Approved characteristics '])),
    udDate: normalize(getField(r, ['Inspection Lot UD Date', 'UD Date'])),
    remarks: normalize(getField(r, ['UD Remarks', 'Remarks', 'Batch Remarks'])),
    udCode: normalize(getField(r, ['UD Code'])),
  })), [deviations])

  // Build every quality lookup in one linear pass. The previous implementation
  // scanned the 263K-row quality array several times and then filtered the whole
  // array again for every deviation, which froze the browser after the fast cache render.
  const qualityIndex = useMemo(() => {
    const byBatch = new Map()
    const byBatchMaterial = new Map()
    const rejected = new Map()
    const approved = new Map()
    const latestByBatchMaterial = new Map()
    const latestByBatchMaterialLot = new Map()
    const seenRejected = new Map()
    const seenApproved = new Map()

    const addCharacteristic = (target, seenMap, key, item) => {
      if (!key || !item.characteristic) return
      const signature = `${item.characteristic}|${item.value}|${item.inspectionLot}|${item.sample}|${item.activity}`
      let seen = seenMap.get(key)
      if (!seen) { seen = new Set(); seenMap.set(key, seen) }
      if (seen.has(signature)) return
      seen.add(signature)
      const list = target.get(key) || []
      list.push(item)
      target.set(key, list)
    }

    qualityRows.forEach(row => {
      const key = batchMaterialKey(row.batch, row.material)
      if (!key) return

      const list = byBatchMaterial.get(key) || []
      list.push(row)
      byBatchMaterial.set(key, list)

      // Batch-only fallback, so a batch whose material is missing or that spans
      // several materials still resolves to its quality rows instead of showing
      // an empty card.
      const batchOnly = normalize(row.batch)
      if (batchOnly) {
        const batchList = byBatch.get(batchOnly) || []
        batchList.push(row)
        byBatch.set(batchOnly, batchList)
      }

      const timestamp = row.date ? new Date(row.date).getTime() : 0
      if (timestamp > (latestByBatchMaterial.get(key)?.timestamp || 0)) {
        latestByBatchMaterial.set(key, { timestamp, date: row.date })
      }
      if (row.inspectionLot) {
        const lotKey = `${key}|${normalize(row.inspectionLot)}`
        if (timestamp > (latestByBatchMaterialLot.get(lotKey)?.timestamp || 0)) {
          latestByBatchMaterialLot.set(lotKey, { timestamp, date: row.date })
        }
      }
      const status = normalize(row.status || row.approval).toLowerCase()
      const isRejected = ['rejection', 'rejected', 'fail', 'failed', 'פסול', 'לא תקין', 'חריג'].some(x => status.includes(x))
      const item = {
        characteristic: row.characteristic,
        value: row.value,
        lower: row.lower,
        upper: row.upper,
        unit: row.unit,
        remarks: row.remarks,
        qualitative: row.qualitative,
        inspectionLot: row.inspectionLot,
        sample: row.sample,
        activity: row.activity,
        date: row.date,
      }
      addCharacteristic(isRejected ? rejected : approved, isRejected ? seenRejected : seenApproved, key, item)
    })

    return {
    byBatch,
    byBatchMaterial,
    rejected,
    approved,
    latestByBatchMaterial,
    latestByBatchMaterialLot,
}
  }, [qualityRows])

  const enrichedDeviationRows = useMemo(() => deviationRows.map(row => {
    const key = batchMaterialKey(row.batch, row.material)
    const lotKey = key && row.inspectionLot ? `${key}|${normalize(row.inspectionLot)}` : ''
    const sampleDate = key
      ? ((lotKey && qualityIndex.latestByBatchMaterialLot.get(lotKey)?.date) || qualityIndex.latestByBatchMaterial.get(key)?.date || null)
      : null
    return {
      ...row,
      sampleDate,
      rejectedCharacteristics: key ? (qualityIndex.rejected.get(key) || []) : [],
      approvedCharacteristics: key ? (qualityIndex.approved.get(key) || []) : [],
    }
  }), [deviationRows, qualityIndex])

  // Plant rule: a quality record is uniquely identified by exact Batch + Material.
  // Order, facility, routing group and inspection lot remain display fields only.
  const selectedBatchData = useMemo(() => {
    const batch = normalize(selectedBatch)
    const requestedMaterial = normalize(selectedBatchMaterial)
    if (!batch) return null

    const batchProductionRows = prod.filter(row => normalize(row.batch) === batch)
    const batchMaterials = [...new Set(batchProductionRows.map(row => normalize(row.material)).filter(Boolean))]
    const material = requestedMaterial || (batchMaterials.length === 1 ? batchMaterials[0] : '')
    const key = batchMaterialKey(batch, material)

    const productionRows = material
      ? batchProductionRows.filter(row => normalize(row.material) === material)
      : batchProductionRows

    // Exact Batch+Material stays the rule. When the batch spans several
    // materials (or the material is blank) we fall back to batch-only instead of
    // rendering an empty card, and flag it so the UI can say so.
    const exactQuality = key ? (qualityIndex.byBatchMaterial.get(key) || []) : []
    const usedBatchOnlyFallback = !key || (!exactQuality.length && batchMaterials.length !== 1)
    const qualityForBatch = usedBatchOnlyFallback
      ? (qualityIndex.byBatch.get(batch) || [])
      : exactQuality

    const deviationForBatch = usedBatchOnlyFallback
      ? enrichedDeviationRows.filter(row => normalize(row.batch) === batch)
      : enrichedDeviationRows.filter(row => batchMaterialKey(row.batch, row.material) === key)

    return {
      batch,
      material,
      materials: batchMaterials,
      usedBatchOnlyFallback,
      production: productionRows,
      quality: qualityForBatch,
      deviations: deviationForBatch,
    }
  }, [selectedBatch, selectedBatchMaterial, prod, qualityIndex, enrichedDeviationRows])

  const openBatchCard = (batch, material = '') => {
    if (!batch) return
    setSelectedBatch(normalize(batch))
    setSelectedBatchMaterial(normalize(material))
  }

  const dataMonths = useMemo(() => [...new Set(prod.map(r => monthKey(r.date)).filter(Boolean))].sort(), [prod])
  const targetMonths = useMemo(() => [...new Set(targets.map(r => r.month).filter(Boolean))].sort(), [targets])
  const availableMonths = useMemo(() => [...new Set([...targetMonths, ...dataMonths])].sort().reverse(), [targetMonths, dataMonths])
  useEffect(() => { if (!planningMonth && availableMonths.length) setPlanningMonth(availableMonths[0]) }, [availableMonths, planningMonth])

  const dateBounds = useMemo(() => {
    let minTime = Infinity
    let maxTime = -Infinity
    const scan = rows => rows.forEach(row => {
      const time = row.date ? new Date(row.date).getTime() : NaN
      if (!Number.isFinite(time)) return
      if (time < minTime) minTime = time
      if (time > maxTime) maxTime = time
    })
    scan(prod); scan(qualityRows); scan(deviationRows)
    return {
      min: Number.isFinite(minTime) ? iso(new Date(minTime)) : '',
      max: Number.isFinite(maxTime) ? iso(new Date(maxTime)) : '',
    }
  }, [prod, qualityRows, deviationRows])

  const baseFiltered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return prod.filter(r => {
      const day = r.productionDay || iso(r.date)
      const inRange = (!from || day >= from) && (!to || day <= to)
      return inRange && (!q || [r.facility, r.order, r.batch, r.material, r.desc].some(v => String(v || '').toLowerCase().includes(q)))
    })
  }, [prod, from, to, query])
  const filtered = useMemo(() => baseFiltered.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.facility)), [baseFiltered, selectedFacilities])

  const udByBatchMaterial = useMemo(() => {
    const exact = new Map()
    const byBatch = new Map()
    const put = (row) => {
      const batch = normalize(row?.batch)
      const material = normalize(row?.material)
      const ud = normalize(row?.udCode || row?.usageDecision)
      if (!batch || !ud) return
      if (material) exact.set(`${batch}|${material}`, ud)
      if (!byBatch.has(batch)) byBatch.set(batch, ud)
    }
    qualityRows.forEach(put)
    deviationRows.forEach(put)
    return { exact, byBatch }
  }, [qualityRows, deviationRows])

  const productionUsageDecision = row => {
    const batch = normalize(row?.batch)
    const material = normalize(row?.material)
    if (!batch) return '—'
    return udByBatchMaterial.exact.get(`${batch}|${material}`) || udByBatchMaterial.byBatch.get(batch) || '—'
  }

  const productionSortValue = (row, key) => {
    if (key === 'date') return row?.date ? new Date(row.date).getTime() : 0
    if (key === 'ud') return productionUsageDecision(row)
    if (key === 'facility') return normalize(row?.facility)
    if (key === 'routingGroup') return normalize(row?.routingGroup)
    if (key === 'order') return normalize(row?.order)
    if (key === 'batch') return normalize(row?.batch)
    if (key === 'material') return normalize(row?.material)
    if (key === 'desc') return normalize(row?.desc)
    if (key === 'qty') return Number(row?.qty) || 0
    return ''
  }

  const sortedRecentProduction = useMemo(() => {
    const rows = filtered.slice()
    const { key, direction } = productionSort
    const factor = direction === 'asc' ? 1 : -1
    rows.sort((a,b) => {
      const av = productionSortValue(a,key), bv = productionSortValue(b,key)
      if (typeof av === 'number' && typeof bv === 'number') return (av-bv)*factor
      return String(av).localeCompare(String(bv), 'he', { numeric:true, sensitivity:'base' }) * factor
    })
    return rows.slice(0,200)
  }, [filtered, productionSort, udByBatchMaterial])

  const toggleProductionSort = key => setProductionSort(current => ({
    key,
    direction: current.key === key && current.direction === 'asc' ? 'desc' : 'asc'
  }))
  const productionSortArrow = key => productionSort.key === key ? (productionSort.direction === 'asc' ? ' ↑' : ' ↓') : ' ↕'


  const discoveredFacilities = useMemo(() => [...new Set([...targets.map(t => t.facility), ...prod.map(r => r.facility)].filter(Boolean))].sort(), [targets, prod])
  const optionalFacilities = useMemo(() => discoveredFacilities.filter(id => !PRIMARY_FACILITIES.includes(id) && !additionalFacilities.includes(id)), [discoveredFacilities, additionalFacilities])
  const facilities = useMemo(() => [...PRIMARY_FACILITIES, ...additionalFacilities], [additionalFacilities])
  const availableYears = useMemo(() => [...new Set(prod.map(r => r.date?.getFullYear()).filter(Boolean))].sort((a,b) => b-a), [prod])

  const planningRows = useMemo(() => buildResourceRows({
    production: prod,
    targets,
    planningMonth,
    fallbackFacilities: facilities,
    manualMappings,
  }), [planningMonth, prod, targets, facilities, manualMappings])


  const mappingTargets = useMemo(() => planningRows.map(row => ({
    key: row.targetKey,
    resource: row.resource,
    facility: row.facility,
    facilities: row.facilities,
    target: row.target,
    actual: row.actual,
    pct: row.pct,
  })), [planningRows])

  const addMappingEvent = (action, mapping, note = '') => {
    setMappingTimeline(current => [{
      id:`${Date.now()}-${Math.random()}`,
      at:new Date().toISOString(), action, note,
      mappingId:mapping?.id || '', material:mapping?.material || '',
      family:mapping?.family || '', targetResource:mapping?.targetResource || '',
      user:currentUser?.email || currentUser?.user_metadata?.full_name || 'מנהל',
    }, ...current].slice(0,500))
  }

  const syncMappingToCloud = async mapping => {
    if (!supabase || !currentUser?.id) return null
    const payload = {
      row_key:mapping.rowKey, material:mapping.material || '', description:mapping.description || '',
      facility_family:mapping.family, target_key:mapping.targetKey, target_resource:mapping.targetResource,
      status:mapping.status, active:mapping.active !== false, created_by:currentUser.id,
      created_by_email:currentUser.email || '', approved_by:mapping.status === 'approved' ? currentUser.id : null,
      approved_at:mapping.status === 'approved' ? new Date().toISOString() : null,
      updated_at:new Date().toISOString(),
    }
    if (mapping.id?.length === 36) payload.id = mapping.id
    const { data, error } = await supabase.from('iml_product_mappings').upsert(payload).select().maybeSingle()
    if (error) return null
    return data
  }

  const openMappingDialog = item => {
    const candidates = mappingTargets.filter(target => (target.facilities || []).map(stationFamily).includes(item.family))
    setMappingDialog({ ...item, candidates })
    setMappingTargetKey(candidates[0]?.key || '')
    setMappingMessage('')
  }

  const savePendingMapping = async () => {
    if (!mappingDialog || !mappingTargetKey) return
    const target = mappingTargets.find(item => item.key === mappingTargetKey)
    if (!target) return
    const row = mappingDialog.row
    const mapping = {
      id:globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`,
      rowKey:productionMappingKey(row), material:row.material || '', description:row.desc || '',
      family:mappingDialog.family, targetKey:mappingTargetKey, targetResource:target.resource,
      status:'pending', active:true, createdBy:currentUser?.email || '', createdAt:new Date().toISOString(),
    }
    setManualMappings(current => [mapping, ...current.filter(item => item.rowKey !== mapping.rowKey || item.status === 'rejected')])
    addMappingEvent('created', mapping, 'נוצר שיוך ממתין לאישור')
    const cloudRow = await syncMappingToCloud(mapping)
    if (cloudRow?.id) setManualMappings(current => current.map(item => item.id === mapping.id ? { ...item, id:cloudRow.id, cloud:true } : item))
    setMappingDialog(null)
    setMappingMessage('השיוך נשמר כממתין לאישור')
  }

  const updateMappingStatus = async (mapping, status) => {
    const updated = { ...mapping, status, active:status !== 'rejected', approvedAt:status === 'approved' ? new Date().toISOString() : mapping.approvedAt }
    setManualMappings(current => current.map(item => item.id === mapping.id ? updated : item))
    addMappingEvent(status, updated, status === 'approved' ? 'השיוך אושר והשפיע על הדשבורד' : 'השיוך נדחה')
    await syncMappingToCloud(updated)
  }

  const rollbackMapping = async mapping => {
    const updated = { ...mapping, active:false, status:'rejected' }
    setManualMappings(current => current.map(item => item.id === mapping.id ? updated : item))
    addMappingEvent('rollback', updated, 'השיוך בוטל והחישוב חזר למצב הקודם')
    await syncMappingToCloud(updated)
  }


  const mappingSimulation = useMemo(() => {
    const assignments = new Map()
    planningRows.forEach(targetRow => {
      ;(targetRow.productionRows || []).forEach(row => {
        const list = assignments.get(row) || []
        list.push({
          resource: targetRow.resource,
          facility: targetRow.facility,
          station: targetRow.station,
          routingGroup: targetRow.routingGroup,
        })
        assignments.set(row, list)
      })
    })

    const rows = filtered.map((row, index) => {
      const matches = assignments.get(row) || []
      const rowKey = productionMappingKey(row)
      const approvedManual = manualMappings.find(mapping => mapping.active !== false && mapping.status === 'approved' && mapping.rowKey === rowKey)
      const pendingManual = manualMappings.find(mapping => mapping.active !== false && mapping.status === 'pending' && mapping.rowKey === rowKey)
      const status = matches.length === 0 ? 'unmatched' : matches.length > 1 ? 'duplicate' : 'matched'
      const stationDigits = String(row.facility || '').replace(/\D/g, '')
      const suffix = stationDigits.length >= 2 ? stationDigits.slice(-2) : ''
      const family = suffix ? `15${suffix}` : row.facility || ''
      return {
        key: `${row.order || ''}-${row.batch || ''}-${index}`,
        row,
        matches,
        status,
        family,
        assignedResource: matches.map(item => item.resource).filter(Boolean).join(' | '),
        approvedManual, pendingManual,
        explanation: status === 'matched'
          ? `שויך ליעד אחד: ${matches[0].resource || matches[0].facility || 'ללא שם'}`
          : status === 'duplicate'
            ? `שויך ל-${matches.length} יעדים — חשש לספירה כפולה`
            : 'לא נמצא יעד תואם לפי כללי המנוע הנוכחיים',
      }
    })

    const summary = rows.reduce((acc, item) => {
      acc.rows += 1
      acc.quantity += Number(item.row.qty) || 0
      if (item.status === 'matched') { acc.matched += 1; acc.matchedQty += Number(item.row.qty) || 0 }
      if (item.status === 'unmatched') { acc.unmatched += 1; acc.unmatchedQty += Number(item.row.qty) || 0 }
      if (item.status === 'duplicate') { acc.duplicate += 1; acc.duplicateQty += Number(item.row.qty) || 0 }
      return acc
    }, { rows:0, quantity:0, matched:0, matchedQty:0, unmatched:0, unmatchedQty:0, duplicate:0, duplicateQty:0 })

    return { rows, summary }
  }, [filtered, planningRows, manualMappings])

  const visibleMappingSimulation = useMemo(() => (
    simulatorOnlyIssues
      ? mappingSimulation.rows.filter(item => item.status !== 'matched')
      : mappingSimulation.rows
  ), [mappingSimulation, simulatorOnlyIssues])

  const exportMappingSimulation = () => {
    const rows = mappingSimulation.rows.map(item => ({
      Date: item.row.productionDay || iso(item.row.date),
      Time: item.row.date ? new Date(item.row.date).toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) : '',
      StorageLocation: item.row.facility,
      FacilityFamily: item.family,
      RoutingGroup: item.row.routingGroup,
      OrderType: item.row.orderType,
      Order: item.row.order,
      Batch: item.row.batch,
      Material: item.row.material,
      Description: item.row.desc,
      Quantity: item.row.qty,
      MatchStatus: item.status,
      AssignedResource: item.assignedResource,
      Explanation: item.explanation,
    }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Mapping Simulator')
    XLSX.writeFile(wb, `IML_Production_Mapping_Simulator_${new Date().toISOString().slice(0,10)}.xlsx`)
  }

  const facilityStats = useMemo(() => facilities.map(id => {
    let rows = baseFiltered.filter(r => r.facility === id)
    if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
    const actual = rows.reduce((s, r) => s + r.qty, 0)
    const plans = planningRows.filter(x => (x.facilities || [x.facility]).includes(id))
    const target = plans.reduce((sum, row) => sum + row.target, 0)
    const forecast = plans.reduce((sum, row) => sum + row.forecast, 0)
    const rank = { risk: 3, warning: 2, 'no-target': 1, good: 0, achieved: 0 }
    const state = plans.sort((a,b) => (rank[b.state] || 0) - (rank[a.state] || 0))[0]?.state || 'no-target'
    return { id, actual, target, pct: target ? actual / target * 100 : 0, orders: new Set(rows.map(r => r.order).filter(Boolean)).size, state, forecast }
  }), [facilities, baseFiltered, planningRows])


  const toggleFacility = (id) => setSelectedFacilities(current => current.includes(id) ? current.filter(x => x !== id) : [...current, id])
  const allFacilitiesSelected = selectedFacilities.length === facilities.length
  const toggleAllFacilities = () => setSelectedFacilities(allFacilitiesSelected ? [] : [...facilities])

  const total = filtered.reduce((s, x) => s + x.qty, 0)
  const activeFacilities = facilityStats.filter(x => x.actual > 0).length
  const morningQty = filtered.filter(r => r.shift?.key === 'morning').reduce((s, r) => s + r.qty, 0)
  const eveningQty = filtered.filter(r => r.shift?.key === 'evening').reduce((s, r) => s + r.qty, 0)
  const nightQty = filtered.filter(r => r.shift?.key === 'night').reduce((s, r) => s + r.qty, 0)

  const shiftAnalysis = useMemo(() => {
    const batchShift = new Map()
    filtered.forEach(r => { if (r.batch && r.shift?.key !== 'unknown') batchShift.set(r.batch, r.shift) })
    const rowsByShift = new Map(['morning','evening','night'].map(key => [key, []]))
    filtered.forEach(r => { if (rowsByShift.has(r.shift?.key)) rowsByShift.get(r.shift.key).push(r) })
    const labels = { morning:'בוקר', evening:'ערב', night:'לילה' }
    const hours = { morning:'07:00–15:00', evening:'15:00–23:00', night:'23:00–07:00' }
    return ['morning','evening','night'].map(key => {
      const rows = rowsByShift.get(key).slice().sort((a,b) => (a.date?.getTime?.() || 0) - (b.date?.getTime?.() || 0))
      const groups = new Map()
      rows.forEach(r => {
        const groupKey = `${r.shift.workDate}|${r.facility}|${r.routingGroup || 'NO-ROUTING'}`
        const list = groups.get(groupKey) || []; list.push(r); groups.set(groupKey, list)
      })
      const changeovers = []
      groups.forEach(list => {
        list.sort((a,b) => a.date - b.date)
        for (let i=1;i<list.length;i++) {
          const prev=list[i-1], curr=list[i]
          if (prev.material && curr.material && prev.material !== curr.material) {
            const minutes = Math.max(0, Math.round((curr.date - prev.date)/60000))
            changeovers.push({ facility:curr.facility, routingGroup:curr.routingGroup, at:curr.date, fromMaterial:prev.material, fromDesc:prev.desc, toMaterial:curr.material, toDesc:curr.desc, minutes })
          }
        }
      })
      const batches = new Set(rows.map(r=>r.batch).filter(Boolean))
      const shiftDeviations = enrichedDeviationRows.filter(d => { const st = normalize(d.status).toLowerCase(); const isOpen = !st || !['approved','closed','מאושר','סגור'].some(x => st.includes(x)); return isOpen && d.batch && batches.has(d.batch) })
      const total = rows.reduce((sum,r)=>sum+r.qty,0)
      return {
        key, label:labels[key], hours:hours[key], total,
        orders:new Set(rows.map(r=>r.order).filter(Boolean)).size,
        batches:batches.size,
        materials:new Set(rows.map(r=>r.material).filter(Boolean)).size,
        changeovers,
        avgChangeover:changeovers.length ? changeovers.reduce((sum,c)=>sum+c.minutes,0)/changeovers.length : 0,
        totalChangeover:changeovers.reduce((sum,c)=>sum+c.minutes,0),
        deviations:shiftDeviations.length,
        avgPerHour:total/8,
        share: total ? total / Math.max(1, filtered.reduce((sum,r)=>sum+r.qty,0))*100 : 0,
      }
    })
  }, [filtered, enrichedDeviationRows])
  // A single date/facility scope is shared by production, quality and deviations.
  // This prevents quality/deviation cards from continuing to show rows outside
  // the selected period (for example when choosing "יום אחרון").
  const filteredQualityRows = useMemo(() => qualityRows.filter(r =>
    (!selectedFacilities.length || selectedFacilities.includes(r.facility)) &&
    matchesDateRange(r.date, from, to)
  ), [qualityRows, selectedFacilities, from, to])

  const filteredDeviationRows = useMemo(() => enrichedDeviationRows.filter(r =>
    (!selectedFacilities.length || selectedFacilities.includes(r.facility)) &&
    matchesDateRange(r.sampleDate || r.date, from, to)
  ), [enrichedDeviationRows, selectedFacilities, from, to])

  const qualityBad = useMemo(() => filteredQualityRows.filter(r => {
    const st = normalize(r.status).toLowerCase()
    return st && !['accepted', 'תקין', 'pass', 'approved', 'מאושר'].some(x => st.includes(x))
  }).sort((a,b) => (b.date?.getTime?.() || new Date(b.date || 0).getTime()) - (a.date?.getTime?.() || new Date(a.date || 0).getTime())), [filteredQualityRows])

  const openDeviations = useMemo(() => filteredDeviationRows.filter(r => {
    const st = normalize(r.status).toLowerCase()
    return !st || !['approved', 'closed', 'מאושר', 'סגור'].some(x => st.includes(x))
  }).sort((a,b) => new Date(b.sampleDate || b.date || 0) - new Date(a.sampleDate || a.date || 0)), [filteredDeviationRows])

  const alerts = useMemo(() => planningRows.filter(r => ['risk', 'warning'].includes(r.state)).sort((a,b) => ({ risk:0, warning:1 }[a.state] - { risk:0, warning:1 }[b.state])), [planningRows])
  const achievedCount = planningRows.filter(r => ['achieved', 'good'].includes(r.state)).length
  const riskCount = planningRows.filter(r => r.state === 'risk').length
  const warningCount = planningRows.filter(r => r.state === 'warning').length
  const targetTotal = planningRows.reduce((s,r) => s + r.target, 0)
  const targetActual = planningRows.reduce((s,r) => s + r.actual, 0)
  const targetForecast = planningRows.reduce((s,r) => s + r.forecast, 0)
  const uniqueOrders = useMemo(() => new Set(filtered.map(r => r.order).filter(Boolean)).size, [filtered])
  const uniqueBatches = useMemo(() => new Set(filtered.map(r => r.batch).filter(Boolean)).size, [filtered])
  const managerInsights = useMemo(() => {
    const insights = []
    const highestRisk = planningRows.filter(r => r.state === 'risk').sort((a,b) => (b.requiredDaily - b.provenMax) - (a.requiredDaily - a.provenMax))[0]
    const warning = planningRows.filter(r => r.state === 'warning').sort((a,b) => a.forecast / Math.max(1,a.target) - b.forecast / Math.max(1,b.target))[0]
    const best = planningRows.filter(r => r.target > 0).sort((a,b) => (b.forecast / b.target) - (a.forecast / a.target))[0]
    if (highestRisk) insights.push({ state:'risk', title:`${planningName(highestRisk)} דורש פעולה`, text:`הקצב הנדרש הוא ${fmt(highestRisk.requiredDaily)} ליום לעומת שיא מוכח של ${fmt(highestRisk.provenMax)}.` })
    if (warning) insights.push({ state:'warning', title:`${planningName(warning)} נמצא בסיכון`, text:`התחזית היא ${fmt(warning.forecast)} מול יעד של ${fmt(warning.target)}. נדרש שיפור בקצב היומי.` })
    if (best && ['good','achieved'].includes(best.state)) insights.push({ state:'good', title:`${planningName(best)} מוביל`, text:`התחזית הנוכחית היא ${pctFmt(best.forecast / best.target * 100)} מהיעד החודשי.` })
    if (openDeviations.length) insights.push({ state:'risk', title:`${openDeviations.length} חריגות איכות פתוחות`, text:'לחיצה על הכרטיס תפתח את פירוט המנות ומאפייני החריגה.' })
    if (!insights.length) insights.push({ state:'good', title:'אין מוקדים חריגים', text:'לפי הנתונים הטעונים, המתקנים נמצאים במסלול תקין ואין חריגות פתוחות.' })
    return insights.slice(0,4)
  }, [planningRows, openDeviations.length])


  const controlTowerFacilities = useMemo(() => {
    const deviationByFacility = new Map()
    openDeviations.forEach(row => { const facility=canonicalFacility(row.facility); if(facility) deviationByFacility.set(facility,(deviationByFacility.get(facility)||0)+1) })
    return planningRows.map(row => {
      const facilitiesForRow=row.facilities||[row.facility].filter(Boolean)
      const deviationsCount=facilitiesForRow.reduce((sum,id)=>sum+(deviationByFacility.get(id)||0),0)
      const forecastPct=row.target?row.forecast/row.target*100:0, actualPct=row.target?row.actual/row.target*100:0
      const planScore=row.target?Math.min(100,forecastPct):70, qualityScore=Math.max(0,100-deviationsCount*8)
      const healthScore=Math.max(0,Math.min(100,Math.round(planScore*.72+qualityScore*.28)))
      const state=!row.target?'no-target':forecastPct>=100?'good':forecastPct>=90?'warning':'risk'
      return { ...row, facility:row.resource||row.facility, facilityIds:facilitiesForRow, deviationsCount, forecastPct, actualPct, healthScore, state, gap:row.forecast-row.target }
    }).filter(row=>row.target>0||row.actual>0||row.deviationsCount>0)
      .sort((a,b)=>({risk:0,warning:1,good:2,'no-target':3}[a.state]-{risk:0,warning:1,good:2,'no-target':3}[b.state])||a.facility.localeCompare(b.facility))
  }, [planningRows, openDeviations])

  const facilityViewMatches = row => {
    if (facilityViewMode === 'target') return Number(row.target || 0) > 0
    if (facilityViewMode === 'active') return Number(row.actual || 0) > 0
    if (facilityViewMode === 'risk') return row.state === 'risk' || row.state === 'warning'
    if (facilityViewMode === 'all') return true
    return Number(row.target || 0) > 0 || Number(row.actual || 0) > 0
  }
  const visibleControlTowerFacilities = useMemo(() => controlTowerFacilities.filter(facilityViewMatches), [controlTowerFacilities, facilityViewMode])
  const visiblePlanningRows = useMemo(() => planningRows.filter(facilityViewMatches), [planningRows, facilityViewMode])
  const facilityViewFilters = <div className="facility-view-filters">
    {[['relevant','יעד + פעילות'],['target','עם יעד'],['active','פעילים'],['risk','בסיכון'],['all','הכול']].map(([value,label]) =>
      <button key={value} type="button" className={facilityViewMode === value ? 'active' : ''} onClick={() => setFacilityViewMode(value)}>{label}</button>
    )}
  </div>

  const controlTowerTrend = useMemo(() => {
    const byDay = new Map()
    prod.filter(row => monthKey(row.date) === planningMonth).forEach(row => {
      const key = iso(row.date)
      if (key) byDay.set(key, (byDay.get(key) || 0) + row.qty)
    })
    return [...byDay.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-7)
  }, [prod, planningMonth])

  const jumpToDetails = (tab) => {
    setActiveTab(tab)
    window.setTimeout(() => document.getElementById('details-section')?.scrollIntoView({ behavior:'smooth', block:'start' }), 50)
  }

  const exportWorkbook = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filtered.map(r => ({ Date: iso(r.date), Time: r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : '', Facility: r.facility, Order: r.order, Batch: r.batch, Material: r.material, Description: r.desc, RoutingGroup: r.routingGroup, RoutingDescription: r.routingDescription, Quantity: r.qty }))), 'Production')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planningRows.map(r => ({ Month: planningMonth, Facility: r.facility, RoutingGroup: r.routingGroup, Station: r.station, Line: r.lineName, Activity: r.activity, MonthlyTarget: r.target, Actual: r.actual, Remaining: r.remaining, RequiredDaily: r.requiredDaily, RecentAverage: r.recentAverage, ProvenMax: r.provenMax, Forecast: r.forecast, Status: r.label }))), 'Planning')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(qualityBad.map(r => ({ Date: iso(r.date), Facility: r.facility, InspectionLot: r.inspectionLot, Order: r.order, Batch: r.batch, Material: r.material, Status: r.status }))), 'Quality')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(openDeviations.map(r => ({ Date: iso(r.date), Facility: r.facility, Batch: r.batch, Material: r.material, Status: r.status, RejectedCount: r.rejectedCount, RejectedCharacteristics: r.rejectedCharacteristics.map(x => `${x.characteristic}: ${x.value || x.qualitative || '-'} (${[x.lower, x.upper].filter(v => v !== '').join('–')} ${x.unit || ''})`).join(' | '), ApprovedCharacteristics: r.approvedCharacteristics.map(x => `${x.characteristic}: ${x.value || x.qualitative || '-'} (${[x.lower, x.upper].filter(v => v !== '').join('–')} ${x.unit || ''})`).join(' | '), Remarks: r.remarks }))), 'Deviations')
    XLSX.writeFile(wb, `IML_Sprint8_Resource_Report_${new Date().toISOString().slice(0,10)}.xlsx`)
  }
  const downloadTargetWorkbook = async () => {
    try {
      const stored = await idbGetKey(TARGET_FILE_KEY)
      if (stored?.bytes) {
        const blob = new Blob([stored.bytes], { type: stored.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        const url = URL.createObjectURL(blob)
        const link = document.createElement('a')
        link.href = url
        link.download = stored.name || `IML_Monthly_Targets_${planningMonth || monthKey(new Date())}.xlsx`
        document.body.appendChild(link)
        link.click()
        link.remove()
        setTimeout(() => URL.revokeObjectURL(url), 1500)
        setStatus(`קובץ היעדים המקורי הורד — ניתן לערוך ב-Excel ולהטעין מחדש`)
        return
      }

      // If no newer target workbook was uploaded in this browser, download
      // the full approved master workbook bundled with the application.
      const response = await fetch('/templates/IML_Targets_Master.xlsx', { cache:'no-store' })
      if (!response.ok) throw new Error('קובץ תבנית היעדים המלא לא נמצא באפליקציה')
      const blob = await response.blob()
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = '26 דוח ביצוע SAP Aug.xlsx'
      document.body.appendChild(link)
      link.click()
      link.remove()
      setTimeout(() => URL.revokeObjectURL(url), 1500)
      setStatus('הורדה תבנית היעדים המלאה. לאחר טעינת קובץ יעדים חדש, הכפתור יוריד את הקובץ האחרון שנטען.')
    } catch (error) {
      console.error('Target workbook download failed', error)
      setStatus(`הורדת קובץ היעדים נכשלה: ${error?.message || 'שגיאה לא ידועה'}`)
    }
  }

  const openPrintReport = (title, subtitle, headers, rows) => {
    const popup = window.open('', '_blank', 'width=1400,height=900')
    if (!popup) { setStatus('הדפדפן חסם חלון הדפסה. יש לאפשר חלונות קופצים לאתר.'); return }
    const esc = value => String(value ?? '').replace(/[&<>\"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[ch]))
    const bodyRows = rows.map((row, index) => `<tr class="${esc(row._state || '')}"><td class="index">${index+1}</td>${headers.map(h => `<td>${esc(row[h.key])}</td>`).join('')}</tr>`).join('')
    popup.document.write(`<!doctype html><html dir="rtl"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
      @page{size:A4 landscape;margin:9mm}*{box-sizing:border-box}body{font-family:Arial,"Segoe UI",sans-serif;color:#16324a;margin:0;background:#f4f8fb}.sheet{background:white;padding:22px}.head{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:4px solid #159b83;padding-bottom:14px;margin-bottom:18px}.head h1{margin:0;font-size:28px}.head p{margin:5px 0 0;color:#64748b}.brand{font-weight:800;color:#159b83;font-size:19px}.summary{display:flex;gap:10px;margin:0 0 16px;flex-wrap:wrap}.chip{background:#eaf7f3;border:1px solid #c7eadf;padding:8px 12px;border-radius:12px;font-weight:700}table{width:100%;border-collapse:separate;border-spacing:0;font-size:11px;overflow:hidden;border:1px solid #dbe5ec;border-radius:12px}th{background:#173b57;color:white;padding:9px 6px}td{padding:8px 6px;border-bottom:1px solid #e6edf2;text-align:center}tr:nth-child(even) td{background:#f8fbfd}tr.good td,tr.achieved td{background:#effaf5}tr.warning td{background:#fff8e6}tr.risk td{background:#fff0f0}.index{font-weight:700;color:#64748b}.footer{margin-top:12px;color:#64748b;font-size:10px}@media print{body{background:white}.sheet{padding:0}.no-print{display:none}}
    </style></head><body><div class="sheet"><div class="head"><div><h1>${esc(title)}</h1><p>${esc(subtitle)}</p></div><div class="brand">IML CONTROL</div></div><div class="summary"><div class="chip">חודש: ${esc(planningMonth || '—')}</div><div class="chip">תאריך הדפסה: ${esc(new Date().toLocaleString('he-IL'))}</div><div class="chip">רשומות: ${rows.length}</div></div><table><thead><tr><th>#</th>${headers.map(h=>`<th>${esc(h.label)}</th>`).join('')}</tr></thead><tbody>${bodyRows}</tbody></table><div class="footer">דוח ניהולי — IML CONTROL</div></div><script>window.onload=()=>setTimeout(()=>window.print(),250)<\/script></body></html>`)
    popup.document.close()
  }

  const planningDisplayStation = row => /^EC\s*\(23\)/i.test(normalize(row?.resource)) ? '1523' : (row?.station || row?.facility || '—')

  const printDailyManagement = () => openPrintReport(
    'Daily Management — ניהול יומי',
    'יעדים, ביצוע, תחזית וקצב נדרש לכל משאב פעיל',
    [
      {key:'resource',label:'משאב יעד'},{key:'station',label:'תחנה'},{key:'line',label:'קו / משאב'},
      {key:'target',label:'יעד חודשי'},{key:'actual',label:'בפועל'},{key:'pct',label:'% ביצוע'},
      {key:'remaining',label:'נותר'},{key:'days',label:'ימים נותרו'},{key:'required',label:'נדרש ליום'},
      {key:'avg7',label:'ממוצע 7 ימים'},{key:'forecast',label:'תחזית'},{key:'status',label:'סטטוס'}
    ],
    planningRows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:planningDisplayStation(r),line:r.lineName||r.routingGroup||'—',target:fmt(r.target),actual:fmt(r.actual),pct:pctFmt(r.pct),remaining:fmt(r.remaining),days:r.remainingWorkdays,required:fmt(r.requiredDaily),avg7:fmt(r.recentAverage),forecast:fmt(r.forecast),status:r.label}))
  )


  const selectedFacilityLabel = () => selectedFacilities.length ? selectedFacilities.join(', ') : 'כל המתקנים'
  const rowMatchesSelectedFacilities = row => {
    if (!selectedFacilities.length) return true
    const ids = row.facilityIds || row.facilities || [row.station, row.facility].filter(Boolean)
    return ids.some(id => selectedFacilities.includes(String(id)))
  }
  const printFacilityOverview = () => {
    const rows = visibleControlTowerFacilities.filter(rowMatchesSelectedFacilities)
    openPrintReport('סקירת מתקנים', `מתקנים: ${selectedFacilityLabel()}`, [
      {key:'resource',label:'משאב / סימניה'},{key:'station',label:'תחנה'},{key:'target',label:'יעד'},{key:'actual',label:'בוצע'},
      {key:'forecast',label:'תחזית'},{key:'gap',label:'פער צפוי'},{key:'required',label:'קצב נדרש'},{key:'health',label:'Health Score'},{key:'deviations',label:'חריגות'}
    ], rows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:(r.facilityIds||r.facilities||[]).join(', ')||r.station||'—',target:fmt(r.target),actual:fmt(r.actual),forecast:fmt(r.forecast),gap:fmt(r.gap),required:fmt(r.requiredDaily),health:r.healthScore,deviations:r.deviationsCount})))
  }
  const printMonthlyForecast = () => {
    const rows = visiblePlanningRows.filter(rowMatchesSelectedFacilities)
    openPrintReport('תחזית חודשית לפי מתקן', `מתקנים: ${selectedFacilityLabel()}`, [
      {key:'resource',label:'משאב / סימניה'},{key:'station',label:'תחנה'},{key:'target',label:'יעד חודשי'},{key:'actual',label:'בוצע'},
      {key:'pct',label:'% ביצוע'},{key:'remaining',label:'נותר'},{key:'required',label:'נדרש ליום'},{key:'avg7',label:'ממוצע 7 ימים'},{key:'max',label:'שיא מוכח'},{key:'forecast',label:'תחזית'},{key:'status',label:'סטטוס'}
    ], rows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:planningDisplayStation(r),target:fmt(r.target),actual:fmt(r.actual),pct:pctFmt(r.pct),remaining:fmt(r.remaining),required:fmt(r.requiredDaily),avg7:fmt(r.recentAverage),max:fmt(r.provenMax),forecast:fmt(r.forecast),status:r.label})))
  }
  const printFacilityPerformance = () => {
    const rows = facilityStats.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.id))
    openPrintReport('ביצועים לפי מתקן בטווח המסונן', `מתקנים: ${selectedFacilityLabel()} · ${from || '—'} עד ${to || '—'}`, [
      {key:'facility',label:'מתקן'},{key:'qty',label:'כמות'},{key:'orders',label:'Orders'},{key:'batches',label:'מנות'}
    ], rows.map(r => ({facility:r.id,qty:fmt(r.actual),orders:r.orders||0,batches:new Set(baseFiltered.filter(x => x.facility === r.id).map(x => x.batch).filter(Boolean)).size})))
  }
  const printRecentProduction = () => {
    const rows = sortedRecentProduction.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.facility))
    openPrintReport('רשומות תפוקה אחרונות', `מתקנים: ${selectedFacilityLabel()} · ${from || '—'} עד ${to || '—'}`, [
      {key:'date',label:'תאריך'},{key:'ud',label:'החלטת שימוש (UD)'},{key:'facility',label:'משאב יעד'},{key:'routing',label:'מתקן / תחנה'},
      {key:'order',label:'הזמנה'},{key:'batch',label:'Batch'},{key:'material',label:'מק״ט חומר'},{key:'desc',label:'תיאור חומר'},{key:'qty',label:'כמות'}
    ], rows.map(r => ({date:iso(r.date),ud:productionUsageDecision(r),facility:r.facility,routing:r.routingGroup||'—',order:r.order||'—',batch:r.batch||'—',material:r.material||'—',desc:r.desc||'—',qty:fmt(r.qty)})))
  }

  const printMonthlyTargets = () => {
    const rows = targets.filter(t => !planningMonth || t.month === planningMonth)
    openPrintReport('קובץ יעדים חודשי', 'תצוגת יעדים ושיוכי תחנות לצורך בקרה ועדכון', [
      {key:'resource',label:'משאב / יעד'},{key:'station',label:'תחנה'},{key:'facilities',label:'מתקנים משויכים'},
      {key:'target',label:'יעד חודשי'},{key:'capacity',label:'קיבולת'},{key:'notes',label:'הערות'}
    ], rows.map(t => ({resource:t.resource||'—',station:t.station||t.facility||'—',facilities:(t.facilities||[t.facility]).filter(Boolean).join(', '),target:fmt(t.target),capacity:fmt(t.capacity),notes:t.notes||''})))
  }

  const clearAllData = async () => {
    if (!window.confirm('למחוק את כל הנתונים המשותפים מהענן? הפעולה תשפיע על כל המשתמשים.')) return
    setBusy(true)
    try {
      await deleteAllCloudDatasets(currentUser)
      setProduction([]); setQuality([]); setDeviations([]); setTargets([]); setDataMeta({ production:null, quality:null, deviations:null, targets:null }); localStorage.removeItem(STORAGE_KEY); await idbClear().catch(console.warn)
      setStatus('כל הנתונים נמחקו מהענן'); setCloudState({ mode:'cloud', lastSync:new Date().toISOString(), message:'מחובר למסד המשותף', latencyMs:cloudState.latencyMs, live:true }); setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPlanningMonth(''); setAdditionalFacilities([]); setFacilityToAdd(''); setPeriodYear(''); setPeriodQuarter('')
    } catch (error) { setStatus(`מחיקת הנתונים מהענן נכשלה: ${error.message}`) }
    finally { setBusy(false); setTimeout(() => setUploadProgress(null), 1800) }
  }
  const dailyTrend = useMemo(() => {
    const map = new Map(); filtered.forEach(r => { const d = iso(r.date); if (d) map.set(d, (map.get(d) || 0) + r.qty) })
    return [...map.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-14)
  }, [filtered])
  const maxDaily = Math.max(1, ...dailyTrend.map(([,v]) => v))
  const setQuickRange = (daysBack) => {
    if (!dateBounds.max) return
    const end = new Date(`${dateBounds.max}T12:00:00`), start = new Date(end); start.setDate(end.getDate() - daysBack + 1)
    setFrom(iso(start)); setTo(dateBounds.max); setPeriodYear(''); setPeriodQuarter('')
  }
  const applyYearFilter = (yearValue) => {
    setPeriodYear(yearValue); setPeriodQuarter('')
    if (!yearValue) { setFrom(''); setTo(''); return }
    setFrom(`${yearValue}-01-01`); setTo(`${yearValue}-12-31`)
  }
  const applyQuarterFilter = (quarterValue) => {
    setPeriodQuarter(quarterValue)
    if (!periodYear || !quarterValue) return
    const q = Number(quarterValue), startMonth = (q - 1) * 3 + 1, endMonth = startMonth + 2
    const lastDay = new Date(Number(periodYear), endMonth, 0).getDate()
    setFrom(`${periodYear}-${String(startMonth).padStart(2,'0')}-01`)
    setTo(`${periodYear}-${String(endMonth).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`)
  }
  const addFacility = () => {
    if (!facilityToAdd || additionalFacilities.includes(facilityToAdd)) return
    setAdditionalFacilities(current => [...current, facilityToAdd]); setFacilityToAdd('')
  }
  const removeAdditionalFacility = (id) => {
    setAdditionalFacilities(current => current.filter(x => x !== id))
    setSelectedFacilities(current => current.filter(x => x !== id))
  }

  const openHomeArea = (area) => {
    setShowHome(false)
    if (area === 'quality') setActiveTab('quality')
    if (area === 'production' || area === 'recent') setActiveTab('production')
    if (area === 'admin') setActiveTab(canManageData ? 'mapping-center' : 'production')
    window.setTimeout(() => {
      const targetId = area === 'planning' ? 'planning-section' : area === 'daily' ? 'daily-management-section' : area === 'quality' || area === 'recent' || area === 'admin' ? 'details-section' : 'control-tower-section'
      document.getElementById(targetId)?.scrollIntoView({ behavior:'smooth', block:'start' })
    }, 80)
  }

  if (showHome) return <div className="command-home" dir="rtl">
    <header className="command-home-header">
      <div className="command-home-brand"><img src="/icons/mark-128.png" alt="IML"/><div><strong>חדר בקרה — מתקני אריזה</strong><span>COMMAND CENTER</span></div></div>
      <div className="command-home-user"><Home size={19}/><b>דף ראשי</b><span></span><div><strong>{isGuest ? 'אורח' : (currentUser?.email || 'משתמש')}</strong><small>{isGuest ? 'צפייה בלבד' : userRole === 'admin' ? 'מנהל מערכת' : userRole === 'manager' ? 'מנהל מתקן' : 'צפייה בלבד'}</small></div></div>
    </header>
    <main className="command-home-main">
      <section className="command-home-status">
        <article><div className={`home-status-icon cloud ${cloudState.mode}`}><Cloud/></div><span>מצב מערכת</span><b>{cloudState.mode === 'cloud' ? 'מחובר לענן' : cloudState.mode === 'connecting' ? 'מתחבר...' : 'מצב מקומי'}</b><small>{cloudState.lastSync ? `עדכון אחרון ${new Date(cloudState.lastSync).toLocaleString('he-IL')}` : 'ממתין לסנכרון'}</small></article>
        <article><Database/><span>כמויות</span><b>{fmt(production.length)}</b><small>רשומות</small></article>
        <article><FlaskConical/><span>איכות</span><b>{fmt(quality.length)}</b><small>תוצאות</small></article>
        <article><AlertTriangle/><span>מנות חריגות</span><b>{fmt(openDeviations.length)}</b><small>פתוחות בטווח</small></article>
        <article className="home-range"><CalendarDays/><span>טווח תאריכים נוכחי</span><b>{from || '—'} — {to || '—'}</b><small>{planningMonth || 'ללא חודש תכנון'}</small></article>
      </section>
      <section className="command-home-welcome"><span>IML CONTROL</span><h1>ברוך הבא לחדר הבקרה</h1><p>בחר אזור לצפייה ולניתוח נתונים</p></section>
      <section className="command-home-cards">
        <button className="home-nav-card teal" onClick={()=>openHomeArea('production')}><div className="home-nav-icon"><BarChart3/></div><h2>סקירת מתקנים</h2><p>סקירה כוללת של כל המתקנים, ביצועים מול יעדים ותחזית חודשית</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card green" onClick={()=>openHomeArea('daily')}><div className="home-nav-icon"><ClipboardList/></div><h2>ניהול יומי</h2><p>ניהול ותפעול יומי של מתקנים, קווים, תקלות והערות</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card purple" onClick={()=>openHomeArea('planning')}><div className="home-nav-icon"><Factory/></div><h2>תחזית חודשית לפי מתקן</h2><p>תחזית ביצועים חודשית לפי מתקן עם ניתוח קצבים ועמידה ביעדים</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card orange" onClick={()=>openHomeArea('quality')}><div className="home-nav-icon"><FlaskConical/></div><h2>איכות</h2><p>סקירת איכות, מנות חריגות, תוצאות מעבדה וכרטיסי מנה</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card blue" onClick={()=>openHomeArea('recent')}><div className="home-nav-icon"><ClipboardList/></div><h2>רשומות תפוקה אחרונות</h2><p>רשומות התפוקה האחרונות ומעקב אחר המנות והחומרים שיוצרו</p><span>כניסה לאזור <ChevronLeft/></span></button>
      </section>
      <section className="command-home-bottom">
        <article className="home-attention"><div className="home-bottom-title"><BellRing/><h3>דורש תשומת לב</h3></div>{openDeviations.length ? <div className="home-alert-row"><AlertTriangle/><b>{openDeviations.length} מנות חריגות פתוחות</b><span>מומלץ לעבור למסך האיכות לבדיקה</span><button onClick={()=>openHomeArea('quality')}>לצפייה <ChevronLeft/></button></div> : <div className="home-ok"><CheckCircle2/> אין כרגע חריגות פתוחות בטווח שנבחר</div>}</article>
        <article className="home-quick"><div className="home-bottom-title"><Activity/><h3>פעולות מהירות</h3></div><div className="home-quick-grid"><button onClick={()=>openHomeArea('production')}><BarChart3/> סקירת מתקנים</button><button onClick={()=>openHomeArea('quality')}><FlaskConical/> איכות</button><button onClick={()=>openHomeArea('planning')}><Target/> יעדים ותחזית</button>{canManageData && <button onClick={()=>openHomeArea('admin')}><Settings2/> מרכז ניהול</button>}</div></article>
      </section>
    </main>
    <footer className="command-home-footer"><span>IML CONTROL © 2026</span><b>{BUILD_LABEL}</b><span className={cloudState.mode === 'cloud' ? 'online' : ''}>● {cloudState.mode === 'cloud' ? 'ONLINE' : 'OFFLINE'}</span></footer>
  </div>

  return <div className={`dashboard ${sidebarCollapsed ? 'sidebar-collapsed' : ''} ${managementMode ? 'management-mode' : ''}`} dir="rtl">
    <aside className="side filter-side">
      <button type="button" className="side-collapse-button" onClick={() => setSidebarCollapsed(v => !v)} title={sidebarCollapsed ? 'פתיחת מסננים' : 'כיווץ מסננים'}>{sidebarCollapsed ? <PanelRightOpen size={18}/> : <PanelRightClose size={18}/>}</button>
      <div className="brand branded"><img src="/icons/mark-128.png" alt="IML"/><div>IML<span>CONTROL</span></div></div>
      <div className="side-filter-title"><Search size={18}/><strong>חיפוש וסינון</strong></div>
      <label className="side-field"><span>Quick Search</span><div><Search size={16}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Order, Batch, חומר..."/></div></label>
      <label className="side-field"><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></label>
      <label className="side-field"><span>מתאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => { setFrom(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <label className="side-field"><span>עד תאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => { setTo(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <div className="side-field facility-multi-field"><span>מתקנים</span><button type="button" className={`facility-multi-trigger ${selectedFacilities.length ? 'active' : ''}`} onClick={() => setFacilityPickerOpen(v => !v)}><Factory size={16}/><b>{selectedFacilities.length ? `${selectedFacilities.length} מתקנים נבחרו` : 'כל המתקנים'}</b><span>⌄</span></button>{facilityPickerOpen && <div className="facility-multi-menu"><div className="facility-multi-menu-head"><strong>בחירת מתקנים</strong><button type="button" onClick={() => setSelectedFacilities([])}>נקה הכול</button></div>{facilities.map(id => <label key={id}><input type="checkbox" checked={selectedFacilities.includes(id)} onChange={() => toggleFacility(id)}/><span>{id}</span></label>)}</div>}</div>
      <div className="side-quick-ranges"><button onClick={() => setQuickRange(1)}>יום</button><button onClick={() => setQuickRange(2)}>יומיים</button><button onClick={() => setQuickRange(30)}>30 יום</button></div>
      <button className="side-clear" onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPeriodYear(''); setPeriodQuarter('') }}><X size={16}/> ניקוי מסננים</button>
      <div className="side-live-stats"><div><Database/><span><b>{fmt(production.length)}</b><small>תפוקה</small></span></div><div><FlaskConical/><span><b>{fmt(quality.length + deviations.length)}</b><small>איכות</small></span></div></div>
      <div className="side-note">{BUILD_LABEL} · {userRole === 'admin' ? 'Admin' : userRole === 'manager' ? 'Manager' : 'Viewer'}</div>
    </aside>

    <main className="main" id="control-tower-section">
      <header className="header">
        <div><h1>חדר בקרה — מתקני אריזה</h1><p>{BUILD_LABEL}</p></div>
        <div className="header-actions">
          <div className="ui-refresh-controls"><button type="button" className="action ui-control home-return" onClick={() => setShowHome(true)}><Home size={18}/><span>דף ראשי</span></button>
            <button type="button" className="action ui-control" onClick={() => setSidebarCollapsed(v => !v)}>{sidebarCollapsed ? <PanelRightOpen size={18}/> : <PanelRightClose size={18}/>}<span>{sidebarCollapsed ? 'פתח מסננים' : 'כווץ מסננים'}</span></button>
            <button type="button" className={`action ui-control ${managementMode ? 'active' : ''}`} onClick={() => setManagementMode(v => !v)}>{managementMode ? <Minimize2 size={18}/> : <Maximize2 size={18}/>}<span>{managementMode ? 'יציאה ממצב ניהולי' : 'מצב ניהולי'}</span></button>
          </div>
          <div className="user-session"><img className="user-brand-avatar" src="/icons/mark-64.png" alt="IML"/><span><b>{isGuest ? 'אורח' : (currentUser?.email || 'משתמש')}</b><small>{isGuest ? 'צפייה בלבד' : userRole === 'admin' ? 'מנהל מערכת' : userRole === 'manager' ? 'מנהל מתקן' : 'צפייה בלבד'}</small></span></div>
          <button className="action secondary" onClick={printMonthlyTargets} disabled={!targets.length}><Printer size={18}/> הדפסת יעדים</button>
          <button className="action secondary" onClick={downloadTargetWorkbook}><FileSpreadsheet size={18}/> הורדת תבנית יעדים</button>
          <button className="action secondary" onClick={exportWorkbook} disabled={!production.length}><Download size={18}/> יצוא Excel</button>
          {canDeleteData && <button className="action danger" onClick={clearAllData} disabled={!production.length && !quality.length && !deviations.length && !targets.length}><Trash2 size={18}/> מחיקה</button>}
          {canManageData && <label className={`upload ${busy ? 'disabled' : ''}`}><Upload size={19}/>{busy ? 'טוען...' : 'טעינת Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e => handleFiles([...e.target.files])}/></label>}
          <button className="action secondary" onClick={onSignOut}><LogOut size={18}/> יציאה</button>
        </div>
      </header>

      <section className="top-status-bar" aria-label="סטטוס מערכת">
        <div className={`top-status-item cloud ${cloudState.mode}`}><span className="status-dot"></span><small>ענן</small><b>{cloudState.mode === 'cloud' ? 'מחובר' : cloudState.mode === 'connecting' ? 'מתחבר' : 'מקומי'}</b></div>
        <div className="top-status-item"><small>עדכון אחרון</small><b>{cloudState.lastSync ? new Date(cloudState.lastSync).toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) : '—'}</b></div>
        <div className="top-status-item"><small>חודש פעיל</small><b>{planningMonth || '—'}</b></div>
        <div className="top-status-item"><small>תפוקה</small><b>{fmt(production.length)}</b></div>
        <div className="top-status-item"><small>איכות</small><b>{fmt(quality.length + deviations.length)}</b></div>
        <div className="top-status-item range"><small>טווח</small><b>{from || '—'} → {to || '—'}</b></div>
        {canManageData && <button type="button" className={`diagnostics-toggle ${diagnosticsOpen ? 'active' : ''}`} onClick={() => setDiagnosticsOpen(v => !v)} title="אבחון מערכת"><Activity size={16}/><span>אבחון</span></button>}
      </section>

      {(busy || cloudState.mode !== 'cloud' || /נכשל|שגיאה|ממתין|מתחבר|בודק|טוען|לא זמין/.test(status)) && <div className={`load-status compact ${cloudState.mode === 'error' || /נכשל|שגיאה/.test(status) ? 'error' : ''}`}><CheckCircle2 size={17}/>{status}</div>}
      {uploadProgress && <section className="upload-progress-card"><div className="upload-progress-head"><strong>{uploadProgress.fileName}</strong><span>{uploadProgress.percent}%</span></div><div className="upload-progress-track"><div style={{width:`${uploadProgress.percent}%`}}/></div><small>{uploadProgress.message}</small></section>}

      {canManageData && diagnosticsOpen && <section className="diagnostics-panel">
        <div className="diagnostics-panel-head"><div><Activity size={18}/><strong>אבחון מערכת</strong></div><button type="button" onClick={() => setDiagnosticsOpen(false)}><X size={16}/> סגור</button></div>
        <div className="diagnostics-grid">
          <article><small>Cache</small><b>{perfStats.cache}</b></article>
          <article><small>Queries</small><b>{perfStats.queries}</b></article>
          <article><small>Production</small><b>{production.length.toLocaleString()}</b></article>
          <article><small>Quality</small><b>{quality.length.toLocaleString()}</b></article>
          <article><small>In range</small><b>{(filtered.length + filteredQualityRows.length + filteredDeviationRows.length).toLocaleString()}</b></article>
          <article><small>Load</small><b>{perfStats.loadMs ? `${perfStats.loadMs}ms` : perfStats.phase}</b></article>
          <article><small>Data source</small><b>{cloudState.mode === 'cloud' ? 'Supabase' : 'Browser Cache'}</b></article>
          <article><small>Realtime</small><b>{cloudState.live ? 'פעיל' : 'לא פעיל'}</b></article>
          <article><small>Latency</small><b>{cloudState.latencyMs != null ? `${cloudState.latencyMs}ms` : '—'}</b></article>
          <article className="wide"><small>Range</small><b>{from || '—'} → {to || '—'}</b></article>
          <article className="wide"><small>Last sync</small><b>{cloudState.lastSync ? new Date(cloudState.lastSync).toLocaleString('he-IL') : '—'}</b></article>
        </div>
      </section>}

      {cloudState.mode !== 'cloud' && <section className={`cloud-status ${cloudState.mode}`}>
        <div className="cloud-status-icon">{cloudState.mode === 'offline' ? <WifiOff/> : <Cloud/>}</div>
        <div><strong>{cloudState.mode === 'connecting' ? 'מתחבר לענן' : cloudState.mode === 'offline' ? 'מצב מקומי זמני' : 'שגיאת סנכרון'}</strong><span>{cloudState.message}</span></div>
        <div className="cloud-status-meta"><small>מקור נתונים</small><b>{cloudState.mode === 'offline' ? 'Browser Cache' : 'Supabase'}</b><small>{cloudState.live ? '● עדכון חי פעיל' : '○ עדכון חי לא פעיל'}</small>{cloudState.latencyMs != null && <small>זמן תגובה: {cloudState.latencyMs}ms</small>}{cloudState.lastSync && <small>סנכרון: {new Date(cloudState.lastSync).toLocaleString('he-IL')}</small>}</div>
      </section>}


      <section className="control-tower-hero">
        <div className="control-tower-title"><div><span className="tower-kicker">PACKAGING CONTROL TOWER</span><h2>תמונת מצב ניהולית בזמן אמת</h2><p>יעדים, תחזית חודשית, איכות והזמנות — במבט אחד</p></div><div className={`tower-online ${cloudState.mode === 'cloud' ? 'online' : ''}`}><span></span>{cloudState.mode === 'cloud' ? 'ONLINE' : 'OFFLINE'}</div></div>
        <div className="tower-kpis">
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Target/><span>יעד חודשי כולל</span><b>{fmt(targetTotal)}</b><small>{planningMonth || 'ללא חודש נבחר'}</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Database/><span>בוצע החודש</span><b>{fmt(targetActual)}</b><small>{targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><TrendingUp/><span>תחזית סוף חודש</span><b>{fmt(targetForecast)}</b><small>{targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><Gauge/><span>קצב יומי נדרש</span><b>{fmt(planningRows.reduce((sum,row)=>sum+row.requiredDaily,0))}</b><small>לכל המתקנים</small></button>
          <button className={targetForecast >= targetTotal && targetTotal ? 'good' : 'bad'} onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><AlertTriangle/><span>פער צפוי</span><b>{targetTotal ? fmt(targetForecast-targetTotal) : '—'}</b><small>{targetForecast >= targetTotal ? 'מעל היעד' : 'נדרש להגביר קצב'}</small></button>
        </div>
      </section>

      <section className="tower-facility-section">
        <div className="panel-head facility-panel-head"><div><Factory/><h2>סקירת מתקנים</h2></div><div className="facility-head-actions">{facilityViewFilters}<span>{visibleControlTowerFacilities.length} מתקנים במעקב</span><button type="button" className="section-print-btn" onClick={printFacilityOverview}><Printer size={16}/> הדפסה</button><button type="button" className="section-toggle" onClick={() => setShowFacilityOverview(v => !v)}>{showFacilityOverview ? 'הסתר' : 'הצג'}</button></div></div>
        {showFacilityOverview && <div className="tower-facility-grid">
          {visibleControlTowerFacilities.map(row => <button key={row.facility} className={`tower-facility-card ${row.state}`} onClick={() => setSelectedResource(row)}>
            <div className="tower-facility-head"><div><i></i><strong>{row.facility}</strong></div><span>{row.state === 'good' ? 'תקין' : row.state === 'warning' ? 'דורש תשומת לב' : row.state === 'risk' ? 'בסיכון' : 'ללא יעד'}</span></div>
            <div className="tower-health"><div><HeartPulse/><span>Health Score</span></div><b>{row.healthScore}<small>/100</small></b></div>
            <div className="tower-progress"><i style={{width:`${Math.min(100,row.actualPct)}%`}}/></div>
            <dl><div><dt>יעד חודשי</dt><dd>{fmt(row.target)}</dd></div><div><dt>בוצע</dt><dd>{fmt(row.actual)}</dd></div><div><dt>תחזית</dt><dd>{fmt(row.forecast)}</dd></div><div><dt>פער צפוי</dt><dd className={row.gap >= 0 ? 'positive' : 'negative'}>{row.gap >= 0 ? '+' : ''}{fmt(row.gap)}</dd></div><div><dt>קצב נדרש</dt><dd>{fmt(row.requiredDaily)}</dd></div><div><dt>חריגות פתוחות</dt><dd>{row.deviationsCount}</dd></div></dl>
            <span className="tower-enter">לפרטים מלאים <ArrowLeft size={16}/></span>
          </button>)}
          {!visibleControlTowerFacilities.length && <div className="empty wide-empty">אין מתקנים התואמים למסנן התצוגה.</div>}
        </div>}
      </section>

      <section className="tower-lower-grid">
        <article className="tower-alerts-card"><div className="panel-head"><div><BellRing/><h2>התראות אחרונות</h2></div><button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}>הצג הכול</button></div><div className="tower-alert-list">{managerInsights.map((item,index)=><button key={index} className={item.state} onClick={() => item.title.includes('חריגות') ? jumpToDetails('deviations') : document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><i>{item.state==='risk'?'!':item.state==='warning'?'⚠':'✓'}</i><div><strong>{item.title}</strong><span>{item.text}</span></div></button>)}</div></article>
        <article className="tower-chart-card"><div className="panel-head"><div><BarChart3/><h2>ביצוע יומי — 7 ימים אחרונים</h2></div><span>{planningMonth}</span></div><div className="tower-mini-chart">{controlTowerTrend.map(([day,value])=>{const max=Math.max(1,...controlTowerTrend.map(x=>x[1])); return <div key={day}><b>{fmt(value)}</b><span><i style={{height:`${Math.max(8,value/max*100)}%`}}/></span><small>{day.slice(5)}</small></div>})}{!controlTowerTrend.length&&<div className="empty">אין נתונים להצגת מגמה</div>}</div></article>
      </section>

      <section className="data-center">
        <div className="panel-head"><div><ShieldCheck/><h2>מרכז נתונים</h2></div><span>4 מקורות מידע</span></div>
        <p className="data-center-help">כל קובץ נבדק בדפדפן ולאחר מכן נשמר ב־Supabase. מרגע שהטעינה מסתיימת, אותו מידע זמין לכל המשתמשים המחוברים.</p>
        <div className="data-source-grid">
          <DataSource title="תפוקות" icon={<Factory/>} meta={dataMeta.production} count={production.length} acceptLabel="טען קובץ תפוקות" busy={busy} onFiles={files => loadFiles(files, 'production')} canManage={canManageData}/>
          <DataSource title="תוצאות איכות" icon={<FlaskConical/>} meta={dataMeta.quality} count={quality.length} acceptLabel="הוסף תוצאות איכות חדשות" busy={busy} onFiles={files => loadFiles(files, 'quality')} canManage={canManageData}/>
          <DataSource title="חריגות איכות" icon={<AlertTriangle/>} meta={dataMeta.deviations} count={deviations.length} acceptLabel="טען קובץ חריגות" busy={busy} onFiles={files => loadFiles(files, 'deviations')} canManage={canManageData}/>
          <DataSource title="יעדים חודשיים" icon={<Target/>} meta={dataMeta.targets} count={targets.length} acceptLabel="טען קובץ יעדים" busy={busy} onFiles={files => loadFiles(files, 'targets')} canManage={canManageData}/>
        </div>
      </section>

      <section className="planning-toolbar">
        <div><Target/><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></div>
        <small>היעד היומי אינו מוזן: הוא מחושב מחדש בכל יום לפי יתרת היעד והימים שנותרו בחודש.</small>
      </section>

      <section className="executive-strip executive-six">
        <Executive icon={<Target/>} title="ביצוע מול יעד" value={targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} sub={`${fmt(targetActual)} מתוך ${fmt(targetTotal)}`} onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<TrendingUp/>} title="תחזית סוף חודש" value={targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} sub={fmt(targetForecast)} onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<CheckCircle2/>} title="במסלול / הושג" value={achievedCount} sub="מתקנים" good onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<AlertTriangle/>} title="בסיכון" value={warningCount} sub="נדרש שיפור" warn onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<BellRing/>} title="חריגות פתוחות" value={openDeviations.length} sub="מנות איכות" bad onClick={() => jumpToDetails('deviations')}/>
        <Executive icon={<ClipboardList/>} title="Orders בטווח" value={uniqueOrders} sub={`${uniqueBatches} מנות`} onClick={() => jumpToDetails('production')}/>
      </section>

      <section className="manager-brief">
        <div className="panel-head"><div><Activity/><h2>תמונת מצב ניהולית</h2></div><span>מבוסס על הנתונים הטעונים</span></div>
        <div className="manager-insights">{managerInsights.map((item,i) => <button key={i} className={`manager-insight ${item.state}`} onClick={() => item.title.includes('חריגות') ? jumpToDetails('deviations') : document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><strong>{item.title}</strong><span>{item.text}</span></button>)}</div>
      </section>

      <section className="filters legacy-filters">
        <label><Search size={17}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש מתקן, Order, Batch או חומר"/></label>
        <label><CalendarDays size={17}/>שנה<select value={periodYear} onChange={e => applyYearFilter(e.target.value)}><option value="">כל השנים</option>{availableYears.map(y => <option key={y} value={y}>{y}</option>)}</select></label>
        <label><CalendarDays size={17}/>רבעון<select value={periodQuarter} disabled={!periodYear} onChange={e => applyQuarterFilter(e.target.value)}><option value="">כל השנה</option><option value="1">רבעון 1</option><option value="2">רבעון 2</option><option value="3">רבעון 3</option><option value="4">רבעון 4</option></select></label>
        <label><CalendarDays size={17}/>מתאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => { setFrom(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
        <label><CalendarDays size={17}/>עד תאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => { setTo(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
        <button onClick={() => setQuickRange(1)}>יום אחרון</button><button onClick={() => setQuickRange(2)}>יומיים</button><button onClick={() => setQuickRange(30)}>30 יום</button>
        <button onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPeriodYear(''); setPeriodQuarter('') }}>נקה</button>
      </section>

      <section className="extra-facilities">
        <div><Factory size={18}/><strong>מתקנים נוספים</strong><span>ברירת המחדל מציגה רק את מתקני הליבה.</span></div>
        <div className="extra-facility-actions"><select value={facilityToAdd} onChange={e => setFacilityToAdd(e.target.value)}><option value="">בחר מתקן נוסף</option>{optionalFacilities.map(id => <option key={id} value={id}>{id}</option>)}</select><button onClick={addFacility} disabled={!facilityToAdd}>הוסף מתקן</button></div>
        {!!additionalFacilities.length && <div className="extra-facility-chips">{additionalFacilities.map(id => <button key={id} onClick={() => removeAdditionalFacility(id)}>{id}<X size={14}/></button>)}</div>}
      </section>

      {selectedFacilities.length > 0 && <div className="selection"><div className="selection-info"><span>מתקנים נבחרים:</span><div className="selection-chips">{selectedFacilities.map(id => <button className="selection-chip" key={id} onClick={() => toggleFacility(id)}>{id}<X size={14}/></button>)}</div></div><button className="clear-selection" onClick={() => setSelectedFacilities([])}><X size={16}/> הסר הכול</button></div>}

      <section className="summary-grid">
        <Summary title="סה״כ תפוקה מסוננת" value={fmt(total)} sub="ליטר / ק״ג לפי הקובץ"/>
        <Summary title="מתקנים פעילים" value={activeFacilities} sub={`מתוך ${facilities.length} מתקנים`}/>
        <Summary title="משמרת בוקר" value={fmt(morningQty)} sub="07:00–15:00"/>
        <Summary title="משמרת ערב" value={fmt(eveningQty)} sub="15:00–23:00"/>
        <Summary title="משמרת לילה" value={fmt(nightQty)} sub="23:00–07:00"/>
        <Summary title="חריגות פתוחות" value={openDeviations.length} sub="לפי קובץ החריגות" warn/>
        <Summary title="איכות לא תקינה" value={qualityBad.length} sub="לפי קובץ האיכות" warn/>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Gauge/><div><h2>תחזית חודשית לפי מתקן</h2><p>יעד חודשי, קצב נדרש, קצב אחרון, שיא מוכח ותחזית</p></div></div><div className="facility-head-actions">{facilityViewFilters}<button type="button" className="section-print-btn" onClick={printMonthlyForecast}><Printer size={16}/> הדפסה</button><button type="button" className="section-toggle" onClick={() => setShowMonthlyForecast(v => !v)}>{showMonthlyForecast ? 'הסתר' : 'הצג'}</button><button className="select-all-facilities" onClick={toggleAllFacilities}><CheckCircle2 size={17}/>{allFacilitiesSelected ? 'ביטול בחירת הכול' : 'בחירת כל המתקנים'}</button></div></div>
      {showMonthlyForecast && <section className="forecast-grid" id="planning-section">
        {visiblePlanningRows.map(row => <ForecastCard key={row.id} {...row} selected={(row.facilities || [row.facility]).some(id => selectedFacilities.includes(id))} onClick={() => setSelectedFacilities(row.facilities || [row.facility])}/>) }
        {!visiblePlanningRows.length && <div className="empty wide-empty">אין מתקנים התואמים למסנן התצוגה.</div>}
      </section>}

      <section className="alert-panel" id="alerts-section">
        <div className="panel-head"><div><BellRing/><h2>מה דורש תשומת לב היום?</h2></div><span>{alerts.length} התראות</span></div>
        <div className="alert-list">
          {alerts.map(r => <div className={`alert-item ${r.state}`} key={r.id}><div className="alert-symbol">{r.state === 'risk' ? '!' : '⚠'}</div><div><strong>{planningName(r)} — {r.label}</strong><p>{r.state === 'risk' ? `נדרש ${fmt(r.requiredDaily)} ליום, אך השיא המוכח הוא ${fmt(r.provenMax)}.` : `התחזית היא ${fmt(r.forecast)} מול יעד ${fmt(r.target)}. נדרש קצב של ${fmt(r.requiredDaily)} ליום.`}</p></div></div>)}
          {!alerts.length && <div className="empty">אין התראות תכנון לחודש הנבחר.</div>}
        </div>
      </section>

      <section className="daily-management" id="daily-management-section">
        <div className="panel-head"><div><CalendarCheck/><h2>Daily Management</h2></div><div className="daily-print-actions"><span>{planningMonth}</span><button type="button" className="action secondary daily-print-btn" onClick={printDailyManagement} disabled={!planningRows.length}><Printer size={17}/> הדפסה צבעונית</button></div></div>
        <div className="table-wrap"><table><thead><tr><th>משאב יעד</th><th>מתקן / תחנה</th><th>תחנה / קו</th><th>פעילות</th><th>יעד חודשי</th><th>בפועל</th><th>% ביצוע</th><th>נותר</th><th>ימים נותרו</th><th>נדרש ליום</th><th>ממוצע 7 ימים</th><th>שיא מוכח</th><th>תחזית</th><th>סטטוס</th></tr></thead><tbody>
          {visiblePlanningRows.map(r => <tr key={r.id}><td><b>{r.resource || r.facility}</b></td><td>{planningDisplayStation(r)}</td><td>{[planningDisplayStation(r), r.lineName].filter(Boolean).join(' · ') || '—'}</td><td>{r.activity}</td><td>{fmt(r.target)}</td><td>{fmt(r.actual)}</td><td>{pctFmt(r.pct)}</td><td>{fmt(r.remaining)}</td><td>{r.remainingWorkdays}</td><td>{fmt(r.requiredDaily)}</td><td>{fmt(r.recentAverage)}</td><td>{fmt(r.provenMax)}</td><td>{fmt(r.forecast)}</td><td><StatusBadge state={r.state} label={r.label}/></td></tr>)}
          {!visiblePlanningRows.length && <tr><td colSpan="14" className="empty">אין מתקנים התואמים למסנן התצוגה</td></tr>}
        </tbody></table></div>
      </section>

      <section className="trend-card">
        <div className="trend-head"><div><h2>מגמת תפוקה יומית</h2><p>14 הימים האחרונים בטווח הנבחר</p></div><Save size={20}/></div>
        <div className="trend-bars">{dailyTrend.map(([date, value]) => <div className="trend-item" key={date} title={`${date}: ${fmt(value)}`}><div className="trend-value">{fmt(value)}</div><div className="trend-track"><i style={{height: `${Math.max(5, value / maxDaily * 100)}%`}}/></div><small>{date.slice(5)}</small></div>)}{!dailyTrend.length && <div className="empty trend-empty">טען קובץ תפוקות להצגת מגמה</div>}</div>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Factory/><div><h2>ביצועים לפי מתקן בטווח המסונן</h2><p>לחיצה על כרטיס מסננת תפוקה, איכות וחריגות</p></div></div><button type="button" className="section-print-btn" onClick={printFacilityPerformance}><Printer size={16}/> הדפסה</button></div>
      <section className="facility-grid">{facilityStats.map(x => <Facility key={x.id} {...x} selected={selectedFacilities.includes(x.id)} onClick={() => toggleFacility(x.id)}/>)}</section>

      <section className="tabs" id="details-section">
        <button className={activeTab === 'production' ? 'active' : ''} onClick={() => setActiveTab('production')}><BarChart3 size={16}/> תפוקה</button>
        <button className={activeTab === 'shifts' ? 'active' : ''} onClick={() => setActiveTab('shifts')}><Clock3 size={16}/> ניתוח משמרות</button>
        <button className={activeTab === 'quality' ? 'active' : ''} onClick={() => setActiveTab('quality')}><FlaskConical size={16}/> איכות ({qualityBad.length})</button>
        <button className={activeTab === 'deviations' ? 'active' : ''} onClick={() => setActiveTab('deviations')}><AlertTriangle size={16}/> מנות חריגות ({openDeviations.length})</button>
        {canManageData && <button className={activeTab === 'mapping-simulator' ? 'active' : ''} onClick={() => setActiveTab('mapping-simulator')}><ClipboardList size={16}/> סימולטור שיוך ({mappingSimulation.summary.unmatched + mappingSimulation.summary.duplicate})</button>}
        {canManageData && <button className={activeTab === 'mapping-center' ? 'active' : ''} onClick={() => setActiveTab('mapping-center')}><ShieldCheck size={16}/> מרכז מיפויים ({manualMappings.filter(item => item.active !== false && item.status === 'pending').length})</button>}
      </section>
      {activeTab === 'production' && <section className="details"><div className="details-title-row"><h2>רשומות תפוקה אחרונות</h2><div className="details-title-actions"><span className="details-note">לחיצה על כותרת עמודה ממיינת מקטן לגדול / מהגדול לקטן</span><button type="button" className="section-print-btn" onClick={printRecentProduction}><Printer size={16}/> הדפסה</button></div></div><div className="table-wrap"><table className="sortable-production-table" data-smart-sum-column="8"><thead><tr><th><button type="button" onClick={()=>toggleProductionSort('date')}>תאריך{productionSortArrow('date')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('ud')}>החלטת שימוש (UD){productionSortArrow('ud')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('facility')}>משאב יעד{productionSortArrow('facility')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('routingGroup')}>מתקן / תחנה{productionSortArrow('routingGroup')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('order')}>הזמנה{productionSortArrow('order')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('batch')}>Batch{productionSortArrow('batch')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('material')}>מק״ט חומר{productionSortArrow('material')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('desc')}>תיאור חומר{productionSortArrow('desc')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('qty')}>כמות{productionSortArrow('qty')}</button></th></tr></thead><tbody>{sortedRecentProduction.map((r, i) => <tr key={`${r.order}-${r.batch}-${i}`}><td>{iso(r.date)}</td><td>{productionUsageDecision(r)}</td><td>{r.facility}</td><td>{r.routingGroup || '—'}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material || '—'}</td><td>{r.desc || '—'}</td><td>{fmt(r.qty)}</td></tr>)}{!sortedRecentProduction.length && <tr className="smart-empty-row"><td colSpan="9" className="empty">אין רשומות להצגה</td></tr>}</tbody></table></div></section>}
      {activeTab === 'mapping-simulator' && canManageData && <section className="details mapping-simulator">
        <div className="mapping-simulator-head"><div><h2>סימולטור שיוך תפוקה</h2><p className="details-note">המסך מציג לאיזה יעד כל רשומת SAP נכנסת. בשורה בעייתית לחץ "שייך" כדי לראות את ההשפעה לפני שמירה.</p></div><div className="mapping-simulator-actions"><label><input type="checkbox" checked={simulatorOnlyIssues} onChange={event => setSimulatorOnlyIssues(event.target.checked)}/> הצג רק בעיות</label><button type="button" onClick={exportMappingSimulation}><Download size={16}/> ייצוא סימולציה</button></div></div>
        {mappingMessage && <div className="mapping-message">{mappingMessage}</div>}
        <div className="mapping-summary-grid"><article><span>רשומות בטווח</span><b>{fmt(mappingSimulation.summary.rows)}</b><small>{fmt(mappingSimulation.summary.quantity)} כמות</small></article><article className="mapping-ok"><span>שויכו תקין</span><b>{fmt(mappingSimulation.summary.matched)}</b><small>{fmt(mappingSimulation.summary.matchedQty)} כמות</small></article><article className="mapping-unmatched"><span>לא שויכו</span><b>{fmt(mappingSimulation.summary.unmatched)}</b><small>{fmt(mappingSimulation.summary.unmatchedQty)} כמות</small></article><article className="mapping-duplicate"><span>שיוך כפול</span><b>{fmt(mappingSimulation.summary.duplicate)}</b><small>{fmt(mappingSimulation.summary.duplicateQty)} כמות</small></article></div>
        <div className="table-wrap mapping-table-wrap"><table><thead><tr><th>סטטוס</th><th>תאריך</th><th>תחנה</th><th>משפחה</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור מוצר</th><th>כמות</th><th>יעד</th><th>מקור</th><th>פעולה</th></tr></thead><tbody>{visibleMappingSimulation.slice(0,1000).map(item => <tr key={item.key} className={`mapping-row mapping-${item.status}`}><td><span className={`mapping-status mapping-status-${item.status}`}>{item.status === 'matched' ? 'תקין' : item.status === 'duplicate' ? 'כפול' : 'לא שויך'}</span></td><td>{item.row.productionDay || iso(item.row.date)}</td><td>{item.row.facility || '—'}</td><td>{item.family || '—'}</td><td>{item.row.order || '—'}</td><td>{item.row.batch || '—'}</td><td>{item.row.material || '—'}</td><td>{item.row.desc || '—'}</td><td>{fmt(item.row.qty)}</td><td>{item.assignedResource || item.pendingManual?.targetResource || '—'}</td><td>{item.approvedManual ? 'מיפוי ידני מאושר' : item.pendingManual ? 'ממתין לאישור' : item.explanation}</td><td>{item.family !== '1542' && item.status !== 'matched' && !item.pendingManual ? <button className="mapping-assign-btn" onClick={() => openMappingDialog(item)}>שייך</button> : item.pendingManual ? <span className="mapping-pending-chip">ממתין</span> : '—'}</td></tr>)}{!visibleMappingSimulation.length && <tr><td colSpan="12" className="empty">לא נמצאו בעיות שיוך בטווח שנבחר</td></tr>}</tbody></table></div>
      </section>}

      {activeTab === 'mapping-center' && canManageData && <section className="details mapping-center">
        <div className="mapping-simulator-head"><div><h2>מרכז בקרת מיפויים</h2><p className="details-note">אישור, דחייה וביטול של שיוכים. שיוך משפיע על הדשבורד רק לאחר אישור.</p></div></div>
        <div className="mapping-summary-grid"><article className="mapping-duplicate"><span>ממתינים לאישור</span><b>{manualMappings.filter(item => item.active !== false && item.status === 'pending').length}</b></article><article className="mapping-ok"><span>מאושרים</span><b>{manualMappings.filter(item => item.active !== false && item.status === 'approved').length}</b></article><article className="mapping-unmatched"><span>נדחו / בוטלו</span><b>{manualMappings.filter(item => item.status === 'rejected' || item.active === false).length}</b></article><article><span>אירועי היסטוריה</span><b>{mappingTimeline.length}</b></article></div>
        <div className="table-wrap"><table><thead><tr><th>סטטוס</th><th>מתקן</th><th>מק״ט</th><th>תיאור</th><th>יעד</th><th>נוצר ע״י</th><th>תאריך</th><th>פעולות</th></tr></thead><tbody>{manualMappings.map(mapping => <tr key={mapping.id}><td><span className={`mapping-status mapping-status-${mapping.status === 'approved' ? 'matched' : mapping.status === 'pending' ? 'duplicate' : 'unmatched'}`}>{mapping.status === 'approved' ? 'מאושר' : mapping.status === 'pending' ? 'ממתין' : 'נדחה'}</span></td><td>{mapping.family}</td><td>{mapping.material || '—'}</td><td>{mapping.description || '—'}</td><td>{mapping.targetResource}</td><td>{mapping.createdBy || 'מנהל'}</td><td>{mapping.createdAt ? new Date(mapping.createdAt).toLocaleString('he-IL') : '—'}</td><td><div className="mapping-row-actions">{mapping.status === 'pending' && <><button onClick={() => updateMappingStatus(mapping,'approved')}>אשר</button><button className="danger" onClick={() => updateMappingStatus(mapping,'rejected')}>דחה</button></>}{mapping.status === 'approved' && mapping.active !== false && <button className="danger" onClick={() => rollbackMapping(mapping)}>בטל שיוך</button>}</div></td></tr>)}{!manualMappings.length && <tr><td colSpan="8" className="empty">עדיין לא נוצרו מיפויים ידניים</td></tr>}</tbody></table></div>
        <div className="mapping-timeline"><h3>Timeline החלטות</h3>{mappingTimeline.slice(0,30).map(event => <div key={event.id}><time>{new Date(event.at).toLocaleString('he-IL')}</time><b>{event.action}</b><span>{event.material || '—'} → {event.targetResource || '—'}</span><small>{event.user} · {event.note}</small></div>)}</div>
      </section>}

      {mappingDialog && <div className="mapping-dialog-backdrop"><div className="mapping-dialog"><button className="mapping-dialog-close" onClick={() => setMappingDialog(null)}><X/></button><h2>שיוך חומר ליעד</h2><div className="mapping-source-card"><b>{mappingDialog.row.material || 'ללא מק״ט'}</b><span>{mappingDialog.row.desc || 'ללא תיאור'}</span><small>תחנה {mappingDialog.row.facility} · משפחה {mappingDialog.family} · כמות {fmt(mappingDialog.row.qty)}</small></div><label>בחר יעד<select value={mappingTargetKey} onChange={event => setMappingTargetKey(event.target.value)}><option value="">בחר יעד...</option>{mappingDialog.candidates.map(target => <option key={target.key} value={target.key}>{target.resource} — יעד {fmt(target.target)} — בוצע {fmt(target.actual)}</option>)}</select></label>{(() => { const target=mappingTargets.find(item => item.key===mappingTargetKey); if(!target) return null; const impacted=mappingSimulation.rows.filter(item => item.family===mappingDialog.family && productionMappingKey(item.row)===productionMappingKey(mappingDialog.row)); const qty=impacted.reduce((sum,item)=>sum+(Number(item.row.qty)||0),0); const before=target.actual||0; const after=before+qty; return <div className="mapping-impact"><h3>סימולציית השפעה</h3><div><span>כמות שתשויך</span><b>{fmt(qty)}</b></div><div><span>ביצוע לפני</span><b>{fmt(before)}</b></div><div><span>ביצוע אחרי</span><b>{fmt(after)}</b></div><div><span>עמידה ביעד</span><b>{target.target ? `${Math.round(before/target.target*100)}% → ${Math.round(after/target.target*100)}%` : 'ללא יעד'}</b></div><p>השינוי לא ישפיע על הדשבורד עד לאישור במרכז המיפויים.</p></div> })()}<button className="mapping-save-btn" disabled={!mappingTargetKey} onClick={savePendingMapping}><Save size={17}/> שמור כממתין לאישור</button></div></div>}

      {activeTab === 'shifts' && <section className="details shift-intelligence"><h2>ניתוח משמרות — בוקר, ערב ולילה</h2><p className="details-note">החלפות מוצר מזוהות לפי שינוי מק״ט באותו מתקן ו-Routing group. זמן המעבר הוא פער זמן משוער בין שני דיווחים עוקבים.</p><div className="shift-card-grid">{shiftAnalysis.map(item => <article className={`shift-analysis-card shift-${item.key}`} key={item.key}><div className="shift-card-head"><div><h3>{item.label}</h3><span>{item.hours}</span></div><strong>{fmt(item.total)}</strong></div><div className="shift-metrics"><div><span>קצב ממוצע לשעה</span><b>{fmt(item.avgPerHour)}</b></div><div><span>Orders</span><b>{item.orders}</b></div><div><span>Batch</span><b>{item.batches}</b></div><div><span>מק״טים</span><b>{item.materials}</b></div><div><span>החלפות מוצר</span><b>{item.changeovers.length}</b></div><div><span>חריגות איכות</span><b>{item.deviations}</b></div><div><span>ממוצע מעבר</span><b>{fmt(item.avgChangeover)} דק׳</b></div><div><span>תרומה לתפוקה</span><b>{pctFmt(item.share)}</b></div></div></article>)}</div><h3 className="shift-subtitle">פירוט החלפות מוצר</h3><div className="table-wrap"><table><thead><tr><th>משמרת</th><th>שעה</th><th>משאב יעד</th><th>מתקן / תחנה</th><th>מוצר קודם</th><th>מוצר חדש</th><th>פער דיווח משוער</th></tr></thead><tbody>{shiftAnalysis.flatMap(item => item.changeovers.map((c,i) => <tr key={`${item.key}-${i}`}><td>{item.label}</td><td>{c.at?.toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}</td><td>{c.facility}</td><td>{c.routingGroup || '—'}</td><td>{c.fromMaterial}{c.fromDesc ? ` · ${c.fromDesc}` : ''}</td><td>{c.toMaterial}{c.toDesc ? ` · ${c.toDesc}` : ''}</td><td>{fmt(c.minutes)} דקות</td></tr>))}{!shiftAnalysis.some(item => item.changeovers.length) && <tr><td colSpan="7" className="empty">לא זוהו החלפות מוצר בטווח שנבחר</td></tr>}</tbody></table></div></section>}
      {activeTab === 'quality' && <section className="details"><h2>תוצאות איכות לא תקינות</h2><div className="table-wrap"><table><thead><tr><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Inspection Lot</th><th>Order</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th></tr></thead><tbody>{qualityBad.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? new Date(r.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.inspectionLot}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'ללא סטטוס'}</span></td></tr>)}{!qualityBad.length && <tr><td colSpan="8" className="empty">לא נמצאו תוצאות איכות לא תקינות</td></tr>}</tbody></table></div></section>}
      {activeTab === 'deviations' && <section className="details"><h2>מנות חריגות פתוחות</h2><p className="details-note">לכל מנה מוצגים מאפייני החריגה ולצדם המאפיינים התקינים שנמשכו מקובץ תוצאות האיכות לפי Batch + מק״ט.</p><div className="table-wrap"><table><thead><tr><th>תאריך חריגה</th><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th><th>מאפייני החריגה</th><th>מאפיינים תקינים</th><th>הערות</th></tr></thead><tbody>{openDeviations.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{iso(r.sampleDate) || '—'}</td><td>{r.sampleDate ? new Date(r.sampleDate).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'פתוח'}</span>{r.udCode && <small className="ud-code">{r.udCode}</small>}</td><td className="deviation-characteristics"><div className="characteristics-count bad-count">{r.rejectedCharacteristics.length} חריגים</div>{r.rejectedCharacteristics.length ? r.rejectedCharacteristics.map((c,j) => <div className="deviation-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו פרטי מאפיינים חריגים בקובץ האיכות{r.rejectedCount ? ` (בקובץ החריגות מופיע מספר: ${r.rejectedCount})` : ''}</span>}</td><td className="deviation-characteristics valid-characteristics"><div className="characteristics-count good-count">{r.approvedCharacteristics.length} תקינים</div>{r.approvedCharacteristics.length ? r.approvedCharacteristics.map((c,j) => <div className="deviation-characteristic valid-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו מאפיינים תקינים למנה בקובץ האיכות</span>}</td><td>{r.remarks || '—'}</td></tr>)}{!openDeviations.length && <tr><td colSpan="10" className="empty">לא נמצאו מנות חריגות פתוחות</td></tr>}</tbody></table></div></section>}
    </main>
    {selectedResource && <ResourceDetailModal resource={selectedResource} onClose={() => setSelectedResource(null)} onOpenBatch={(batch, material='') => { setSelectedResource(null); openBatchCard(batch, material) }}/>}
    {selectedBatchData && <BatchControlCard data={selectedBatchData} onClose={() => { setSelectedBatch(''); setSelectedBatchMaterial('') }}/>}
  </div>
}

function ResourceDetailModal({ resource, onClose, onOpenBatch }) {
  useEffect(() => {
    const previous = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const close = event => { if (event.key === 'Escape') onClose() }
    window.addEventListener('keydown', close)
    return () => { document.body.style.overflow = previous; window.removeEventListener('keydown', close) }
  }, [onClose])
  const rows = resource.productionRows || []
  const batches = [...new Set(rows.map(row => row.batch).filter(Boolean))]
  const orders = [...new Set(rows.map(row => row.order).filter(Boolean))]
  const materials = [...rows.reduce((map,row) => {
    const key = normalize(row.material) || 'ללא מק״ט'
    const current = map.get(key) || { material:key, description:row.desc || '', qty:0, batches:new Set(), orders:new Set(), rows:0 }
    current.qty += Number(row.qty) || 0
    current.rows += 1
    if (row.batch) current.batches.add(row.batch)
    if (row.order) current.orders.add(row.order)
    if (!current.description && row.desc) current.description = row.desc
    map.set(key,current)
    return map
  }, new Map()).values()].sort((a,b)=>b.qty-a.qty)
  const calculatedActual = rows.reduce((sum,row)=>sum+(Number(row.qty)||0),0)
  const progress = resource.target ? Math.min(100, resource.actual / resource.target * 100) : 0
  const printRows = () => {
    const popup = window.open('', '_blank', 'width=1400,height=900')
    if (!popup) return
    const esc = value => String(value ?? '').replace(/[&<>\"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[ch]))
    const body = rows.slice().sort((a,b)=>String(b.productionDay||'').localeCompare(String(a.productionDay||''))).map((row,index)=>`<tr><td>${index+1}</td><td>${esc(row.productionDay || iso(row.date) || '—')}</td><td>${esc(row.facility||'—')}</td><td>${esc(row.routingGroup||'—')}</td><td>${esc(row.orderType||'—')}</td><td>${esc(row.order||'—')}</td><td>${esc(row.batch||'—')}</td><td>${esc(row.material||'—')}</td><td>${esc(row.desc||'—')}</td><td>${esc(fmt(row.qty))}</td></tr>`).join('')
    popup.document.write(`<!doctype html><html dir="rtl"><head><meta charset="utf-8"><title>${esc(resource.resource)}</title><style>@page{size:A4 landscape;margin:8mm}body{font-family:Arial,sans-serif;color:#173b57}.head{display:flex;justify-content:space-between;border-bottom:4px solid #159b83;padding-bottom:12px;margin-bottom:14px}.head h1{margin:0}.meta{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}.meta span{background:#eaf7f3;padding:7px 10px;border-radius:10px;font-weight:700}table{width:100%;border-collapse:collapse;font-size:10px}th{background:#173b57;color:#fff;padding:7px}td{padding:6px;border-bottom:1px solid #dce7ec;text-align:center}tr:nth-child(even) td{background:#f7fafc}</style></head><body><div class="head"><div><h1>${esc(resource.resource)}</h1><p>${esc([resource.station&&`Station ${resource.station}`,resource.description,resource.lineName].filter(Boolean).join(' · '))}</p></div><b>IML CONTROL</b></div><div class="meta"><span>יעד ${esc(fmt(resource.target))}</span><span>בוצע ${esc(fmt(resource.actual))}</span><span>תחזית ${esc(fmt(resource.forecast))}</span><span>מק״טים ${materials.length}</span><span>מנות ${batches.length}</span></div><table><thead><tr><th>#</th><th>תאריך</th><th>תחנה</th><th>Routing</th><th>Order Type</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור</th><th>כמות</th></tr></thead><tbody>${body}</tbody></table><script>window.onload=()=>setTimeout(()=>window.print(),200)<\/script></body></html>`)
    popup.document.close()
  }

  const exportRows = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(materials.map(item => ({
      Material:item.material, Description:item.description, Quantity:item.qty, Batches:item.batches.size, Orders:item.orders.size, Rows:item.rows,
    }))), 'Materials')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(row => ({
      Date:row.productionDay || iso(row.date), Facility:row.facility, RoutingGroup:row.routingGroup, OrderType:row.orderType,
      Order:row.order, Batch:row.batch, Material:row.material, Description:row.desc, Quantity:row.qty,
    }))), 'Calculation Rows')
    XLSX.writeFile(wb, `IML_${String(resource.resource || 'resource').replace(/[^a-zA-Z0-9_-]+/g,'_')}_${planningSafeMonth(resource)}.xlsx`)
  }
  return <div className="resource-modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) onClose() }}>
    <section className="resource-detail-modal resource-detail-modal-audit" role="dialog" aria-modal="true" aria-label={`פירוט ${resource.resource}`}>
      <header><div><span>RESOURCE CONTROL CENTER</span><h2>{resource.resource}</h2><p>{[resource.station && `Station ${resource.station}`, resource.description, resource.lineName].filter(Boolean).join(' · ')}</p></div><div className="resource-head-actions"><button className="resource-export-btn" onClick={printRows}><Printer size={17}/> הדפסת סימניה</button><button className="resource-export-btn" onClick={exportRows}><FileSpreadsheet size={17}/> ייצוא חישוב</button><button onClick={onClose} aria-label="סגירה"><X/></button></div></header>
      <div className={`resource-status-banner ${resource.state}`}><strong>{resource.label}</strong><span>תחזית {fmt(resource.forecast)} מול יעד {fmt(resource.target)}</span></div>
      <div className="resource-detail-kpis">
        <BatchMetric label="יעד חודשי" value={fmt(resource.target)}/><BatchMetric label="בוצע" value={fmt(resource.actual)}/><BatchMetric label="סכום רשומות" value={fmt(calculatedActual)}/><BatchMetric label="מק״טים" value={materials.length}/><BatchMetric label="מנות" value={batches.length}/><BatchMetric label="Orders" value={orders.length}/><BatchMetric label="תחזית" value={fmt(resource.forecast)}/><BatchMetric label="קצב נדרש" value={fmt(resource.requiredDaily)}/>
      </div>
      <div className="resource-detail-progress"><div><span>התקדמות חודשית</span><b>{Math.round(progress)}%</b></div><i><em style={{width:`${progress}%`}}/></i></div>
      <section className="resource-audit-section">
        <div className="resource-audit-title"><div><h3>מק״טים שנכנסו לחישוב</h3><p>הסכום בטבלה הוא בדיוק מקור נתון ה״בוצע״ בכרטיס.</p></div><span>{materials.length} מק״טים · {fmt(calculatedActual)}</span></div>
        <div className="table-wrap resource-audit-table"><table><thead><tr><th>מק״ט</th><th>תיאור</th><th>כמות בחישוב</th><th>מנות</th><th>Orders</th><th>רשומות</th></tr></thead><tbody>{materials.map(item=><tr key={item.material}><td><b>{item.material}</b></td><td>{item.description || '—'}</td><td><b>{fmt(item.qty)}</b></td><td>{item.batches.size}</td><td>{item.orders.size}</td><td>{item.rows}</td></tr>)}{!materials.length&&<tr><td colSpan="6" className="empty">לא נמצאו מק״טים בחישוב.</td></tr>}</tbody></table></div>
      </section>
      <section className="resource-audit-section">
        <div className="resource-audit-title"><div><h3>פירוט מלא — מנות ורשומות תפוקה</h3><p>כל שורה שהמנוע כלל בחישוב המשאב.</p></div><span>{rows.length} רשומות</span></div>
        <div className="table-wrap resource-audit-table resource-row-audit"><table><thead><tr><th>תאריך</th><th>תחנה</th><th>Routing</th><th>Order Type</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור</th><th>כמות</th></tr></thead><tbody>{rows.slice().sort((a,b)=>String(b.productionDay||'').localeCompare(String(a.productionDay||''))).map((row,index)=><tr key={`${row.order}-${row.batch}-${row.material}-${index}`}><td>{row.productionDay || iso(row.date) || '—'}</td><td>{row.facility || '—'}</td><td>{row.routingGroup || '—'}</td><td>{row.orderType || '—'}</td><td>{row.order || '—'}</td><td>{row.batch ? <button type="button" className="batch-link" onClick={()=>onOpenBatch(row.batch,row.material)}>{row.batch}</button> : '—'}</td><td>{row.material || '—'}</td><td>{row.desc || '—'}</td><td><b>{fmt(row.qty)}</b></td></tr>)}{!rows.length&&<tr><td colSpan="9" className="empty">לא נמצאו רשומות בחישוב.</td></tr>}</tbody></table></div>
      </section>
    </section>
  </div>
}

const planningSafeMonth = resource => String(resource?.month || new Date().toISOString().slice(0,7))

const BATCH_MODAL_WIDE_STYLES = `
.batch-scope-note{margin:0;padding:6px 18px 0;font-size:12.5px;color:#8a6d1f;background:transparent}
.batch-verdict-strip{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:8px 18px 0;padding:9px 14px;border:1px solid #dce7ec;border-radius:12px;background:#f6fafb;font-size:13px}
.batch-verdict-strip>span{color:#5b7480;font-weight:600}
.batch-verdict-strip>b.good{color:#0f7a52}
.batch-verdict-strip>b.bad{color:#c0392b}
.batch-verdict-strip>em{font-style:normal;color:#7a8f99;font-size:12.5px}
.batch-kpi small{display:block;margin-top:2px;font-size:11.5px;color:#8098a4;font-weight:500}
.quality-lot-filter{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:0 0 10px}
.quality-lot-filter>span{font-size:12.5px;color:#5b7480;font-weight:600}
.quality-lot-filter button{border:1px solid #dce7ec;background:#fff;border-radius:999px;padding:4px 12px;font-size:12.5px;cursor:pointer;color:#37515c}
.quality-lot-filter button.active{background:#0f5c6b;border-color:#0f5c6b;color:#fff}
.lot-kind{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11.5px;font-weight:600;white-space:nowrap}
.lot-kind.final{background:#e4f1ff;color:#12508f}
.lot-kind.inprocess{background:#f1efe6;color:#7a6a3a}
.lot-kind.unknown{background:#f0f3f5;color:#7a8f99}
.quality-status.unknown{background:#f0f3f5;color:#7a8f99}
.batch-modal-backdrop{padding:10px!important;align-items:stretch!important;justify-content:stretch!important}
.batch-control-card.batch-control-card-wide{width:calc(100vw - 20px)!important;max-width:none!important;height:calc(100vh - 20px)!important;max-height:none!important;margin:0!important;border-radius:22px!important;display:flex!important;flex-direction:column!important;overflow:hidden!important}
.batch-control-card-wide .batch-card-head{flex:0 0 auto}
.batch-control-card-wide .batch-summary-grid{grid-template-columns:repeat(8,minmax(130px,1fr))!important;gap:10px!important;padding-inline:18px!important}
.batch-control-card-wide .batch-kpi-row{grid-template-columns:repeat(4,minmax(170px,1fr))!important;padding-inline:18px!important}
.batch-control-card-wide .batch-card-body.batch-card-body-wide{display:block!important;flex:1 1 auto!important;min-height:0!important;padding:0 18px 18px!important;overflow:hidden!important}
.batch-workspace-tabs{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:10px 0 12px;border-bottom:1px solid #dce7ec;margin-bottom:12px}
.batch-workspace-tabs button{border:1px solid #cbdbe2;background:#fff;color:#24445a;border-radius:12px;padding:10px 18px;font-weight:800;cursor:pointer;display:flex;align-items:center;gap:8px}
.batch-workspace-tabs button.active{background:#0d766f;color:#fff;border-color:#0d766f;box-shadow:0 8px 18px rgba(13,118,111,.18)}
.batch-workspace-tabs button.bad.active{background:#c93636;border-color:#c93636}
.batch-workspace{height:calc(100% - 58px);min-height:0;overflow:auto}
.batch-workspace .batch-panel{height:auto;min-height:100%;margin:0!important}
.batch-workspace .table-wrap{max-height:none!important;height:auto!important;overflow:auto!important}
.batch-workspace table{width:100%!important;min-width:1380px!important;table-layout:auto!important}
.batch-workspace th,.batch-workspace td{padding:12px 14px!important;white-space:nowrap;vertical-align:top}
.batch-workspace th:nth-child(2),.batch-workspace td:nth-child(2){min-width:300px;white-space:normal}
.batch-workspace th:last-child,.batch-workspace td:last-child{min-width:260px;white-space:normal}
.batch-deviation-table th:nth-child(7),.batch-deviation-table td:nth-child(7){min-width:360px;white-space:normal}
.batch-deviation-table th:nth-child(8),.batch-deviation-table td:nth-child(8){min-width:320px;white-space:normal}
.batch-deviation-table .deviation-characteristic{margin-bottom:8px;padding:8px 10px;border-radius:10px;background:#fff6f6;border:1px solid #ffd7d7}
.batch-deviation-table .deviation-characteristic strong{display:block;margin-bottom:4px}
.batch-support-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}
.batch-support-grid .batch-panel{min-height:0!important}
@media(max-width:1200px){.batch-control-card-wide .batch-summary-grid{grid-template-columns:repeat(4,minmax(130px,1fr))!important}.batch-support-grid{grid-template-columns:1fr}.batch-workspace table{min-width:1100px!important}}
`;

function BatchControlCard({ data, onClose }) {
  const [qualityFilter, setQualityFilter] = useState('all')
  const [lotFilter, setLotFilter] = useState('all')
  const [activeBatchTab, setActiveBatchTab] = useState('quality')
  useEffect(() => {
    const previous = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const onKey = (event) => { if (event.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => { document.body.style.overflow = previous; window.removeEventListener('keydown', onKey) }
  }, [onClose])
  const productionRows = data.production || []
  const qualityRows = data.quality || []
  const deviationRows = data.deviations || []
  const firstProduction = productionRows[0] || {}
  const firstQuality = qualityRows[0] || {}
  const facilities = [...new Set([...productionRows.map(r=>r.facility), ...qualityRows.map(r=>r.facility), ...deviationRows.map(r=>r.facility)].filter(Boolean))]
  const orders = [...new Set([...productionRows.map(r=>r.order), ...qualityRows.map(r=>r.order)].filter(Boolean))]
  const materials = [...new Set([...productionRows.map(r=>r.material), ...qualityRows.map(r=>r.material), ...deviationRows.map(r=>r.material)].filter(Boolean))]
  const descriptions = [...new Set(productionRows.map(r=>r.desc).filter(Boolean))]
  const routingGroups = [...new Set(productionRows.map(r=>r.routingGroup).filter(Boolean))]
  const totalQty = productionRows.reduce((sum,r)=>sum+num(r.qty),0)
  const prodDates = productionRows.map(r=>r.date).filter(Boolean).sort((a,b)=>a-b)
  const qualityDates = qualityRows.map(r=>r.date).filter(Boolean).sort((a,b)=>a-b)
  const REJECT_WORDS = ['rejection','rejected','fail','failed','פסול','לא תקין','חריג']
  const ACCEPT_WORDS = ['acceptance','accepted','approved','pass','passed','תקין','מאושר']
  const allQuality = qualityRows.filter(r=>r.characteristic).map((r,index) => {
    const status = normalize(r.status || r.approval).toLowerCase()
    const rejected = REJECT_WORDS.some(x=>status.includes(x))
    const approved = !rejected && ACCEPT_WORDS.some(x=>status.includes(x))
    return {
      ...r,
      rejected,
      approved,
      unknown: !rejected && !approved,
      lotKind: lotKind(r.inspectionLot),
      // Sample # + Operation Activity make the key unique per real measurement,
      // so repeat samples of the same characteristic no longer share a React key.
      _key:`${r.inspectionLot}-${r.sample}-${r.activity}-${r.characteristic}-${index}`,
    }
  })
  const badCount = allQuality.filter(r=>r.rejected).length
  const goodCount = allQuality.filter(r=>r.approved).length
  const unknownCount = allQuality.filter(r=>r.unknown).length
  const inspectionLots = [...new Set(qualityRows.map(r=>r.inspectionLot).filter(Boolean))]
  const finalLots = inspectionLots.filter(l=>lotKind(l)===LOT_KIND_FINAL)
  const inProcessLots = inspectionLots.filter(l=>lotKind(l)===LOT_KIND_INPROCESS)
  const finalQuality = allQuality.filter(r=>r.lotKind===LOT_KIND_FINAL)
  const finalBad = finalQuality.filter(r=>r.rejected).length
  const shownQuality = allQuality
    .filter(r=>qualityFilter==='all' || (qualityFilter==='bad' ? r.rejected : !r.rejected))
    .filter(r=>lotFilter==='all' || r.lotKind===lotFilter)
  // Success % is measured against rows that actually carry a verdict, so rows
  // with a blank status no longer silently inflate the approved side.
  const judged = goodCount + badCount
  const qualityPct = judged ? Math.round(goodCount/judged*100) : 0

  // The usage decision comes from the usage-decision file (SAP UD), never from
  // concatenated Result Status values on the quality file.
  const decidedRows = deviationRows.filter(r=>!isOpenDecision(r))
  const openDecisionRows = deviationRows.filter(isOpenDecision)
  const finalDecision = decidedRows
    .slice()
    .sort((a,b)=>new Date(b.date||0)-new Date(a.date||0))[0] || null
  const udCodes = [...new Set(decidedRows.map(r=>r.udCode).filter(Boolean))]
  const qaLabel = finalDecision
    ? [finalDecision.udCode, finalDecision.status].filter(Boolean).join(' · ')
    : 'ממתין להחלטת שימוש'
  const released = udCodes.some(v => /^A1|release/i.test(normalize(v)))
  const decisionDate = finalDecision?.date || null
  const steps = [
    {label:'ייצור / אריזה', done:productionRows.length>0, date:prodDates[0]},
    {label:'בדיקות איכות', done:qualityRows.length>0, date:qualityDates[0]},
    {label:'החלטת QA', done:Boolean(finalDecision), date:decisionDate},
    {label:'שחרור', done:released, date:released ? decisionDate : null},
  ]
  const exportBatch = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(productionRows.map(r=>({Date:iso(r.date),Facility:r.facility,RoutingGroup:r.routingGroup,Order:r.order,Batch:r.batch,Material:r.material,Description:r.desc,Quantity:r.qty}))), 'Production')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allQuality.map(r=>({Date:iso(r.date),InspectionLot:r.inspectionLot,LotKind:lotKindLabel(r.lotKind),Sample:r.sample,Activity:r.activity,Material:r.material||materials.join(', '),Characteristic:r.characteristic,Result:r.value||r.qualitative,Lower:r.lower,Upper:r.upper,Unit:r.unit,Status:r.rejected?'חריג':r.approved?'תקין':'ללא סטטוס',Remarks:r.remarks}))), 'Quality')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deviationRows.map(r=>({Date:iso(r.date),InspectionLot:r.inspectionLot,Facility:r.facility,Material:r.material||materials.join(', '),Status:r.status,UDCode:r.udCode,UDDate:r.udDate,Rejected:r.rejectedCount,Approved:r.approvedCount,Open:isOpenDecision(r)?'פתוח':'הוכרע',Remarks:r.remarks}))), 'Deviations')
    XLSX.writeFile(wb, `Batch_${data.batch}.xlsx`)
  }
  return <div className="batch-modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget) onClose()}}>
    <style>{BATCH_MODAL_WIDE_STYLES}</style>
    <section className="batch-control-card batch-control-card-wide" role="dialog" aria-modal="true" aria-label={`כרטיס מנה ${data.batch}`}>
      <header className="batch-card-head"><div><span className="batch-eyebrow">BATCH CONTROL CENTER</span><h2>כרטיס מנה — {data.batch}</h2><p>{descriptions.join(' · ') || materials.join(', ') || 'ללא תיאור חומר'}</p></div><div className="batch-head-actions"><button type="button" className="batch-export" onClick={exportBatch}><Download size={17}/> ייצוא מנה</button><button type="button" className="batch-close" onClick={onClose} aria-label="סגירה"><X/></button></div></header>
      <div className="batch-summary-grid">
        <BatchMetric label="Batch" value={data.batch}/><BatchMetric label="Order" value={orders.join(', ') || '—'}/><BatchMetric label="מק״ט חומר" value={materials.join(', ') || '—'}/><BatchMetric label="מתקן" value={facilities.join(', ') || '—'}/><BatchMetric label="Routing group" value={routingGroups.join(', ') || '—'}/><BatchMetric label="כמות ארוזה" value={fmt(totalQty)}/><BatchMetric label="Inspection Lot" value={inspectionLots.join(', ') || '—'}/><BatchMetric label="QA" value={qaLabel}/>
      </div>
      {data.usedBatchOnlyFallback && (data.materials?.length || 0) > 1 && <p className="batch-scope-note">המנה מופיעה על {data.materials.length} מק״טים — מוצגות תוצאות כל המק״טים של המנה.</p>}
      <div className="batch-kpi-row"><div className="batch-kpi good"><span>מאפיינים תקינים</span><b>{goodCount}</b></div><div className="batch-kpi bad"><span>מאפיינים חריגים</span><b>{badCount}</b></div><div className="batch-kpi"><span>אחוז הצלחה</span><b>{qualityPct}%</b>{unknownCount>0&&<small>{unknownCount} ללא סטטוס — לא נספרו</small>}</div><div className="batch-kpi"><span>חריגות פתוחות</span><b>{openDecisionRows.length}</b>{decidedRows.length>0&&<small>{decidedRows.length} הוכרעו</small>}</div></div>
      {finalLots.length>0 && <div className="batch-verdict-strip"><span>בדיקה סופית ({finalLots.join(', ')})</span><b className={finalBad?'bad':'good'}>{finalBad?`${finalBad} מתוך ${finalQuality.length} מאפיינים חריגים`:`כל ${finalQuality.length} המאפיינים תקינים`}</b>{inProcessLots.length>0&&<em>+ {allQuality.length-finalQuality.length} בדיקות תוך-תהליכיות ({inProcessLots.join(', ')}) — ללא החלטת שימוש</em>}</div>}
      <div className="batch-card-body batch-card-body-wide">
        <nav className="batch-workspace-tabs" aria-label="לשוניות כרטיס מנה">
          <button type="button" className={activeBatchTab==='quality'?'active':''} onClick={()=>setActiveBatchTab('quality')}><FlaskConical size={17}/> תוצאות איכות ({allQuality.length})</button>
          <button type="button" className={`bad ${activeBatchTab==='deviations'?'active':''}`} onClick={()=>setActiveBatchTab('deviations')}><AlertTriangle size={17}/> חריגים ({openDecisionRows.length})</button>
          <button type="button" className={activeBatchTab==='production'?'active':''} onClick={()=>setActiveBatchTab('production')}><Factory size={17}/> ייצור ואריזה ({productionRows.length})</button>
          <button type="button" className={activeBatchTab==='timeline'?'active':''} onClick={()=>setActiveBatchTab('timeline')}><Clock3 size={17}/> Timeline והערות</button>
        </nav>
        <div className="batch-workspace">
          {activeBatchTab==='quality' && <section className="batch-panel quality-panel"><div className="batch-panel-title"><div><h3>תוצאות איכות</h3><p>תצוגה רחבה של כל המאפיינים שנמצאו למנה בקובץ האיכות</p></div><div className="quality-filter"><button className={qualityFilter==='all'?'active':''} onClick={()=>setQualityFilter('all')}>הכול {allQuality.length}</button><button className={qualityFilter==='good'?'active good':''} onClick={()=>setQualityFilter('good')}>תקינים {goodCount}</button><button className={qualityFilter==='bad'?'active bad':''} onClick={()=>setQualityFilter('bad')}>חריגים {badCount}</button></div></div>{inspectionLots.length>1&&<div className="quality-lot-filter"><span>Inspection Lot:</span><button className={lotFilter==='all'?'active':''} onClick={()=>setLotFilter('all')}>הכול {allQuality.length}</button><button className={lotFilter===LOT_KIND_FINAL?'active':''} onClick={()=>setLotFilter(LOT_KIND_FINAL)}>סופי {finalQuality.length}</button><button className={lotFilter===LOT_KIND_INPROCESS?'active':''} onClick={()=>setLotFilter(LOT_KIND_INPROCESS)}>תוך-תהליכי {allQuality.filter(r=>r.lotKind===LOT_KIND_INPROCESS).length}</button></div>}<div className="table-wrap batch-quality-table"><table><thead><tr><th>מק״ט חומר</th><th>Inspection Lot</th><th>סוג בדיקה</th><th>Sample #</th><th>שלב</th><th>מאפיין</th><th>תוצאה</th><th>גבול תחתון</th><th>גבול עליון</th><th>יחידה</th><th>סטטוס</th><th>הערה</th></tr></thead><tbody>{shownQuality.map(r=><tr key={r._key} className={r.rejected?'quality-row-bad':''}><td>{r.material || materials.join(', ') || '—'}</td><td>{r.inspectionLot||'—'}</td><td><span className={`lot-kind ${r.lotKind||'unknown'}`}>{lotKindLabel(r.lotKind)}</span></td><td>{r.sample||'—'}</td><td>{r.activity||'—'}</td><td><b>{r.characteristic}</b></td><td>{r.value||r.qualitative||'—'}</td><td>{r.lower||'—'}</td><td>{r.upper||'—'}</td><td>{r.unit||'—'}</td><td><span className={`quality-status ${r.rejected?'bad':r.approved?'good':'unknown'}`}>{r.rejected?'חריג':r.approved?'תקין':'ללא סטטוס'}</span></td><td>{r.remarks||'—'}</td></tr>)}{!shownQuality.length&&<tr><td colSpan="12" className="empty">לא נמצאו תוצאות במסנן שנבחר</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='deviations' && <section className="batch-panel deviation-panel"><div className="batch-panel-title"><div><h3>חריגים</h3><p>תצוגה רחבה של כל החריגות, המאפיינים החריגים וההערות המקושרות למנה</p></div><span className="production-record-count">{openDecisionRows.length} פתוחות · {decidedRows.length} הוכרעו</span></div><div className="table-wrap batch-deviation-table"><table><thead><tr><th>תאריך חריגה</th><th>Inspection Lot</th><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>סטטוס</th><th>UD Code</th><th>תאריך החלטה</th><th>מאפייני חריגה</th><th>הערות</th></tr></thead><tbody>{deviationRows.map((r,i)=><tr key={`${r.batch}-${r.inspectionLot||i}-${i}`} className={isOpenDecision(r)?'quality-row-bad':''}><td>{iso(r.date)||'—'}</td><td>{r.inspectionLot||'—'}</td><td>{iso(r.sampleDate)||'—'}</td><td>{r.sampleDate?new Date(r.sampleDate).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}):'—'}</td><td>{r.facility||facilities.join(', ')||'—'}</td><td><span className={`quality-status ${isOpenDecision(r)?'bad':'good'}`}>{r.status||'ממתין להחלטת שימוש'}</span></td><td>{r.udCode||'—'}</td><td>{r.udDate||'—'}</td><td>{r.rejectedCharacteristics?.length ? r.rejectedCharacteristics.map((c,j)=><div className="deviation-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value||c.qualitative||'—'}{c.unit?` ${c.unit}`:''}</b></span><span>מפרט: {c.lower!==''||c.upper!==''?`${c.lower||'—'} עד ${c.upper||'—'}${c.unit?` ${c.unit}`:''}`:'—'}</span>{c.remarks&&c.remarks!=='N/A'&&<small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו פרטי מאפיינים חריגים בקובץ האיכות{r.rejectedCount?` (בקובץ החריגות מופיע מספר: ${r.rejectedCount})`:''}</span>}</td><td>{r.remarks||'—'}</td></tr>)}{!deviationRows.length&&<tr><td colSpan="10" className="empty">לא נמצאו חריגות למנה</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='production' && <section className="batch-panel production-panel"><div className="batch-panel-title"><div><h3>נתוני ייצור ואריזה</h3><p>כל דיווחי התפוקה המקושרים למנה</p></div><span className="production-record-count">{productionRows.length} דיווחים</span></div><div className="table-wrap batch-production-table"><table><thead><tr><th>תאריך</th><th>שעה</th><th>משאב יעד</th><th>מתקן / תחנה</th><th>Order</th><th>מק״ט</th><th>תיאור</th><th>כמות</th><th>משמרת</th></tr></thead><tbody>{productionRows.slice().sort((a,b)=>new Date(b.date||0)-new Date(a.date||0)).map((r,index)=><tr key={`${r.order}-${r.date?.getTime?.()||index}-${index}`}><td>{iso(r.date)||'—'}</td><td>{r.date?new Date(r.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}):'—'}</td><td>{r.facility||'—'}</td><td>{r.routingGroup||'—'}</td><td>{r.order||'—'}</td><td>{r.material||'—'}</td><td>{r.desc||'—'}</td><td><b>{fmt(r.qty)}</b></td><td>{shiftInfo(r.date).label}</td></tr>)}{!productionRows.length&&<tr><td colSpan="9" className="empty">לא נמצאו דיווחי ייצור או אריזה למנה</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='timeline' && <div className="batch-support-grid"><section className="batch-panel"><div className="batch-panel-title"><div><h3>Timeline</h3><p>מצב התקדמות המנה</p></div></div><div className="batch-timeline">{steps.map((step,i)=><div className={`timeline-step ${step.done?'done':''}`} key={step.label}><i>{step.done?<CheckCircle2 size={18}/>:<Clock3 size={18}/>}</i><div><b>{step.label}</b><span>{step.date?`${iso(step.date)} ${new Date(step.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}`:'טרם הושלם'}</span></div>{i<steps.length-1&&<em/>}</div>)}</div></section><section className="batch-panel"><div className="batch-panel-title"><div><h3>חריגות והערות</h3><p>{deviationRows.length} רשומות מקושרות למנה</p></div></div><div className="batch-deviations">{deviationRows.map((r,i)=><article key={i}><div><span className="quality-status bad">{r.status||'חריגה'}</span>{r.udCode&&<b>{r.udCode}</b>}</div><p><b>מק״ט חומר: {r.material || materials.join(', ') || '—'}</b></p><p>{r.remarks||'לא הוזנה הערה'}</p><small>{iso(r.date)||'ללא תאריך'}</small></article>)}{!deviationRows.length&&<div className="batch-empty-good"><CheckCircle2/> לא נמצאו חריגות למנה</div>}</div></section></div>}
        </div>
      </div>
    </section>
  </div>
}
function BatchMetric({label,value}) { return <div className="batch-metric"><span>{label}</span><b>{value}</b></div> }

function DataSource({ title, icon, meta, count, acceptLabel, busy, onFiles, canManage }) {
  const loaded = Boolean(meta || count)
  const loadedAt = meta?.loadedAt ? new Date(meta.loadedAt).toLocaleString('he-IL') : 'טרם נטען'
  return <article className={`data-source ${loaded ? 'ready' : ''}`}>
    <div className="data-source-head"><div className="data-source-icon">{icon}</div><div><h3>{title}</h3><span>{loaded ? 'תקין וזמין' : 'ממתין לקובץ'}</span></div></div>
    <div className="data-source-count"><b>{fmt(count)}</b><span>רשומות פעילות</span></div>
    <div className="data-source-meta"><small title={meta?.fileName || ''}>{meta?.fileName || 'לא נבחר קובץ'}</small><small>{loadedAt}</small>{meta?.source === 'cloud' && <small className="cloud-source-label">מקור: Supabase{meta?.loadedBy ? ` · ${meta.loadedBy}` : ''}</small>}{meta?.facilities ? <small>{meta.facilities} מתקנים זוהו במדגם</small> : null}</div>
    {canManage ? <label className={`source-upload ${busy ? 'disabled' : ''}`}><RefreshCw size={16}/>{acceptLabel}<input type="file" accept=".xlsx,.xls" disabled={busy} onChange={e => { const files=[...e.target.files]; e.target.value=''; onFiles(files) }}/></label> : <div className="viewer-lock"><ShieldCheck size={16}/> צפייה בלבד</div>}
  </article>
}

function Summary({ title, value, sub, warn }) { return <div className={`summary ${warn ? 'warn' : ''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Executive({ icon, title, value, sub, good, warn, bad, onClick }) { return <button type="button" className={`executive ${good?'good':''} ${warn?'warn':''} ${bad?'bad':''} ${onClick?'clickable':''}`} onClick={onClick}><div className="executive-icon">{icon}</div><div><span>{title}</span><b>{value}</b><small>{sub}</small></div></button> }
function StatusBadge({ state, label }) { return <span className={`status-pill ${state}`}>{label}</span> }
function ForecastCard({ facility, facilities, resource, packagingType, routingGroup, station, lineName, target, actual, pct, remaining, requiredDaily, recentAverage, provenMax, forecast, remainingWorkdays, state, label, selected, onClick }) {
  return <article className={`forecast-card ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0">
    <div className="forecast-head"><div><small>משאב / מתקן</small><h3>{resource || facility}</h3>{(facilities || []).includes('1542') && packagingType && <div className="forecast-packaging-type"><span>קו אריזה</span><b>{packagingType}</b></div>}{(routingGroup || station) && <div className="forecast-resource"><b>{station || routingGroup}</b><span>{lineName || routingGroup}</span>{routingGroup && <small>{routingGroup}</small>}</div>}</div><StatusBadge state={state} label={label}/></div>
    <div className="forecast-main"><div><span>ביצוע</span><b>{pctFmt(pct)}</b></div><div><span>תחזית</span><b>{target ? pctFmt(forecast / target * 100) : '—'}</b></div></div>
    <div className="bar"><i style={{width:`${Math.min(100, pct)}%`}}/></div>
    <div className="forecast-values"><span>יעד<strong>{fmt(target)}</strong></span><span>בפועל<strong>{fmt(actual)}</strong></span><span>נותר<strong>{fmt(remaining)}</strong></span></div>
    <div className="forecast-metrics"><div><span>ימים נותרו</span><b>{remainingWorkdays}</b></div><div><span>נדרש ליום</span><b>{fmt(requiredDaily)}</b></div><div><span>7 ימים</span><b>{fmt(recentAverage)}</b></div><div><span>שיא מוכח</span><b>{fmt(provenMax)}</b></div></div>
  </article>
}
function Facility({ id, actual, target, pct, orders, selected, onClick }) {
  const state = !target ? 'no-target' : pct >= 100 ? 'good' : pct >= 75 ? 'warning' : 'risk'
  return <article className={`facility ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0"><div className="facility-top"><div><small>מתקן</small><h3>{id}</h3></div><b>{target ? pctFmt(pct) : '—'}</b></div><div className="bar"><i style={{width:`${Math.min(100,pct)}%`}}/></div><div className="facility-numbers"><span>בפועל<strong>{fmt(actual)}</strong></span><span>יעד חודשי<strong>{fmt(target)}</strong></span><span>הזמנות<strong>{orders}</strong></span></div></article>
}