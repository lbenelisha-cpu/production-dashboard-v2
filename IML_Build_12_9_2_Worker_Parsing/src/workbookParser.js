// Shared workbook parsing + mapping.
//
// This module is the single source of truth for turning an Excel export into the
// compact row shapes the app stores. It is imported by DashboardApp.jsx AND by
// workbookWorker.js, so the mapping can run off the main thread without the two
// copies drifting apart.
//
// Why it matters: a 40MB / 280K-row SAP quality export costs ~1.1GB of JS heap to
// parse. Doing that on the main thread, alongside React state and the quality rows
// already in memory, is what pushed the tab past its limit. When the mapping runs
// in a worker, only the compact rows (an order of magnitude smaller) cross back,
// and an out-of-memory failure kills the worker instead of the tab.
import * as XLSX from 'xlsx'

export const FACILITY_ALIASES = {
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
export const normalize = (v) => String(v ?? '').trim()
export const normalizeRouting = (v) => normalize(v).toUpperCase()
export const normKey = (v) => normalize(v).toLowerCase().replace(/[\s_\-./()]+/g, '')
export const num = (v) => {
  const n = Number(String(v ?? '').replace(/,/g, '').replace(/\s/g, ''))
  return Number.isFinite(n) ? n : 0
}
export const excelDate = (v) => {
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
export const localDateTimeString = date => {
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
export const localDateOnlyString = value => {
  const date = excelDate(value)
  if (!date) return ''
  const pad = number => String(number).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}
export const excelTime = (v) => {
  if (v === '' || v === null || v === undefined) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return { h: v.getHours(), m: v.getMinutes(), s: v.getSeconds() }
  if (typeof v === 'number') {
    const seconds = Math.round((v % 1) * 86400) % 86400
    return { h: Math.floor(seconds / 3600), m: Math.floor((seconds % 3600) / 60), s: seconds % 60 }
  }
  const match = normalize(v).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/)
  return match ? { h: Number(match[1]), m: Number(match[2]), s: Number(match[3] || 0) } : null
}
export const combineExcelDateTime = (dateValue, timeValue, fallbackValue = '') => {
  const date = excelDate(dateValue) || excelDate(fallbackValue)
  if (!date) return null
  const time = excelTime(timeValue)
  const result = new Date(date)
  if (time) result.setHours(time.h, time.m, time.s, 0)
  return result
}
export const getField = (row, names) => {
  const map = new Map(Object.keys(row || {}).map(k => [normKey(k), row[k]]))
  for (const name of names) {
    const value = map.get(normKey(name))
    if (value !== undefined) return value
  }
  return ''
}
export const canonicalFacility = (value) => {
  const clean = normalize(value).toUpperCase()
  for (const [id, aliases] of Object.entries(FACILITY_ALIASES)) {
    if (aliases.some(alias => normalize(alias).toUpperCase() === clean)) return id
  }
  const digits = clean.match(/15\d{2}/)?.[0]
  return digits || clean
}
export async function readWorkbook(file) {
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
export function mapRowsInPlace(rows, mapFn) {
  const out = []
  for (let i = 0; i < rows.length; i++) {
    const mapped = mapFn(rows[i], i)
    rows[i] = null
    if (mapped) out.push(mapped)
  }
  rows.length = 0
  return out
}
export function classifyFile(rows) {
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

// Required-column check, shared by the main thread and the worker.
export function validateRows(kind, rows) {
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

const MATERIAL_KEYS = ['Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט', 'Material']

export const compactProductionRow = r => ({
  __compactProduction: true,
  facility: canonicalFacility(getField(r, ['Storage Location', 'Storage location'])),
  productionDay: localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date'])),
  finishDate: localDateTimeString(combineExcelDateTime(
    getField(r, ['Actual finish date', 'Actual Finish Date']),
    getField(r, ['Actual Finish Time', 'Actual finish time']),
    getField(r, ['Release date (actual)', 'Time Stamp'])
  )),
  qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
  order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
  batch: normalize(getField(r, ['Batch', 'Batch Number'])),
  material: normalize(getField(r, ['Material', ...MATERIAL_KEYS])),
  desc: normalize(getField(r, ['Material description', 'Material Description'])),
  orderType: normalize(getField(r, ['Order Type'])),
  routingGroup: normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup'])),
  routingDescription: normalize(getField(r, ['Description', 'Routing Description'])),
})

export const compactQualityRow = r => ({
  __compactQuality: true,
  facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
  date: combineExcelDateTime(
    getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
    getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
    getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
  ),
  batch: normalize(getField(r, ['Batch', 'Batch Number'])),
  material: normalize(getField(r, MATERIAL_KEYS)),
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
  udCode: normalize(getField(r, ['UD Code', 'Usage Decision', 'Usage decision', 'החלטת שימוש'])),
})

const FACILITY_SAMPLE_KEYS = ['Storage Location', 'Inspection Lot Storage Location', 'Process Order Storage Location', 'Facility', 'Production Line', 'מתקן']

// Consumes `rows` (see mapRowsInPlace) and returns everything loadFiles needs.
export function buildDataset(kind, rows) {
  const rawRows = rows.length
  const facilities = new Set(rows.slice(0, 5000)
    .map(r => canonicalFacility(getField(r, FACILITY_SAMPLE_KEYS))).filter(Boolean)).size
  if (kind === 'production') {
    const compact = mapRowsInPlace(rows, compactProductionRow).filter(r => r.facility && (r.qty || r.order || r.batch))
    return { kind, rows: compact, rawRows, facilities }
  }
  if (kind === 'quality') {
    const compact = mapRowsInPlace(rows, compactQualityRow).filter(r => r.batch || r.inspectionLot)
    return { kind, rows: compact, rawRows, facilities }
  }
  return { kind, rows, rawRows, facilities }
}

// Full pipeline: parse -> classify -> validate -> compact. Runs identically on the
// main thread and inside the worker.
export async function parseWorkbookFile(file, forcedKind = '') {
  const rows = await readWorkbook(file)
  if (!rows.length) {
    const error = new Error('EMPTY_WORKBOOK')
    error.code = 'EMPTY_WORKBOOK'
    throw error
  }
  const detected = classifyFile(rows)
  const kind = forcedKind || detected
  if (kind === 'unknown') {
    const error = new Error('UNKNOWN_KIND')
    error.code = 'UNKNOWN_KIND'
    error.headers = Object.keys(rows[0] || {}).filter(k => k !== '__sheet').slice(0, 6)
    throw error
  }
  const missing = validateRows(kind, rows)
  if (missing.length) {
    const error = new Error('MISSING_COLUMNS')
    error.code = 'MISSING_COLUMNS'
    error.missing = missing
    throw error
  }
  return { ...buildDataset(kind, rows), detected }
}
