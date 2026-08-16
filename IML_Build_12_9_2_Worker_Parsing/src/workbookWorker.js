// Off-main-thread workbook parsing.
//
// The worker receives the File handle, runs the full parse -> classify -> compact
// pipeline in its OWN heap, and posts back only the compact rows. Two consequences
// that matter for the 40MB / 280K-row SAP quality export:
//
//   1. The ~1.1GB parse never touches the main thread's heap, which is already
//      holding React state and the quality rows loaded so far.
//   2. If the parse still runs out of memory, the worker dies and we catch it.
//      Previously the same failure took the whole tab down, or surfaced as the
//      misleading "file type not recognized" error.
//
// The UI thread also stays responsive, so the 30+ second parse no longer freezes
// the page.
import { parseWorkbookFile } from './workbookParser'

self.onmessage = async (event) => {
  const { file, forcedKind, requestId } = event.data || {}
  try {
    self.postMessage({ requestId, type: 'progress', phase: 'parse', message: 'קורא את הקובץ' })
    const result = await parseWorkbookFile(file, forcedKind)
    self.postMessage({
      requestId,
      type: 'done',
      kind: result.kind,
      detected: result.detected,
      rows: result.rows,
      rawRows: result.rawRows,
      facilities: result.facilities,
    })
  } catch (error) {
    self.postMessage({
      requestId,
      type: 'error',
      code: error?.code || '',
      message: error?.message || 'שגיאה לא ידועה',
      missing: error?.missing || null,
      headers: error?.headers || null,
    })
  }
}
