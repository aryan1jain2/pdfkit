// ============================================================
// XLSX PARSER — pdfkit-client
//
// An .xlsx file is a ZIP containing:
//
//   xl/
//     workbook.xml              ← sheet list, defined names
//     worksheets/
//       sheet1.xml              ← one file per sheet
//     sharedStrings.xml         ← ALL text lives here, cells just reference indices
//     styles.xml                ← cell formatting (fonts, fills, borders, numFmts)
//     _rels/workbook.xml.rels   ← maps sheet rId to file path
//
// Cell value types (t attribute on <c>):
//   s  → shared string (value is index into sharedStrings)
//   n  → number (or absent = number by default)
//   b  → boolean (0/1)
//   e  → error (#REF!, #VALUE! etc.)
//   str → formula result string
//   inlineStr → inline string (rare)
//
// Column addresses:
//   "A1" → col=0, row=0
//   "Z1" → col=25, row=0
//   "AA1"→ col=26, row=0
//   Uses base-26 encoding
// ============================================================

import type {
  IRDocument, IRPage, TableElement, TableRow, TableCell, CellStyle,
  ParseResult, Rgba
} from '../ir/types'

import { loadZip, isParseError, parseXml, getEl, getEls, attr, numAttr } from '../utils/xml'
import { hexToRgba, cmToPt, BLACK, WHITE } from '../utils/units'

// ============================================================
// ENTRY POINT
// ============================================================

export async function parseXlsx(file: File | ArrayBuffer): Promise<ParseResult> {
  const zip = await loadZip(file)
  if (isParseError(zip)) return { ok: false, error: zip }

  // Load shared strings (MUST be read before sheets)
  const sharedStrings = await loadSharedStrings(zip)

  // Load styles (for cell colours/fonts)
  const styles = await loadStyles(zip)

  // Read workbook to get sheet metadata
  const wbXml = await zip.getText('xl/workbook.xml')
  if (!wbXml) return { ok: false, error: { code: 'MISSING_ENTRY', message: 'xl/workbook.xml not found' } }

  const wbDoc = parseXml(wbXml)
  if (!wbDoc) return { ok: false, error: { code: 'XML_PARSE_FAILED', message: 'Could not parse workbook.xml' } }

  // Sheet order and names
  const sheets = getEls(wbDoc, 'sheet').map(s => ({
    name:  attr(s, 'name'),
    rId:   attr(s, 'r:id') || attr(s, 'id'),
  }))

  // Map rId → file path via relationships
  const relsXml = await zip.getText('xl/_rels/workbook.xml.rels')
  const pathMap = buildSheetPathMap(relsXml)

  const pages: IRPage[] = []

  for (const sheet of sheets) {
    const path = pathMap.get(sheet.rId)
    if (!path) continue

    const sheetXml = await zip.getText(`xl/${path}`)
    if (!sheetXml) continue

    const sheetDoc = parseXml(sheetXml)
    if (!sheetDoc) continue

    const page = parseSheet(sheetDoc, sharedStrings, styles, sheet.name)
    pages.push(page)
  }

  return {
    ok: true,
    doc: { format: 'xlsx', pages, metadata: {} }
  }
}

// ============================================================
// SHARED STRINGS
// ============================================================

async function loadSharedStrings(zip: any): Promise<string[]> {
  const xml = await zip.getText('xl/sharedStrings.xml')
  if (!xml) return []

  const doc = parseXml(xml)
  if (!doc) return []

  // Each <si> is a string entry. It may contain a single <t> or multiple <r><t> runs.
  return getEls(doc, 'si').map(si => {
    const runs = getEls(si, 't')
    return runs.map(t => t.textContent ?? '').join('')
  })
}

// ============================================================
// STYLES
// ============================================================

interface XlsxStyles {
  fills:   Array<{ fgColor?: Rgba }>
  fonts:   Array<{ bold: boolean; italic: boolean; color?: Rgba; size: number; name: string }>
  borders: Array<{ top: boolean; bottom: boolean; left: boolean; right: boolean }>
  cellXfs: Array<{ fontId: number; fillId: number; borderId: number; numFmtId: number; alignH?: string }>
}

async function loadStyles(zip: any): Promise<XlsxStyles> {
  const empty: XlsxStyles = { fills: [], fonts: [], borders: [], cellXfs: [] }
  const xml = await zip.getText('xl/styles.xml')
  if (!xml) return empty

  const doc = parseXml(xml)
  if (!doc) return empty

  // Fills
  const fills = getEls(doc, 'fill').map(fill => {
    const fgColor = getEl(fill, 'fgColor')
    const rgb = fgColor ? attr(fgColor, 'rgb') : ''
    return { fgColor: rgb ? hexToRgba(rgb.slice(2)) : undefined }  // ARGB → skip alpha prefix
  })

  // Fonts
  const fonts = getEls(doc, 'font').map(font => ({
    bold:   !!getEl(font, 'b'),
    italic: !!getEl(font, 'i'),
    color:  (() => {
      const c = getEl(font, 'color')
      const rgb = c ? attr(c, 'rgb') : ''
      return rgb ? hexToRgba(rgb.slice(2)) : undefined
    })(),
    size:   numAttr(getEl(font, 'sz')!, 'val', 11),
    name:   attr(getEl(font, 'name')!, 'val') || 'Calibri',
  }))

  // Borders
  const borders = getEls(doc, 'border').map(b => ({
    top:    !!getEl(b, 'top'),
    bottom: !!getEl(b, 'bottom'),
    left:   !!getEl(b, 'left'),
    right:  !!getEl(b, 'right'),
  }))

  // Cell format cross-reference table (xf entries under cellXfs)
  const cellXfs = getEls(getEl(doc, 'cellXfs')!, 'xf').map(xf => ({
    fontId:   numAttr(xf, 'fontId'),
    fillId:   numAttr(xf, 'fillId'),
    borderId: numAttr(xf, 'borderId'),
    numFmtId: numAttr(xf, 'numFmtId'),
    alignH:   (() => {
      const align = getEl(xf, 'alignment')
      return align ? attr(align, 'horizontal') : undefined
    })()
  }))

  return { fills, fonts, borders, cellXfs }
}

// ============================================================
// SHEET PATH MAP
// ============================================================

function buildSheetPathMap(relsXml: string | null): Map<string, string> {
  const map = new Map<string, string>()
  if (!relsXml) return map

  const doc = parseXml(relsXml)
  if (!doc) return map

  for (const rel of getEls(doc, 'Relationship')) {
    if (!attr(rel, 'Type').includes('worksheet')) continue
    map.set(attr(rel, 'Id'), attr(rel, 'Target'))
  }

  return map
}

// ============================================================
// SHEET PARSING
// ============================================================

function parseSheet(
  sheetDoc: Document,
  sharedStrings: string[],
  styles: XlsxStyles,
  sheetName: string
): IRPage {
  // Collect all cells indexed by [row][col]
  const cellGrid: Map<number, Map<number, { value: string; style: CellStyle }>> = new Map()

  let maxRow = 0
  let maxCol = 0

  for (const row of getEls(sheetDoc, 'row')) {
    const rowIdx = numAttr(row, 'r', 0) - 1  // 1-indexed → 0-indexed

    for (const c of getEls(row, 'c')) {
      const ref = attr(c, 'r')              // e.g. "B3"
      const { col } = cellRefToIndex(ref)
      const type    = attr(c, 't')
      const styleId = numAttr(c, 's')

      // Resolve value
      const vEl  = getEl(c, 'v')
      const raw  = vEl?.textContent ?? ''
      let value  = raw

      if (type === 's') {
        // Shared string index
        value = sharedStrings[parseInt(raw)] ?? ''
      } else if (type === 'b') {
        value = raw === '1' ? 'TRUE' : 'FALSE'
      } else if (type === 'e') {
        value = raw  // error string e.g. #REF!
      }

      // Resolve style
      const cellStyle = resolveStyle(styles, styleId)

      if (!cellGrid.has(rowIdx)) cellGrid.set(rowIdx, new Map())
      cellGrid.get(rowIdx)!.set(col, { value, style: cellStyle })

      maxRow = Math.max(maxRow, rowIdx)
      maxCol = Math.max(maxCol, col)
    }
  }

  // Build rows/cols arrays for the table element
  const numCols = maxCol + 1
  const colWidth = cmToPt(2.5)  // default column width ~2.5cm

  const tableRows: TableRow[] = []
  for (let r = 0; r <= maxRow; r++) {
    const cells: TableCell[] = []
    for (let c = 0; c < numCols; c++) {
      const cell = cellGrid.get(r)?.get(c)
      cells.push({
        value:   cell?.value ?? '',
        colspan: 1,
        rowspan: 1,
        style:   cell?.style ?? {}
      })
    }
    tableRows.push({ cells, height: cmToPt(0.6) })
  }

  const tableWidth = numCols * colWidth
  const tableHeight = tableRows.reduce((s, r) => s + r.height, 0)

  // Page size: fit the table with margins
  const margin = cmToPt(1)
  const pageWidth  = Math.max(tableWidth  + margin * 2, 595.28)   // min A4
  const pageHeight = Math.max(tableHeight + margin * 2, 841.89)

  const table: TableElement = {
    type: 'table',
    x: margin,
    y: margin,
    width: tableWidth,
    colWidths: Array(numCols).fill(colWidth),
    rows: tableRows
  }

  return {
    width: pageWidth,
    height: pageHeight,
    background: WHITE,
    elements: [table],
    label: sheetName
  }
}

// ============================================================
// HELPERS
// ============================================================

/** Convert a cell reference like "AB12" to { col, row } (0-indexed) */
export function cellRefToIndex(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) return { col: 0, row: 0 }

  const colStr = match[1]
  const rowStr = match[2]

  let col = 0
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64)
  }

  return { col: col - 1, row: parseInt(rowStr) - 1 }
}

function resolveStyle(styles: XlsxStyles, styleId: number): CellStyle {
  const xf = styles.cellXfs[styleId]
  if (!xf) return {}

  const font   = styles.fonts[xf.fontId]
  const fill   = styles.fills[xf.fillId]
  const border = styles.borders[xf.borderId]

  const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
    left: 'left', center: 'center', right: 'right', general: 'left'
  }

  return {
    fill:         fill?.fgColor,
    color:        font?.color,
    fontSize:     font?.size,
    fontWeight:   font?.bold ? 'bold' : 'normal',
    fontStyle:    font?.italic ? 'italic' : 'normal',
    align:        alignMap[xf.alignH ?? 'general'] ?? 'left',
    borderTop:    border?.top,
    borderBottom: border?.bottom,
    borderLeft:   border?.left,
    borderRight:  border?.right,
  }
}
