// ============================================================
// PDF GENERATOR — pdfkit-client
// ============================================================

import { jsPDF } from 'jspdf'
import type { IRDocument, IRPage, IRElement, TextElement, ImageElement, ShapeElement, TableElement, Rgba } from '../ir/types'

export interface PDFGeneratorOptions {
  pageNumbers?: boolean
  title?:  string
  author?: string
}

export interface GeneratedPDF {
  download(filename?: string): void
  toBlob(): Blob
  toDataUrl(): string
}

// ============================================================
// ENTRY POINT
// ============================================================

export function generatePDF(doc: IRDocument, opts: PDFGeneratorOptions = {}): GeneratedPDF {
  if (doc.pages.length === 0) throw new Error('Document has no pages')

  const firstPage = doc.pages[0]!
  const pdf = new jsPDF({
    orientation: firstPage.width > firstPage.height ? 'landscape' : 'portrait',
    unit:        'pt',
    format:      [firstPage.width, firstPage.height],
  })

  if (opts.title)  pdf.setProperties({ title:  opts.title  })
  if (opts.author) pdf.setProperties({ author: opts.author })

  doc.pages.forEach((page, i) => {
    if (i > 0) pdf.addPage([page.width, page.height], page.width > page.height ? 'landscape' : 'portrait')
    renderPage(pdf, page)
    if (opts.pageNumbers) renderPageNumber(pdf, i + 1, doc.pages.length, page)
  })

  return {
    download(filename = 'output.pdf') { pdf.save(filename) },
    toBlob():    Blob   { return pdf.output('blob') },
    toDataUrl(): string { return pdf.output('datauristring') },
  }
}

// ============================================================
// PAGE
// ============================================================

function renderPage(pdf: any, page: IRPage): void {
  // Background
  if (page.background.a > 0) {
    pdf.setFillColor(page.background.r, page.background.g, page.background.b)
    pdf.rect(0, 0, page.width, page.height, 'F')
  }

  // Build background colour map for contrast correction
  const bgMap = buildBackgroundMap(page)

  // Render in Z-order: shapes → images → tables → text (text always on top)
  for (const el of page.elements) {
    if (el.type === 'shape') renderShape(pdf, el)
  }
  for (const el of page.elements) {
    if (el.type === 'image') renderImage(pdf, el)
  }
  for (const el of page.elements) {
    if (el.type === 'table') renderTable(pdf, el)
  }
  for (let i = 0; i < page.elements.length; i++) {
    const el = page.elements[i]!
    if (el.type === 'text') {
      renderText(pdf, el as TextElement, bgMap.get(i) ?? page.background)
    }
  }
}

// ============================================================
// BACKGROUND MAP
// For each text element, find the dominant background colour
// from any overlapping filled shape. Used for contrast checks.
// ============================================================

function buildBackgroundMap(page: IRPage): Map<number, Rgba> {
  const map = new Map<number, Rgba>()

  // Collect all filled shapes
  const filledShapes = page.elements
    .filter(el => el.type === 'shape')
    .map(el => el as ShapeElement)
    .filter(s => s.fill && s.fill.a > 0.1 && s.kind !== 'line' && s.kind !== 'arrow')

  for (let i = 0; i < page.elements.length; i++) {
    const el = page.elements[i]!
    if (el.type !== 'text') continue

    const t = el as TextElement
    let bestFill: Rgba | null = null
    let bestArea = 0

    for (const shape of filledShapes) {
      const area = overlapArea(
        t.x, t.y, t.x + t.width, t.y + t.height,
        shape.x, shape.y, shape.x + shape.width, shape.y + shape.height
      )
      if (area > bestArea) {
        bestArea = area
        bestFill = shape.fill!
      }
    }

    if (bestFill && bestArea > 0) map.set(i, bestFill)
  }

  return map
}

/** Area of the intersection of two axis-aligned rectangles. Returns 0 if no overlap. */
function overlapArea(
  ax1: number, ay1: number, ax2: number, ay2: number,
  bx1: number, by1: number, bx2: number, by2: number
): number {
  const ix1 = Math.max(ax1, bx1)
  const iy1 = Math.max(ay1, by1)
  const ix2 = Math.min(ax2, bx2)
  const iy2 = Math.min(ay2, by2)
  if (ix2 <= ix1 || iy2 <= iy1) return 0
  return (ix2 - ix1) * (iy2 - iy1)
}

// ============================================================
// CONTRAST HELPERS (WCAG relative luminance)
// ============================================================

/** sRGB channel → linear light value */
function linearize(c: number): number {
  const s = c / 255
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4)
}

/** Relative luminance of an Rgba colour (0 = black, 1 = white) */
function luminance(c: Rgba): number {
  return 0.2126 * linearize(c.r) + 0.7152 * linearize(c.g) + 0.0722 * linearize(c.b)
}

/** WCAG contrast ratio between two colours (1:1 to 21:1) */
function contrastRatio(a: Rgba, b: Rgba): number {
  const la = luminance(a)
  const lb = luminance(b)
  const lighter = Math.max(la, lb)
  const darker  = Math.min(la, lb)
  return (lighter + 0.05) / (darker + 0.05)
}

const RGBA_BLACK: Rgba = { r: 0,   g: 0,   b: 0,   a: 1 }
const RGBA_WHITE: Rgba = { r: 255, g: 255, b: 255, a: 1 }

/**
 * Given a text colour and its effective background, return the colour
 * that should actually be rendered.
 * Rules:
 *   1. If the original colour has alpha < 0.5, it's intentionally decorative — leave it.
 *      (Ghost/watermark titles use low alpha; correcting them makes ghosts visible.)
 *   2. If contrast >= 3.0, colour is already readable — keep it.
 *   3. Otherwise pick black or white, whichever contrasts better with the background.
 */
function ensureContrast(textColor: Rgba, bg: Rgba): Rgba {
  // Respect intentionally semi-transparent colours — don't make them opaque
  if (textColor.a < 0.5) return textColor
  if (contrastRatio(textColor, bg) >= 3.0) return textColor
  const blackContrast = contrastRatio(RGBA_BLACK, bg)
  const whiteContrast = contrastRatio(RGBA_WHITE, bg)
  return whiteContrast >= blackContrast ? RGBA_WHITE : RGBA_BLACK
}

function renderElement(pdf: any, el: IRElement): void {
  switch (el.type) {
    case 'text':   renderText(pdf, el); break
    case 'image':  return renderImage(pdf, el)
    case 'shape':  return renderShape(pdf, el)
    case 'table':  return renderTable(pdf, el)
  }
}

// ============================================================
// TEXT
// The core challenge: jsPDF.text() with maxWidth wraps visually
// but doesn't tell us how many lines were produced. We must
// calculate wrap count ourselves using splitTextToSize() so we
// can advance cursorY correctly and avoid overlap.
// ============================================================

function renderText(
  pdf: any,
  el: TextElement,
  effectiveBg: Rgba = RGBA_WHITE
): number {
  let result = el.y
  try {
    result = _renderTextInner(pdf, el, effectiveBg)
  } catch (e) {
    console.error('[pdfkit-client] renderText failed:', e, '\nRuns:', el.runs.length, 'at', el.x, el.y)
  }
  return result
}

function _renderTextInner(
  pdf: any,
  el: TextElement,
  effectiveBg: Rgba = RGBA_WHITE
): number {
  if (el.runs.length === 0) return el.y

  // Draw shape background fill (passed as hint from parser)
  const bgFill = (el as any).bgFill as Rgba | undefined
  if (bgFill && bgFill.a > 0.05) {
    pdf.setFillColor(bgFill.r, bgFill.g, bgFill.b)
    pdf.rect(el.x, el.y, el.width, el.height, 'F')
  }

  // The actual background this text sits on — use bgFill if present, else the
  // overlapping shape colour passed in from buildBackgroundMap
  const activeBg = (bgFill && bgFill.a > 0.05) ? bgFill : effectiveBg

  // Skip ghost/decorative titles — all runs near-transparent AND no solid bg
  const textRuns = el.runs.filter(r => r.text !== '\n')
  const allFaded = textRuns.length > 0 && textRuns.every(r => r.color.a < 0.5)
  if (allFaded && !bgFill) {
    console.debug(`[faded] skipping element at y=${el.y} — all runs alpha<0.5`)
    return el.y
  }

  // Group runs into paragraphs (split by \n sentinels)
  const paragraphs: typeof el.runs[] = []
  let current: typeof el.runs = []
  for (const run of el.runs) {
    if (run.text === '\n') { paragraphs.push(current); current = [] }
    else current.push(run)
  }
  if (current.length > 0) paragraphs.push(current)

  let cursorY = el.y

  for (const para of paragraphs) {
    if (para.length === 0) {
      // Empty paragraph — blank line using last known size
      cursorY += 14 * el.lineHeight
      continue
    }

    // Don't clip — PPTX text boxes can legitimately overflow their declared height
    // (e.g. when titles wrap). Let content render naturally.

    // Dominant size for this paragraph
    const paraSize = para.reduce((max, r) => Math.max(max, r.fontSize), 12)
    const lineH    = paraSize * el.lineHeight

    // Baseline = top of line + fontSize
    const baseline = cursorY + paraSize

    let cursorX = el.x
    let maxLinesInPara = 1  // track how many wrapped lines this para needs

    for (const run of para) {
      if (!run.text) continue

      // Sanitize bullet characters — Symbol/Wingdings chars corrupt in jsPDF
      const text = sanitizeText(run.text)
      if (!text) continue

      const fontStyle = getFontStyle(run.fontWeight, run.fontStyle)
      try { pdf.setFont(mapFont(run.fontFamily), fontStyle) }
      catch { pdf.setFont('helvetica', fontStyle) }
      const safeFontSize = (run.fontSize > 0 && isFinite(run.fontSize)) ? run.fontSize : 12
      pdf.setFontSize(safeFontSize)

      // jsPDF setTextColor does not support alpha — all text renders fully opaque.
      // Skip any run that was designed to be semi-transparent (ghost/watermark titles).
      // Threshold 0.5: runs with a < 0.5 were intentionally faded and would look
      // wrong rendered solid. Runs with a >= 0.5 were intended to be visible.
      if (run.color.a < 0.5) continue

      // Ensure text is readable against its background
      const safeColor = ensureContrast(run.color, activeBg)
      pdf.setTextColor(safeColor.r, safeColor.g, safeColor.b)

      const availWidth = el.width - (cursorX - el.x)
      if (availWidth < 4) {
        // No room left on this line — move to next line
        cursorX = el.x
        cursorY += lineH
        continue
      }

      // Calculate how many lines this run will produce when wrapped
      const lines: string[] = pdf.splitTextToSize(text, availWidth)
        .filter((l: string) => typeof l === 'string' && l.length > 0)
      if (lines.length === 0) continue
      maxLinesInPara = Math.max(maxLinesInPara, lines.length)

      if (lines.length === 1) {
        // Single line — draw and advance X
        if (typeof lines[0] !== 'string' || lines[0].length === 0) { continue }
        if (!isFinite(cursorX) || !isFinite(baseline)) { continue }
        pdf.text(lines[0]!, cursorX, baseline)
        const lineW = pdf.getStringUnitWidth(lines[0]!) * (safeFontSize / pdf.internal.scaleFactor)
        // Add clickable link area if run has a URL
        if (run.url) {
          pdf.link(cursorX, baseline - run.fontSize, lineW, run.fontSize * 1.2, { url: run.url })
        }
        cursorX += lineW
      } else {
        // Multi-line — draw first line at cursorX, remaining at el.x
        if (typeof lines[0] === 'string' && lines[0].length > 0 && isFinite(cursorX) && isFinite(baseline)) {
          pdf.text(lines[0]!, cursorX, baseline)
        }
        for (let li = 1; li < lines.length; li++) {
          const lineText = lines[li]!
          const lineY = baseline + li * lineH
          if (typeof lineText === 'string' && lineText.length > 0 && isFinite(el.x) && isFinite(lineY)) {
            pdf.text(lineText, el.x, lineY)
          }
        }
        // After wrapping, X resets to where last line ends
        const lastLine = lines[lines.length - 1]!
        cursorX = el.x + pdf.getStringUnitWidth(lastLine) * (safeFontSize / pdf.internal.scaleFactor)
        maxLinesInPara = Math.max(maxLinesInPara, lines.length)
      }
    }

    // Advance Y by however many lines this paragraph actually used
    cursorY += lineH * maxLinesInPara
  }

  pdf.setTextColor(0, 0, 0)
  return cursorY
}

/**
 * Replace Symbol/Wingdings bullet characters and other non-printable
 * chars that jsPDF can't render in its built-in fonts.
 * These show up as corrupted glyphs like "%Ṗ" in output.
 */
function sanitizeText(text: string): string {
  // Single-pass replacement of characters jsPDF cannot render in built-in fonts.
  // The key offenders are Symbol/Wingdings glyphs in the Private Use Area (U+F000-U+F0FF)
  // which PPTX stores as bullet characters.
  // We MUST NOT touch real Unicode like U+2022 (•) — jsPDF handles those fine.
  return text.replace(/[\uF000-\uF0FF\u0080-\u009F]/g, ch => {
    // Wingdings / Symbol PUA — map to safe ASCII bullet
    if (ch >= '\uF000' && ch <= '\uF0FF') return '• '
    // C1 control characters — drop
    return ''
  })
}

function getFontStyle(weight: string, style: string): string {
  if (weight === 'bold' && style === 'italic') return 'bolditalic'
  if (weight === 'bold')   return 'bold'
  if (style  === 'italic') return 'italic'
  return 'normal'
}

function mapFont(family: string): string {
  const lower = family.toLowerCase()
  if (lower.includes('arial') || lower.includes('helvetica') || lower.includes('calibri')) return 'helvetica'
  if (lower.includes('times') || lower.includes('georgia') || lower.includes('garamond')) return 'times'
  if (lower.includes('courier') || lower.includes('consolas') || lower.includes('mono'))  return 'courier'
  return 'helvetica'
}

// ============================================================
// IMAGE
// ============================================================

function renderImage(pdf: any, el: ImageElement): void {
  try {
    const format = el.src.includes('png')  ? 'PNG'
                 : el.src.includes('jpeg') || el.src.includes('jpg') ? 'JPEG'
                 : 'PNG'
    pdf.addImage(el.src, format, el.x, el.y, el.width, el.height)
  } catch (e) {
    console.warn('[pdfkit-client] Failed to render image:', e)
  }
}

// ============================================================
// SHAPE
// ============================================================

function renderShape(pdf: any, el: ShapeElement): void {
  const hasFill   = el.fill   && el.fill.a   > 0.05
  const hasStroke = el.stroke && el.stroke.a > 0.05
  if (!hasFill && !hasStroke) return

  if (hasFill)   pdf.setFillColor(el.fill!.r,   el.fill!.g,   el.fill!.b)
  if (hasStroke) {
    pdf.setDrawColor(el.stroke!.r, el.stroke!.g, el.stroke!.b)
    pdf.setLineWidth(el.strokeWidth || 0.75)
  }

  const drawMode = hasFill && hasStroke ? 'FD' : hasFill ? 'F' : 'D'

  switch (el.kind) {
    case 'ellipse':
      pdf.ellipse(el.x + el.width / 2, el.y + el.height / 2, el.width / 2, el.height / 2, drawMode)
      break
    case 'line':
    case 'arrow': {
      if (!hasStroke) break
      // flipH/flipV (set by connector parser) change which corner is start/end
      const flipH = (el as any).flipH === true
      const flipV = (el as any).flipV === true
      const sx = flipH ? el.x + el.width  : el.x
      const sy = flipV ? el.y + el.height : el.y
      const ex = flipH ? el.x             : el.x + el.width
      const ey = flipV ? el.y             : el.y + el.height
      pdf.line(sx, sy, ex, ey)
      if (el.kind === 'arrow') {
        const angle   = Math.atan2(ey - sy, ex - sx)
        const headLen = Math.min(8, Math.sqrt((ex - sx) ** 2 + (ey - sy) ** 2) * 0.25)
        pdf.line(ex, ey, ex - headLen * Math.cos(angle - 0.4), ey - headLen * Math.sin(angle - 0.4))
        pdf.line(ex, ey, ex - headLen * Math.cos(angle + 0.4), ey - headLen * Math.sin(angle + 0.4))
      }
      break
    }
    default:
      pdf.rect(el.x, el.y, el.width, el.height, drawMode)
  }
}

// ============================================================
// TABLE
// ============================================================

function renderTable(pdf: any, el: TableElement): void {
  let rowY = el.y

  for (const row of el.rows) {
    let cellX = el.x

    for (let c = 0; c < row.cells.length; c++) {
      const cell      = row.cells[c]!
      const cellWidth = el.colWidths[c] ?? 50

      if (cell.style.fill && cell.style.fill.a > 0.05) {
        pdf.setFillColor(cell.style.fill.r, cell.style.fill.g, cell.style.fill.b)
        pdf.rect(cellX, rowY, cellWidth, row.height, 'F')
      }

      pdf.setDrawColor(200, 200, 200)
      pdf.setLineWidth(0.5)
      pdf.rect(cellX, rowY, cellWidth, row.height, 'D')

      if (cell.value) {
        const fontSize   = cell.style.fontSize ?? 10
        const fontWeight = cell.style.fontWeight ?? 'normal'
        const fontStyle  = cell.style.fontStyle  ?? 'normal'
        pdf.setFont('helvetica', getFontStyle(fontWeight, fontStyle))
        pdf.setFontSize(fontSize)
        const color = cell.style.color
        if (color) pdf.setTextColor(color.r, color.g, color.b)
        else       pdf.setTextColor(0, 0, 0)
        const padding = 3
        pdf.text(
          cell.value,
          cellX + padding,
          rowY + row.height / 2 + fontSize / 3,
          { maxWidth: cellWidth - padding * 2 }
        )
      }

      cellX += cellWidth
    }

    rowY += row.height
  }
}

// ============================================================
// PAGE NUMBERS
// ============================================================

function renderPageNumber(pdf: any, current: number, total: number, page: IRPage): void {
  pdf.setFont('helvetica', 'normal')
  pdf.setFontSize(9)
  pdf.setTextColor(150, 150, 150)
  pdf.text(`${current} / ${total}`, page.width / 2, page.height - 12, { align: 'center' })
  pdf.setTextColor(0, 0, 0)
}
