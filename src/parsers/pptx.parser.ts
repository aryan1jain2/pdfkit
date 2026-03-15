// ============================================================
// PPTX PARSER — pdfkit-client
// ============================================================

import type {
  IRDocument, IRPage, IRElement,
  TextElement, TextRun, ImageElement, ShapeElement, TableElement,
  TableRow, TableCell, CellStyle,
  ParseResult, Rgba
} from '../ir/types'

import { loadZip, isParseError, parseXml, getEl, getEls, attr, numAttr } from '../utils/xml'
import { emuToPt, halfPtToPt, hexToRgba, BLACK, WHITE, TRANSPARENT, PAGE_SIZES } from '../utils/units'

type ThemeColorMap = Map<string, Rgba>

// ============================================================
// ENTRY POINT
// ============================================================

export async function parsePptx(file: File | ArrayBuffer): Promise<ParseResult> {
  const zip = await loadZip(file)
  if (isParseError(zip)) return { ok: false, error: zip }

  const presXml = await zip.getText('ppt/presentation.xml')
  if (!presXml) return { ok: false, error: { code: 'MISSING_ENTRY', message: 'ppt/presentation.xml not found' } }

  const presDoc = parseXml(presXml)
  if (!presDoc) return { ok: false, error: { code: 'XML_PARSE_FAILED', message: 'Could not parse presentation.xml' } }

  const slideSize   = readSlideSize(presDoc)
  const themeXml    = await zip.getText('ppt/theme/theme1.xml')
  const themeColors = readThemeColors(themeXml)
  const relsXml     = await zip.getText('ppt/_rels/presentation.xml.rels')
  const slideOrder  = readSlideOrder(relsXml)

  const pages: IRPage[] = []

  for (const slidePath of slideOrder) {
    const slideXml = await zip.getText(`ppt/${slidePath}`)
    if (!slideXml) continue
    const slideDoc = parseXml(slideXml)
    if (!slideDoc) continue

    const slideFile   = slidePath.split('/').pop()!
    const slideRelsXml = await zip.getText(`ppt/slides/_rels/${slideFile}.rels`)
    const { images: imageMap, hyperlinks } = await buildSlideRels(zip, slideRelsXml)
    pages.push(parseSlide(slideDoc, slideSize, imageMap, hyperlinks, themeColors))
  }

  return { ok: true, doc: { format: 'pptx', pages, metadata: {} } }
}

// ============================================================
// SLIDE SIZE
// ============================================================

function readSlideSize(presDoc: Document): { width: number; height: number } {
  const sldSz = getEl(presDoc, 'sldSz')
  if (!sldSz) return PAGE_SIZES.SLIDE_WIDESCREEN
  return {
    width:  emuToPt(numAttr(sldSz, 'cx', 6858000)),
    height: emuToPt(numAttr(sldSz, 'cy', 3858750))
  }
}

// ============================================================
// THEME COLOURS
// Resolves schemeClr names (dk1, lt1, accent1 etc.) to Rgba
// ============================================================

function readThemeColors(themeXml: string | null): ThemeColorMap {
  const map: ThemeColorMap = new Map()
  if (!themeXml) return map

  const doc = parseXml(themeXml)
  if (!doc) return map

  const clrScheme = getEl(doc, 'clrScheme')
  if (!clrScheme) return map

  for (let i = 0; i < clrScheme.children.length; i++) {
    const slot = clrScheme.children[i]!
    const name = slot.localName

    const srgb = getEl(slot, 'srgbClr')
    if (srgb) { map.set(name, hexToRgba(attr(srgb, 'val'))); continue }

    const sys = getEl(slot, 'sysClr')
    if (sys) {
      const lastClr = attr(sys, 'lastClr')
      if (lastClr) map.set(name, hexToRgba(lastClr))
    }
  }

  return map
}

function resolveSchemeColor(el: Element, themeColors: ThemeColorMap): Rgba {
  const val = attr(el, 'val')
  const aliasMap: Record<string, string> = {
    tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2',
  }
  const key = aliasMap[val] ?? val
  let color = themeColors.get(key) ?? BLACK

  // Apply colour modifiers from DIRECT children only.
  // Using getEl() would do a deep search and pick up modifiers from sibling nodes.
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i]!
    const name  = child.localName

    if (name === 'lumMod') {
      const mod = numAttr(child, 'val', 100000) / 100000
      color = { ...color, r: Math.min(255, Math.round(color.r * mod)), g: Math.min(255, Math.round(color.g * mod)), b: Math.min(255, Math.round(color.b * mod)) }
    } else if (name === 'lumOff') {
      const off = numAttr(child, 'val', 0) / 100000
      color = { ...color, r: Math.min(255, Math.round(color.r + 255 * off)), g: Math.min(255, Math.round(color.g + 255 * off)), b: Math.min(255, Math.round(color.b + 255 * off)) }
    } else if (name === 'alpha') {
      color = { ...color, a: Math.min(1, numAttr(child, 'val', 100000) / 100000) }
    } else if (name === 'shade') {
      const shade = numAttr(child, 'val', 100000) / 100000
      color = { ...color, r: Math.round(color.r * shade), g: Math.round(color.g * shade), b: Math.round(color.b * shade) }
    } else if (name === 'tint') {
      const tint = numAttr(child, 'val', 100000) / 100000
      color = { ...color, r: Math.round(color.r + (255 - color.r) * (1 - tint)), g: Math.round(color.g + (255 - color.g) * (1 - tint)), b: Math.round(color.b + (255 - color.b) * (1 - tint)) }
    }
  }

  return color
}

// ============================================================
// SLIDE ORDER
// ============================================================

function readSlideOrder(relsXml: string | null): string[] {
  if (!relsXml) return []
  const doc = parseXml(relsXml)
  if (!doc) return []

  return getEls(doc, 'Relationship')
    .filter(el => attr(el, 'Type').includes('/slide"') || attr(el, 'Type').endsWith('/slide'))
    .sort((a, b) => {
      const aId = parseInt(attr(a, 'Id').replace('rId', ''))
      const bId = parseInt(attr(b, 'Id').replace('rId', ''))
      return aId - bId
    })
    .map(el => attr(el, 'Target'))
}

// ============================================================
// IMAGE MAP
// ============================================================

interface SlideRels {
  images: Map<string, string>    // rId → base64 data URL
  hyperlinks: Map<string, string> // rId → URL string
}

async function buildSlideRels(zip: any, relsXml: string | null): Promise<SlideRels> {
  const images    = new Map<string, string>()
  const hyperlinks = new Map<string, string>()
  if (!relsXml) return { images, hyperlinks }

  const doc = parseXml(relsXml)
  if (!doc) return { images, hyperlinks }

  const MIME: Record<string, string> = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp',
    emf: 'image/emf', wmf: 'image/wmf',
  }

  for (const rel of getEls(doc, 'Relationship')) {
    const type   = attr(rel, 'Type')
    const rId    = attr(rel, 'Id')
    const target = attr(rel, 'Target')

    if (type.includes('image')) {
      const path    = target.replace('../', 'ppt/')
      const ext     = path.split('.').pop()?.toLowerCase() ?? ''
      const mime    = MIME[ext] ?? 'image/png'
      const dataUrl = await zip.getDataUrl(path, mime)
      if (dataUrl) images.set(rId, dataUrl)
    } else if (type.includes('hyperlink')) {
      // External hyperlinks have TargetMode="External"
      if (target.startsWith('http') || attr(rel, 'TargetMode') === 'External') {
        hyperlinks.set(rId, target)
      }
    }
  }

  return { images, hyperlinks }
}

// ============================================================
// SLIDE PARSING
// ============================================================

function parseSlide(
  slideDoc: Document,
  size: { width: number; height: number },
  imageMap: Map<string, string>,
  hyperlinks: Map<string, string>,
  themeColors: ThemeColorMap
): IRPage {
  const elements: IRElement[] = []
  let background: Rgba = WHITE

  const bgEl = getEl(slideDoc, 'bg')
  if (bgEl) {
    const solidFill = getEl(bgEl, 'solidFill')
    if (solidFill) background = readColour(solidFill, themeColors) ?? WHITE
  }

  const spTree = getEl(slideDoc, 'spTree')
  if (!spTree) return { ...size, background, elements }

  // Regular shapes and text boxes
  for (const sp of getEls(spTree, 'sp')) {
    const el = parseShape(sp, themeColors, hyperlinks)
    if (el) elements.push(el)
  }

  // Connectors / arrows: <p:cxnSp>
  for (const cxn of getEls(spTree, 'cxnSp')) {
    const el = parseConnector(cxn, themeColors)
    if (el) elements.push(el)
  }

  // Images
  for (const pic of getEls(spTree, 'pic')) {
    const el = parsePicture(pic, imageMap)
    if (el) elements.push(el)
  }

  // Tables (and charts, which we skip gracefully)
  for (const gf of getEls(spTree, 'graphicFrame')) {
    const tbl = getEl(gf, 'tbl')
    if (tbl) {
      const el = parseTable(gf, tbl, themeColors)
      if (el) elements.push(el)
    }
  }

  return { width: size.width, height: size.height, background, elements }
}

// ============================================================
// SHAPE: <p:sp>
// ============================================================

function parseShape(sp: Element, themeColors: ThemeColorMap, hyperlinks: Map<string, string> = new Map()): IRElement | null {
  const xfrm = getEl(sp, 'xfrm')
  if (!xfrm) return null

  const off = getEl(xfrm, 'off')
  const ext = getEl(xfrm, 'ext')
  if (!off || !ext) return null

  const x = emuToPt(numAttr(off, 'x'))
  const y = emuToPt(numAttr(off, 'y'))
  const w = emuToPt(numAttr(ext, 'cx'))
  const h = emuToPt(numAttr(ext, 'cy'))

  const txBody = getEl(sp, 'txBody')
  const spPr   = getEl(sp, 'spPr')

  // Shapes with text: render background fill first, then text on top
  if (txBody) {
    return parseTextBox(txBody, x, y, w, h, themeColors, spPr, hyperlinks)
  }

  if (!spPr) return null

  // No fill = noFill element present, or absent solidFill = transparent
  const noFill    = getEl(spPr, 'noFill')
  const solidFill = getEl(spPr, 'solidFill')
  const fill      = noFill ? TRANSPARENT : solidFill ? (readColour(solidFill, themeColors) ?? TRANSPARENT) : TRANSPARENT

  const ln          = getEl(spPr, 'ln')
  const lnNoFill    = ln ? getEl(ln, 'noFill') : null
  const lnSolidFill = ln && !lnNoFill ? getEl(ln, 'solidFill') : null
  const stroke      = lnSolidFill ? (readColour(lnSolidFill, themeColors) ?? undefined) : undefined
  const strokeWidth = ln && !lnNoFill ? Math.max(emuToPt(numAttr(ln, 'w', 12700)), 0.5) : 0

  const prstGeom = getEl(spPr, 'prstGeom')
  const prst     = prstGeom ? attr(prstGeom, 'prst') : ''
  const kind     = prst === 'ellipse' ? 'ellipse' : 'rect'

  return { type: 'shape', kind, x, y, width: w, height: h, fill, stroke, strokeWidth } satisfies ShapeElement
}

// ============================================================
// CONNECTOR / ARROW: <p:cxnSp>
// Same structure as <p:sp> but always a line, never has txBody
// ============================================================

// Connector flip info — tells the renderer which corner is the start point
export interface ConnectorFlip { flipH: boolean; flipV: boolean }

function parseConnector(cxn: Element, themeColors: ThemeColorMap): ShapeElement | null {
  const xfrm = getEl(cxn, 'xfrm')
  if (!xfrm) return null

  const off = getEl(xfrm, 'off')
  const ext = getEl(xfrm, 'ext')
  if (!off || !ext) return null

  const x = emuToPt(numAttr(off, 'x'))
  const y = emuToPt(numAttr(off, 'y'))
  const w = emuToPt(numAttr(ext, 'cx'))
  const h = emuToPt(numAttr(ext, 'cy'))

  // flipH/flipV change which corner is start vs end of the line
  // flipH=1 means line goes from right-to-left (x+w → x)
  // flipV=1 means line goes from bottom-to-top (y+h → y)
  const flipH = attr(xfrm, 'flipH') === '1'
  const flipV = attr(xfrm, 'flipV') === '1'

  const spPr        = getEl(cxn, 'spPr')
  const ln          = spPr ? getEl(spPr, 'ln') : null
  const lnSolidFill = ln ? getEl(ln, 'solidFill') : null
  const stroke      = lnSolidFill ? (readColour(lnSolidFill, themeColors) ?? BLACK) : BLACK
  const strokeWidth = ln ? Math.max(emuToPt(numAttr(ln, 'w', 12700)), 0.5) : 0.75

  const tailEnd  = ln ? getEl(ln, 'tailEnd') : null
  const headEnd  = ln ? getEl(ln, 'headEnd') : null
  const hasArrow = (tailEnd && attr(tailEnd, 'type') !== 'none') ||
                   (headEnd && attr(headEnd, 'type') !== 'none')

  // Encode flip as metadata on the shape for the renderer
  const shape: ShapeElement & { flipH?: boolean; flipV?: boolean } = {
    type: 'shape',
    kind: hasArrow ? 'arrow' : 'line',
    x, y, width: w, height: h,
    fill: TRANSPARENT,
    stroke,
    strokeWidth,
  }
  if (flipH) shape.flipH = true
  if (flipV) shape.flipV = true
  return shape
}

// ============================================================
// TEXT BOX: <p:txBody>
// ============================================================

function parseTextBox(
  txBody: Element,
  x: number, y: number, w: number, h: number,
  themeColors: ThemeColorMap,
  spPr: Element | null,
  hyperlinks: Map<string, string> = new Map()
): TextElement {
  const runs: TextRun[] = []

  // Background fill from shape properties — stored as a hint for the renderer
  let bgFill: Rgba | undefined
  if (spPr) {
    const solidFill = getEl(spPr, 'solidFill')
    if (solidFill) bgFill = readColour(solidFill, themeColors) ?? undefined
  }

  // Body-level default text props
  const lstStyle   = getEl(txBody, 'lstStyle')
  const bodyPr     = getEl(txBody, 'bodyPr')

  const paragraphs = getEls(txBody, 'p')

  // Track list counter across paragraphs for auto-numbered lists
  let listCounter = 0
  let lastIndent  = -1

  for (const para of paragraphs) {
    const pPr    = getEl(para, 'pPr')
    const indent = pPr ? numAttr(pPr, 'lvl', 0) : 0

    // Reset counter when indent level changes
    if (indent !== lastIndent) {
      listCounter = 0
      lastIndent  = indent
    }

    // ── Bullet / list prefix ──────────────────────────────
    let prefix = ''

    if (pPr) {
      const buNone    = getEl(pPr, 'buNone')    // explicit no-bullet
      const buChar    = getEl(pPr, 'buChar')    // custom char bullet e.g. •
      const buAutoNum = getEl(pPr, 'buAutoNum') // numbered list
      const buFont    = getEl(pPr, 'buFont')    // wingdings etc.

      if (!buNone) {
        if (buChar) {
          const raw = attr(buChar, 'char')
          // Map Symbol/Wingdings private-use chars to standard bullet
          // These render as "%Ṗ" corruption in jsPDF's built-in fonts
          const isSafeChar = raw && raw.charCodeAt(0) >= 0x20 && raw.charCodeAt(0) <= 0x7E
          prefix = isSafeChar ? raw + ' ' : '• '
        } else if (buAutoNum) {
          listCounter++
          const type = attr(buAutoNum, 'type')
          // arabicPeriod = 1. 2. 3. | arabicParenR = 1) 2) | alphaLcPeriod = a. b.
          if (type.startsWith('alphaLc')) {
            prefix = String.fromCharCode(96 + listCounter) + '. '
          } else if (type.startsWith('alphaUc')) {
            prefix = String.fromCharCode(64 + listCounter) + '. '
          } else if (type.includes('ParenR')) {
            prefix = listCounter + ') '
          } else {
            prefix = listCounter + '. '   // default: 1. 2. 3.
          }
        }
      }
    }

    // ── Get default run props from first run for prefix styling ──
    const firstRpr    = getEl(getEls(para, 'r')[0] ?? para, 'rPr')
    const defaultSize = firstRpr ? halfPtToPt(numAttr(firstRpr, 'sz', 1800)) : 18
    const defaultColor = readRunColour(firstRpr, themeColors) ?? (bgFill ? WHITE : BLACK)

    if (prefix) {
      runs.push({
        text:          prefix,
        fontSize:      defaultSize,
        fontFamily:    'sans-serif',
        fontWeight:    'normal',
        fontStyle:     'normal',
        color:         defaultColor,
        underline:     false,
        strikethrough: false,
      })
    }

    // ── Text runs ─────────────────────────────────────────
    for (const r of getEls(para, 'r')) {
      const rPr = getEl(r, 'rPr')
      const t   = getEl(r, 't')
      if (!t?.textContent) continue

      // Determine text colour — if no explicit colour and shape has a dark bg,
      // default to white so it's visible
      let color = readRunColour(rPr, themeColors)
      if (!color) {
        color = bgFill ? WHITE : BLACK
      }

      // Hyperlink: <a:hlinkClick r:id="rId1"/> inside rPr
      let url: string | undefined
      if (rPr) {
        const hlink = getEl(rPr, 'hlinkClick')
        if (hlink) {
          const rId = attr(hlink, 'r:id') || attr(hlink, 'id')
          url = hyperlinks.get(rId)
        }
      }

      runs.push({
        text:          t.textContent,
        fontSize:      rPr ? halfPtToPt(numAttr(rPr, 'sz', 1800)) : 18,
        fontFamily:    readFontFamily(rPr) ?? 'sans-serif',
        fontWeight:    rPr && attr(rPr, 'b') === '1' ? 'bold' : 'normal',
        fontStyle:     rPr && attr(rPr, 'i') === '1' ? 'italic' : 'normal',
        color,
        underline:     rPr ? (attr(rPr, 'u') !== 'none' && attr(rPr, 'u') !== '') : false,
        strikethrough: rPr ? attr(rPr, 'strike') === 'sngStrike' : false,
        ...(url ? { url } : {})
      })
    }

    // Paragraph separator
    runs.push({
      text: '\n', fontSize: defaultSize, fontFamily: 'sans-serif',
      fontWeight: 'normal', fontStyle: 'normal',
      color: TRANSPARENT, underline: false, strikethrough: false
    })
  }

  return {
    type: 'text', x, y, width: w, height: h,
    align: 'left',
    runs,
    lineHeight: 1.2,
    // Pass background colour as hint so generator can draw box bg
    ...(bgFill ? { bgFill } : {})
  } as TextElement & { bgFill?: Rgba }
}

// ============================================================
// IMAGE: <p:pic>
// ============================================================

function parsePicture(pic: Element, imageMap: Map<string, string>): ImageElement | null {
  const blip = getEl(pic, 'blip')
  if (!blip) return null

  const rId = attr(blip, 'r:embed') || attr(blip, 'embed')
  if (!rId) return null

  const src = imageMap.get(rId)
  if (!src) return null

  const xfrm = getEl(pic, 'xfrm')
  if (!xfrm) return null

  const off = getEl(xfrm, 'off')
  const ext = getEl(xfrm, 'ext')
  if (!off || !ext) return null

  return {
    type: 'image',
    x:      emuToPt(numAttr(off, 'x')),
    y:      emuToPt(numAttr(off, 'y')),
    width:  emuToPt(numAttr(ext, 'cx')),
    height: emuToPt(numAttr(ext, 'cy')),
    src
  }
}

// ============================================================
// TABLE: <a:tbl>
// ============================================================

function parseTable(graphicFrame: Element, tbl: Element, themeColors: ThemeColorMap): TableElement | null {
  const xfrm = getEl(graphicFrame, 'xfrm')
  if (!xfrm) return null

  const off = getEl(xfrm, 'off')
  const ext = getEl(xfrm, 'ext')
  if (!off || !ext) return null

  const x = emuToPt(numAttr(off, 'x'))
  const y = emuToPt(numAttr(off, 'y'))
  const w = emuToPt(numAttr(ext, 'cx'))

  const colWidths = getEls(tbl, 'gridCol').map(c => emuToPt(numAttr(c, 'w')))
  const rows: TableRow[] = []

  for (const tr of getEls(tbl, 'tr')) {
    const rowHeight = emuToPt(numAttr(tr, 'h'))
    const cells: TableCell[] = []

    for (const tc of getEls(tr, 'tc')) {
      const texts   = getEls(tc, 't').map(t => t.textContent ?? '').join('')
      const tcPr    = getEl(tc, 'tcPr')
      const colspan = numAttr(tc, 'gridSpan', 1)
      const rowspan = numAttr(tc, 'rowSpan', 1)

      const solidFill = tcPr ? getEl(tcPr, 'solidFill') : null
      const style: CellStyle = {
        fill:  solidFill ? readColour(solidFill, themeColors) ?? undefined : undefined,
        align: 'left',
      }

      cells.push({ value: texts, colspan, rowspan, style })
    }

    rows.push({ cells, height: rowHeight })
  }

  return { type: 'table', x, y, width: w, colWidths, rows }
}

// ============================================================
// COLOUR HELPERS
// ============================================================

function readColour(parent: Element, themeColors: ThemeColorMap = new Map()): Rgba | null {
  const srgb = getEl(parent, 'srgbClr')
  if (srgb) {
    const rgba = hexToRgba(attr(srgb, 'val'))
    return applyAlpha(rgba, srgb)
  }

  const schemeClr = getEl(parent, 'schemeClr')
  if (schemeClr) {
    const rgba = resolveSchemeColor(schemeClr, themeColors)
    return applyAlpha(rgba, schemeClr)
  }

  return null
}

/** Read <a:alpha val="N"/> where N is in thousandths of a percent (100000 = opaque).
 *  Must check DIRECT children only — getEl() is a deep search and would pick up
 *  alpha elements from sibling/nested nodes and corrupt unrelated colours. */
function applyAlpha(color: Rgba, el: Element): Rgba {
  // Only iterate direct children of the colour element
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i]!
    if (child.localName === 'alpha') {
      const val = numAttr(child, 'val', 100000)
      return { ...color, a: Math.min(1, val / 100000) }
    }
  }
  return color
}

function readRunColour(rPr: Element | null, themeColors: ThemeColorMap = new Map()): Rgba | null {
  if (!rPr) return null
  const solidFill = getEl(rPr, 'solidFill')
  return solidFill ? readColour(solidFill, themeColors) : null
}

function readFontFamily(rPr: Element | null): string | null {
  if (!rPr) return null
  const latin = getEl(rPr, 'latin')
  return latin ? attr(latin, 'typeface') || null : null
}

function mapAlign(algn: string): 'left' | 'center' | 'right' | 'justify' {
  const map: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
    l: 'left', ctr: 'center', r: 'right', just: 'justify'
  }
  return map[algn] ?? 'left'
}
