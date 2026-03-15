// ============================================================
// INTERMEDIATE REPRESENTATION (IR) — pdfkit-client
// This is the language-neutral data model that sits between
// Office XML parsing and PDF rendering. All parsers produce
// this format. All renderers consume it.
// ============================================================

// ------ Units -----------------------------------------------
// All measurements in our IR are in POINTS (pt)
// 1 pt = 1/72 inch
// Conversion helpers live in src/utils/units.ts

export type Pt = number  // points
export type Rgba = { r: number; g: number; b: number; a: number }

// ------ Text ------------------------------------------------

export type FontWeight = 'normal' | 'bold'
export type FontStyle  = 'normal' | 'italic'
export type TextAlign  = 'left' | 'center' | 'right' | 'justify'

export interface TextRun {
  text: string
  fontSize:   Pt
  fontFamily: string
  fontWeight: FontWeight
  fontStyle:  FontStyle
  color:      Rgba
  underline:  boolean
  strikethrough: boolean
  url?: string   // hyperlink URL if this run is a clickable link
}

export interface TextElement {
  type: 'text'
  x: Pt
  y: Pt
  width:  Pt
  height: Pt
  align:  TextAlign
  runs:   TextRun[]      // A paragraph can mix bold/italic/color mid-line
  lineHeight: number     // multiplier, e.g. 1.2
}

// ------ Images ----------------------------------------------

export interface ImageElement {
  type: 'image'
  x: Pt
  y: Pt
  width:  Pt
  height: Pt
  src: string            // base64 data URL  e.g. "data:image/png;base64,..."
  alt?: string
}

// ------ Shapes ----------------------------------------------

export type ShapeKind = 'rect' | 'ellipse' | 'line' | 'arrow'

export interface ShapeElement {
  type:    'shape'
  kind:    ShapeKind
  x: Pt
  y: Pt
  width:  Pt
  height: Pt
  fill?:   Rgba
  stroke?: Rgba
  strokeWidth: Pt
}

// ------ Table (Excel / PPTX tables) -------------------------

export interface CellStyle {
  fill?:       Rgba
  color?:      Rgba
  fontSize?:   Pt
  fontWeight?: FontWeight
  fontStyle?:  FontStyle
  align?:      TextAlign
  borderTop?:    boolean
  borderBottom?: boolean
  borderLeft?:   boolean
  borderRight?:  boolean
}

export interface TableCell {
  value:       string    // plain text content
  colspan:     number
  rowspan:     number
  style:       CellStyle
}

export interface TableRow {
  cells:  TableCell[]
  height: Pt
}

export interface TableElement {
  type:    'table'
  x: Pt
  y: Pt
  width:   Pt
  colWidths: Pt[]        // width of each column
  rows:    TableRow[]
}

// ------ Union of all elements --------------------------------

export type IRElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | TableElement

// ------ Page / Slide ----------------------------------------

export interface IRPage {
  width:      Pt
  height:     Pt
  background: Rgba
  elements:   IRElement[]

  // metadata (optional, used for Excel sheet names etc.)
  label?: string
}

// ------ Document (top-level) --------------------------------

export type SourceFormat = 'xlsx' | 'pptx'

export interface IRDocument {
  format:  SourceFormat
  pages:   IRPage[]
  metadata: {
    title?:   string
    author?:  string
    created?: string
  }
}

// ------ Parser result ---------------------------------------
// Parsers return this — either success with IR or a structured error

export type ParseResult =
  | { ok: true;  doc: IRDocument }
  | { ok: false; error: ParseError }

export interface ParseError {
  code:    ParseErrorCode
  message: string
  detail?: string
}

export type ParseErrorCode =
  | 'INVALID_FILE'
  | 'UNSUPPORTED_FORMAT'
  | 'CORRUPT_ZIP'
  | 'MISSING_ENTRY'
  | 'XML_PARSE_FAILED'
  | 'UNKNOWN'

// ── Hyperlink support ─────────────────────────────────────────
// Added to TextRun so individual runs can be clickable links
