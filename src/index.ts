// ============================================================
// PUBLIC API — pdfkit-client
//
// This is the only file consumers need to import from.
// Everything else is internal implementation.
//
// Usage:
//   import { toPDF, toIR } from 'pdfkit-client'
//
//   const result = await toPDF(file)
//   if (result.ok) result.pdf.download('output.pdf')
//   else console.error(result.error)
// ============================================================

import { parsePptx } from './parsers/pptx.parser'
import { parseXlsx }  from './parsers/xlsx.parser'
import { generatePDF, type PDFGeneratorOptions, type GeneratedPDF } from './pdf/generator'
import type { IRDocument, ParseResult } from './ir/types'

export type { IRDocument, IRPage, IRElement } from './ir/types'
export type { GeneratedPDF, PDFGeneratorOptions } from './pdf/generator'

// ============================================================
// toPDF — main entry point
// ============================================================

export type ToPDFResult =
  | { ok: true;  pdf: GeneratedPDF }
  | { ok: false; error: string }

export interface ToPDFOptions extends PDFGeneratorOptions {
  /** Subset of page indices to include (0-indexed). Default: all pages */
  pages?: number[]
}

/**
 * Convert a .pptx or .xlsx File to a PDF.
 *
 * @example
 * const result = await toPDF(file)
 * if (result.ok) {
 *   result.pdf.download('slides.pdf')
 * }
 */
export async function toPDF(
  file: File,
  opts: ToPDFOptions = {}
): Promise<ToPDFResult> {
  try {
    // 1. Parse to IR
    const parseResult = await toIR(file)
    if (!parseResult.ok) {
      return { ok: false, error: parseResult.error.message }
    }

    let doc = parseResult.doc

    // 2. Filter pages if requested
    if (opts.pages && opts.pages.length > 0) {
      doc = {
        ...doc,
        pages: opts.pages
          .filter(i => i >= 0 && i < doc.pages.length)
          .map(i => doc.pages[i]!)
      }
    }

    // 3. Generate PDF
    const pdf = generatePDF(doc, opts)
    return { ok: true, pdf }

  } catch (err) {
    return { ok: false, error: String(err) }
  }
}

// ============================================================
// toIR — exposes the intermediate representation
// Useful for debugging, testing, or custom rendering
// ============================================================

/**
 * Parse a file to the intermediate representation without generating a PDF.
 * Useful for inspecting the parsed structure.
 */
export async function toIR(file: File): Promise<ParseResult> {
  const ext = file.name.split('.').pop()?.toLowerCase()

  switch (ext) {
    case 'pptx': return parsePptx(file)
    case 'xlsx': return parseXlsx(file)
    default:
      return {
        ok: false,
        error: {
          code:    'UNSUPPORTED_FORMAT',
          message: `Unsupported file format: .${ext}`,
          detail:  'Supported formats: .pptx, .xlsx'
        }
      }
  }
}

// ============================================================
// Supported formats
// ============================================================

export const SUPPORTED_FORMATS = ['.pptx', '.xlsx'] as const
export type SupportedFormat = typeof SUPPORTED_FORMATS[number]

export function isSupportedFile(file: File): boolean {
  return SUPPORTED_FORMATS.some(ext => file.name.toLowerCase().endsWith(ext))
}
