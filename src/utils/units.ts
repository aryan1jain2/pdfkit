// ============================================================
// UNIT CONVERSIONS — pdfkit-client
//
// Office formats use different unit systems. We normalise
// everything to POINTS (pt) in our IR.
//
// Reference:
//   1 inch      = 72 pt
//   1 inch      = 96 px  (screen)
//   1 inch      = 914400 EMU  (PPTX/DOCX English Metric Units)
//   1 inch      = 9144 twips  (older Word formats)
//   1 cm        = 28.3465 pt
// ============================================================

const PT_PER_INCH   = 72
const EMU_PER_INCH  = 914400
const PX_PER_INCH   = 96
const TWIP_PER_INCH = 1440

// ------ EMU (PPTX, DOCX) ------------------------------------

/** Convert EMUs (English Metric Units) to points */
export function emuToPt(emu: number): number {
  return (emu / EMU_PER_INCH) * PT_PER_INCH
}

/** Convert points to EMUs */
export function ptToEmu(pt: number): number {
  return (pt / PT_PER_INCH) * EMU_PER_INCH
}

// ------ Pixels ----------------------------------------------

/** Convert CSS pixels to points (assumes 96dpi screen) */
export function pxToPt(px: number): number {
  return (px / PX_PER_INCH) * PT_PER_INCH
}

/** Convert points to CSS pixels */
export function ptToPx(pt: number): number {
  return (pt / PT_PER_INCH) * PX_PER_INCH
}

// ------ Twips (legacy Word) ---------------------------------

/** Convert twips to points */
export function twipToPt(twip: number): number {
  return (twip / TWIP_PER_INCH) * PT_PER_INCH
}

// ------ Half-points (PPTX font sizes) -----------------------

/** PPTX stores font sizes in half-points (hundredths of a point).
 *  e.g. <a:rPr sz="2400" /> means 24pt */
export function halfPtToPt(halfPt: number): number {
  return halfPt / 100
}

// ------ Centimetres -----------------------------------------

export function cmToPt(cm: number): number {
  return cm * 28.3465
}

// ------ Colour helpers --------------------------------------

import type { Rgba } from '../ir/types'

/** Parse a 6-char hex colour string (with or without #) to Rgba */
export function hexToRgba(hex: string, alpha = 1): Rgba {
  const clean = hex.replace('#', '')
  const r = parseInt(clean.slice(0, 2), 16)
  const g = parseInt(clean.slice(2, 4), 16)
  const b = parseInt(clean.slice(4, 6), 16)
  return { r, g, b, a: alpha }
}

/** Default fallback colours */
export const BLACK: Rgba = { r: 0,   g: 0,   b: 0,   a: 1 }
export const WHITE: Rgba = { r: 255, g: 255, b: 255, a: 1 }
export const TRANSPARENT: Rgba = { r: 0, g: 0, b: 0, a: 0 }

/** Convert Rgba to CSS rgba() string */
export function rgbaToCss(c: Rgba): string {
  return `rgba(${c.r},${c.g},${c.b},${c.a})`
}

/** Convert Rgba to hex string (loses alpha) */
export function rgbaToHex(c: Rgba): string {
  return (
    '#' +
    [c.r, c.g, c.b]
      .map(v => v.toString(16).padStart(2, '0'))
      .join('')
  )
}

// ------ Page size presets -----------------------------------

export const PAGE_SIZES = {
  A4:     { width: 595.28, height: 841.89 },  // pt
  LETTER: { width: 612,    height: 792     },
  A3:     { width: 841.89, height: 1190.55 },
  SLIDE_WIDESCREEN: { width: 720, height: 405 },   // 16:9 PPTX default
  SLIDE_STANDARD:   { width: 720, height: 540 },   // 4:3 PPTX default
} as const

export type PageSizeName = keyof typeof PAGE_SIZES
