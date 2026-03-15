// ============================================================
// TESTS — pdfkit-client
// Run with: npx vitest run
// Run with coverage: npx vitest run --coverage
// ============================================================

import { describe, it, expect } from 'vitest'
import { cellRefToIndex } from '../src/parsers/xlsx.parser'
import {
  emuToPt, ptToEmu, halfPtToPt,
  pxToPt, ptToPx,
  hexToRgba, rgbaToCss, rgbaToHex,
  cmToPt, twipToPt,
  PAGE_SIZES
} from '../src/utils/units'
import { parseXml, getText, getEl, getEls, attr, numAttr } from '../src/utils/xml'
import { isSupportedFile, SUPPORTED_FORMATS } from '../src/index'

// ============================================================
// UNIT CONVERSIONS
// ============================================================

describe('emuToPt', () => {
  it('1 inch = 914400 EMU = 72pt', () =>
    expect(emuToPt(914400)).toBeCloseTo(72))
  it('0 EMU = 0pt', () => expect(emuToPt(0)).toBe(0))
  it('typical PPTX slide width = 540pt', () =>
    expect(emuToPt(6858000)).toBeCloseTo(540, 0))
  it('is reversible with ptToEmu', () =>
    expect(emuToPt(ptToEmu(72))).toBeCloseTo(72))
})

describe('halfPtToPt', () => {
  it('2400 → 24pt', () => expect(halfPtToPt(2400)).toBe(24))
  it('1800 → 18pt', () => expect(halfPtToPt(1800)).toBe(18))
  it('1100 → 11pt', () => expect(halfPtToPt(1100)).toBe(11))
  it('0 → 0',       () => expect(halfPtToPt(0)).toBe(0))
})

describe('pxToPt / ptToPx', () => {
  it('96px = 72pt', () => expect(pxToPt(96)).toBeCloseTo(72))
  it('is reversible', () => expect(pxToPt(ptToPx(36))).toBeCloseTo(36))
})

describe('cmToPt', () => {
  it('1cm ≈ 28.35pt', () => expect(cmToPt(1)).toBeCloseTo(28.35, 1))
  it('2.54cm ≈ 72pt (1 inch)', () => expect(cmToPt(2.54)).toBeCloseTo(72, 0))
})

describe('twipToPt', () => {
  it('1440 twips = 72pt', () => expect(twipToPt(1440)).toBeCloseTo(72))
})

// ============================================================
// COLOUR HELPERS
// ============================================================

describe('hexToRgba', () => {
  it('parses white',         () => expect(hexToRgba('FFFFFF')).toEqual({ r: 255, g: 255, b: 255, a: 1 }))
  it('parses black',         () => expect(hexToRgba('000000')).toEqual({ r: 0, g: 0, b: 0, a: 1 }))
  it('parses red',           () => expect(hexToRgba('FF0000')).toEqual({ r: 255, g: 0, b: 0, a: 1 }))
  it('strips leading #',     () => expect(hexToRgba('#00FF00')).toEqual({ r: 0, g: 255, b: 0, a: 1 }))
  it('accepts custom alpha', () => expect(hexToRgba('0000FF', 0.5)).toEqual({ r: 0, g: 0, b: 255, a: 0.5 }))
  it('parses lowercase',     () => expect(hexToRgba('ff6b35')).toEqual({ r: 255, g: 107, b: 53, a: 1 }))
})

describe('rgbaToCss', () => {
  it('produces rgba string', () =>
    expect(rgbaToCss({ r: 255, g: 0, b: 0, a: 1 })).toBe('rgba(255,0,0,1)'))
  it('includes alpha', () =>
    expect(rgbaToCss({ r: 0, g: 0, b: 0, a: 0.5 })).toBe('rgba(0,0,0,0.5)'))
})

describe('rgbaToHex', () => {
  it('converts red',   () => expect(rgbaToHex({ r: 255, g: 0,   b: 0,   a: 1 })).toBe('#ff0000'))
  it('converts white', () => expect(rgbaToHex({ r: 255, g: 255, b: 255, a: 1 })).toBe('#ffffff'))
  it('converts black', () => expect(rgbaToHex({ r: 0,   g: 0,   b: 0,   a: 1 })).toBe('#000000'))
})

// ============================================================
// PAGE SIZES
// ============================================================

describe('PAGE_SIZES', () => {
  it('A4 is portrait', () =>
    expect(PAGE_SIZES.A4.height).toBeGreaterThan(PAGE_SIZES.A4.width))
  it('SLIDE_WIDESCREEN is 16:9', () => {
    const ratio = PAGE_SIZES.SLIDE_WIDESCREEN.width / PAGE_SIZES.SLIDE_WIDESCREEN.height
    expect(ratio).toBeCloseTo(16 / 9, 1)
  })
  it('SLIDE_STANDARD is 4:3', () => {
    const ratio = PAGE_SIZES.SLIDE_STANDARD.width / PAGE_SIZES.SLIDE_STANDARD.height
    expect(ratio).toBeCloseTo(4 / 3, 1)
  })
})

// ============================================================
// XML UTILITIES
// ============================================================

describe('parseXml', () => {
  it('parses valid XML', () => {
    const doc = parseXml('<root><child attr="val">text</child></root>')
    expect(doc).not.toBeNull()
    expect(doc!.documentElement.tagName).toBe('root')
  })
  it('does not throw on bad XML', () => {
    expect(() => parseXml('<unclosed>')).not.toThrow()
  })
  it('returns null for empty string', () => {
    expect(parseXml('')).toBeNull()
  })
})

describe('getEl / getEls', () => {
  it('finds element by local name across namespaces', () => {
    const doc = parseXml('<root><p:sp xmlns:p="urn:ppt"><a:t xmlns:a="urn:dml">hi</a:t></p:sp></root>')!
    expect(getEl(doc, 'sp')).not.toBeNull()
  })
  it('returns null for missing element', () => {
    expect(getEl(parseXml('<root/>')!, 'missing')).toBeNull()
  })
  it('returns all matching elements', () => {
    const doc = parseXml('<root><item/><item/><item/></root>')!
    expect(getEls(doc, 'item')).toHaveLength(3)
  })
  it('returns empty array when none found', () => {
    expect(getEls(parseXml('<root/>')!, 'x')).toEqual([])
  })
})

describe('attr / numAttr', () => {
  const el = parseXml('<el x="100" y="abc" />')!.documentElement
  it('reads string attribute',          () => expect(attr(el, 'x')).toBe('100'))
  it('returns empty string if missing', () => expect(attr(el, 'z')).toBe(''))
  it('reads numeric attribute',         () => expect(numAttr(el, 'x')).toBe(100))
  it('uses fallback for non-numeric',   () => expect(numAttr(el, 'y', 42)).toBe(42))
  it('uses fallback for missing attr',  () => expect(numAttr(el, 'z', 99)).toBe(99))
})

describe('getText', () => {
  it('returns text content of child', () => {
    const doc = parseXml('<root><title>Hello</title></root>')!
    expect(getText(doc, 'title')).toBe('Hello')
  })
  it('returns empty string when missing', () => {
    expect(getText(parseXml('<root/>')!, 'title')).toBe('')
  })
})

// ============================================================
// CELL REFERENCE PARSING
// ============================================================

describe('cellRefToIndex', () => {
  it('A1  → col 0,   row 0',   () => expect(cellRefToIndex('A1')).toEqual({ col: 0, row: 0 }))
  it('B2  → col 1,   row 1',   () => expect(cellRefToIndex('B2')).toEqual({ col: 1, row: 1 }))
  it('Z1  → col 25,  row 0',   () => expect(cellRefToIndex('Z1')).toEqual({ col: 25, row: 0 }))
  it('AA1 → col 26,  row 0',   () => expect(cellRefToIndex('AA1')).toEqual({ col: 26, row: 0 }))
  it('AB3 → col 27,  row 2',   () => expect(cellRefToIndex('AB3')).toEqual({ col: 27, row: 2 }))
  it('AZ1 → col 51,  row 0',   () => expect(cellRefToIndex('AZ1')).toEqual({ col: 51, row: 0 }))
  it('BA1 → col 52,  row 0',   () => expect(cellRefToIndex('BA1')).toEqual({ col: 52, row: 0 }))
  it('ZZ1 → col 701, row 0',   () => expect(cellRefToIndex('ZZ1')).toEqual({ col: 701, row: 0 }))
  it('A100  → row 99',         () => expect(cellRefToIndex('A100')).toEqual({ col: 0, row: 99 }))
  it('A1000 → row 999',        () => expect(cellRefToIndex('A1000')).toEqual({ col: 0, row: 999 }))
  it('handles empty string',   () => expect(cellRefToIndex('')).toEqual({ col: 0, row: 0 }))
})

// ============================================================
// PUBLIC API
// ============================================================

describe('isSupportedFile', () => {
  const f = (name: string) => new File([], name)
  it('accepts .pptx',       () => expect(isSupportedFile(f('deck.pptx'))).toBe(true))
  it('accepts .xlsx',       () => expect(isSupportedFile(f('data.xlsx'))).toBe(true))
  it('rejects .docx',       () => expect(isSupportedFile(f('doc.docx'))).toBe(false))
  it('rejects .pdf',        () => expect(isSupportedFile(f('file.pdf'))).toBe(false))
  it('rejects .png',        () => expect(isSupportedFile(f('img.png'))).toBe(false))
  it('is case-insensitive', () => expect(isSupportedFile(f('DECK.PPTX'))).toBe(true))
})

describe('SUPPORTED_FORMATS', () => {
  it('contains .pptx and .xlsx', () => {
    expect(SUPPORTED_FORMATS).toContain('.pptx')
    expect(SUPPORTED_FORMATS).toContain('.xlsx')
  })
})
