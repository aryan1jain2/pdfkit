// ============================================================
// ZIP + XML UTILITIES — pdfkit-client
//
// Both .xlsx and .pptx are ZIP archives containing XML files.
// These helpers abstract the unzipping and XML parsing so
// individual parsers stay focused on format-specific logic.
//
// Depends on: JSZip (loaded via CDN or npm)
// ============================================================

import JSZip from 'jszip'
import type { ParseError } from '../ir/types'

// ------ ZIP -------------------------------------------------

export interface ZipArchive {
  /** Get text content of a file by path, or null if missing */
  getText(path: string): Promise<string | null>
  /** Get binary content as Uint8Array, or null if missing */
  getBytes(path: string): Promise<Uint8Array | null>
  /** Get file as base64 data URL */
  getDataUrl(path: string, mimeType: string): Promise<string | null>
  /** List all file paths in the archive */
  listFiles(): string[]
}

/**
 * Load a File or ArrayBuffer as a ZIP archive.
 * Returns a thin wrapper with convenient read methods.
 */
export async function loadZip(
  input: File | ArrayBuffer
): Promise<ZipArchive | ParseError> {
  try {
    const zip = await JSZip.loadAsync(input)

    return {
      async getText(path: string): Promise<string | null> {
        const file = zip.file(path)
        if (!file) return null
        return file.async('string')
      },

      async getBytes(path: string): Promise<Uint8Array | null> {
        const file = zip.file(path)
        if (!file) return null
        return file.async('uint8array')
      },

      async getDataUrl(path: string, mimeType: string): Promise<string | null> {
        const file = zip.file(path)
        if (!file) return null
        const b64 = await file.async('base64')
        return `data:${mimeType};base64,${b64}`
      },

      listFiles(): string[] {
        return Object.keys(zip.files).filter(k => !zip.files[k].dir)
      }
    }
  } catch (e) {
    return {
      code: 'CORRUPT_ZIP',
      message: 'Failed to open file as ZIP archive',
      detail: String(e)
    }
  }
}

export function isParseError(x: unknown): x is ParseError {
  return typeof x === 'object' && x !== null && 'code' in x
}

// ------ XML -------------------------------------------------

/**
 * Parse an XML string to a Document.
 * Returns null if parsing fails.
 */
export function parseXml(xmlString: string): Document | null {
  try {
    const parser = new DOMParser()
    const doc = parser.parseFromString(xmlString, 'application/xml')
    // DOMParser signals errors via a <parsererror> element
    if (doc.querySelector('parsererror')) return null
    return doc
  } catch {
    return null
  }
}

/**
 * Get the text content of the first matching element.
 * Namespace-agnostic: matches local name only.
 */
export function getText(parent: Element | Document, localName: string): string {
  const el = getEl(parent, localName)
  return el?.textContent?.trim() ?? ''
}

/**
 * Get the first element matching a local name (ignores namespace prefix).
 */
export function getEl(
  parent: Element | Document,
  localName: string
): Element | null {
  // getElementsByTagNameNS('*', ...) finds across all namespaces
  const results = (parent as Element).getElementsByTagNameNS?.('*', localName)
    ?? (parent as Document).getElementsByTagNameNS('*', localName)
  return results.length > 0 ? results[0] : null
}

/**
 * Get all elements matching a local name.
 */
export function getEls(
  parent: Element | Document,
  localName: string
): Element[] {
  const results = (parent as Element).getElementsByTagNameNS?.('*', localName)
    ?? (parent as Document).getElementsByTagNameNS('*', localName)
  return Array.from(results)
}

/**
 * Get an attribute value. Handles plain and namespaced (r:embed etc.)
 * attributes by trying getAttribute first, then scanning by local name.
 */
export function attr(el: Element, name: string): string {
  const direct = el.getAttribute(name)
  if (direct !== null) return direct

  // Fallback: match by local name only (handles r:embed, p:val, etc.)
  const localName = name.includes(':') ? name.split(':')[1]! : name
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i]!
    if (a.localName === localName) return a.value
  }
  return ''
}

/**
 * Get a numeric attribute, returning a fallback if missing/NaN.
 */
export function numAttr(el: Element, name: string, fallback = 0): number {
  const v = parseFloat(el.getAttribute(name) ?? '')
  return isNaN(v) ? fallback : v
}
