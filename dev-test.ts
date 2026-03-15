import { toPDF, toIR, isSupportedFile } from './src/index'

const dropZone    = document.getElementById('dropZone')!
const fileInput   = document.getElementById('fileInput') as HTMLInputElement
const logEl       = document.getElementById('log')!
const preview     = document.getElementById('preview')!
const pdfFrame    = document.getElementById('pdfFrame') as HTMLIFrameElement
const downloadBtn = document.getElementById('downloadBtn') as HTMLButtonElement

// Track blob URL so we can revoke it
let currentBlobUrl: string | null = null

function log(msg: string, type: 'ok' | 'err' | 'info' | 'dim' = 'info') {
  const span = document.createElement('span')
  span.className = `log-${type}`
  span.textContent = msg + '\n'
  if (logEl.querySelector('.log-dim')?.textContent?.startsWith('//')) {
    logEl.innerHTML = ''
  }
  logEl.appendChild(span)
  logEl.scrollTop = logEl.scrollHeight
}

function logSection(title: string) {
  log(`\n── ${title} ──────────────────────────`, 'dim')
}

async function handleFile(file: File) {
  logEl.innerHTML = ''
  preview.style.display = 'none'
  downloadBtn.disabled = true

  // Revoke previous blob URL to free memory
  if (currentBlobUrl) {
    URL.revokeObjectURL(currentBlobUrl)
    currentBlobUrl = null
  }

  log(`File: ${file.name}`, 'info')
  log(`Size: ${(file.size / 1024).toFixed(1)} KB`, 'dim')
  log(`Type: ${file.type || '(no mime type)'}`, 'dim')

  if (!isSupportedFile(file)) {
    log(`✗ Unsupported format. Expected .pptx or .xlsx`, 'err')
    return
  }

  // ── Step 1: Parse to IR ──────────────────────────────────
  logSection('Step 1: Parse → IR')
  const t0 = performance.now()
  const ir = await toIR(file)
  const parseMs = (performance.now() - t0).toFixed(1)

  if (!ir.ok) {
    log(`✗ Parse failed [${ir.error.code}]: ${ir.error.message}`, 'err')
    if (ir.error.detail) log(`  detail: ${ir.error.detail}`, 'err')
    return
  }

  log(`✓ Parsed in ${parseMs}ms`, 'ok')
  log(`  Format : ${ir.doc.format}`, 'dim')
  log(`  Pages  : ${ir.doc.pages.length}`, 'dim')

  ir.doc.pages.forEach((page, i) => {
    const counts = page.elements.reduce((acc, el) => {
      acc[el.type] = (acc[el.type] ?? 0) + 1
      return acc
    }, {} as Record<string, number>)
    const label = page.label ? ` "${page.label}"` : ''
    const breakdown = Object.entries(counts).map(([k, v]) => `${v} ${k}`).join(', ') || 'no elements'
    log(`  Page ${i + 1}${label}: ${page.width.toFixed(0)}×${page.height.toFixed(0)}pt  [${breakdown}]`, 'dim')
  })

  // ── Step 2: Generate PDF ──────────────────────────────────
  logSection('Step 2: IR → PDF')
  const t1 = performance.now()
  const result = await toPDF(file, { pageNumbers: true })
  const pdfMs = (performance.now() - t1).toFixed(1)

  if (!result.ok) {
    log(`✗ PDF generation failed: ${result.error}`, 'err')
    return
  }

  log(`✓ PDF generated in ${pdfMs}ms`, 'ok')

  // ── Step 3: Preview ────────────────────────────────────────
  logSection('Step 3: Preview')

  const blob = result.pdf.toBlob()
  log(`  PDF size: ${(blob.size / 1024).toFixed(1)} KB`, 'dim')

  if (blob.size < 1000) {
    log(`  ⚠ PDF is very small — may be empty`, 'err')
  }

  // Use blob URL instead of data URI — browsers block data: URIs in iframes
  currentBlobUrl = URL.createObjectURL(blob)
  pdfFrame.src = currentBlobUrl
  preview.style.display = 'block'

  downloadBtn.disabled = false
  downloadBtn.onclick = () => {
    const name = file.name.replace(/\.(pptx|xlsx)$/i, '.pdf')
    result.pdf.download(name)
  }

  log(`✓ Done`, 'ok')
}

// ── Drag & drop ──────────────────────────────────────────────
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault()
  dropZone.classList.add('drag-over')
})
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over')
})
dropZone.addEventListener('drop', (e) => {
  e.preventDefault()
  dropZone.classList.remove('drag-over')
  const file = e.dataTransfer?.files[0]
  if (file) handleFile(file)
})
fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0]
  if (file) handleFile(file)
})
