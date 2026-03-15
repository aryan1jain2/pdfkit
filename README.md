# pdfkit-client

> Client-side `.pptx` and `.xlsx` → PDF conversion. No server. No uploads. Pure browser.

[![npm version](https://badge.fury.io/js/%40ozh1n%2Fpdfkit.svg)](https://badge.fury.io/js/@ozh1n/pdfkit)

---

## Why

Every other solution either requires a server, sends your files to a third party, or is abandonware. This library runs entirely in the browser using the File API, JSZip, and jsPDF.

**Trade-off you should know about:** Fidelity is ~70-80% on complex files. Charts, advanced animations, and exotic fonts won't transfer. Simple-to-moderate presentations and spreadsheets work well.

---

## Install

```bash
npm install @ozh1n/pdfkit jszip jspdf
```

Or via CDN:
```html
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="https://unpkg.com/jspdf/dist/jspdf.umd.min.js"></script>
<script src="https://unpkg.com/@ozh1n/pdfkit/dist/index.umd.js"></script>
```

---

## Usage

### Basic

```js
import { toPDF } from 'pdfkit-client'

const input = document.querySelector('input[type="file"]')
input.addEventListener('change', async () => {
  const file = input.files[0]
  const result = await toPDF(file)

  if (result.ok) {
    result.pdf.download('output.pdf')   // triggers browser download
  } else {
    console.error(result.error)
  }
})
```

### With options

```js
const result = await toPDF(file, {
  pages:       [0, 1, 2],       // only first 3 pages/sheets
  pageNumbers: true,            // add page numbers
  title:       'My Document',  // PDF metadata
})
```

### Access the PDF as a Blob (for upload, preview, etc.)

```js
const result = await toPDF(file)
if (result.ok) {
  const blob    = result.pdf.toBlob()      // Blob
  const dataUrl = result.pdf.toDataUrl()   // base64 string for <iframe src>
}
```

### Inspect the Intermediate Representation

```js
import { toIR } from 'pdfkit-client'

const ir = await toIR(file)
if (ir.ok) {
  console.log(ir.doc.pages.length, 'pages')
  console.log(ir.doc.pages[0].elements)   // all parsed elements
}
```

---

## Supported formats

| Format | Support | Notes |
|--------|---------|-------|
| `.pptx` | ✅ | Text, images, shapes, tables |
| `.xlsx` | ✅ | Cell data, basic styles, multiple sheets |
| Charts | ⚠️ planned v0.2 | |
| Animations | ❌ never | PDF is static |
| Complex formulas | ⚠️ values only | Formula results are preserved |

---

## Architecture

```
File (.xlsx / .pptx)
    ↓
JSZip → unzip in memory
    ↓
Format Parser → Intermediate Representation (IR)
    ↓
PDF Generator (jsPDF) → PDF Blob
    ↓
Download / Blob / DataURL
```

The IR is a format-neutral data model. This separation means:
- Parsers are independently testable
- You can write custom renderers (e.g. HTML preview, PNG export)
- The public `toIR()` API lets you build on top of the parsed data

---

## Contributing

```bash
git clone https://github.com/aryan1jain2/pdfkit
npm install
npm run dev      # demo app at localhost:5173
npm test         # unit tests
npm run build    # produces dist/
```

PRs welcome, especially for:
- Chart rendering
- Better font substitution
- PPTX theme colour resolution
- Merged cell support in XLSX

---

## License

MIT
