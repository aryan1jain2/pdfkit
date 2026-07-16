import { defineConfig } from 'vite'
import { resolve } from 'path'

// Builds the drag-and-drop demo page (index.html + dev-test.ts) as a static
// site for GitHub Pages — separate from the library build in vite.config.ts.
// `base` must match the repo name so asset URLs resolve at aryan1jain2.github.io/pdfkit/.
export default defineConfig({
  base: '/pdfkit/',
  build: {
    outDir: 'demo-dist',
    emptyOutDir: true,
  },
  resolve: {
    alias: { '@': resolve(__dirname, 'src') },
  },
})
