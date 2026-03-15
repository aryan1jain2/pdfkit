import { defineConfig } from 'vite'
import dts from 'vite-plugin-dts'
import { resolve } from 'path'

export default defineConfig({
  plugins: [
    dts({
      include:     ['src/**/*.ts'],
      outDir:      'dist',
      rollupTypes: true,       // bundles into single index.d.ts
    })
  ],

  build: {
    lib: {
      entry:    resolve(__dirname, 'src/index.ts'),
      name:     'PdfkitClient',
      fileName: (format) => {
        if (format === 'es')  return 'index.esm.js'
        if (format === 'umd') return 'index.umd.js'
        return `index.${format}.js`
      },
      formats: ['es', 'umd'],
    },

    rollupOptions: {
      // jszip + jspdf are peer deps — users bring their own copy
      external: ['jszip', 'jspdf'],
      output: {
        globals: {
          jszip: 'JSZip',
          jspdf: 'jsPDF',
        },
        sourcemap: true,
      }
    },

    minify:      false,   // let consumers minify their final bundle
    outDir:      'dist',
    emptyOutDir: true,
  },

  resolve: {
    alias: { '@': resolve(__dirname, 'src') }
  }
})
