import { defineConfig } from 'vitest/config'
import { resolve } from 'path'

export default defineConfig({
  test: {
    // jsdom gives us DOMParser, File, Blob, TextEncoder in Node
    environment: 'jsdom',

    include:  ['tests/**/*.test.ts'],

    coverage: {
      provider:  'v8',
      reporter:  ['text', 'html'],
      include:   ['src/**/*.ts'],
      exclude:   ['src/index.ts'],
      thresholds: {
        statements: 70,
        branches:   60,
        functions:  70,
        lines:      70,
      }
    },

    reporter: 'verbose',
  },

  resolve: {
    alias: { '@': resolve(__dirname, 'src') }
  }
})
