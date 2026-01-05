import { defineConfig } from 'vitest/config'
import { resolve } from 'path'

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['src/**/*.test.ts', 'test/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.test.ts', 'src/**/*.d.ts'],
    },
    testTimeout: 10000,
  },
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
      '@/cell': resolve(__dirname, 'src/cell'),
      '@/sheet': resolve(__dirname, 'src/sheet'),
      '@/workbook': resolve(__dirname, 'src/workbook'),
      '@/formula': resolve(__dirname, 'src/formula'),
      '@/types': resolve(__dirname, 'src/types'),
    },
  },
})
