import { defineConfig } from 'tsup'

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    'cell/index': 'src/cell/index.ts',
    'sheet/index': 'src/sheet/index.ts',
    'workbook/index': 'src/workbook/index.ts',
    'formula/index': 'src/formula/index.ts',
  },
  format: ['esm'],
  dts: true,
  sourcemap: true,
  clean: true,
  splitting: false,
  treeshake: true,
  external: ['mongo.do', 'gitx.do'],
})
