import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs', 'esm'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  minify: false,
  treeshake: true,
  external: ['jszip'],
  globalName: 'XmlXlsxLite',
  banner: {
    js: '// xml-xlsx-lite – Minimal XLSX writer using raw XML + JSZip\n// https://github.com/mikemikex1/xml-xlsx-lite\n',
  },
})
