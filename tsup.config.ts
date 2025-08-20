import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs', 'esm'],
  dts: {
    resolve: true,
  },
  splitting: false,
  sourcemap: true,
  clean: true,
  minify: false,
  treeshake: true,
  external: ['jszip'],
  globalName: 'XmlXlsxLite',
  outExtension({ format }) {
    return {
      js: format === 'esm' ? '.esm.js' : '.js',
    }
  },
  banner: {
    js: '// xml-xlsx-lite – Minimal XLSX writer using raw XML + JSZip\n// https://github.com/mikemikex1/xml-xlsx-lite\n',
  },
})
