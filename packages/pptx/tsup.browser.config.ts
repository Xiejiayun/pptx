import { defineConfig } from 'tsup';

export default defineConfig({
  entry: { browser: 'src/browser.ts' },
  format: ['esm'],
  platform: 'browser',
  target: 'es2022',
  dts: false,
  splitting: false,
  clean: false,
  minify: false,
  sourcemap: false,
  treeshake: true,
  noExternal: [/.*/],
});
