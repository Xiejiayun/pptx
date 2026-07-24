import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    cli: 'src/cli.ts',
  },
  format: ['esm'],
  platform: 'node',
  target: 'node20',
  dts: false,
  splitting: true,
  clean: true,
  minify: false,
  sourcemap: false,
  treeshake: true,
  noExternal: [/.*/],
  banner: {
    js: "import { createRequire as __pptxCreateRequire } from 'node:module'; const require = __pptxCreateRequire(import.meta.url);",
  },
});
