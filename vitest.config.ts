import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const root = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      '@pptx/lossless-xml': `${root}packages/lossless-xml/src/index.ts`,
      '@pptx/opc': `${root}packages/opc/src/index.ts`,
      '@pptx/codecs': `${root}packages/codecs/src/index.ts`,
      '@pptx/model': `${root}packages/model/src/index.ts`,
      '@pptx/pptxgenjs-adapter': `${root}packages/pptxgenjs-adapter/src/index.ts`,
      '@pptx/validator': `${root}packages/validator/src/index.ts`,
      '@pptx/sdk': `${root}packages/sdk/src/index.ts`,
      '@pptx/testkit': `${root}packages/testkit/src/index.ts`,
    },
  },
  test: {
    include: ['packages/**/*.test.ts'],
    coverage: { reporter: ['text', 'json-summary'] },
  },
});
