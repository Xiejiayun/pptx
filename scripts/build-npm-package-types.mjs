import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const outputRoot = join(repositoryRoot, 'packages/pptx/dist');
const typesRoot = join(outputRoot, 'types');
const publicInternalDeclarations = new Set([
  'chart-diagnostics.internal.d.ts',
  'chart-state.internal.d.ts',
  'chart-workbook.internal.d.ts',
  'media-state.internal.d.ts',
  'media-timing-state.internal.d.ts',
  'presentation-slide-number.internal.d.ts',
  'slide-background.internal.d.ts',
  'slide-number.internal.d.ts',
]);

const packages = new Map([
  ['lossless-xml', 'lossless-xml'],
  ['opc', 'opc'],
  ['codecs', 'codecs'],
  ['model', 'model'],
  ['validator', 'validator'],
  ['sdk', 'sdk'],
  ['pptxgenjs-adapter', 'pptxgenjs-adapter'],
  ['plugin-transitions', 'plugin-transitions'],
  ['plugin-animations', 'plugin-animations'],
  ['plugin-advanced-charts', 'plugin-advanced-charts'],
  ['plugin-smartart', 'plugin-smartart'],
]);

await rm(typesRoot, { recursive: true, force: true });

for (const [packageName, outputName] of packages) {
  const sourceDirectory = packageName.startsWith('plugin-')
    ? join(repositoryRoot, 'plugins', packageName.slice('plugin-'.length), 'dist')
    : join(repositoryRoot, 'packages', packageName, 'dist');
  const outputDirectory = join(typesRoot, outputName);
  await copyDeclarations(sourceDirectory, outputDirectory);
}

const requiredPublicDeclarations = new Map([
  [join(typesRoot, 'model/scheme-color.d.ts'), [
    'SCHEME_COLORS',
    'SchemeColor',
  ]],
  [join(typesRoot, 'model/placeholder.d.ts'), [
    'PLACEHOLDER_TYPES',
    'PlaceholderType',
    'PlaceholderSelector',
  ]],
  [join(typesRoot, 'sdk/master-layout.d.ts'), [
    'DefineSlideMasterOptions',
    'SlideLayoutModel',
    'SlideMasterModel',
  ]],
]);
for (const [declaration, exports] of requiredPublicDeclarations) {
  const source = await readFile(declaration, 'utf8');
  for (const name of exports) {
    if (!source.includes(name)) {
      throw new Error(`Packed declaration ${relative(repositoryRoot, declaration)} is missing ${name}`);
    }
  }
}

await writeFile(
  join(outputRoot, 'index.d.ts'),
  [
    "export * from './types/sdk/index.js';",
    "export * from './types/pptxgenjs-adapter/index.js';",
    "export * as transitions from './types/plugin-transitions/index.js';",
    "export * as animations from './types/plugin-animations/index.js';",
    "export * as advancedCharts from './types/plugin-advanced-charts/index.js';",
    "export * as smartArt from './types/plugin-smartart/index.js';",
    '',
  ].join('\n'),
);

async function copyDeclarations(sourceDirectory, outputDirectory) {
  await mkdir(outputDirectory, { recursive: true });
  const entries = await readdir(sourceDirectory, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = join(sourceDirectory, entry.name);
    const outputPath = join(outputDirectory, entry.name);
    if (entry.isDirectory()) {
      await copyDeclarations(sourcePath, outputPath);
      continue;
    }
    if (!entry.name.endsWith('.d.ts') || entry.name.includes('.test.') ||
        (entry.name.includes('.internal.') && !publicInternalDeclarations.has(entry.name))) continue;
    const source = await readFile(sourcePath, 'utf8');
    const rewritten = source.replace(/(['"])@pptx\/([a-z0-9-]+)\1/g, (match, quote, dependency) => {
      const dependencyOutput = packages.get(dependency);
      if (!dependencyOutput) throw new Error(`Unknown internal type dependency: @pptx/${dependency}`);
      let target = relative(dirname(outputPath), join(typesRoot, dependencyOutput, 'index.js')).split(sep).join('/');
      if (!target.startsWith('.')) target = `./${target}`;
      return `${quote}${target}${quote}`;
    });
    await writeFile(outputPath, rewritten);
  }
}
