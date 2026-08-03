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

const shapeDeclaration = join(typesRoot, 'model/shapes.d.ts');
const shapeDeclarationSource = await readFile(shapeDeclaration, 'utf8');
const slideDeclaration = join(typesRoot, 'model/slide.d.ts');
const slideDeclarationSource = await readFile(slideDeclaration, 'utf8');
const addTableCellOptionsStart = slideDeclarationSource.indexOf(
  'export interface AddTableCellOptions',
);
const addTableCellStart = slideDeclarationSource.indexOf(
  'export interface AddTableCell',
  addTableCellOptionsStart + 1,
);
const addTableCellOptionsDeclaration = addTableCellOptionsStart >= 0 &&
    addTableCellStart > addTableCellOptionsStart
  ? slideDeclarationSource.slice(addTableCellOptionsStart, addTableCellStart)
  : '';
if (!addTableCellOptionsDeclaration.includes('readonly hyperlink?: Hyperlink;')) {
  throw new Error('Packed AddTableCellOptions declaration is missing table-cell hyperlink');
}
const tableCellStart = shapeDeclarationSource.indexOf('export interface TableCell {');
const tableRowStart = shapeDeclarationSource.indexOf(
  'export interface TableRow',
  tableCellStart,
);
const tableCellDeclaration = tableCellStart >= 0 && tableRowStart > tableCellStart
  ? shapeDeclarationSource.slice(tableCellStart, tableRowStart)
  : '';
if (!tableCellDeclaration.includes('readonly hyperlink?: Hyperlink;')) {
  throw new Error('Packed TableCell declaration is missing table-cell hyperlink snapshot');
}
const tableModelStart = shapeDeclarationSource.indexOf('export declare class TableModel');
const chartModelStart = shapeDeclarationSource.indexOf(
  'export declare class ChartModel',
  tableModelStart,
);
const tableModelDeclaration = tableModelStart >= 0 && chartModelStart > tableModelStart
  ? shapeDeclarationSource.slice(tableModelStart, chartModelStart)
  : '';
if (!tableModelDeclaration.includes(
  'get verticalAlignment(): TextBoxVerticalAlignment | undefined;',
) || !tableModelDeclaration.includes(
  'set verticalAlignment(value: TextBoxVerticalAlignment | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level vertical alignment');
}
if (!tableModelDeclaration.includes(
  'get textDirection(): TableCellTextDirection | undefined;',
) || !tableModelDeclaration.includes(
  'set textDirection(value: TableCellTextDirection | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level text direction');
}
if (!tableModelDeclaration.includes(
  'get horizontalAlignment(): TextAlignment | undefined;',
) || !tableModelDeclaration.includes(
  'set horizontalAlignment(value: TextAlignment | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level horizontal alignment');
}
if (!tableModelDeclaration.includes(
  'get margins(): TextBoxMargins | undefined;',
) || !tableModelDeclaration.includes(
  'set margins(value: TextBoxMarginInput | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level margins');
}
if (!tableModelDeclaration.includes(
  'get fill(): TableCellFill | undefined;',
) || !tableModelDeclaration.includes(
  'set fill(value: TableCellFill | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level fill');
}
if (!tableModelDeclaration.includes(
  'get borders(): TableCellBorders | undefined;',
) || !tableModelDeclaration.includes(
  'set borders(value: TableCellBorderInput | undefined);',
)) {
  throw new Error('Packed TableModel declaration is missing table-level borders');
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
