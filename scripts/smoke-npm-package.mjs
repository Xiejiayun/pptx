import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const tarball = resolve(process.argv[2] ?? '');
if (!tarball.endsWith('.tgz')) throw new Error('Usage: node scripts/smoke-npm-package.mjs <package.tgz>');
const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const directory = await mkdtemp(join(tmpdir(), 'jiayunxie-pptx-smoke-'));
try {
  await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  run('npm', ['install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false', tarball], directory);

  const installed = join(directory, 'node_modules', '@jiayunxie', 'pptx');
  const manifest = JSON.parse(await readFile(join(installed, 'package.json'), 'utf8'));
  if (manifest.name !== '@jiayunxie/pptx' || manifest.version !== '0.1.0') {
    throw new Error(`Unexpected package identity: ${manifest.name}@${manifest.version}`);
  }
  if (JSON.stringify(manifest).includes('workspace:')) throw new Error('Packed manifest contains workspace protocol');
  if (manifest.exports?.['.']?.browser !== './dist/browser.js') {
    throw new Error('Packed manifest is missing the browser conditional export');
  }
  for (const section of ['dependencies', 'optionalDependencies', 'peerDependencies']) {
    for (const name of Object.keys(manifest[section] ?? {})) {
      if (name.startsWith('@pptx/')) throw new Error(`Packed manifest contains internal runtime dependency: ${name}`);
    }
  }

  const distFiles = (await listFiles(join(installed, 'dist'))).filter((name) => /\.(?:js|d\.ts)$/.test(name));
  for (const file of distFiles) {
    const source = await readFile(file, 'utf8');
    if (source.includes('@pptx/')) throw new Error(`Bundled output contains internal import: ${file}`);
  }
  const browserSource = await readFile(join(installed, 'dist/browser.js'), 'utf8');
  if (/\b(?:from|import)\s*['"]node:/.test(browserSource)) {
    throw new Error('Browser bundle contains a static node: import');
  }

  await writeFile(
    join(directory, 'smoke.mjs'),
    `import { inches, PptxDocument, GradientCodec, importPptxGenJS, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const created = PptxDocument.create();
const createdText = created.addSlide().addText('Smoke\\n\\nParagraph', { align: 'center', bullet: true, level: 2, spacing: { before: 4, after: 6, line: { kind: 'exact', points: 20 } }, tabStops: [{ position: 1.25 }, { position: 2.5, alignment: 'right' }] });
createdText.text = 'Updated\\r\\nParagraph';
const richText = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: 1, spacing: { line: { kind: 'multiple', factor: 1.5 } }, tabStops: [{ position: 1.5, alignment: 'center' }], runs: [{ text: 'Bold', style: { bold: true, fontSize: 18, color: { kind: 'srgb', value: 'ff0000' }, glow: { color: { kind: 'srgb', value: '00ff00' }, opacity: 0.5, size: 8 }, highlight: { kind: 'srgb', value: 'ffff00' }, outline: { color: { kind: 'srgb', value: '0000ff' }, size: 1.5 }, underline: true, strike: true } }, { text: 'Blue', softBreakBefore: true, style: { color: { kind: 'scheme', value: 'accent1' }, glow: { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2.5 }, highlight: { kind: 'scheme', value: 'accent2' }, outline: { color: { kind: 'scheme', value: 'accent3' }, size: 2 }, underline: { style: 'dbl', color: { kind: 'srgb', value: '00ff00' } }, strike: 'dblStrike' } }] }]);
richText.richText = [{ align: 'justify', bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 }, level: 3, spacing: { before: 5, after: 7, line: { kind: 'exact', points: 22 } }, tabStops: [{ position: 2.75, alignment: 'decimal' }], runs: [{ text: 'Updated rich', style: { baseline: 'superscript', characterSpacing: 2.5, italic: true, glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 }, highlight: { kind: 'srgb', value: '00ff00' }, outline: { color: { kind: 'scheme', value: 'accent1' }, size: 0.75 }, underline: { style: 'wavyHeavy', color: { kind: 'scheme', value: 'accent2' } }, strike: false } }] }];
const custom = PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
custom.slideSize = { width: inches(10), height: inches(7.5) };
const customXml = new TextDecoder().decode(custom.opcPackage.requirePart('/ppt/presentation.xml').bytes);
const checks = {
  PptxDocument: typeof PptxDocument === 'function',
  createText: createdText.text === 'Updated\\nParagraph' && createdText.richText.every(({ align, bullet, level, spacing, tabStops }) => align === 'center' && bullet?.kind === 'bullet' && bullet.indent === 27 && level === 2 && spacing?.line?.kind === 'exact' && Array.isArray(tabStops) && tabStops[0]?.position === 1.25 && tabStops[1]?.alignment === 'right') && created.slides[0].shapes[0] === createdText,
  richText: richText.text === 'Updated rich' && richText.richText[0].align === 'justify' && richText.richText[0].bullet.style === 'romanUcPeriod' && richText.richText[0].level === 3 && richText.richText[0].spacing.line.kind === 'exact' && Array.isArray(richText.richText[0].tabStops) && richText.richText[0].tabStops[0].alignment === 'decimal' && richText.richText[0].runs[0].style.baseline === 'superscript' && richText.richText[0].runs[0].style.characterSpacing === 2.5 && richText.richText[0].runs[0].style.italic === true && richText.richText[0].runs[0].style.glow.color.value === 'accent3' && richText.richText[0].runs[0].style.glow.opacity === 0.25 && richText.richText[0].runs[0].style.glow.size === 6 && richText.richText[0].runs[0].style.highlight.value === '00FF00' && richText.richText[0].runs[0].style.outline.color.value === 'accent1' && richText.richText[0].runs[0].style.outline.size === 0.75 && richText.richText[0].runs[0].style.underline.style === 'wavyHeavy' && richText.richText[0].runs[0].style.underline.color.value === 'accent2' && richText.richText[0].runs[0].style.strike === false,
  customSlideSize: custom.slideSize.width === inches(10) && customXml.includes('<p:sldSz cx="9144000" cy="6858000"/>'),
  GradientCodec: typeof GradientCodec === 'function',
  importPptxGenJS: typeof importPptxGenJS === 'function',
  transitions: typeof transitions.TransitionCodec === 'function',
  animations: typeof animations.AnimationTimingCodec === 'function',
  advancedCharts: typeof advancedCharts.AdvancedChartCodec === 'function',
  smartArt: typeof smartArt.SmartArtDiagramCodec === 'function',
};
if (Object.values(checks).some((value) => !value)) throw new Error(JSON.stringify(checks));
process.stdout.write(JSON.stringify(checks));
`,
  );
  const apiResult = run(process.execPath, ['smoke.mjs'], directory);
  const apiChecks = JSON.parse(apiResult.stdout);

  await writeFile(
    join(directory, 'browser-smoke.mjs'),
    `import { inches, PptxDocument, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const resolved = import.meta.resolve('@jiayunxie/pptx');
if (!resolved.endsWith('/dist/browser.js')) throw new Error('Browser condition resolved to ' + resolved);
const checks = [PptxDocument, transitions.TransitionCodec, animations.AnimationTimingCodec, advancedCharts.AdvancedChartCodec, smartArt.SmartArtDiagramCodec];
if (checks.some((value) => typeof value !== 'function')) throw new Error('Browser API surface is incomplete');
const created = PptxDocument.create({ slideSize: '16:9' });
if (created.addSlide().addText('Browser\\nText', { align: 'center', bullet: true, level: 2, spacing: { line: { kind: 'multiple', factor: 1.25 } }, tabStops: [{ position: 1.25 }] }).richText[0].tabStops[0].position !== 1.25) throw new Error('Browser create-text API failed');
const browserRich = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'number', style: 'alphaUcPeriod' }, level: 3, spacing: { before: 4, after: 6 }, tabStops: [{ position: 2.5, alignment: 'decimal' }], runs: [{ text: 'Rich', style: { baseline: 'subscript', characterSpacing: 0, bold: true, glow: { opacity: 0.75, size: 4 }, highlight: { kind: 'scheme', value: 'accent1' }, outline: { color: { kind: 'srgb', value: 'ff0000' }, size: 1.25 }, underline: { style: 'wavyDbl' }, strike: 'dblStrike' } }] }]).richText[0];
if (browserRich.tabStops[0].alignment !== 'decimal' || browserRich.runs[0].style.baseline !== 'subscript' || browserRich.runs[0].style.characterSpacing !== 0 || browserRich.runs[0].style.glow.color.value !== 'FFFFFF' || browserRich.runs[0].style.glow.opacity !== 0.75 || browserRich.runs[0].style.glow.size !== 4 || browserRich.runs[0].style.highlight.value !== 'accent1' || browserRich.runs[0].style.outline.color.value !== 'FF0000' || browserRich.runs[0].style.outline.size !== 1.25 || browserRich.runs[0].style.underline.style !== 'wavyDbl' || browserRich.runs[0].style.strike !== 'dblStrike') throw new Error('Browser rich-text API failed');
PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
created.slideSize = { width: inches(10), height: inches(7.5) };
process.stdout.write(resolved);
`,
  );
  run(process.execPath, ['--conditions=browser', 'browser-smoke.mjs'], directory);

  await writeFile(
    join(directory, 'smoke.ts'),
    `import {
  PptxDocument,
  inches,
  type CustomSlideSize,
  type RichTextParagraph,
  type TextAlignment,
  type NumberingStyle,
  type ParagraphBullet,
  type ParagraphLineSpacing,
  type ParagraphSpacing,
  type ParagraphTabStop,
  type ParagraphTabStopAlignment,
  type RichTextUnderline,
  type RichTextUnderlineStyle,
  type RichTextStrikeStyle,
  type RichTextOutline,
  type RichTextGlow,
  type RichTextBaseline,
  type RichTextRunStyle,
  GradientCodec,
  importPptxGenJS,
  transitions,
  animations,
  advancedCharts,
  smartArt,
} from '@jiayunxie/pptx';

const documentPromise: Promise<PptxDocument> = PptxDocument.open(new Uint8Array());
const createdDocument: PptxDocument = PptxDocument.create({ format: 'pptx', slideSize: 'wide' });
const customSlideSize: CustomSlideSize = { width: inches(11.7), height: inches(8.3) };
const customDocument: PptxDocument = PptxDocument.create({ slideSize: customSlideSize });
customDocument.slideSize = { width: inches(10), height: inches(7.5) };
const alignment: TextAlignment = 'center';
const bullet: ParagraphBullet = { kind: 'bullet', character: '◆', indent: 18 };
const numbering: NumberingStyle = 'romanUcPeriod';
const lineSpacing: ParagraphLineSpacing = { kind: 'multiple', factor: 1.5 };
const spacing: ParagraphSpacing = { before: 4, after: 6, line: lineSpacing };
const tabAlignment: ParagraphTabStopAlignment = 'right';
const tabStops: readonly ParagraphTabStop[] = [{ position: 1.25 }, { position: 2.5, alignment: tabAlignment }];
const underlineStyle: RichTextUnderlineStyle = 'dotDashHeavy';
const underline: RichTextUnderline = { style: underlineStyle, color: { kind: 'srgb', value: 'FF0000' } };
const strike: RichTextStrikeStyle = 'dblStrike';
const outline: RichTextOutline = { color: { kind: 'scheme', value: 'accent1' }, size: 1.5 };
const glow: RichTextGlow = { color: { kind: 'scheme', value: 'accent2' }, opacity: 0.5, size: 8 };
const baseline: RichTextBaseline = 'superscript';
const characterStyle: RichTextRunStyle = { baseline, characterSpacing: 2.5 };
const createdText = createdDocument.addSlide().addText('Typed\\ntext', { align: alignment, bullet, level: 2, spacing, tabStops });
createdText.text = 'Updated\\n\\ntyped text';
const paragraphs: readonly RichTextParagraph[] = [{ align: 'justify', bullet: { kind: 'number', style: numbering, startAt: 3 }, level: 3, spacing: { line: { kind: 'exact', points: 20 } }, tabStops, runs: [{ text: 'Typed rich', style: { ...characterStyle, fontSize: 12.5, bold: true, color: { kind: 'scheme', value: 'tx1' }, glow, highlight: { kind: 'srgb', value: 'FFFF00' }, outline, underline, strike } }] }];
const richText = createdDocument.slides[0].addRichText(paragraphs);
richText.richText = paragraphs;
const gradientConstructor: typeof GradientCodec = GradientCodec;
const adapter: typeof importPptxGenJS = importPptxGenJS;
const transition: transitions.SlideTransition = { effect: 'fade' };
const animationConstructor: typeof animations.AnimationTimingCodec = animations.AnimationTimingCodec;
const chartConstructor: typeof advancedCharts.AdvancedChartCodec = advancedCharts.AdvancedChartCodec;
const smartArtConstructor: typeof smartArt.SmartArtDiagramCodec = smartArt.SmartArtDiagramCodec;
documentPromise.then((document) => {
  transitions.installTransitionPlugin(document);
  animations.installAnimationPlugin(document);
  advancedCharts.installAdvancedChartPlugin(document);
  smartArt.installSmartArtPlugin(document);
});
void [documentPromise, createdDocument, customDocument, createdText, richText, gradientConstructor, adapter, transition, animationConstructor, chartConstructor, smartArtConstructor];
`,
  );
  run(
    process.execPath,
    [
      join(repositoryRoot, 'node_modules/typescript/bin/tsc'),
      'smoke.ts',
      '--noEmit',
      '--strict',
      '--target',
      'ES2022',
      '--module',
      'NodeNext',
      '--moduleResolution',
      'NodeNext',
      '--typeRoots',
      join(repositoryRoot, 'node_modules/@types'),
      '--types',
      'node',
    ],
    directory,
  );

  const bin = join(directory, 'node_modules', '.bin', process.platform === 'win32' ? 'pptx-inspect.cmd' : 'pptx-inspect');
  const cliResult = run(bin, ['--json', 'doctor'], directory);
  const doctor = JSON.parse(cliResult.stdout);
  if (!doctor.ok || doctor.data?.version !== '0.1.0') throw new Error(`CLI smoke failed: ${cliResult.stdout}`);

  process.stdout.write(
    `${JSON.stringify({ ok: true, tarball: basename(tarball), api: apiChecks, types: true, cli: doctor.data.version })}\n`,
  );
} finally {
  await rm(directory, { recursive: true, force: true });
}

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8' });
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} failed (${result.status})\n${result.stdout}\n${result.stderr}`);
  }
  return result;
}

async function listFiles(directory) {
  const files = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await listFiles(path)));
    else files.push(path);
  }
  return files;
}
