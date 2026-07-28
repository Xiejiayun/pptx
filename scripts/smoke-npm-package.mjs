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
const created = PptxDocument.create({ rtlMode: true });
const createdText = created.addSlide().addText('Smoke\\n\\nParagraph', { align: 'center', fit: 'shrink', valign: 'top', vert: 'vert270', wrap: false, bullet: true, level: 2, margin: 10, rtlMode: true, spacing: { before: 4, after: 6, line: { kind: 'exact', points: 20 } }, tabStops: [{ position: 1.25 }, { position: 2.5, alignment: 'right' }] });
const initialTextWrap = createdText.textWrap;
const initialTextDirection = createdText.textDirection;
const initialTextFit = createdText.textFit;
createdText.text = 'Updated\\r\\nParagraph';
createdText.textMargins = { top: 4, left: 8 };
createdText.verticalAlignment = 'bottom';
createdText.textWrap = true;
const updatedTextWrap = createdText.textWrap;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
const updatedTextDirection = createdText.textDirection;
createdText.textDirection = undefined;
createdText.textFit = 'resize';
const updatedTextFit = createdText.textFit;
createdText.textFit = 'none';
createdText.textFit = undefined;
const richText = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: 1, spacing: { line: { kind: 'multiple', factor: 1.5 } }, tabStops: [{ position: 1.5, alignment: 'center' }], runs: [{ text: 'Bold', style: { bold: true, fontSize: 18, color: { kind: 'srgb', value: 'ff0000' }, glow: { color: { kind: 'srgb', value: '00ff00' }, opacity: 0.5, size: 8 }, highlight: { kind: 'srgb', value: 'ffff00' }, outline: { color: { kind: 'srgb', value: '0000ff' }, size: 1.5 }, underline: true, strike: true } }, { text: 'Blue', softBreakBefore: true, style: { lang: 'de-DE', color: { kind: 'scheme', value: 'accent1' }, glow: { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2.5 }, highlight: { kind: 'scheme', value: 'accent2' }, outline: { color: { kind: 'scheme', value: 'accent3' }, size: 2 }, underline: { style: 'dbl', color: { kind: 'srgb', value: '00ff00' } }, strike: 'dblStrike' } }] }, { rtl: false, runs: [{ text: 'LTR override' }] }], { lang: 'fr-CA', rtlMode: true });
const marginText = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, marginLeft: false, runs: [{ text: 'Bullet' }] }, { marginLeft: false, runs: [{ text: 'Absent' }] }], { paragraphMarginLeft: 24 });
const initialParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const bulletMarginIsolation = marginText.richText[1].marginLeft === undefined && marginText.richText[1].bullet.indent === 27;
marginText.richText = [{ marginLeft: 6, runs: [{ text: 'Six' }] }, { marginLeft: 0, runs: [{ text: 'Zero' }] }, { marginLeft: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }];
const updatedParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const rightMarginText = created.slides[0].addRichText([{ marginRight: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, runs: [{ text: 'Bullet' }] }, { marginRight: false, runs: [{ text: 'Absent' }] }], { paragraphMarginRight: 24 });
const initialParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const bulletRightMarginCoexistence = rightMarginText.richText[1].marginRight === 24 && rightMarginText.richText[1].bullet.indent === 27;
rightMarginText.richText = [{ marginRight: 6, runs: [{ text: 'Six' }] }, { marginRight: 0, runs: [{ text: 'Zero' }] }, { marginRight: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }, { bullet: true, marginRight: 9, runs: [{ text: 'Bullet' }] }];
const updatedParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const inheritedLanguage = richText.richText[0].runs[0].style.lang;
const localLanguage = richText.richText[0].runs[1].style.lang;
const initialRtl = richText.richText.map(({ rtl }) => rtl);
const presentationRtlEnabled = created.rtlMode;
created.rtlMode = false;
const presentationRtlDisabled = created.rtlMode;
created.rtlMode = undefined;
const presentationRtlCleared = created.rtlMode;
const paragraphRtlAfterGlobalClear = richText.richText.map(({ rtl }) => rtl);
richText.richText = [{ align: 'justify', bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 }, level: 3, spacing: { before: 5, after: 7, line: { kind: 'exact', points: 22 } }, tabStops: [{ position: 2.75, alignment: 'decimal' }], runs: [{ text: 'Updated rich', style: { lang: 'ja-JP', baseline: 'superscript', characterSpacing: 2.5, italic: true, glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 }, highlight: { kind: 'srgb', value: '00ff00' }, outline: { color: { kind: 'scheme', value: 'accent1' }, size: 0.75 }, underline: { style: 'wavyHeavy', color: { kind: 'scheme', value: 'accent2' } }, strike: false } }] }];
const custom = PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
custom.slideSize = { width: inches(10), height: inches(7.5) };
const customXml = new TextDecoder().decode(custom.opcPackage.requirePart('/ppt/presentation.xml').bytes);
const checks = {
  PptxDocument: typeof PptxDocument === 'function',
  presentationRtl: presentationRtlEnabled === true && presentationRtlDisabled === false && presentationRtlCleared === undefined && paragraphRtlAfterGlobalClear[0] === true && paragraphRtlAfterGlobalClear[1] === false,
  paragraphMarginLeft: initialParagraphMargins[0] === 12 && initialParagraphMargins[1] === undefined && initialParagraphMargins[2] === undefined && bulletMarginIsolation && updatedParagraphMargins[0] === 6 && updatedParagraphMargins[1] === 0 && updatedParagraphMargins[2] === undefined && updatedParagraphMargins[3] === undefined,
  paragraphMarginRight: initialParagraphRightMargins[0] === 12 && initialParagraphRightMargins[1] === 24 && initialParagraphRightMargins[2] === undefined && bulletRightMarginCoexistence && updatedParagraphRightMargins[0] === 6 && updatedParagraphRightMargins[1] === 0 && updatedParagraphRightMargins[2] === undefined && updatedParagraphRightMargins[3] === undefined && updatedParagraphRightMargins[4] === 9,
  createText: createdText.text === 'Updated\\nParagraph' && initialTextWrap === false && updatedTextWrap === true && createdText.textWrap === undefined && initialTextDirection === 'vert270' && updatedTextDirection === 'wordArtVert' && createdText.textDirection === undefined && initialTextFit === 'shrink' && updatedTextFit === 'resize' && createdText.textFit === undefined && createdText.verticalAlignment === 'bottom' && createdText.textMargins.top === 4 && createdText.textMargins.left === 8 && createdText.textMargins.right === undefined && createdText.richText.every(({ align, bullet, level, rtl, spacing, tabStops }) => align === 'center' && bullet?.kind === 'bullet' && bullet.indent === 27 && level === 2 && rtl === true && spacing?.line?.kind === 'exact' && Array.isArray(tabStops) && tabStops[0]?.position === 1.25 && tabStops[1]?.alignment === 'right') && created.slides[0].shapes[0] === createdText,
  richText: inheritedLanguage === 'fr-CA' && localLanguage === 'de-DE' && initialRtl[0] === true && initialRtl[1] === false && richText.text === 'Updated rich' && richText.richText[0].rtl === undefined && richText.richText[0].align === 'justify' && richText.richText[0].bullet.style === 'romanUcPeriod' && richText.richText[0].level === 3 && richText.richText[0].spacing.line.kind === 'exact' && Array.isArray(richText.richText[0].tabStops) && richText.richText[0].tabStops[0].alignment === 'decimal' && richText.richText[0].runs[0].style.lang === 'ja-JP' && richText.richText[0].runs[0].style.baseline === 'superscript' && richText.richText[0].runs[0].style.characterSpacing === 2.5 && richText.richText[0].runs[0].style.italic === true && richText.richText[0].runs[0].style.glow.color.value === 'accent3' && richText.richText[0].runs[0].style.glow.opacity === 0.25 && richText.richText[0].runs[0].style.glow.size === 6 && richText.richText[0].runs[0].style.highlight.value === '00FF00' && richText.richText[0].runs[0].style.outline.color.value === 'accent1' && richText.richText[0].runs[0].style.outline.size === 0.75 && richText.richText[0].runs[0].style.underline.style === 'wavyHeavy' && richText.richText[0].runs[0].style.underline.color.value === 'accent2' && richText.richText[0].runs[0].style.strike === false,
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
const created = PptxDocument.create({ rtlMode: true, slideSize: '16:9' });
const browserText = created.addSlide().addText('Browser\\nText', { align: 'center', fit: 'resize', valign: 'bottom', vert: 'vert', wrap: false, bullet: true, level: 2, margin: [0, 0, 0, 0], rtlMode: true, spacing: { line: { kind: 'multiple', factor: 1.25 } }, tabStops: [{ position: 1.25 }] });
if (browserText.textWrap !== false || browserText.verticalAlignment !== 'bottom' || browserText.textDirection !== 'vert' || browserText.textFit !== 'resize' || browserText.richText.some(({ rtl }) => rtl !== true) || browserText.richText[0].tabStops[0].position !== 1.25 || browserText.textMargins.top !== 0 || browserText.textMargins.right !== 0 || browserText.textMargins.bottom !== 0 || browserText.textMargins.left !== 0) throw new Error('Browser create-text API failed');
const browserRich = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'number', style: 'alphaUcPeriod' }, level: 3, spacing: { before: 4, after: 6 }, tabStops: [{ position: 2.5, alignment: 'decimal' }], runs: [{ text: 'Rich', style: { lang: 'ja-JP', baseline: 'subscript', characterSpacing: 0, bold: true, glow: { opacity: 0.75, size: 4 }, highlight: { kind: 'scheme', value: 'accent1' }, outline: { color: { kind: 'srgb', value: 'ff0000' }, size: 1.25 }, underline: { style: 'wavyDbl' }, strike: 'dblStrike' } }] }], { rtlMode: true }).richText[0];
if (browserRich.rtl !== true || browserRich.tabStops[0].alignment !== 'decimal' || browserRich.runs[0].style.lang !== 'ja-JP' || browserRich.runs[0].style.baseline !== 'subscript' || browserRich.runs[0].style.characterSpacing !== 0 || browserRich.runs[0].style.glow.color.value !== 'FFFFFF' || browserRich.runs[0].style.glow.opacity !== 0.75 || browserRich.runs[0].style.glow.size !== 4 || browserRich.runs[0].style.highlight.value !== 'accent1' || browserRich.runs[0].style.outline.color.value !== 'FF0000' || browserRich.runs[0].style.outline.size !== 1.25 || browserRich.runs[0].style.underline.style !== 'wavyDbl' || browserRich.runs[0].style.strike !== 'dblStrike') throw new Error('Browser rich-text API failed');
const browserMargin = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginLeft: 24 });
if (browserMargin.richText[0].marginLeft !== 12) throw new Error('Browser paragraph margin create failed');
browserMargin.richText = [{ marginLeft: false, runs: [{ text: 'Cleared' }] }];
if (browserMargin.richText[0].marginLeft !== undefined) throw new Error('Browser paragraph margin clear failed');
const browserRightMargin = created.slides[0].addRichText([{ bullet: true, marginRight: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginRight: 24 });
if (browserRightMargin.richText[0].marginRight !== 12 || browserRightMargin.richText[0].bullet.indent !== 27) throw new Error('Browser paragraph right margin create failed');
browserRightMargin.richText = [{ marginRight: false, runs: [{ text: 'Cleared' }] }];
if (browserRightMargin.richText[0].marginRight !== undefined) throw new Error('Browser paragraph right margin clear failed');
if (created.rtlMode !== true) throw new Error('Browser presentation RTL create failed');
created.rtlMode = false;
if (created.rtlMode !== false || browserRich.rtl !== true) throw new Error('Browser presentation RTL edit failed');
created.rtlMode = undefined;
if (created.rtlMode !== undefined || browserRich.rtl !== true) throw new Error('Browser presentation RTL clear failed');
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
  type TextBoxMarginInput,
  type TextBoxMargins,
  type TextBoxFit,
  type TextBoxTextDirection,
  type TextBoxVerticalAlignment,
  GradientCodec,
  importPptxGenJS,
  transitions,
  animations,
  advancedCharts,
  smartArt,
} from '@jiayunxie/pptx';

const documentPromise: Promise<PptxDocument> = PptxDocument.open(new Uint8Array());
const createdDocument: PptxDocument = PptxDocument.create({ format: 'pptx', slideSize: 'wide' });
const globalRtl: PptxDocument = PptxDocument.create({ rtlMode: true });
const globalRtlSnapshot: boolean | undefined = globalRtl.rtlMode;
globalRtl.rtlMode = false;
globalRtl.rtlMode = undefined;
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
const characterStyle: RichTextRunStyle = { baseline, characterSpacing: 2.5, lang: 'de-DE' };
const margin: TextBoxMarginInput = [4, 8, 4, 8];
const verticalAlignment: TextBoxVerticalAlignment = 'top';
const direction: TextBoxTextDirection = 'vert270';
const fit: TextBoxFit = 'shrink';
const createdText = createdDocument.addSlide().addText('Typed\\ntext', { align: alignment, fit, valign: verticalAlignment, vert: direction, wrap: true, bullet, level: 2, margin, spacing, tabStops });
const marginSnapshot: TextBoxMargins | undefined = createdText.textMargins;
const wrapSnapshot: boolean | undefined = createdText.textWrap;
const directionSnapshot: TextBoxTextDirection | undefined = createdText.textDirection;
const fitSnapshot: TextBoxFit | undefined = createdText.textFit;
createdText.textMargins = { top: 3, left: 6 };
createdText.verticalAlignment = 'bottom';
createdText.verticalAlignment = undefined;
createdText.textWrap = false;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
createdText.textDirection = undefined;
createdText.textFit = 'resize';
createdText.textFit = 'none';
createdText.textFit = undefined;
createdText.text = 'Updated\\n\\ntyped text';
const paragraphs: readonly RichTextParagraph[] = [{ align: 'justify', bullet: { kind: 'number', style: numbering, startAt: 3 }, level: 3, spacing: { line: { kind: 'exact', points: 20 } }, tabStops, runs: [{ text: 'Typed rich', style: { ...characterStyle, fontSize: 12.5, bold: true, color: { kind: 'scheme', value: 'tx1' }, glow, highlight: { kind: 'srgb', value: 'FFFF00' }, outline, underline, strike } }] }];
const richText = createdDocument.slides[0].addRichText(paragraphs, { lang: 'fr-CA' });
richText.richText = paragraphs;
const rtlParagraphs: readonly RichTextParagraph[] = [
  { rtl: true, runs: [{ text: 'RTL' }] },
  { rtl: false, runs: [{ text: 'LTR' }] },
];
createdDocument.addSlide().addRichText(rtlParagraphs, { rtlMode: true });
const paragraphMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginLeft: 12, runs: [{ text: 'Override' }] },
  { marginLeft: false, runs: [{ text: 'Suppressed' }] },
];
createdDocument.addSlide().addRichText(paragraphMargins, { paragraphMarginLeft: 24 });
const paragraphRightMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginRight: 12, runs: [{ text: 'Override' }] },
  { marginRight: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, marginRight: 18, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphRightMargins, { paragraphMarginRight: 24 });
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
void [documentPromise, createdDocument, globalRtl, globalRtlSnapshot, customDocument, createdText, marginSnapshot, wrapSnapshot, directionSnapshot, fitSnapshot, fit, direction, verticalAlignment, richText, rtlParagraphs, paragraphMargins, paragraphRightMargins, gradientConstructor, adapter, transition, animationConstructor, chartConstructor, smartArtConstructor];
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
