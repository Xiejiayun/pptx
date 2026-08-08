#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pptxModule = pathToFileURL(path.join(repoRoot, 'packages', 'pptx', 'dist', 'index.js')).href;
const {
  assertDeckSpec,
  connectorTransform,
  degrees,
  inches,
  PptxDocument,
} = await import(pptxModule);

const DEFAULT_THEME = Object.freeze({
  deep: '0B2E22', primary: '123D2A', secondary: '3D6B45', accent: 'E0A83E',
  background: 'F3F0E4', surface: 'DDE8D3', text: '173126', mutedText: '607267',
  contrast: 'FFFFFF', danger: 'B94732', cool: '4A91A8', font: 'Arial',
});

const solid = (value, transparency) => ({
  kind: 'solid', color: { kind: 'srgb', value },
  ...(transparency === undefined ? {} : { transparency }),
});
const stroke = (value, width = 1, transparency) => ({
  kind: 'line', color: { kind: 'srgb', value }, width,
  ...(transparency === undefined ? {} : { transparency }),
});

function text(slide, theme, value, x, y, width, height, fontSize, color, options = {}) {
  return slide.addRichText([{ runs: [{
    text: String(value),
    style: {
      fontFamily: theme.font,
      fontSize,
      color: { kind: 'srgb', value: color },
      bold: options.bold,
      italic: options.italic,
    },
  }], align: options.align ?? 'left' }], {
    x: inches(x), y: inches(y), width: inches(width), height: inches(height),
    margin: 0, wrap: options.wrap ?? true, fit: options.fit ?? 'shrink',
    valign: options.valign ?? 'top', align: options.align ?? 'left',
  });
}

function title(slide, theme, value, dark = false) {
  text(slide, theme, value, 0.65, 0.45, 12, 0.7, 38, dark ? theme.contrast : theme.text, {
    bold: true, wrap: false,
  });
}

function kicker(slide, theme, value, color = theme.secondary) {
  text(slide, theme, String(value).toUpperCase(), 0.67, 1.16, 7.5, 0.24, 10.5, color, {
    bold: true, wrap: false,
  });
}

function slideNumber(slide, theme, number, dark = false) {
  text(slide, theme, String(number).padStart(2, '0'), 12.15, 7.02, 0.5, 0.18, 10,
    dark ? theme.surface : theme.mutedText, { align: 'right', wrap: false });
}

function leaf(slide, theme, x, y, width, height, color, rotation = 0, transparency = 0) {
  slide.addShape('ellipse', {
    x: inches(x), y: inches(y), width: inches(width), height: inches(height),
    rotation: degrees(rotation), fill: solid(color, transparency), line: { kind: 'none' },
  });
  const transform = connectorTransform(
    { x: inches(x + width * 0.22), y: inches(y + height * 0.5) },
    { x: inches(x + width * 0.78), y: inches(y + height * 0.5) },
  );
  slide.addShape('line', {
    ...transform,
    line: stroke(theme.deep, 0.8, 55),
  });
}

function notes(slide, sources = []) {
  slide.addNotes(`[Sources]\n${sources.join('\n')}`);
}

function buildDeckSpec(content) {
  return {
    schemaVersion: 1,
    slideSize: { width: inches(13.333), height: inches(7.5) },
    safeArea: { top: inches(0.4), right: inches(0.4), bottom: inches(0.4), left: inches(0.4) },
    gap: inches(0.3),
    fontSafetyFactor: 1.1,
    slides: content.slides.map((slide, index) => {
      const cover = slide.family === 'cover';
      const bounds = cover
        ? { x: inches(0.72), y: inches(3.85), width: inches(11.8), height: inches(0.9) }
        : { x: inches(0.65), y: inches(0.45), width: inches(12), height: inches(0.7) };
      return {
        id: slide.id ?? `slide-${index + 1}`,
        family: slide.family,
        regions: [{ id: 'title-region', bounds, collision: 'exclusive' }],
        elements: [{
          kind: 'text', id: 'title', regionId: 'title-region', bounds,
          role: 'title', text: slide.title, fontFamily: content.theme?.font ?? DEFAULT_THEME.font,
          fontSize: cover ? 54 : 38, bold: true, wrap: false, maxLines: 1,
          fit: 'error', minFontSize: cover ? 50 : 36,
        }],
      };
    }),
  };
}

function addCover(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.deep);
  slide.addShape('ellipse', { x: inches(0.1), y: inches(0.1), width: inches(6.4), height: inches(3.2), fill: solid(theme.secondary, 25), line: { kind: 'none' } });
  slide.addShape('ellipse', { x: inches(6.7), y: inches(0.15), width: inches(6.4), height: inches(3.15), fill: solid(theme.primary, 35), line: { kind: 'none' } });
  for (let index = 0; index < 10; index += 1) {
    leaf(slide, theme, 0.45 + index * 1.2, 0.55 + (index % 3) * 0.32, 1.35, 0.5,
      index % 2 ? theme.surface : theme.secondary, index % 2 ? 25 : -25, 18 + index * 3);
  }
  text(slide, theme, spec.title, 0.72, 3.85, 11.8, 0.9, 54, theme.contrast, { bold: true, wrap: false });
  text(slide, theme, spec.subtitle ?? '', 0.76, 4.8, 9.9, 0.7, 19, theme.surface);
  kicker(slide, theme, spec.kicker ?? '', theme.accent);
  notes(slide, spec.sources);
  slideNumber(slide, theme, number, true);
}

function addBands(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'A layered system');
  const rows = spec.rows.slice(0, 4);
  const colors = [theme.surface, theme.secondary, theme.primary, theme.deep];
  rows.forEach((row, index) => {
    const y = 1.55 + index * 1.2;
    const dark = index > 0;
    slide.addShape('rect', { x: inches(0.65), y: inches(y), width: inches(12.03), height: inches(1.08), fill: solid(colors[index]), line: { kind: 'none' } });
    text(slide, theme, row.heading, 0.9, y + 0.18, 2.45, 0.34, 21, dark ? theme.contrast : theme.deep, { bold: true, wrap: false });
    text(slide, theme, row.body, 3.5, y + 0.16, 3.8, 0.7, 15, dark ? theme.contrast : theme.text, { bold: true });
    text(slide, theme, row.detail ?? '', 7.55, y + 0.2, 4.6, 0.5, 14, dark ? theme.surface : theme.text);
  });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addSpotlight(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'A keystone relationship');
  slide.addShape('ellipse', { x: inches(0.7), y: inches(1.55), width: inches(4.8), height: inches(4.8), fill: solid(theme.accent), line: { kind: 'none' } });
  text(slide, theme, spec.hero.heading, 1.1, 2.15, 4.0, 0.7, 40, theme.deep, { bold: true, align: 'center', wrap: false });
  text(slide, theme, spec.hero.subheading ?? '', 1.25, 2.95, 3.7, 0.4, 19, theme.deep, { bold: true, align: 'center' });
  text(slide, theme, spec.hero.body, 1.25, 3.55, 3.7, 1.2, 16, theme.text, { align: 'center', valign: 'middle' });
  spec.items.slice(0, 3).forEach((item, index) => {
    const y = 1.65 + index * 1.55;
    text(slide, theme, item.heading, 6.05, y, 5.8, 0.4, 23, index === 0 ? theme.cool : theme.secondary, { bold: true, wrap: false });
    text(slide, theme, item.body, 6.05, y + 0.52, 5.8, 0.7, 16, theme.text);
    if (index < 2) slide.addShape('line', { x: inches(6.05), y: inches(y + 1.35), width: inches(5.6), height: 1, line: stroke(theme.secondary, 1, 55) });
  });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addRoles(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.deep);
  title(slide, theme, spec.title, true);
  kicker(slide, theme, spec.kicker ?? 'Invisible work', theme.accent);
  const positions = [[0.95, 1.9], [9.0, 1.9], [0.95, 4.85], [9.0, 4.85]];
  spec.items.slice(0, 4).forEach((item, index) => {
    const [x, y] = positions[index];
    text(slide, theme, item.heading, x, y, 3.35, 0.4, 21, [theme.accent, theme.cool, theme.danger, theme.surface][index], { bold: true, wrap: false });
    text(slide, theme, item.body, x, y + 0.5, 3.45, 0.88, 15, theme.surface);
  });
  slide.addShape('ellipse', { x: inches(5.55), y: inches(2.75), width: inches(2.2), height: inches(1.35), fill: solid(theme.accent), line: { kind: 'none' } });
  for (const [x, y, rotation] of [[4.9, 2.0, -25], [7.0, 2.0, 25], [4.9, 4.0, 25], [7.0, 4.0, -25]]) {
    slide.addShape('ellipse', { x: inches(x), y: inches(y), width: inches(1.55), height: inches(1.05), rotation: degrees(rotation), fill: solid(theme.secondary), line: { kind: 'none' } });
  }
  text(slide, theme, spec.footer ?? '', 4.2, 6.35, 5.0, 0.36, 16, theme.contrast, { bold: true, align: 'center' });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number, true);
}

function addBranches(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'A living architecture');
  slide.addShape('line', { x: inches(6.55), y: inches(1.55), width: 1, height: inches(4.85), line: stroke(theme.deep, 6) });
  for (let index = 0; index < 12; index += 1) {
    const side = index % 2 ? 1 : -1;
    leaf(slide, theme, 6.48 + side * (0.18 + (index % 3) * 0.35), 1.45 + index * 0.37,
      1.0, 0.34, index < 4 ? theme.surface : index < 8 ? theme.secondary : theme.primary,
      side > 0 ? 22 : -22, index * 2);
  }
  const positions = [[0.78, 1.75], [8.1, 2.1], [0.78, 4.7]];
  spec.items.slice(0, 3).forEach((item, index) => {
    const [x, y] = positions[index];
    text(slide, theme, item.heading, x, y, 4.3, 0.42, 24, [theme.primary, theme.secondary, theme.cool][index], { bold: true, wrap: false });
    text(slide, theme, item.body, x, y + 0.58, 4.45, 0.9, 16, theme.text);
  });
  text(slide, theme, spec.callout ?? '', 8.1, 5.1, 4.2, 0.8, 22, theme.danger, { bold: true });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addStats(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'Three signals');
  const colors = [theme.primary, theme.cool, theme.accent];
  spec.items.slice(0, 3).forEach((item, index) => {
    const x = 0.78 + index * 3.98;
    slide.addShape('ellipse', { x: inches(x), y: inches(1.65), width: inches(3.2), height: inches(3.2), fill: solid(colors[index]), line: { kind: 'none' } });
    text(slide, theme, item.value, x + 0.2, 2.15, 2.8, 0.7, 38, index === 2 ? theme.deep : theme.contrast, { bold: true, align: 'center', wrap: false });
    text(slide, theme, item.unit ?? '', x + 0.3, 2.95, 2.6, 0.36, 15, index === 2 ? theme.deep : theme.contrast, { bold: true, align: 'center' });
    text(slide, theme, item.heading, x - 0.05, 5.08, 3.4, 0.4, 20, colors[index], { bold: true, align: 'center', wrap: false });
    text(slide, theme, item.body, x - 0.05, 5.58, 3.4, 0.8, 14, theme.text, { align: 'center' });
  });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

async function addChart(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'Evidence over time');
  const chart = await slide.addChart('bar', [{
    name: spec.chart.name,
    categories: spec.chart.categories,
    values: spec.chart.values,
  }], { x: inches(0.72), y: inches(1.65), width: inches(8.3), height: inches(4.9) });
  await chart.replaceDefinition({
    groups: [{
      type: 'bar',
      series: [{ name: spec.chart.name, categories: spec.chart.categories, values: spec.chart.values }],
      options: { direction: 'column', gapWidth: 58, dataLabels: { showValue: true, position: 'outsideEnd', face: theme.font, size: 11, color: { kind: 'srgb', value: theme.text } }, series: [{ fill: solid(theme.secondary), line: { kind: 'none' } }] },
    }],
    options: {
      legend: { visible: false },
      chartArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
      plotArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
      categoryAxis: { face: theme.font, size: 11, color: { kind: 'srgb', value: theme.text }, line: stroke(theme.mutedText, 1, 55), majorTickMark: 'none' },
      valueAxis: { visible: false, majorGridLine: stroke(theme.mutedText, 1, 75) },
    },
  });
  slide.addShape('rect', { x: inches(9.45), y: inches(1.78), width: inches(3.08), height: inches(4.62), fill: solid(theme.deep), line: { kind: 'none' } });
  text(slide, theme, spec.callout.value, 9.78, 2.25, 2.45, 0.75, 42, theme.accent, { bold: true, align: 'center', wrap: false });
  text(slide, theme, spec.callout.heading, 9.8, 3.2, 2.4, 0.72, 16, theme.contrast, { bold: true, align: 'center' });
  text(slide, theme, spec.callout.body, 9.78, 4.55, 2.45, 1.0, 17, theme.surface, { align: 'center', valign: 'middle' });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addProcess(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'A chain reaction', theme.danger);
  const steps = spec.items.slice(0, 4);
  const positions = [0.82, 3.93, 7.04, 10.15];
  for (let index = 0; index < steps.length - 1; index += 1) {
    const transform = connectorTransform(
      { x: inches(positions[index] + 2.2), y: inches(3.28) },
      { x: inches(positions[index + 1] - 0.12), y: inches(3.28) },
    );
    slide.addShape('line', { ...transform, line: stroke(theme.danger, 3), arrows: { end: 'triangle' } });
  }
  steps.forEach((item, index) => {
    const x = positions[index];
    const color = [theme.accent, theme.danger, '9E3A2A', theme.deep][index];
    slide.addShape('roundRect', { x: inches(x), y: inches(2.08), width: inches(2.2), height: inches(2.35), fill: solid(color), line: { kind: 'none' } });
    text(slide, theme, item.heading, x + 0.2, 2.28, 1.8, 0.4, 21, index === 0 ? theme.deep : theme.contrast, { bold: true, wrap: false });
    text(slide, theme, item.body, x + 0.2, 2.9, 1.8, 1.05, 14, index === 0 ? theme.deep : theme.contrast);
  });
  text(slide, theme, spec.footer ?? '', 1.6, 5.4, 10.1, 0.8, 21, theme.primary, { bold: true, align: 'center' });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addActions(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.deep);
  title(slide, theme, spec.title, true);
  kicker(slide, theme, spec.kicker ?? 'What changes the trajectory', theme.accent);
  spec.items.slice(0, 3).forEach((item, index) => {
    const y = 1.7 + index * 1.48;
    text(slide, theme, String(index + 1).padStart(2, '0'), 0.85, y, 0.55, 0.4, 20, theme.accent, { bold: true });
    text(slide, theme, item.heading, 1.65, y, 4.5, 0.42, 21, theme.contrast, { bold: true, wrap: false });
    text(slide, theme, item.body, 6.2, y, 5.7, 0.72, 15, theme.surface);
    if (index < 2) slide.addShape('line', { x: inches(1.65), y: inches(y + 1.05), width: inches(10.2), height: 1, line: stroke(theme.secondary, 1, 55) });
  });
  for (let index = 0; index < 7; index += 1) leaf(slide, theme, 0.55 + index * 1.75, 6.35, 1.3, 0.42, index % 2 ? theme.secondary : theme.surface, index % 2 ? 20 : -20, 30);
  notes(slide, spec.sources);
  slideNumber(slide, theme, number, true);
}

export async function createFastPresentation(content, output, deckSpecOutput) {
  if (!Array.isArray(content.slides) || content.slides.length < 1) throw new Error('Content spec requires slides');
  const theme = { ...DEFAULT_THEME, ...(content.theme ?? {}) };
  const deckSpec = buildDeckSpec({ ...content, theme });
  assertDeckSpec(deckSpec);
  if (deckSpecOutput) await writeFile(deckSpecOutput, `${JSON.stringify(deckSpec, null, 2)}\n`);
  const document = PptxDocument.create({
    author: content.author ?? 'Codex', company: content.company ?? '',
    subject: content.subject ?? content.title, title: content.title, slideSize: 'wide',
  });
  for (let index = 0; index < content.slides.length; index += 1) {
    const slide = content.slides[index];
    const number = index + 1;
    if (slide.family === 'cover') addCover(document, theme, slide, number);
    else if (slide.family === 'bands') addBands(document, theme, slide, number);
    else if (slide.family === 'spotlight') addSpotlight(document, theme, slide, number);
    else if (slide.family === 'roles') addRoles(document, theme, slide, number);
    else if (slide.family === 'branches') addBranches(document, theme, slide, number);
    else if (slide.family === 'stats') addStats(document, theme, slide, number);
    else if (slide.family === 'chart') await addChart(document, theme, slide, number);
    else if (slide.family === 'process') addProcess(document, theme, slide, number);
    else if (slide.family === 'actions') addActions(document, theme, slide, number);
    else throw new Error(`Unsupported fast layout family: ${slide.family}`);
  }
  await document.writeFile(output, { compression: true });
  return { output, deckSpec, slideCount: content.slides.length };
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];
    if (value === '--input') result.input = argv[++index];
    else if (value === '--output') result.output = argv[++index];
    else if (value === '--deck-spec-out') result.deckSpecOutput = argv[++index];
    else throw new Error(`Unknown argument: ${value}`);
  }
  if (!result.input || !result.output || !result.deckSpecOutput) {
    throw new Error('Usage: ppt-fast-create.mjs --input content.json --output deck.pptx --deck-spec-out deck-spec.json');
  }
  return result;
}

export async function main(argv = process.argv.slice(2)) {
  try {
    const args = parseArgs(argv);
    const content = JSON.parse(await readFile(path.resolve(args.input), 'utf8'));
    const result = await createFastPresentation(
      content,
      path.resolve(args.output),
      path.resolve(args.deckSpecOutput),
    );
    process.stdout.write(`${JSON.stringify({ ok: true, ...result, deckSpec: undefined })}\n`);
  } catch (error) {
    process.stdout.write(`${JSON.stringify({ ok: false, error: String(error) })}\n`);
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
