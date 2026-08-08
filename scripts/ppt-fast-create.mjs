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

const PROCESS_LAYOUT = Object.freeze({
  positions: Object.freeze([0.82, 3.93, 7.04, 10.15]),
  card: Object.freeze({ y: 2.08, width: 2.45, height: 2.7 }),
  heading: Object.freeze({ dx: 0.2, y: 2.25, width: 2.05, height: 0.82, fontSize: 24 }),
  body: Object.freeze({ dx: 0.2, y: 3.2, width: 2.05, height: 1.3, fontSize: 16 }),
});

const BRANCH_POSITIONS = Object.freeze([
  Object.freeze([0.78, 1.75]), Object.freeze([8.1, 1.75]),
  Object.freeze([0.78, 4.55]), Object.freeze([8.1, 4.55]),
]);

const ACTION_LAYOUT = Object.freeze({ startY: 1.65, stepY: 1.18 });

const DEFAULT_KICKERS = Object.freeze({
  cover: '', bands: 'A layered system', spotlight: 'A keystone relationship',
  roles: 'Invisible work', branches: 'A living architecture', stats: 'Three signals',
  chart: 'Evidence over time', process: 'A chain reaction', actions: 'What changes the trajectory',
});

function luminance(hex) {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16) / 255);
  return channels.reduce((sum, channel, index) => {
    const linear = channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
    return sum + linear * [0.2126, 0.7152, 0.0722][index];
  }, 0);
}

export function contrastRatio(first, second) {
  const lighter = Math.max(luminance(first), luminance(second));
  const darker = Math.min(luminance(first), luminance(second));
  return (lighter + 0.05) / (darker + 0.05);
}

export function accessibleTextColor(background, preferred, alternatives = [], minimum = 4.5) {
  const candidates = [...new Set([preferred, ...alternatives, 'FFFFFF', '000000'])];
  const passing = candidates.find((candidate) => contrastRatio(background, candidate) >= minimum);
  if (passing) return passing;
  return candidates.sort((left, right) => (
    contrastRatio(background, right) - contrastRatio(background, left)
  ))[0];
}

function mixWithWhite(hex, amount) {
  return [0, 2, 4].map((offset) => {
    const channel = Number.parseInt(hex.slice(offset, offset + 2), 16);
    return Math.round(channel + (255 - channel) * amount).toString(16).padStart(2, '0');
  }).join('').toUpperCase();
}

export function normalizeFastTheme(input = {}) {
  const theme = { ...DEFAULT_THEME, ...input };
  const colorKeys = [
    'deep', 'primary', 'secondary', 'accent', 'background', 'surface',
    'text', 'mutedText', 'contrast', 'danger', 'cool',
  ];
  for (const key of colorKeys) {
    if (!/^[0-9A-F]{6}$/iu.test(theme[key])) {
      throw new Error(`Theme ${key} must be a six-digit hex color`);
    }
    theme[key] = theme[key].toUpperCase();
  }
  return {
    ...theme,
    background: luminance(theme.background) < 0.72 ? mixWithWhite(theme.background, 0.88) : theme.background.toUpperCase(),
    surface: luminance(theme.surface) < 0.52 ? mixWithWhite(theme.surface, 0.72) : theme.surface.toUpperCase(),
  };
}

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
  const background = dark ? theme.deep : theme.background;
  const preferred = dark ? theme.contrast : theme.text;
  text(slide, theme, value, 0.65, 0.45, 12, 0.7, 38,
    accessibleTextColor(background, preferred, [theme.deep, theme.contrast, theme.surface]), {
    bold: true, wrap: false,
  });
}

function kicker(slide, theme, value, color = theme.secondary, background = theme.background) {
  text(slide, theme, String(value).toUpperCase(), 0.67, 1.16, 7.5, 0.24, 10.5,
    accessibleTextColor(background, color, [theme.primary, theme.deep, theme.contrast, theme.surface]), {
    bold: true, wrap: false,
  });
}

function slideNumber(slide, theme, number, dark = false) {
  const background = dark ? theme.deep : theme.background;
  text(slide, theme, String(number).padStart(2, '0'), 12.15, 7.02, 0.5, 0.18, 10,
    accessibleTextColor(background, dark ? theme.surface : theme.mutedText,
      [theme.contrast, theme.text, theme.deep]), { align: 'right', wrap: false });
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

function deckText(id, value, fontFamily, x, y, width, height, fontSize, options = {}) {
  return {
    kind: 'text', id, role: options.role ?? 'body', text: String(value ?? ''),
    fontFamily, fontSize, bold: options.bold, wrap: options.wrap ?? true,
    maxLines: options.maxLines ?? 1, fit: 'error', minFontSize: fontSize,
    bounds: { x: inches(x), y: inches(y), width: inches(width), height: inches(height) },
    ...(options.regionId ? { regionId: options.regionId } : {}),
    ...(options.allowBleed ? { allowBleed: true } : {}),
  };
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
      const fontFamily = content.theme?.font ?? DEFAULT_THEME.font;
      const bounds = cover
        ? { x: inches(0.72), y: inches(3.85), width: inches(11.8), height: inches(0.9) }
        : { x: inches(0.65), y: inches(0.45), width: inches(12), height: inches(0.7) };
      const regions = [{ id: 'title-region', bounds, collision: 'exclusive' }];
      const elements = [{
        kind: 'text', id: 'title', regionId: 'title-region', bounds,
        role: 'title', text: slide.title, fontFamily,
        fontSize: cover ? 54 : 38, bold: true, wrap: false, maxLines: 1,
        fit: 'error', minFontSize: cover ? 50 : 36,
      }];
      elements.push(deckText(
        'kicker', slide.kicker ?? DEFAULT_KICKERS[slide.family] ?? '', fontFamily,
        0.67, 1.16, 7.5, 0.24, 10.5,
        { role: 'caption', bold: true, wrap: false, maxLines: 1 },
      ));
      if (slide.family === 'cover') {
        elements.push(deckText('subtitle', slide.subtitle, fontFamily, 0.76, 4.8, 9.9, 0.7, 19, {
          role: 'subtitle', maxLines: 2,
        }));
      } else if (slide.family === 'bands') {
        slide.rows.slice(0, 4).forEach((row, rowIndex) => {
          const y = 1.55 + rowIndex * 1.2;
          elements.push(
            deckText(`row-${rowIndex + 1}-heading`, row.heading, fontFamily, 0.9, y + 0.18, 2.45, 0.4, 24,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`row-${rowIndex + 1}-body`, row.body, fontFamily, 3.5, y + 0.14, 3.8, 0.78, 16,
              { role: 'body', bold: true, maxLines: 3 }),
            deckText(`row-${rowIndex + 1}-detail`, row.detail, fontFamily, 7.55, y + 0.16, 4.6, 0.6, 16,
              { role: 'caption', maxLines: 1 }),
          );
        });
      } else if (slide.family === 'spotlight') {
        elements.push(
          deckText('hero-heading', slide.hero.heading, fontFamily, 1.1, 2.15, 4, 0.7, 40,
            { role: 'label', bold: true, wrap: false, maxLines: 1 }),
          deckText('hero-subheading', slide.hero.subheading, fontFamily, 1.25, 2.95, 3.7, 0.5, 24,
            { role: 'label', bold: true, maxLines: 1 }),
          deckText('hero-body', slide.hero.body, fontFamily, 1.25, 3.55, 3.7, 1.2, 16,
            { role: 'body', maxLines: 4 }),
        );
        slide.items.slice(0, 3).forEach((item, itemIndex) => {
          const y = 1.65 + itemIndex * 1.55;
          elements.push(
            deckText(`item-${itemIndex + 1}-heading`, item.heading, fontFamily, 6.05, y, 5.8, 0.42, 24,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-body`, item.body, fontFamily, 6.05, y + 0.52, 5.8, 0.7, 16,
              { role: 'body', maxLines: 2 }),
          );
        });
      } else if (slide.family === 'roles') {
        const positions = [[0.95, 1.9], [9, 1.9], [0.95, 4.85], [9, 4.85]];
        slide.items.slice(0, 4).forEach((item, itemIndex) => {
          const [x, y] = positions[itemIndex];
          elements.push(
            deckText(`item-${itemIndex + 1}-heading`, item.heading, fontFamily, x, y, 3.35, 0.42, 24,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-body`, item.body, fontFamily, x, y + 0.5, 3.45, 0.88, 16,
              { role: 'body', maxLines: 3 }),
          );
        });
        elements.push(deckText('footer', slide.footer, fontFamily, 3.85, 6.35, 5.7, 0.36, 16,
          { role: 'subtitle', bold: true, wrap: false, maxLines: 1 }));
      } else if (slide.family === 'branches') {
        slide.items.slice(0, 4).forEach((item, itemIndex) => {
          const [x, y] = BRANCH_POSITIONS[itemIndex];
          elements.push(
            deckText(`item-${itemIndex + 1}-heading`, item.heading, fontFamily, x, y, 4.3, 0.42, 24,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-body`, item.body, fontFamily, x, y + 0.58, 4.45, 0.9, 16,
              { role: 'body', maxLines: 3 }),
          );
        });
        if (slide.items.length < 4) {
          elements.push(deckText('callout', slide.callout, fontFamily, 8.1, 5.1, 4.2, 0.8, 24,
            { role: 'label', bold: true, maxLines: 2 }));
        }
      } else if (slide.family === 'stats') {
        slide.items.slice(0, 3).forEach((item, itemIndex) => {
          const x = 0.78 + itemIndex * 3.98;
          elements.push(
            deckText(`item-${itemIndex + 1}-value`, item.value, fontFamily, x + 0.2, 2.15, 2.8, 0.7, 38,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-unit`, item.unit, fontFamily, x + 0.3, 2.95, 2.6, 0.36, 16,
              { role: 'caption', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-heading`, item.heading, fontFamily, x - 0.05, 5.08, 3.4, 0.42, 24,
              { role: 'label', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-body`, item.body, fontFamily, x - 0.05, 5.58, 3.4, 0.8, 16,
              { role: 'body', maxLines: 3 }),
          );
        });
      } else if (slide.family === 'chart') {
        elements.push(
          deckText('chart-name', slide.chart.name, fontFamily, 0.82, 1.5, 8, 0.34, 16,
            { role: 'caption', bold: true, wrap: false, maxLines: 1 }),
          deckText('callout-value', slide.callout.value, fontFamily, 9.5, 2.25, 2.98, 0.75, 34,
            { role: 'label', bold: true, wrap: false, maxLines: 1 }),
          deckText('callout-heading', slide.callout.heading, fontFamily, 9.72, 3.14, 2.56, 0.86, 24,
            { role: 'label', bold: true, maxLines: 2 }),
          deckText('callout-body', slide.callout.body, fontFamily, 9.7, 4.45, 2.58, 1.25, 17,
            { role: 'body', maxLines: 4 }),
        );
      }
      if (slide.family === 'process') {
        slide.items.slice(0, 4).forEach((item, itemIndex) => {
          const x = PROCESS_LAYOUT.positions[itemIndex];
          const regionId = `step-${itemIndex + 1}-region`;
          const headingBounds = {
            x: inches(x + PROCESS_LAYOUT.heading.dx), y: inches(PROCESS_LAYOUT.heading.y),
            width: inches(PROCESS_LAYOUT.heading.width), height: inches(PROCESS_LAYOUT.heading.height),
          };
          const bodyBounds = {
            x: inches(x + PROCESS_LAYOUT.body.dx), y: inches(PROCESS_LAYOUT.body.y),
            width: inches(PROCESS_LAYOUT.body.width), height: inches(PROCESS_LAYOUT.body.height),
          };
          regions.push({
            id: regionId,
            bounds: {
              x: inches(x), y: inches(PROCESS_LAYOUT.card.y),
              width: inches(PROCESS_LAYOUT.card.width), height: inches(PROCESS_LAYOUT.card.height),
            },
            collision: 'exclusive',
          });
          elements.push({
            kind: 'text', id: `step-${itemIndex + 1}-heading`, regionId, bounds: headingBounds,
            role: 'label', text: item.heading, fontFamily,
            fontSize: PROCESS_LAYOUT.heading.fontSize, bold: true, wrap: true, maxLines: 2,
            fit: 'error', minFontSize: PROCESS_LAYOUT.heading.fontSize,
          }, {
            kind: 'text', id: `step-${itemIndex + 1}-body`, regionId, bounds: bodyBounds,
            role: 'body', text: item.body, fontFamily,
            fontSize: PROCESS_LAYOUT.body.fontSize, wrap: true, maxLines: 5,
            fit: 'error', minFontSize: PROCESS_LAYOUT.body.fontSize,
          });
        });
        elements.push(deckText('footer', slide.footer, fontFamily, 1.6, 5.4, 10.1, 0.8, 24,
          { role: 'subtitle', bold: true, maxLines: 2 }));
      } else if (slide.family === 'actions') {
        slide.items.slice(0, 4).forEach((item, itemIndex) => {
          const y = ACTION_LAYOUT.startY + itemIndex * ACTION_LAYOUT.stepY;
          elements.push(
            deckText(`item-${itemIndex + 1}-number`, String(itemIndex + 1).padStart(2, '0'), fontFamily,
              0.85, y, 0.55, 0.4, 20, { role: 'caption', bold: true, wrap: false, maxLines: 1 }),
            deckText(`item-${itemIndex + 1}-heading`, item.heading, fontFamily, 1.65, y, 4.5, 0.8, 24,
              { role: 'label', bold: true, maxLines: 2 }),
            deckText(`item-${itemIndex + 1}-body`, item.body, fontFamily, 6.2, y, 5.7, 0.72, 16,
              { role: 'body', maxLines: 2 }),
          );
        });
      }
      elements.push(deckText('slide-number', String(index + 1).padStart(2, '0'), fontFamily,
        12.15, 7.02, 0.5, 0.18, 10, { role: 'caption', wrap: false, maxLines: 1, allowBleed: true }));
      return {
        id: slide.id ?? `slide-${index + 1}`,
        family: slide.family,
        regions,
        elements,
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
  text(slide, theme, spec.title, 0.72, 3.85, 11.8, 0.9, 54,
    accessibleTextColor(theme.deep, theme.contrast, [theme.surface]), { bold: true, wrap: false });
  text(slide, theme, spec.subtitle ?? '', 0.76, 4.8, 9.9, 0.7, 19,
    accessibleTextColor(theme.deep, theme.surface, [theme.contrast]));
  kicker(slide, theme, spec.kicker ?? '',
    accessibleTextColor(theme.deep, theme.accent, [theme.contrast, theme.surface]), theme.deep);
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
    const background = colors[index];
    const primaryText = accessibleTextColor(background, index > 0 ? theme.contrast : theme.deep,
      [theme.text, theme.surface, theme.contrast]);
    const secondaryText = accessibleTextColor(background, index > 0 ? theme.surface : theme.text,
      [theme.contrast, theme.deep]);
    slide.addShape('rect', { x: inches(0.65), y: inches(y), width: inches(12.03), height: inches(1.08), fill: solid(background), line: { kind: 'none' } });
    text(slide, theme, row.heading, 0.9, y + 0.18, 2.45, 0.4, 24, primaryText, { bold: true, wrap: false });
    text(slide, theme, row.body, 3.5, y + 0.14, 3.8, 0.78, 16, primaryText, { bold: true });
    text(slide, theme, row.detail ?? '', 7.55, y + 0.16, 4.6, 0.6, 16, secondaryText);
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
  const heroText = accessibleTextColor(theme.accent, theme.deep, [theme.text, theme.contrast]);
  text(slide, theme, spec.hero.heading, 1.1, 2.15, 4.0, 0.7, 40, heroText, { bold: true, align: 'center', wrap: false });
  text(slide, theme, spec.hero.subheading ?? '', 1.25, 2.95, 3.7, 0.5, 24, heroText, { bold: true, align: 'center' });
  text(slide, theme, spec.hero.body, 1.25, 3.55, 3.7, 1.2, 16, heroText, { align: 'center', valign: 'middle' });
  spec.items.slice(0, 3).forEach((item, index) => {
    const y = 1.65 + index * 1.55;
    text(slide, theme, item.heading, 6.05, y, 5.8, 0.42, 24,
      accessibleTextColor(theme.background, index === 0 ? theme.cool : theme.secondary,
        [theme.primary, theme.deep, theme.text]), { bold: true, wrap: false });
    text(slide, theme, item.body, 6.05, y + 0.52, 5.8, 0.7, 16,
      accessibleTextColor(theme.background, theme.text, [theme.deep]));
    if (index < 2) slide.addShape('line', { x: inches(6.05), y: inches(y + 1.35), width: inches(5.6), height: 1, line: stroke(theme.secondary, 1, 55) });
  });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addRoles(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.deep);
  title(slide, theme, spec.title, true);
  kicker(slide, theme, spec.kicker ?? 'Invisible work', theme.accent, theme.deep);
  const positions = [[0.95, 1.9], [9.0, 1.9], [0.95, 4.85], [9.0, 4.85]];
  spec.items.slice(0, 4).forEach((item, index) => {
    const [x, y] = positions[index];
    const heading = [theme.accent, theme.cool, theme.danger, theme.surface][index];
    text(slide, theme, item.heading, x, y, 3.35, 0.42, 24,
      accessibleTextColor(theme.deep, heading, [theme.contrast, theme.surface, theme.accent]),
      { bold: true, wrap: false });
    text(slide, theme, item.body, x, y + 0.5, 3.45, 0.88, 16,
      accessibleTextColor(theme.deep, theme.surface, [theme.contrast]));
  });
  slide.addShape('ellipse', { x: inches(5.55), y: inches(2.75), width: inches(2.2), height: inches(1.35), fill: solid(theme.accent), line: { kind: 'none' } });
  for (const [x, y, rotation] of [[4.9, 2.0, -25], [7.0, 2.0, 25], [4.9, 4.0, 25], [7.0, 4.0, -25]]) {
    slide.addShape('ellipse', { x: inches(x), y: inches(y), width: inches(1.55), height: inches(1.05), rotation: degrees(rotation), fill: solid(theme.secondary), line: { kind: 'none' } });
  }
  text(slide, theme, spec.footer ?? '', 3.85, 6.35, 5.7, 0.36, 16,
    accessibleTextColor(theme.deep, theme.contrast, [theme.surface]), { bold: true, align: 'center' });
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
  const items = spec.items.slice(0, 4);
  items.forEach((item, index) => {
    const [x, y] = BRANCH_POSITIONS[index];
    text(slide, theme, item.heading, x, y, 4.3, 0.42, 24,
      accessibleTextColor(theme.background,
        [theme.primary, theme.secondary, theme.cool, theme.danger][index],
        [theme.deep, theme.text]),
      { bold: true, wrap: false });
    text(slide, theme, item.body, x, y + 0.58, 4.45, 0.9, 16,
      accessibleTextColor(theme.background, theme.text, [theme.deep]));
  });
  if (items.length < 4) {
    text(slide, theme, spec.callout ?? '', 8.1, 5.1, 4.2, 0.8, 24,
      accessibleTextColor(theme.background, theme.danger, [theme.deep, theme.text]), { bold: true });
  }
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
    const circleColor = colors[index];
    const circleText = accessibleTextColor(circleColor,
      index === 2 ? theme.deep : theme.contrast, [theme.text, theme.contrast]);
    slide.addShape('ellipse', { x: inches(x), y: inches(1.65), width: inches(3.2), height: inches(3.2), fill: solid(colors[index]), line: { kind: 'none' } });
    text(slide, theme, item.value, x + 0.2, 2.15, 2.8, 0.7, 38, circleText, { bold: true, align: 'center', wrap: false });
    text(slide, theme, item.unit ?? '', x + 0.3, 2.95, 2.6, 0.36, 16, circleText, { bold: true, align: 'center' });
    text(slide, theme, item.heading, x - 0.05, 5.08, 3.4, 0.42, 24,
      accessibleTextColor(theme.background, circleColor, [theme.deep, theme.text]),
      { bold: true, align: 'center', wrap: false });
    text(slide, theme, item.body, x - 0.05, 5.58, 3.4, 0.8, 16,
      accessibleTextColor(theme.background, theme.text, [theme.deep]), { align: 'center' });
  });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

async function addChart(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'Evidence over time');
  text(slide, theme, spec.chart.name, 0.82, 1.5, 8, 0.34, 16,
    accessibleTextColor(theme.background, theme.mutedText, [theme.text, theme.deep]),
    { bold: true, wrap: false });
  const chart = await slide.addChart('bar', [{
    name: spec.chart.name,
    categories: spec.chart.categories,
    values: spec.chart.values,
  }], { x: inches(0.72), y: inches(1.84), width: inches(8.3), height: inches(4.71) });
  await chart.replaceDefinition({
    groups: [{
      type: 'bar',
      series: [{ name: spec.chart.name, categories: spec.chart.categories, values: spec.chart.values }],
      options: { direction: 'column', gapWidth: 58, dataLabels: { showValue: true, position: 'outsideEnd', face: theme.font, size: 11, color: { kind: 'srgb', value: accessibleTextColor(theme.background, theme.text, [theme.deep]) } }, series: [{ fill: solid(theme.secondary), line: { kind: 'none' } }] },
    }],
    options: {
      legend: { visible: false },
      chartArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
      plotArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
      categoryAxis: { face: theme.font, size: 11, color: { kind: 'srgb', value: accessibleTextColor(theme.background, theme.text, [theme.deep]) }, line: stroke(theme.mutedText, 1, 55), majorTickMark: 'none' },
      valueAxis: { visible: false, majorGridLine: stroke(theme.mutedText, 1, 75) },
    },
  });
  slide.addShape('rect', { x: inches(9.45), y: inches(1.78), width: inches(3.08), height: inches(4.62), fill: solid(theme.deep), line: { kind: 'none' } });
  text(slide, theme, spec.callout.value, 9.5, 2.25, 2.98, 0.75, 34,
    accessibleTextColor(theme.deep, theme.accent, [theme.contrast, theme.surface]),
    { bold: true, align: 'center', wrap: false });
  text(slide, theme, spec.callout.heading, 9.72, 3.14, 2.56, 0.86, 24,
    accessibleTextColor(theme.deep, theme.contrast, [theme.surface]), { bold: true, align: 'center' });
  text(slide, theme, spec.callout.body, 9.7, 4.45, 2.58, 1.25, 17,
    accessibleTextColor(theme.deep, theme.surface, [theme.contrast]),
    { align: 'center', valign: 'middle' });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addProcess(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.background);
  title(slide, theme, spec.title);
  kicker(slide, theme, spec.kicker ?? 'A chain reaction', theme.danger);
  const steps = spec.items.slice(0, 4);
  for (let index = 0; index < steps.length - 1; index += 1) {
    const transform = connectorTransform(
      { x: inches(PROCESS_LAYOUT.positions[index] + PROCESS_LAYOUT.card.width), y: inches(3.28) },
      { x: inches(PROCESS_LAYOUT.positions[index + 1] - 0.12), y: inches(3.28) },
    );
    slide.addShape('line', { ...transform, line: stroke(theme.danger, 3), arrows: { end: 'triangle' } });
  }
  steps.forEach((item, index) => {
    const x = PROCESS_LAYOUT.positions[index];
    const color = [theme.accent, theme.danger, '9E3A2A', theme.deep][index];
    const cardText = accessibleTextColor(color,
      index === 0 ? theme.deep : theme.contrast, [theme.text, theme.surface, theme.contrast]);
    slide.addShape('roundRect', {
      x: inches(x), y: inches(PROCESS_LAYOUT.card.y),
      width: inches(PROCESS_LAYOUT.card.width), height: inches(PROCESS_LAYOUT.card.height),
      fill: solid(color), line: { kind: 'none' },
    });
    text(slide, theme, item.heading,
      x + PROCESS_LAYOUT.heading.dx, PROCESS_LAYOUT.heading.y,
      PROCESS_LAYOUT.heading.width, PROCESS_LAYOUT.heading.height,
      PROCESS_LAYOUT.heading.fontSize, cardText,
      { bold: true, wrap: true, fit: 'none' });
    text(slide, theme, item.body,
      x + PROCESS_LAYOUT.body.dx, PROCESS_LAYOUT.body.y,
      PROCESS_LAYOUT.body.width, PROCESS_LAYOUT.body.height,
      PROCESS_LAYOUT.body.fontSize, cardText,
      { fit: 'none' });
  });
  text(slide, theme, spec.footer ?? '', 1.6, 5.4, 10.1, 0.8, 24,
    accessibleTextColor(theme.background, theme.primary, [theme.deep, theme.text]),
    { bold: true, align: 'center' });
  notes(slide, spec.sources);
  slideNumber(slide, theme, number);
}

function addActions(document, theme, spec, number) {
  const slide = document.addSlide();
  slide.background = solid(theme.deep);
  title(slide, theme, spec.title, true);
  kicker(slide, theme, spec.kicker ?? 'What changes the trajectory', theme.accent, theme.deep);
  const items = spec.items.slice(0, 4);
  items.forEach((item, index) => {
    const y = ACTION_LAYOUT.startY + index * ACTION_LAYOUT.stepY;
    text(slide, theme, String(index + 1).padStart(2, '0'), 0.85, y, 0.55, 0.4, 20,
      accessibleTextColor(theme.deep, theme.accent, [theme.contrast, theme.surface]), { bold: true });
    text(slide, theme, item.heading, 1.65, y, 4.5, 0.8, 24,
      accessibleTextColor(theme.deep, theme.contrast, [theme.surface]), { bold: true });
    text(slide, theme, item.body, 6.2, y, 5.7, 0.72, 16,
      accessibleTextColor(theme.deep, theme.surface, [theme.contrast]));
    if (index < items.length - 1) slide.addShape('line', { x: inches(1.65), y: inches(y + 0.98), width: inches(10.2), height: 1, line: stroke(theme.secondary, 1, 55) });
  });
  for (let index = 0; index < 7; index += 1) leaf(slide, theme, 0.55 + index * 1.75, 6.35, 1.3, 0.42, index % 2 ? theme.secondary : theme.surface, index % 2 ? 20 : -20, 30);
  notes(slide, spec.sources);
  slideNumber(slide, theme, number, true);
}

export async function createFastPresentation(content, output, deckSpecOutput) {
  if (!Array.isArray(content.slides) || content.slides.length < 1) throw new Error('Content spec requires slides');
  const theme = normalizeFastTheme(content.theme);
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
