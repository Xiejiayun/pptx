import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { PptxDocument } from '../packages/pptx/dist/index.js';
import {
  accessibleTextColor,
  contrastRatio,
  createFastPresentation,
  normalizeFastTheme,
} from './ppt-fast-create.mjs';

function textShape(slide, value) {
  return slide.shapes.find((shape) => (shape.richText ?? [])
    .flatMap((paragraph) => paragraph.runs)
    .map((run) => run.text)
    .join('') === value);
}

function textColor(slide, value) {
  const shape = textShape(slide, value);
  assert.ok(shape, `Expected rendered text shape: ${value}`);
  const color = shape.richText[0]?.runs[0]?.style?.color;
  assert.equal(color?.kind, 'srgb');
  return color.value;
}

test('fast compiler normalizes dark content backgrounds to a readable light field', () => {
  const theme = normalizeFastTheme({ background: '071F17', surface: '123D2A' });
  assert.ok(Number.parseInt(theme.background.slice(0, 2), 16) > 200);
  assert.ok(Number.parseInt(theme.surface.slice(2, 4), 16) > 150);
  assert.equal(theme.deep, '0B2E22');
});

test('accessible text selection guarantees WCAG AA contrast for cold-run theme colors', () => {
  const theme = normalizeFastTheme({
    deep: '062C20', background: 'EFF5E9', accent: 'F2C14E',
    cool: '2A7F8E', danger: 'B4472D', surface: 'CFE4C3',
  });
  for (const preferred of [theme.cool, theme.danger, theme.surface]) {
    const selected = accessibleTextColor(theme.deep, preferred, [theme.contrast, theme.accent]);
    assert.ok(contrastRatio(theme.deep, selected) >= 4.5);
  }
  const statsHeading = accessibleTextColor(theme.background, theme.accent, [theme.deep, theme.text]);
  assert.ok(contrastRatio(theme.background, statsHeading) >= 4.5);
});

test('rendered text across every layout family retains accessible contrast', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-contrast-'));
  const inputTheme = {
    deep: '062C20', primary: '0B5B3A', secondary: '2F7D4A', accent: 'F2C14E',
    background: 'EFF5E9', surface: 'CFE4C3', text: '143126', mutedText: '5A6F63',
    contrast: 'FFFFFA', danger: 'B4472D', cool: '2A7F8E', font: 'Arial',
  };
  const item = (heading) => ({ heading, body: `${heading} supports the system.` });
  const slides = [
    { family: 'cover', title: 'LIVING SYSTEMS', kicker: 'FIELD GUIDE', subtitle: 'A connected natural world.' },
    { family: 'roles', title: 'Roles sustain the system', kicker: 'SYSTEM ROLES', items: ['RAIN', 'CARBON', 'RIVERS', 'CLIMATE'].map(item), footer: 'Benefits cross borders.' },
    { family: 'stats', title: 'Three signals reveal the system', kicker: 'SYSTEM SIGNALS', items: [
      { ...item('GLOBAL SHARE'), value: '10%', unit: 'species' },
      { ...item('INSECT REALM'), value: '2.5M', unit: 'species' },
      { ...item('TREE DIVERSITY'), value: '16,000', unit: 'tree species' },
    ] },
    { family: 'spotlight', title: 'One actor anchors the story', kicker: 'KEYSTONE SPECIES', hero: { heading: 'JAGUAR', subheading: 'Apex predator', body: 'A wide-ranging forest hunter.' }, items: ['TAPIR', 'DOLPHIN', 'MONKEY'].map(item) },
    { family: 'branches', title: 'Small species power large systems', kicker: 'HIDDEN WORKERS', items: ['ANTS', 'BEES', 'BEETLES', 'BUTTERFLIES'].map(item) },
    { family: 'bands', title: 'Layers organize the forest', kicker: 'VERTICAL SYSTEM', rows: [
      { heading: 'EMERGENT', body: 'Crowns meet open sun.', detail: 'Eagles patrol above.' },
      { heading: 'CANOPY', body: 'Leaves capture light.', detail: 'Fruit feeds animals.' },
      { heading: 'UNDERSTORY', body: 'Saplings wait below.', detail: 'Shade shapes growth.' },
      { heading: 'FLOOR', body: 'Fungi recycle matter.', detail: 'Nutrients return.' },
    ] },
    { family: 'chart', title: 'Loss changes over time', kicker: 'ANNUAL TREND', chart: { name: 'Annual loss (km²)', categories: ['2022', '2023', '2024'], values: [11, 9, 6] }, callout: { value: '6 km²', heading: 'Lower, not over', body: 'The remaining loss still matters.' } },
    { family: 'process', title: 'Pressure compounds in sequence', kicker: 'CHAIN REACTION', items: ['ACCESS', 'CLEARING', 'FIRE', 'DECLINE'].map(item), footer: 'Prevention interrupts the chain.' },
    { family: 'actions', title: 'Four moves protect the system', kicker: 'ACT TOGETHER', items: ['PROTECT', 'ENFORCE', 'RESTORE', 'FINANCE'].map(item) },
  ];
  const output = path.join(directory, 'contrast.pptx');
  await createFastPresentation({ title: 'Contrast', theme: inputTheme, slides }, output, path.join(directory, 'deck-spec.json'));
  const reopened = await PptxDocument.open(output);
  const theme = normalizeFastTheme(inputTheme);
  const expectContrast = (slideIndex, value, background) => {
    assert.ok(contrastRatio(background, textColor(reopened.slides[slideIndex], value)) >= 4.5,
      `Expected ${value} to meet contrast on slide ${slideIndex + 1}`);
  };

  for (const value of ['LIVING SYSTEMS', 'FIELD GUIDE', 'A connected natural world.']) {
    expectContrast(0, value, theme.deep);
  }
  for (const heading of ['RAIN', 'CARBON', 'RIVERS', 'CLIMATE']) {
    expectContrast(1, heading, theme.deep);
    expectContrast(1, `${heading} supports the system.`, theme.deep);
  }
  expectContrast(1, 'Benefits cross borders.', theme.deep);
  for (const heading of ['GLOBAL SHARE', 'INSECT REALM', 'TREE DIVERSITY']) {
    expectContrast(2, heading, theme.background);
    expectContrast(2, `${heading} supports the system.`, theme.background);
  }
  for (const [value, background] of [['10%', theme.primary], ['2.5M', theme.cool], ['16,000', theme.accent]]) {
    expectContrast(2, value, background);
  }
  for (const value of ['JAGUAR', 'Apex predator', 'A wide-ranging forest hunter.']) {
    expectContrast(3, value, theme.accent);
  }
  for (const heading of ['TAPIR', 'DOLPHIN', 'MONKEY', 'ANTS', 'BEES', 'BEETLES', 'BUTTERFLIES']) {
    const slideIndex = ['TAPIR', 'DOLPHIN', 'MONKEY'].includes(heading) ? 3 : 4;
    expectContrast(slideIndex, heading, theme.background);
    expectContrast(slideIndex, `${heading} supports the system.`, theme.background);
  }
  const bandBackgrounds = [theme.surface, theme.secondary, theme.primary, theme.deep];
  for (const [index, row] of slides[5].rows.entries()) {
    for (const value of [row.heading, row.body, row.detail]) expectContrast(5, value, bandBackgrounds[index]);
  }
  for (const value of ['Annual loss (km²)']) expectContrast(6, value, theme.background);
  for (const value of ['6 km²', 'Lower, not over', 'The remaining loss still matters.']) {
    expectContrast(6, value, theme.deep);
  }
  const processBackgrounds = [theme.accent, theme.danger, '9E3A2A', theme.deep];
  for (const [index, processItem] of slides[7].items.entries()) {
    expectContrast(7, processItem.heading, processBackgrounds[index]);
    expectContrast(7, processItem.body, processBackgrounds[index]);
  }
  expectContrast(7, 'Prevention interrupts the chain.', theme.background);
  for (const actionItem of slides[8].items) {
    expectContrast(8, actionItem.heading, theme.deep);
    expectContrast(8, actionItem.body, theme.deep);
  }
  expectContrast(8, '01', theme.deep);

  const titleBackgrounds = [theme.deep, theme.deep, theme.background, theme.background, theme.background,
    theme.background, theme.background, theme.background, theme.deep];
  slides.forEach((slide, index) => expectContrast(index, slide.title, titleBackgrounds[index]));
  slides.forEach((slide, index) => expectContrast(index, slide.kicker,
    index === 0 || index === 1 || index === 8 ? theme.deep : theme.background));
});

test('fast compiler creates and reopens a preflighted deck', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-create-'));
  const output = path.join(directory, 'deck.pptx');
  const deckSpecOutput = path.join(directory, 'deck-spec.json');
  const content = {
    title: 'Forest systems',
    slides: [
      {
        family: 'cover', title: 'FOREST SYSTEMS', subtitle: 'A connected living network',
        kicker: 'FIELD GUIDE', sources: ['https://example.org/cover'],
      },
      {
        family: 'stats', title: 'Three signals reveal the system', kicker: 'SYSTEM SIGNALS',
        items: [
          { value: '150', unit: 'units', heading: 'STORED', body: 'A concise explanation.' },
          { value: '50%', unit: 'share', heading: 'RECYCLED', body: 'A concise explanation.' },
          { value: '10%', unit: 'share', heading: 'REPRESENTED', body: 'A concise explanation.' },
        ],
        sources: ['https://example.org/stats'],
      },
    ],
  };
  const result = await createFastPresentation(content, output, deckSpecOutput);
  assert.equal(result.slideCount, 2);
  assert.equal(result.deckSpec.slides.length, 2);
  const savedSpec = JSON.parse(await readFile(deckSpecOutput, 'utf8'));
  assert.equal(savedSpec.slides[1].family, 'stats');
  const reopened = await PptxDocument.open(output);
  assert.equal(reopened.slides.length, 2);
  assert.equal(reopened.slides[0].title.text, 'FOREST SYSTEMS');
  assert.match(reopened.slides[1].notes, /^\[Sources\]/u);
});

test('fast compiler rejects an unsafe title before writing', async () => {
  const content = {
    title: 'Unsafe',
    slides: [{ family: 'stats', title: 'A title that is intentionally far too long to fit safely on one line in a presentation', items: [] }],
  };
  await assert.rejects(createFastPresentation(content, '/tmp/never-written.pptx'), /TEXT_HORIZONTAL_OVERFLOW/);
});

test('process layout preflights and renders a two-line Edges dry and burn heading', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-process-'));
  const output = path.join(directory, 'deck.pptx');
  const deckSpecOutput = path.join(directory, 'deck-spec.json');
  const items = [
    { heading: 'Roads open access', body: 'New corridors fragment habitat and expose intact forest edges.' },
    { heading: 'Forest is cleared', body: 'Agricultural expansion replaces complex ecosystems.' },
    { heading: 'Edges dry and burn', body: 'Hotter, windier boundaries make degradation and fire more likely.' },
    { heading: 'Resilience falls', body: 'Species loss and weaker moisture recycling raise dieback risk.' },
  ];
  const result = await createFastPresentation({
    title: 'Amazon pressure',
    slides: [{ family: 'process', title: 'Deforestation unravels a living network', items }],
  }, output, deckSpecOutput);
  const processSpec = result.deckSpec.slides[0];
  const thirdHeading = processSpec.elements.find((element) => element.id === 'step-3-heading');
  assert.equal(thirdHeading.text, 'Edges dry and burn');
  assert.equal(thirdHeading.fontSize, 24);
  assert.equal(thirdHeading.wrap, true);
  assert.equal(thirdHeading.maxLines, 2);
  const reopened = await PptxDocument.open(output);
  assert.equal(reopened.slides.length, 1);
  assert.equal(reopened.slides[0].title.text, 'Deforestation unravels a living network');
});

test('process layout rejects step copy that exceeds its declared text budget', async () => {
  const item = { heading: 'A heading that cannot fit in the compact process card even on two lines', body: 'Short body.' };
  await assert.rejects(createFastPresentation({
    title: 'Unsafe process',
    slides: [{ family: 'process', title: 'Pressure compounds through a chain', items: [item] }],
  }, '/tmp/never-written-process.pptx'), /step-1-heading/);
});

test('fast compiler rejects unsafe body copy outside the process family', async () => {
  const item = {
    value: '42%', unit: 'share', heading: 'Signal',
    body: 'This intentionally excessive explanation repeats far beyond the three-line body budget. '.repeat(6),
  };
  await assert.rejects(createFastPresentation({
    title: 'Unsafe stats',
    slides: [{ family: 'stats', title: 'Three signals reveal the system', items: [item] }],
  }, '/tmp/never-written-stats.pptx'), /item-1-body/);
});

test('cold-run bands, roles footer, chart labels, and four-action copy pass full text preflight', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-amazon-copy-'));
  const content = {
    title: 'Amazon biodiversity',
    slides: [
      {
        family: 'bands', title: 'A vertical city of plants',
        rows: [
          { heading: 'Emergents', body: 'Kapok and Brazil-nut crowns rise into fierce sun and wind.', detail: 'Harpy eagles patrol the upper air.' },
          { heading: 'Canopy', body: 'A dense green roof where leaves, fruit and epiphytes concentrate life.', detail: 'Monkeys, sloths, birds, insects.' },
          { heading: 'Understory', body: 'Palms and saplings thrive in filtered light with oversized leaves.', detail: 'Frogs, cats, palms, and saplings.' },
          { heading: 'Forest floor', body: 'Fungi and roots rapidly reclaim nutrients from fallen material.', detail: 'Roots and fungi recycle nutrients.' },
        ],
      },
      {
        family: 'roles', title: 'The Amazon moves water and carbon',
        items: [
          { heading: 'RECYCLE RAIN', body: 'Trees release water vapor that helps sustain rainfall across the basin.' },
          { heading: 'STORE CARBON', body: 'Wood, roots, and soils hold carbon accumulated over decades.' },
          { heading: 'COOL THE LAND', body: 'Evapotranspiration moves heat from the surface into the atmosphere.' },
          { heading: 'LINK CONTINENTS', body: 'Moisture transport influences farms, cities, and rivers beyond the forest.' },
        ],
        footer: 'The forest is living climate infrastructure.',
      },
      {
        family: 'chart', title: 'Brazilian Amazon loss remains immense',
        chart: { name: 'Annual deforestation (km²)', categories: ['2021', '2022', '2023', '2024'], values: [13038, 11594, 9064, 6288] },
        callout: { value: '6,288 km²', heading: 'Three-year decline', body: 'Progress that still leaves a vast annual scar.' },
      },
      {
        family: 'actions', title: 'Keep the forest standing — and thriving',
        items: [
          { heading: 'Secure Indigenous rights', body: 'Support territorial governance and locally led stewardship.' },
          { heading: 'Protect connected habitat', body: 'Expand and enforce reserves, corridors and river safeguards.' },
          { heading: 'Transform supply chains', body: 'Trace commodities and eliminate conversion from production.' },
          { heading: 'Restore strategically', body: 'Reconnect fragments with native species while preventing new loss.' },
        ],
      },
    ],
  };
  const result = await createFastPresentation(
    content,
    path.join(directory, 'deck.pptx'),
    path.join(directory, 'deck-spec.json'),
  );
  assert.equal(result.slideCount, 4);
  const bandsDetail = result.deckSpec.slides[0].elements.find((element) => element.id === 'row-1-detail');
  assert.equal(bandsDetail.bounds.height, 548640);
  const rolesFooter = result.deckSpec.slides[1].elements.find((element) => element.id === 'footer');
  assert.equal(rolesFooter.bounds.width, 5212080);
  const chartName = result.deckSpec.slides[2].elements.find((element) => element.id === 'chart-name');
  assert.equal(chartName.text, 'Annual deforestation (km²)');
  const chartHeading = result.deckSpec.slides[2].elements.find((element) => element.id === 'callout-heading');
  assert.equal(chartHeading.bounds.height, 786384);
  assert.equal(result.deckSpec.slides[3].elements.some((element) => element.id === 'item-4-heading'), true);

  const reopened = await PptxDocument.open(path.join(directory, 'deck.pptx'));
  const chartText = reopened.slides[2].shapes
    .flatMap((shape) => shape.richText ?? [])
    .flatMap((paragraph) => paragraph.runs)
    .map((run) => run.text)
    .join(' ');
  assert.match(chartText, /Annual deforestation \(km²\)/u);
  assert.match(chartText, /6,288 km²/u);
});

test('documented fast-path character targets compile at their stated limits', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-budget-limits-'));
  const copyAt = (length, phrase = 'a forest ') => {
    const repeated = phrase.repeat(Math.ceil(length / phrase.length) + 1);
    return `${repeated.slice(0, length - 1)}x`;
  };
  const item = (headingLength, bodyLength) => ({
    heading: copyAt(headingLength),
    body: copyAt(bodyLength),
  });
  const title = copyAt(40);
  const kicker = copyAt(36);
  const slides = [
    { family: 'cover', title: copyAt(30), subtitle: copyAt(90) },
    { family: 'bands', title, kicker, rows: Array.from({ length: 4 }, () => ({
      ...item(13, 64), detail: copyAt(34),
    })) },
    { family: 'spotlight', title, kicker,
      hero: { ...item(13, 90), subheading: copyAt(20) },
      items: Array.from({ length: 3 }, () => item(22, 74)) },
    { family: 'roles', title, kicker, items: Array.from({ length: 4 }, () => item(18, 64)), footer: copyAt(45) },
    { family: 'branches', title, kicker, items: Array.from({ length: 3 }, () => item(22, 74)), callout: copyAt(48) },
    { family: 'stats', title, kicker, items: Array.from({ length: 3 }, () => ({
      ...item(19, 62), value: '111111111', unit: copyAt(20),
    })) },
    { family: 'chart', title, kicker,
      chart: { name: copyAt(30), categories: [copyAt(10), copyAt(10), copyAt(10)], values: [1, 2, 3] },
      callout: { value: '111111111', heading: copyAt(22), body: copyAt(62) } },
    { family: 'process', title, kicker, items: Array.from({ length: 4 }, () => item(20, 65)), footer: copyAt(75) },
    { family: 'actions', title, kicker, items: Array.from({ length: 4 }, () => item(30, 72)) },
  ];
  const output = path.join(directory, 'budget-limits.pptx');
  const result = await createFastPresentation(
    { title: 'Budget limits', slides }, output, path.join(directory, 'deck-spec.json'),
  );
  assert.equal(result.slideCount, 9);
  const reopened = await PptxDocument.open(output);
  assert.equal(reopened.slides.length, 9);
});

test('fast compiler materializes every registered layout family', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-families-'));
  const item = { heading: 'ROLE', body: 'Concise supporting explanation.' };
  const slides = [
    { family: 'cover', title: 'SYSTEMS', subtitle: 'A short frame' },
    { family: 'bands', title: 'Layers organize the system', rows: Array.from({ length: 4 }, () => ({ ...item, detail: 'Example' })) },
    { family: 'spotlight', title: 'One actor anchors the story', hero: { heading: 'HERO', body: 'Core role.' }, items: [item, item, item] },
    { family: 'roles', title: 'Four roles keep work moving', items: [item, item, item, item] },
    { family: 'branches', title: 'Structure creates new habitat', items: [
      item, item, item, { heading: 'FOURTH BRANCH', body: 'Fourth branch remains visible.' },
    ] },
    { family: 'stats', title: 'Three signals reveal the system', items: Array.from({ length: 3 }, () => ({ ...item, value: '42%', unit: 'share' })) },
    { family: 'chart', title: 'The trend changes over time', chart: { name: 'Trend', categories: ['A', 'B'], values: [1, 2] }, callout: { value: '+1', heading: 'change', body: 'Meaning.' } },
    { family: 'process', title: 'Pressure compounds through a chain', items: [item, item, item, item] },
    { family: 'actions', title: 'Four moves protect the system', items: [
      item, item, item, { heading: 'FOURTH ACTION', body: 'Fourth action remains visible.' },
    ] },
  ];
  const output = path.join(directory, 'families.pptx');
  await createFastPresentation({ title: 'Families', slides }, output, path.join(directory, 'deck-spec.json'));
  const reopened = await PptxDocument.open(output);
  assert.equal(reopened.slides.length, 9);
  assert.deepEqual(reopened.slides.map((slide) => slide.title.text), slides.map((slide) => slide.title));
  const slideText = (slide) => slide.shapes
    .flatMap((shape) => shape.richText ?? [])
    .flatMap((paragraph) => paragraph.runs)
    .map((run) => run.text)
    .join(' ');
  assert.match(slideText(reopened.slides[4]), /FOURTH BRANCH/);
  assert.match(slideText(reopened.slides[8]), /FOURTH ACTION/);
});
