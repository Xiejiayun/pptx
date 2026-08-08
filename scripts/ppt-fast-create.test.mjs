import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { PptxDocument } from '../packages/pptx/dist/index.js';
import { createFastPresentation, normalizeFastTheme } from './ppt-fast-create.mjs';

test('fast compiler normalizes dark content backgrounds to a readable light field', () => {
  const theme = normalizeFastTheme({ background: '071F17', surface: '123D2A' });
  assert.ok(Number.parseInt(theme.background.slice(0, 2), 16) > 200);
  assert.ok(Number.parseInt(theme.surface.slice(2, 4), 16) > 150);
  assert.equal(theme.deep, '0B2E22');
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

test('Amazon-sized bands, chart, and four-action copy passes full text preflight', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-amazon-copy-'));
  const content = {
    title: 'Amazon biodiversity',
    slides: [
      {
        family: 'bands', title: 'A vertical city of plants',
        rows: [
          { heading: 'Emergents', body: 'Kapok and Brazil-nut crowns rise into fierce sun and wind.', detail: '45–60 m' },
          { heading: 'Canopy', body: 'A dense green roof where leaves, fruit and epiphytes concentrate life.', detail: 'Primary engine' },
          { heading: 'Understory', body: 'Palms and saplings thrive in filtered light with oversized leaves.', detail: 'Low light' },
          { heading: 'Forest floor', body: 'Fungi and roots rapidly reclaim nutrients from fallen material.', detail: 'Fast cycling' },
        ],
      },
      {
        family: 'chart', title: 'Brazilian Amazon loss remains immense',
        chart: { name: 'Deforested area', categories: ['2020', '2021', '2022', '2023'], values: [10851, 13038, 11594, 9001] },
        callout: { value: '9,001 km²', heading: 'Lost in 2023', body: 'A sharp decline from 2022, yet an area larger than many major cities.' },
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
  assert.equal(result.slideCount, 3);
  assert.equal(result.deckSpec.slides[2].elements.some((element) => element.id === 'item-4-heading'), true);
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
