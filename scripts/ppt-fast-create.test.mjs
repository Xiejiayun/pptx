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

test('fast compiler materializes every registered layout family', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-families-'));
  const item = { heading: 'ROLE', body: 'Concise supporting explanation.' };
  const slides = [
    { family: 'cover', title: 'SYSTEMS', subtitle: 'A short frame' },
    { family: 'bands', title: 'Layers organize the system', rows: Array.from({ length: 4 }, () => ({ ...item, detail: 'Example' })) },
    { family: 'spotlight', title: 'One actor anchors the story', hero: { heading: 'HERO', body: 'Core role.' }, items: [item, item, item] },
    { family: 'roles', title: 'Four roles keep work moving', items: [item, item, item, item] },
    { family: 'branches', title: 'Structure creates new habitat', items: [item, item, item] },
    { family: 'stats', title: 'Three signals reveal the system', items: Array.from({ length: 3 }, () => ({ ...item, value: '42%', unit: 'share' })) },
    { family: 'chart', title: 'The trend changes over time', chart: { name: 'Trend', categories: ['A', 'B'], values: [1, 2] }, callout: { value: '+1', heading: 'change', body: 'Meaning.' } },
    { family: 'process', title: 'Pressure compounds through a chain', items: [item, item, item, item] },
    { family: 'actions', title: 'Three moves protect the system', items: [item, item, item] },
  ];
  const output = path.join(directory, 'families.pptx');
  await createFastPresentation({ title: 'Families', slides }, output, path.join(directory, 'deck-spec.json'));
  const reopened = await PptxDocument.open(output);
  assert.equal(reopened.slides.length, 9);
  assert.deepEqual(reopened.slides.map((slide) => slide.title.text), slides.map((slide) => slide.title));
});
