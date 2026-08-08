import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  assertDeckSpec,
  ChartModel,
  inches,
  PptxDocument,
  preflightDeckSpec,
} from '../packages/pptx/dist/index.js';

const root = new URL('..', import.meta.url);
const skillDir = new URL('../skills/ppt/', import.meta.url);

test('ppt skill has the exact documented file surface', async () => {
  const entries = (await readdir(skillDir)).sort();
  assert.deepEqual(entries, ['SKILL.md', 'pptx.md']);

  const skill = await readFile(new URL('SKILL.md', skillDir), 'utf8');
  const reference = await readFile(new URL('pptx.md', skillDir), 'utf8');
  const combined = `${skill}\n${reference}`;

  assert.match(skill, /^---\nname: ppt\ndescription: .+\n---/);
  assert.match(skill, /does not require reading \[pptx\.md\]\(pptx\.md\)/);
  assert.match(skill, /Read `pptx\.md` completely for edits/);
  assert.match(skill, /Create:|Edit:|Inspect:/);
  assert.match(skill, /separate output path/);
  assert.match(skill, /semantic APIs/);
  assert.match(skill, /\[Sources\]/);
  assert.match(skill, /Render every slide/i);
  assert.match(combined, /ThemeSpec/);
  assert.match(combined, /primary.*secondary.*accent.*background.*surface.*text.*mutedText/s);
  assert.match(combined, /Cambria \+ Arial/);
  assert.match(combined, /Cambria \+ Calibri/);
  assert.match(combined, /Arial \+ Arial/);
  assert.match(combined, /Calibri \+ Calibri/);
  assert.match(combined, /Do not default to Aptos/);
  assert.match(combined, /Forest canopy/);
  assert.match(combined, /Deep ocean/);
  assert.match(combined, /Editorial clay/);
  assert.match(combined, /Night sky/);
  assert.match(combined, /never as fixed templates/);
  assert.match(combined, /cinematic cover/);
  assert.match(combined, /section divider/);
  assert.match(combined, /statement plus hero visual/);
  assert.match(combined, /asymmetric two-column/);
  assert.match(combined, /large statistic plus explanation/);
  assert.match(combined, /comparison/);
  assert.match(combined, /process or timeline/);
  assert.match(combined, /native chart or table plus takeaway/);
  assert.match(combined, /180 seconds/);
  assert.match(combined, /Narrative outline.*20 seconds/s);
  assert.match(combined, /Theme, DeckSpec, and layout preflight.*25 seconds/s);
  assert.match(combined, /generator execution.*75 seconds/s);
  assert.match(combined, /Reopen, validate, render, and inspect in parallel.*45 seconds/s);
  assert.match(combined, /repair and targeted recheck buffer.*15 seconds/si);
  assert.match(combined, /Do not browse, search, scrape, or download/);
  assert.match(combined, /at most one targeted repair/);
  assert.match(combined, /DeckSpec/);
  assert.match(combined, /assertDeckSpec\(\)/);
  assert.match(combined, /connectorTransform\(\)/);
  assert.match(combined, /scripts\/ppt-fast-qa\.mjs/);
  assert.match(combined, /scripts\/ppt-fast-create\.mjs/);
  assert.match(combined, /task-cold clock/);
  assert.match(combined, /If a concrete defect exists/);
  assert.match(combined, /rerender only affected slides/);
  assert.doesNotMatch(combined, /at least one concrete (?:correction|improvement)/i);
  assert.match(reference, /PptxDocument\.create/);
  assert.match(reference, /PptxDocument\.open/);
  assert.match(reference, /tableToSlides/);
  assert.match(reference, /package validate/);
  assert.match(reference, /Deliberate boundaries/);
  assert.doesNotMatch(combined, /\b(?:TODO|TBD)\b|\[TODO|placeholder content/i);
});

test('documented preflight rejects the Amazon clipped-title regression', () => {
  const spec = {
    schemaVersion: 1,
    slideSize: { width: inches(13.333), height: inches(7.5) },
    safeArea: { top: inches(0.4), right: inches(0.4), bottom: inches(0.4), left: inches(0.4) },
    gap: inches(0.3),
    slides: [{
      id: 'slide-7', family: 'chart-or-table',
      regions: [{
        id: 'title-region', collision: 'exclusive',
        bounds: { x: inches(0.62), y: inches(0.42), width: inches(11.9), height: inches(0.64) },
      }],
      elements: [{
        kind: 'text', id: 'title', regionId: 'title-region', role: 'title',
        bounds: { x: inches(0.62), y: inches(0.42), width: inches(11.9), height: inches(0.64) },
        text: 'Deforestation is falling — annual loss remains vast',
        fontFamily: 'Arial', fontSize: 38, wrap: false, maxLines: 1,
        fit: 'error', minFontSize: 36,
      }],
    }],
  };
  const report = preflightDeckSpec(spec);
  assert.equal(report.ok, false);
  assert.ok(report.diagnostics.some((diagnostic) => diagnostic.code === 'TEXT_HORIZONTAL_OVERFLOW'));
  assert.throws(() => assertDeckSpec(spec), /TEXT_HORIZONTAL_OVERFLOW/);
});

test('documented core creation and edit path works at runtime', async () => {
  const temp = await mkdtemp(path.join(tmpdir(), 'ppt-skill-'));
  const created = path.join(temp, 'created.pptx');
  const edited = path.join(temp, 'edited.pptx');

  try {
    const themeSpec = {
      background: 'F3F0E4', surface: 'DDE8D3', secondary: '3D6B45', text: '173126',
      displayFont: 'Cambria', sizes: { title: 40 }, margin: 0.6,
    };
    const document = PptxDocument.create({ author: 'Skill test', title: 'Native creation', slideSize: 'wide' });
    const slide = document.addSlide();
    slide.background = { kind: 'solid', color: { kind: 'srgb', value: themeSpec.background } };
    const surface = slide.addShape('roundRect', {
      x: inches(themeSpec.margin), y: inches(2.2), width: inches(5.8), height: inches(3.8),
      fill: { kind: 'solid', color: { kind: 'srgb', value: themeSpec.surface } },
      line: { kind: 'line', color: { kind: 'srgb', value: themeSpec.secondary }, width: 1 },
    });
    const rich = slide.addRichText([{ runs: [{
      text: 'Native creation',
      style: {
        fontFamily: themeSpec.displayFont, fontSize: themeSpec.sizes.title, bold: true,
        color: { kind: 'srgb', value: themeSpec.text },
      },
    }] }], {
      x: inches(1), y: inches(1), width: inches(6), height: inches(1),
    });
    const body = slide.addRichText([
      { runs: [
        { text: '42%', style: { fontSize: 18, bold: true, color: { kind: 'srgb', value: 'E4A83D' } } },
        { text: ' under pressure.', style: { fontSize: 18 } },
      ] },
      { bullet: true, runs: [{ text: 'Protect habitat.', style: { fontSize: 18 } }] },
    ], {
      x: inches(7.2), y: inches(1), width: inches(4.5), height: inches(1.5), fit: 'shrink',
    });
    slide.slideNumber = {
      x: inches(12), y: inches(7), width: inches(0.5), height: inches(0.25),
      align: 'right', style: { fontSize: 10 },
    };
    const chart = await slide.addChart('bar', [
      { name: 'Area', categories: ['2000', '2010', '2020'], values: [100, 84, 71] },
    ], {
      x: inches(1), y: inches(2.3), width: inches(6), height: inches(3),
    });
    await chart.replaceDefinition({
      groups: [{
        type: 'bar',
        series: [{ name: 'Area', categories: ['2000', '2010', '2020'], values: [100, 84, 71] }],
        options: { dataLabels: { showValue: true, position: 'outsideEnd' } },
      }],
      options: { legend: { visible: true, position: 'bottom' } },
    });
    slide.addNotes('[Sources]\nhttps://example.org');
    await document.writeFile(created, { compression: true });

    assert.equal(rich.richText[0].runs[0].style.fontFamily, 'Cambria');
    assert.equal(rich.richText[0].runs[0].style.fontSize, 40);
    assert.equal(rich.richText[0].runs[0].style.bold, true);
    assert.deepEqual(rich.richText[0].runs[0].style.color, { kind: 'srgb', value: themeSpec.text });
    assert.deepEqual(slide.background, { kind: 'solid', color: { kind: 'srgb', value: themeSpec.background } });
    assert.deepEqual(surface.fill, { kind: 'solid', color: { kind: 'srgb', value: themeSpec.surface } });
    assert.deepEqual(surface.line, {
      kind: 'line', color: { kind: 'srgb', value: themeSpec.secondary }, width: 1, dash: 'solid',
    });
    assert.equal(body.richText[0].runs[0].style.fontSize, 18);
    assert.equal(body.richText[0].runs[1].style.fontSize, 18);
    assert.equal(body.richText[1].runs[0].style.fontSize, 18);
    assert.equal(chart.definition.groups[0].series[0].categories[1], '2010');
    assert.equal(chart.definition.groups[0].options.dataLabels.showValue, true);
    assert.equal(chart.definition.options.legend.position, 'bottom');
    assert.equal(slide.slideNumber.style.fontSize, 10);

    const reopened = await PptxDocument.open(created);
    assert.equal(reopened.slides.length, 1);
    assert.equal(reopened.slides[0].title.text, 'Native creation');
    assert.equal(reopened.slides[0].notes, '[Sources]\nhttps://example.org');
    assert.equal(reopened.slides[0].slideNumber.style.fontSize, 10);
    const reopenedChart = reopened.slides[0].shapes.find((shape) => shape instanceof ChartModel);
    assert.ok(reopenedChart);
    assert.equal(reopenedChart.definition.groups[0].series[0].categories[1], '2010');
    assert.equal(reopenedChart.definition.groups[0].options.dataLabels.showValue, true);
    reopened.slides[0].title.text = 'Semantic edit';
    await reopened.writeFile(edited, { compression: true });

    const verified = await PptxDocument.open(edited);
    assert.equal(verified.slides[0].title.text, 'Semantic edit');
    assert.equal(verified.slides[0].notes, '[Sources]\nhttps://example.org');
  } finally {
    await rm(temp, { recursive: true, force: true });
  }
});

void root;
