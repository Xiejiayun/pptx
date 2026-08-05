import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { ChartModel, inches, PptxDocument } from '../packages/pptx/dist/index.js';

const root = new URL('..', import.meta.url);
const skillDir = new URL('../skills/ppt/', import.meta.url);

test('ppt skill has the exact documented file surface', async () => {
  const entries = (await readdir(skillDir)).sort();
  assert.deepEqual(entries, ['SKILL.md', 'pptx.md']);

  const skill = await readFile(new URL('SKILL.md', skillDir), 'utf8');
  const reference = await readFile(new URL('pptx.md', skillDir), 'utf8');
  const combined = `${skill}\n${reference}`;

  assert.match(skill, /^---\nname: ppt\ndescription: .+\n---/);
  assert.match(skill, /Read \[pptx\.md\]\(pptx\.md\) completely/);
  assert.match(skill, /Create:|Edit:|Inspect:/);
  assert.match(skill, /separate output path/);
  assert.match(skill, /semantic APIs/);
  assert.match(skill, /\[Sources\]/);
  assert.match(skill, /Render every slide/);
  assert.match(skill, /at least one concrete correction/);
  assert.match(reference, /PptxDocument\.create/);
  assert.match(reference, /PptxDocument\.open/);
  assert.match(reference, /tableToSlides/);
  assert.match(reference, /package validate/);
  assert.match(reference, /Deliberate boundaries/);
  assert.doesNotMatch(combined, /\b(?:TODO|TBD)\b|\[TODO|placeholder content/i);
});

test('documented core creation and edit path works at runtime', async () => {
  const temp = await mkdtemp(path.join(tmpdir(), 'ppt-skill-'));
  const created = path.join(temp, 'created.pptx');
  const edited = path.join(temp, 'edited.pptx');

  try {
    const document = PptxDocument.create({ author: 'Skill test', title: 'Native creation', slideSize: 'wide' });
    const slide = document.addSlide();
    const rich = slide.addRichText([{ runs: [{
      text: 'Native creation',
      style: {
        fontFamily: 'Aptos Display', fontSize: 28, bold: true,
        color: { kind: 'srgb', value: '113D2C' },
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

    assert.equal(rich.richText[0].runs[0].style.fontFamily, 'Aptos Display');
    assert.equal(rich.richText[0].runs[0].style.fontSize, 28);
    assert.equal(rich.richText[0].runs[0].style.bold, true);
    assert.deepEqual(rich.richText[0].runs[0].style.color, { kind: 'srgb', value: '113D2C' });
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
