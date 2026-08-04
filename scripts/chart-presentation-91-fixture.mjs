import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const require = createRequire(
  new URL('../packages/pptxgenjs-adapter/package.json', import.meta.url),
);
const PptxGenJS = require('pptxgenjs');

export const CHART_PRESENTATION_PPTXGENJS_VERSION = JSON.parse(readFileSync(
  resolve(dirname(require.resolve('pptxgenjs')), '../package.json'),
  'utf8',
)).version;

const CASES = Object.freeze([
  {
    type: 'pie',
    name: 'pie-bestFit-label-true-percent-false',
    options: { dataLabelPosition: 'bestFit', showLabel: true, showPercent: false },
  },
  {
    type: 'pie',
    name: 'pie-bestFit-label-false-percent-true',
    options: { dataLabelPosition: 'bestFit', showLabel: false, showPercent: true },
  },
  {
    type: 'pie',
    name: 'pie-bestFit-label-false-percent-false',
    options: { dataLabelPosition: 'bestFit', showLabel: false, showPercent: false },
  },
  {
    type: 'doughnut',
    name: 'doughnut-label-true-percent-false',
    options: { dataLabelPosition: 'bestFit', showLabel: true, showPercent: false },
  },
  {
    type: 'scatter',
    name: 'scatter-label-true-custom-control',
    scatter: true,
    options: {
      dataLabelPosition: 't',
      dataLabelFormatScatter: 'custom',
      showLabel: true,
    },
  },
]);

export async function createChartPresentation91Fixture() {
  const presentation = new PptxGenJS();
  for (const entry of CASES) {
    const data = entry.scatter
      ? [
          { name: 'X', values: [1, 2, 3] },
          { name: entry.name, labels: ['A', 'B', 'C'], values: [10, 20, 30] },
        ]
      : [{ name: entry.name, labels: ['A', 'B', 'C'], values: [10, 20, 30] }];
    presentation.addSlide().addChart(presentation.ChartType[entry.type], data, {
      x: 0.5,
      y: 0.5,
      w: 5,
      h: 3,
      ...entry.options,
    });
  }
  const output = await presentation.write({
    outputType: 'uint8array',
    compression: true,
  });
  if (output instanceof Uint8Array) return output;
  if (output instanceof ArrayBuffer) return new Uint8Array(output);
  throw new TypeError('PptxGenJS did not return Uint8Array fixture bytes');
}
