import { describe, expect, it } from 'vitest';
import { CodecRegistry } from '@pptx/codecs';
import {
  ChartModel,
  PRESENTATION_FORMAT_PROFILES,
  PresentationModel,
  chartWorkbookMatches,
} from '@pptx/model';
import { OpcPackage } from '@pptx/opc';
import { installAdvancedChartPlugin } from './index.js';

describe('AdvancedChartCodec', () => {
  it('edits combination-chart features and reports workbook divergence', async () => {
    const pkg = emptyPresentationPackage();
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const chart = await slide.addChart([
      {
        type: 'bar',
        series: [{ name: 'Sales', categories: ['Q1', 'Q2'], values: [10, 15] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 16] }],
      },
    ]);
    const chartPartUri = chart.chartPartUri!;
    const codec = installAdvancedChartPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    expect(codec.inspect(chartPartUri)).toMatchObject({
      combination: true,
      chartTypes: ['barChart', 'lineChart'],
    });
    await codec.setSeriesValues(chartPartUri, 0, [20, 30]);
    expect(await chartWorkbookMatches(
      pkg.requirePart(chart.workbookPartUri!).bytes,
      chart.definition!,
    )).toBe(true);
    const synchronizedWorkbook = pkg.requirePart(chart.workbookPartUri!).bytes.slice();
    codec.addTrendline(chartPartUri, 0);
    codec.addErrorBars(chartPartUri, 0, 2);
    codec.enableDataLabels(chartPartUri, 0);
    expect(codec.inspect(chartPartUri).series[0]).toMatchObject({
      values: [20, 30],
      hasTrendline: true,
      hasErrorBars: true,
      hasDataLabels: true,
    });
    expect(await codec.diagnostics(chartPartUri)).toEqual([]);

    await codec.setSeriesValues(chartPartUri, 0, [20, 31]);
    chart.setXml(chart.xml.replace('<c:v>31</c:v>', '<c:v>30</c:v>'));
    expect((await codec.diagnostics(chartPartUri))[0]?.code)
      .toBe('CHART_WORKBOOK_CACHE_DIVERGENCE');
    codec.replaceEmbeddedWorkbook(chartPartUri, synchronizedWorkbook);
    expect(await codec.diagnostics(chartPartUri)).toEqual([]);
    expect(codec.inspect(chartPartUri).series[0]).toMatchObject({
      values: [20, 30],
      hasTrendline: true,
      hasErrorBars: true,
      hasDataLabels: true,
    });
    expect(codec.createImageFallback(chartPartUri, new Uint8Array([7]))).toMatch(/chartFallback/);
    expect(slide.shapes.find(({ id }) => id === chart.id)).toBeInstanceOf(ChartModel);
  });
});

function emptyPresentationPackage(): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.transaction(() => {
    pkg.setPart(
      '/ppt/presentation.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + '<p:sldIdLst/><p:sldSz cx="9144000" cy="5143500"/>'
        + '<p:notesSz cx="5143500" cy="9144000"/></p:presentation>',
      PRESENTATION_FORMAT_PROFILES.pptx.presentationContentType,
    );
    pkg.addRelationship('/', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'ppt/presentation.xml',
    });
  });
  return pkg;
}
