import { describe, expect, it } from 'vitest';
import { CodecRegistry } from '@pptx/codecs';
import { OpcPackage } from '@pptx/opc';
import { createMinimalPptx } from '@pptx/testkit';
import { installAdvancedChartPlugin } from './index.js';

describe('AdvancedChartCodec', () => {
  it('edits combination-chart features and reports workbook divergence', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    pkg.setPart('/ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/></c:barChart><c:lineChart><c:axId val="1"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    pkg.setPart('/ppt/embeddings/workbook1.xlsx', new Uint8Array([1, 2, 3]), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    pkg.addRelationship('/ppt/charts/chart1.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
      target: '../embeddings/workbook1.xlsx',
    });
    const codec = installAdvancedChartPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    expect(codec.inspect('/ppt/charts/chart1.xml')).toMatchObject({ combination: true, chartTypes: ['barChart', 'lineChart'] });
    codec.setSeriesValues('/ppt/charts/chart1.xml', 0, [20, 30]);
    codec.addTrendline('/ppt/charts/chart1.xml', 0);
    codec.addErrorBars('/ppt/charts/chart1.xml', 0, 2);
    codec.enableDataLabels('/ppt/charts/chart1.xml', 0);
    expect(codec.inspect('/ppt/charts/chart1.xml').series[0]).toMatchObject({
      values: [20, 30],
      hasTrendline: true,
      hasErrorBars: true,
      hasDataLabels: true,
    });
    expect(codec.diagnostics('/ppt/charts/chart1.xml')[0]?.code).toBe('CHART_WORKBOOK_CACHE_DIVERGENCE');
    codec.replaceEmbeddedWorkbook('/ppt/charts/chart1.xml', new Uint8Array([4, 5, 6]));
    expect(codec.diagnostics('/ppt/charts/chart1.xml')).toEqual([]);
    expect(codec.createImageFallback('/ppt/charts/chart1.xml', new Uint8Array([7]))).toMatch(/chartFallback/);
  });
});

