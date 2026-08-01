import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PRESENTATION_FORMAT_PROFILES } from './format.js';
import { PresentationModel } from './presentation.js';
import { chartDiagnostics } from './chart-diagnostics.internal.js';

function emptyPresentationModel(): { pkg: OpcPackage; model: PresentationModel } {
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
  return { pkg, model: new PresentationModel(pkg) };
}

async function chartFixture() {
  const { pkg, model } = emptyPresentationModel();
  const slide = model.addSlide();
  const chart = await slide.addChart('bar', [{
    name: 'Revenue',
    categories: ['Q1', 'Q2'],
    values: [10, 20],
  }]);
  return { pkg, model, slide, chart };
}

describe('chart diagnostics', () => {
  it('reports healthy, divergent, and cache-only chart states without package mutation', async () => {
    const healthy = await chartFixture();
    expect(await chartDiagnostics(healthy.pkg, healthy.slide.partUri)).toEqual([]);
    expect(await healthy.chart.diagnostics()).toEqual([]);

    healthy.chart.setXml(healthy.chart.xml.replace('<c:v>20</c:v>', '<c:v>21</c:v>'));
    expect(await healthy.chart.diagnostics()).toMatchObject([{
      severity: 'error',
      code: 'CHART_WORKBOOK_CACHE_DIVERGENCE',
      partUri: healthy.chart.chartPartUri,
      objectId: String(healthy.chart.id),
    }]);

    const cacheOnly = await chartFixture();
    cacheOnly.chart.setXml(cacheOnly.chart.xml.replace(/<c:externalData[\s\S]*?<\/c:externalData>/, ''));
    expect(await cacheOnly.chart.diagnostics()).toMatchObject([{
      severity: 'warning',
      code: 'CHART_WORKBOOK_MISSING',
    }]);
  });

  it.each([
    [
      'CHART_CACHE_INVALID',
      (xml: string) => xml.replace('<c:ptCount val="2"/>', '<c:ptCount val="3"/>'),
    ],
    [
      'CHART_CACHE_INVALID',
      (xml: string) => xml.replace('Sheet1!$B$2:$B$3', 'OFFSET(A1,0,0)'),
    ],
    [
      'CHART_AXIS_INVALID',
      (xml: string) => xml.replace('<c:crossAx val="10000002"/>', '<c:crossAx val="999"/>'),
    ],
    [
      'CHART_STRUCTURE_AMBIGUOUS',
      (xml: string) => xml.replace('</c:plotArea>', '</c:plotArea><c:plotArea/>'),
    ],
    [
      'CHART_STRUCTURE_UNSUPPORTED',
      (xml: string) => xml.replaceAll('barChart', 'surfaceChart'),
    ],
  ])('classifies malformed standard charts as %s', async (code, mutate) => {
    const { chart } = await chartFixture();
    chart.setXml(mutate(chart.xml));
    expect(await chart.diagnostics()).toMatchObject([{ severity: 'error', code }]);
  });

  it('reports invalid frame relationships and modern chart extensions precisely', async () => {
    const invalid = await chartFixture();
    const relationship = invalid.slide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget === invalid.chart.chartPartUri,
    )!;
    invalid.pkg.removeRelationship(invalid.slide.partUri, relationship.id);
    expect(await chartDiagnostics(invalid.pkg, invalid.slide.partUri)).toMatchObject([{
      severity: 'error',
      code: 'CHART_RELATIONSHIP_INVALID',
      partUri: invalid.slide.partUri,
    }]);

    const modern = await chartFixture();
    const modernRelationship = modern.slide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget === modern.chart.chartPartUri,
    )!;
    modern.pkg.updateRelationship(modern.slide.partUri, modernRelationship.id, {
      type: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
    });
    modern.chart.setXml(
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>',
    );
    expect(await modern.chart.diagnostics()).toMatchObject([{
      severity: 'info',
      code: 'MODERN_CHART_EXTENSION',
    }]);
  });
});
