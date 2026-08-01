import { describe, expect, it, vi } from 'vitest';
import { OpcPackage, relativeRelationshipTarget } from '@pptx/opc';
import { PRESENTATION_FORMAT_PROFILES } from './format.js';
import { PresentationModel } from './presentation.js';
import { ChartModel } from './shapes.js';
import { readChartState } from './chart-state.internal.js';
import { chartWorkbookMatches } from './chart-workbook.internal.js';

const PACKAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

describe('chart semantic editing', () => {
  it('no-ops synchronized definitions and replaces categorical data without losing style spans', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('bar', [{
      name: 'Revenue',
      categories: ['Q1', 'Q2'],
      values: [10, 20],
    }]);
    updateChartXml(pkg, chart, (xml) => xml
      .replace(
        '<c:order val="0"/>',
        '<c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/>'
          + '</a:solidFill></c:spPr>',
      )
      .replace(
        '</c:chartSpace>',
        '<c:extLst><c:ext uri="urn:test"><x:keep xmlns:x="urn:test">KEEP</x:keep>'
          + '</c:ext></c:extLst></c:chartSpace>',
      ));

    const before = packageSnapshot(pkg);
    await expect(chart.replaceDefinition(chart.definition!)).resolves.toBe(chart);
    expect(packageSnapshot(pkg)).toEqual(before);

    const chartPartUri = chart.chartPartUri!;
    const workbookPartUri = chart.workbookPartUri!;
    await expect(chart.replaceSeries([{
      name: 'Net revenue',
      categories: ['Q1', 'Q2', 'Q3'],
      values: [12, 24, 36],
    }])).resolves.toBe(chart);

    expect(chart.chartPartUri).toBe(chartPartUri);
    expect(chart.workbookPartUri).toBe(workbookPartUri);
    expect(chart.definition).toEqual({
      groups: [{
        type: 'bar',
        axis: 'primary',
        options: {
          series: [{
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'ABCDEF' },
            },
          }],
        },
        series: [{
          name: 'Net revenue',
          categories: ['Q1', 'Q2', 'Q3'],
          values: [12, 24, 36],
        }],
      }],
      options: {},
    });
    expect(chart.xml).toContain('<a:srgbClr val="ABCDEF"/>');
    expect(chart.xml).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(chart.xml).toContain('Sheet1!$B$2:$B$4');
    expect(await chartWorkbookMatches(
      pkg.requirePart(chart.workbookPartUri!).bytes,
      chart.definition!,
    )).toBe(true);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedChart = reopened.slides[0]!.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!;
    expect(reopenedChart.definition).toEqual(chart.definition);
    expect(await chartWorkbookMatches(
      reopened.opcPackage.requirePart(reopenedChart.workbookPartUri!).bytes,
      reopenedChart.definition!,
    )).toBe(true);
  });

  it('replaces chart options without rewriting workbook data or unowned extensions', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('bar', [{
      name: 'Revenue',
      categories: ['Q1', 'Q2'],
      values: [10, 20],
    }]);
    updateChartXml(pkg, chart, (xml) => xml
      .replace('</c:ser>', '<c:extLst><c:ext uri="urn:series-keep"/></c:extLst></c:ser>')
      .replace(
        '</c:chartSpace>',
        '<c:extLst><c:ext uri="urn:root-keep"/></c:extLst></c:chartSpace>',
      ));
    const workbookPartUri = chart.workbookPartUri!;
    const workbookBefore = pkg.requirePart(workbookPartUri).bytes.slice();

    await chart.replaceDefinition({
      groups: [{
        type: 'bar',
        axis: 'primary',
        series: chart.definition!.groups[0]!.series,
        options: {
          direction: 'bar',
          grouping: 'stacked',
          gapWidth: 0,
          overlap: -25,
          dataLabels: { showValue: true, position: 'insideEnd' },
          series: [{
            fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4' } },
          }],
        },
      }],
      options: {
        language: 'zh-CN',
        roundedCorners: true,
        title: {
          text: 'Revenue',
          overlay: true,
          position: { x: 0.1, y: 0.2 },
          size: 18,
        },
        legend: { position: 'topRight', overlay: true },
        categoryAxis: { labelRotation: -45 },
        valueAxis: { minimum: 0, maximum: 100, numberFormat: '#,##0' },
        dataTable: { showHorizontalBorder: false, showLegendKeys: false },
      },
    });

    expect(pkg.requirePart(workbookPartUri).bytes).toEqual(workbookBefore);
    expect(chart.definition).toMatchObject({
      groups: [{
        options: {
          direction: 'bar',
          grouping: 'stacked',
          gapWidth: 0,
          overlap: -25,
          dataLabels: { showValue: true, position: 'insideEnd' },
        },
      }],
      options: {
        language: 'zh-CN',
        roundedCorners: true,
        title: { text: 'Revenue', overlay: true, position: { x: 0.1, y: 0.2 }, size: 18 },
        legend: { position: 'topRight', overlay: true },
        categoryAxis: { labelRotation: -45 },
        valueAxis: { minimum: 0, maximum: 100, numberFormat: '#,##0' },
        dataTable: { showHorizontalBorder: false, showLegendKeys: false },
      },
    });
    expect(chart.xml).toContain('uri="urn:series-keep"');
    expect(chart.xml).toContain('uri="urn:root-keep"');

    const beforeNoOp = packageSnapshot(pkg);
    await chart.replaceDefinition(chart.definition!);
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    await chart.replaceDefinition({
      groups: [{
        type: 'bar',
        axis: 'primary',
        series: chart.definition!.groups[0]!.series,
      }],
    });
    expect(pkg.requirePart(workbookPartUri).bytes).toEqual(workbookBefore);
    expect(chart.definition?.options).toEqual({});
    expect(chart.definition?.groups[0]?.options).toBeUndefined();
    expect(chart.xml).toContain('uri="urn:series-keep"');
    expect(chart.xml).toContain('uri="urn:root-keep"');
  });

  it('treats explicit chart option defaults as exact semantic no-ops', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('bar', [{
      name: 'Revenue', categories: ['Q1'], values: [10],
    }]);
    const before = packageSnapshot(pkg);

    await chart.replaceDefinition({
      groups: [{
        type: 'bar',
        axis: 'primary',
        series: chart.definition!.groups[0]!.series,
        options: {
          direction: 'column',
          grouping: 'clustered',
          gapWidth: 150,
          overlap: 0,
          varyColors: false,
        },
      }],
      options: {
        roundedCorners: false,
        displayBlanksAs: 'gap',
        title: { visible: false },
        legend: { visible: false },
        dataTable: { visible: false },
      },
    });

    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('converts chart structures while preserving unowned plot and chart-space children', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }]);
    updateChartXml(pkg, chart, (xml) => xml
      .replace(
        '</c:plotArea>',
        '<c:extLst><c:ext uri="urn:plot-keep"/></c:extLst></c:plotArea>',
      )
      .replace(
        '</c:chartSpace>',
        '<c:extLst><c:ext uri="urn:test"><x:keep xmlns:x="urn:test"/>'
          + '</c:ext></c:extLst></c:chartSpace>',
      ));

    await chart.replaceDefinition({ groups: [{
      type: 'scatter',
      series: [{ name: 'Points', xValues: [1, 2], values: [3, 4] }],
    }] });
    expect(chart.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['scatter', 'primary'],
    ]);
    expect(chart.xml).toContain('<c:ext uri="urn:plot-keep"/>');
    expect(chart.xml).toContain('<x:keep xmlns:x="urn:test"/>');

    await chart.replaceDefinition({ groups: [{
      type: 'bubble',
      series: [{ name: 'Bubbles', xValues: [1, 2], values: [3, 4], sizes: [5, 6] }],
    }] });
    expect(chart.definition?.groups[0]?.type).toBe('bubble');

    await chart.replaceDefinition({ groups: [
      {
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ] });
    expect(chart.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['bar', 'primary'],
      ['line', 'secondary'],
    ]);
    expect(chart.xml).toContain('<c:ext uri="urn:plot-keep"/>');
    expect(await chartWorkbookMatches(
      pkg.requirePart(chart.workbookPartUri!).bytes,
      chart.definition!,
    )).toBe(true);

    const before = packageSnapshot(pkg);
    await expect(chart.replaceSeries([{
      name: 'Unsafe', categories: ['Q1'], values: [1],
    }])).rejects.toThrow(/exactly one group/);
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('upgrades cache-only charts with a canonical workbook and ordered externalData', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('line', [{
      name: 'Plan', categories: ['Q1', 'Q2'], values: [8, 9],
    }]);
    const chartPartUri = chart.chartPartUri!;
    const workbookPartUri = chart.workbookPartUri!;
    pkg.transaction(() => {
      updateChartXml(pkg, chart, (xml) => xml
        .replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/, '')
        .replace(
          '</c:chartSpace>',
          '<c:printSettings><c:headerFooter/></c:printSettings></c:chartSpace>',
        ));
      pkg.deletePart(workbookPartUri);
    });
    expect(readChartState(pkg, chartPartUri)).toMatchObject({ status: 'cache-only' });

    await chart.replaceSeries([{
      name: 'Plan updated', categories: ['Q1', 'Q2'], values: [18, 19],
    }]);
    expect(chart.chartPartUri).toBe(chartPartUri);
    expect(chart.workbookPartUri).toBeDefined();
    expect(readChartState(pkg, chartPartUri)).toMatchObject({ status: 'recognized' });
    expect(chart.xml.indexOf('<c:externalData')).toBeLessThan(
      chart.xml.indexOf('<c:printSettings>'),
    );
    expect(await chartWorkbookMatches(
      pkg.requirePart(chart.workbookPartUri!).bytes,
      chart.definition!,
    )).toBe(true);
  });

  it('isolates shared chart and workbook targets before editing', async () => {
    const { pkg, slide } = emptyPresentation();
    const first = await slide.addChart('bar', [{
      name: 'First', categories: ['Q1', 'Q2'], values: [1, 2],
    }]);
    const second = await slide.addChart('bar', [{
      name: 'Second', categories: ['Q1', 'Q2'], values: [3, 4],
    }]);
    const firstChartUri = first.chartPartUri!;
    const firstWorkbookUri = first.workbookPartUri!;
    const firstChartBytes = pkg.requirePart(firstChartUri).bytes;
    const firstWorkbookBytes = pkg.requirePart(firstWorkbookUri).bytes;
    const secondRelationship = slide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget === second.chartPartUri,
    )!;
    pkg.updateRelationship(slide.partUri, secondRelationship.id, {
      target: relativeRelationshipTarget(slide.partUri, firstChartUri),
      targetMode: 'Internal',
    });
    expect(second.chartPartUri).toBe(firstChartUri);

    await second.replaceSeries([{
      name: 'Second edited', categories: ['Q1', 'Q2'], values: [30, 40],
    }]);
    expect(first.chartPartUri).toBe(firstChartUri);
    expect(first.workbookPartUri).toBe(firstWorkbookUri);
    expect(pkg.requirePart(firstChartUri).bytes).toEqual(firstChartBytes);
    expect(pkg.requirePart(firstWorkbookUri).bytes).toEqual(firstWorkbookBytes);
    expect(first.series[0]?.values).toEqual([1, 2]);
    expect(second.chartPartUri).not.toBe(firstChartUri);
    expect(second.workbookPartUri).not.toBe(firstWorkbookUri);
    expect(second.series[0]?.values).toEqual([30, 40]);

    const third = await slide.addChart('bar', [{
      name: 'Third', categories: ['Q1', 'Q2'], values: [5, 6],
    }]);
    const thirdChartUri = third.chartPartUri!;
    const thirdWorkbookRelationship = pkg.relationships(thirdChartUri).find(
      ({ type }) => type === PACKAGE_RELATIONSHIP,
    )!;
    pkg.updateRelationship(thirdChartUri, thirdWorkbookRelationship.id, {
      target: relativeRelationshipTarget(thirdChartUri, firstWorkbookUri),
      targetMode: 'Internal',
    });
    expect(third.workbookPartUri).toBe(firstWorkbookUri);

    await third.replaceSeries([{
      name: 'Third edited', categories: ['Q1', 'Q2'], values: [50, 60],
    }]);
    expect(third.chartPartUri).toBe(thirdChartUri);
    expect(third.workbookPartUri).not.toBe(firstWorkbookUri);
    expect(pkg.requirePart(firstWorkbookUri).bytes).toEqual(firstWorkbookBytes);
    expect(third.series[0]?.values).toEqual([50, 60]);
  });

  it('rolls back every semantic replacement failure boundary', async () => {
    const stages = [
      'clone allocation',
      'workbook write',
      'chart write',
      'relationship retarget',
      'slide XML',
      'validation',
      'outer transaction',
    ] as const;
    for (const stage of stages) {
      const { pkg, slide, selected } = await sharedReferenceFixture();
      const before = packageSnapshot(pkg);
      const selectedPartUri = selected.chartPartUri;
      const restore = injectReplacementFailure(pkg, slide.partUri, stage);
      try {
        await expect(selected.replaceSeries([{
          name: 'Edited', categories: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30],
        }]), stage).rejects.toThrow();
        expect(packageSnapshot(pkg), stage).toEqual(before);
        expect(selected.chartPartUri, stage).toBe(selectedPartUri);
        expect(slide.shapes.find(({ id }) => id === selected.id), stage).toBe(selected);
      } finally {
        restore();
      }
    }
  });
});

function emptyPresentation(): {
  readonly pkg: OpcPackage;
  readonly model: PresentationModel;
  readonly slide: ReturnType<PresentationModel['addSlide']>;
} {
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
  const model = new PresentationModel(pkg);
  return { pkg, model, slide: model.addSlide() };
}

async function sharedReferenceFixture(): Promise<{
  readonly pkg: OpcPackage;
  readonly slide: ReturnType<PresentationModel['addSlide']>;
  readonly selected: ChartModel;
}> {
  const { pkg, slide } = emptyPresentation();
  const first = await slide.addChart('bar', [{
    name: 'First', categories: ['Q1', 'Q2'], values: [1, 2],
  }]);
  const selected = await slide.addChart('bar', [{
    name: 'Second', categories: ['Q1', 'Q2'], values: [3, 4],
  }]);
  const firstRelationship = slide.relationships.find(
    ({ resolvedTarget }) => resolvedTarget === first.chartPartUri,
  )!;
  const selectedRelationship = slide.relationships.find(
    ({ resolvedTarget }) => resolvedTarget === selected.chartPartUri,
  )!;
  pkg.transaction(() => {
    const part = pkg.requirePart(slide.partUri);
    const xml = new TextDecoder().decode(part.bytes).replace(
      `<c:chart r:id="${selectedRelationship.id}"/>`,
      `<c:chart r:id="${firstRelationship.id}"/>`,
    );
    pkg.setPart(slide.partUri, xml, part.contentType);
    pkg.removeRelationship(slide.partUri, selectedRelationship.id);
  });
  expect(selected.chartPartUri).toBe(first.chartPartUri);
  return { pkg, slide, selected };
}

function injectReplacementFailure(
  pkg: OpcPackage,
  slidePartUri: string,
  stage:
    | 'clone allocation'
    | 'workbook write'
    | 'chart write'
    | 'relationship retarget'
    | 'slide XML'
    | 'validation'
    | 'outer transaction',
): () => void {
  if (stage === 'relationship retarget') {
    const original = pkg.addRelationship.bind(pkg);
    const spy = vi.spyOn(pkg, 'addRelationship').mockImplementation((sourceUri, input) => {
      if (sourceUri === slidePartUri) throw new Error(`injected ${stage}`);
      return original(sourceUri, input);
    });
    return () => spy.mockRestore();
  }
  if (stage === 'outer transaction') {
    const original = pkg.transaction.bind(pkg);
    const spy = vi.spyOn(pkg, 'transaction').mockImplementation(((operation: () => unknown) =>
      original(() => {
        operation();
        throw new Error(`injected ${stage}`);
      })) as typeof pkg.transaction);
    return () => spy.mockRestore();
  }

  const original = pkg.setPart.bind(pkg);
  let chartWrites = 0;
  let workbookWrites = 0;
  const spy = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
    if (stage === 'clone allocation' && uri.startsWith('/ppt/charts/chart') && !pkg.hasPart(uri)) {
      throw new Error(`injected ${stage}`);
    }
    if (uri.startsWith('/ppt/charts/chart') && uri.endsWith('.xml')) {
      chartWrites += 1;
      if (stage === 'chart write' && chartWrites === 2) throw new Error(`injected ${stage}`);
      if (stage === 'validation' && chartWrites === 2) {
        const text = typeof bytes === 'string' ? bytes : new TextDecoder().decode(bytes);
        return original(
          uri,
          text.replace(/<c:ptCount val="\d+"\/>/, '<c:ptCount val="999"/>'),
          contentType,
        );
      }
    }
    if (uri.startsWith('/ppt/embeddings/')) {
      workbookWrites += 1;
      if (stage === 'workbook write' && workbookWrites === 2) {
        throw new Error(`injected ${stage}`);
      }
    }
    if (stage === 'slide XML' && uri === slidePartUri) throw new Error(`injected ${stage}`);
    return original(uri, bytes, contentType);
  });
  return () => spy.mockRestore();
}

function updateChartXml(
  pkg: OpcPackage,
  chart: ChartModel,
  update: (xml: string) => string,
): void {
  const uri = chart.chartPartUri!;
  const part = pkg.requirePart(uri);
  pkg.setPart(uri, update(new TextDecoder().decode(part.bytes)), part.contentType);
}

function packageSnapshot(pkg: OpcPackage): object {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
      relationships,
    })),
    graph: pkg.graph,
    mutations: [...pkg.mutations],
  };
}
