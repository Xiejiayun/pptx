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

  it('edits residual layout without rewriting title, point, label, shape, effect, or extension owners', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart([{
      type: 'bar3D',
      series: [{ name: 'Revenue', categories: ['Q1'], values: [10] }],
      options: { shape: 'coneToMax' },
    }]);
    await chart.replaceDefinition({
      groups: chart.definition!.groups,
      options: {
        layout: { x: 0.1, y: 0.2, width: 0.7, height: 0.6 },
        title: { text: 'Revenue', align: 'right' },
      },
    });
    const pointPayload = '<c:dPt><c:idx val="0"/><c:spPr><a:effectLst/>'
      + '<a:extLst><a:ext uri="urn:point-keep"/></a:extLst></c:spPr></c:dPt>';
    const labelPayload = '<c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich>'
      + '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p>'
      + '</c:rich></c:tx><c:extLst><c:ext uri="urn:label-keep"/></c:extLst>'
      + '</c:dLbl></c:dLbls>';
    const seriesShape = '<c:spPr><a:effectLst/><a:extLst>'
      + '<a:ext uri="urn:series-shape-keep"/></a:extLst></c:spPr>';
    updateChartXml(pkg, chart, (xml) => xml
      .replace('<c:order val="0"/>', `<c:order val="0"/>${pointPayload}${labelPayload}${seriesShape}`)
      .replace('</c:title>', '<c:extLst><c:ext uri="urn:title-keep"/></c:extLst></c:title>'));

    const imported = chart.definition!;
    expect(imported.options.layout).toEqual({ x: 0.1, y: 0.2, width: 0.7, height: 0.6 });
    await chart.replaceDefinition({
      groups: imported.groups,
      options: {
        ...imported.options,
        layout: { x: 0.15, y: 0.25, width: 0.65, height: 0.55 },
      },
    });

    expect(chart.xml).toContain(pointPayload);
    expect(chart.xml).toContain(labelPayload);
    expect(chart.xml).toContain(seriesShape);
    expect(chart.xml).toContain('<c:ext uri="urn:title-keep"/>');
    expect(chart.definition?.options.layout).toEqual({
      x: 0.15, y: 0.25, width: 0.65, height: 0.55,
    });
  });

  it('edits promoted labels only when requested and preserves unsafe custom point labels', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('pie', [{
      name: 'Share', categories: ['A', 'B'], values: [60, 40],
    }]);
    const safeLabels = pptxGenJsSeriesLabels();
    updateChartXml(pkg, chart, (xml) => xml
      .replace('<c:cat>', `${safeLabels}<c:cat>`)
      .replace('</c:ser>', '<c:extLst><c:ext uri="urn:series-keep"/></c:extLst></c:ser>'));
    const workbookPartUri = chart.workbookPartUri!;
    const workbookBefore = pkg.requirePart(workbookPartUri).bytes.slice();

    expect(chart.definition?.groups[0]?.options?.dataLabels).toMatchObject({
      position: 'bestFit',
      showCategoryName: true,
    });
    expect(chart.definition?.groups[0]?.options?.dataLabels?.showPercent).toBeUndefined();

    const imported = chart.definition!;
    const importedGroup = imported.groups[0];
    if (importedGroup?.type !== 'pie') throw new Error('Expected imported pie chart');
    await chart.replaceDefinition({
      groups: [{
        type: 'pie',
        series: importedGroup.series,
        options: {
          ...importedGroup.options,
          dataLabels: {
            ...importedGroup.options!.dataLabels,
            showPercent: false,
          },
        },
      }],
      options: { ...imported.options, title: { text: 'Unrelated title' } },
    });
    expect(chart.xml).toContain(safeLabels);
    expect(chart.xml.match(/<c:dLbls>/gu)).toHaveLength(1);
    expect(chart.xml).toContain('uri="urn:series-keep"');
    expect(pkg.requirePart(workbookPartUri).bytes).toEqual(workbookBefore);

    const titled = chart.definition!;
    await chart.replaceDefinition({
      groups: [{
        ...titled.groups[0]!,
        options: { dataLabels: { position: 'center', showPercent: true } },
      }],
      options: titled.options,
    });
    expect(chart.xml).not.toContain(safeLabels);
    expect(chart.xml.match(/<c:dLbls>/gu)).toHaveLength(1);
    expect(chart.definition?.groups[0]?.options?.dataLabels).toEqual({
      showPercent: true,
      position: 'center',
    });
    expect(chart.xml).toContain('uri="urn:series-keep"');
    expect(pkg.requirePart(workbookPartUri).bytes).toEqual(workbookBefore);

    const edited = chart.definition!;
    await chart.replaceDefinition({
      groups: [{
        type: edited.groups[0]!.type,
        series: edited.groups[0]!.series,
      }],
      options: edited.options,
    });
    expect(chart.xml).not.toContain('<c:dLbls>');
    expect(chart.definition?.groups[0]?.options?.dataLabels).toBeUndefined();
    expect(chart.xml).toContain('uri="urn:series-keep"');
    expect(pkg.requirePart(workbookPartUri).bytes).toEqual(workbookBefore);

    const custom = await slide.addChart('scatter', [{
      name: 'Custom', xValues: [1, 2], values: [3, 4],
    }]);
    const customLabels = pptxGenJsSeriesLabels().replace(
      '<c:idx val="0"/>',
      '<c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>'
        + '<a:r><a:t>KEEP CUSTOM</a:t></a:r></a:p></c:rich></c:tx>',
    );
    updateChartXml(pkg, custom, (xml) => xml.replace('<c:xVal>', `${customLabels}<c:xVal>`));
    expect(custom.definition?.groups[0]?.options?.dataLabels).toBeUndefined();

    const customImported = custom.definition!;
    await custom.replaceDefinition({
      groups: customImported.groups,
      options: { ...customImported.options, title: { text: 'Custom title' } },
    });
    expect(custom.xml).toContain(customLabels);

    const customTitled = custom.definition!;
    await custom.replaceDefinition({
      groups: [{
        ...customTitled.groups[0]!,
        options: { dataLabels: { position: 'top', showCategoryName: true } },
      }],
      options: customTitled.options,
    });
    expect(custom.xml).toContain(customLabels);
    expect(custom.xml).toContain('KEEP CUSTOM');
    expect(custom.xml.match(/<c:dLbls>/gu)).toHaveLength(2);
    expect(custom.definition?.groups[0]?.options?.dataLabels).toEqual({
      showCategoryName: true,
      position: 'top',
    });

    await custom.replaceSeries([{
      name: 'Custom edited', xValues: [10, 20], values: [30, 40],
    }]);
    expect(custom.xml).toContain(customLabels);
    expect(custom.xml).toContain('KEEP CUSTOM');
    expect(custom.definition?.groups[0]?.options?.dataLabels).toEqual({
      showCategoryName: true,
      position: 'top',
    });

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedCustom = reopened.slides[0]!.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel && shape.id === custom.id,
    )!;
    expect(reopenedCustom.xml).toContain(customLabels);
    expect(reopenedCustom.definition?.groups[0]?.options?.dataLabels).toEqual({
      showCategoryName: true,
      position: 'top',
    });
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

  it('edits only the selected indexed point style and preserves sibling bytes', async () => {
    const { slide } = emptyPresentation();
    const chart = await slide.addChart('pie', [{
      name: 'Share', categories: ['A', 'B', 'C'], values: [1, 2, 3],
    }]);
    const base = chart.definition!;
    const points = [0, 1, 2].map((index) => ({
      index,
      fill: {
        kind: 'solid' as const,
        color: { kind: 'srgb' as const, value: index === 1 ? '00AA00' : 'AA0000' },
      },
      shadow: { kind: 'outer' as const },
    }));
    await chart.replaceDefinition({ groups: [{
      type: 'pie',
      series: base.groups[0]!.series,
      options: { series: [{ points }] },
    }] });
    const fragments = (xml: string) => xml.match(/<c:dPt>[\s\S]*?<\/c:dPt>/gu) ?? [];
    const before = fragments(chart.xml);
    const current = chart.definition!;
    const currentOptions = current.groups[0]!.options!.series![0]!;
    await chart.replaceDefinition({
      groups: [{
        type: 'pie',
        series: current.groups[0]!.series,
        options: { series: [{
          ...currentOptions,
          points: currentOptions.points!.map((point) => point.index === 1
            ? {
                ...point,
                fill: { kind: 'solid', color: { kind: 'srgb', value: 'ABCDEF' } },
              }
            : point),
        }] },
      }],
      options: current.options,
    });
    const after = fragments(chart.xml);

    expect(after).toHaveLength(3);
    expect(after[0]).toBe(before[0]);
    expect(after[1]).not.toBe(before[1]);
    expect(after[2]).toBe(before[2]);
    expect(chart.definition?.groups[0]?.options?.series?.[0]?.points?.[1]?.fill)
      .toMatchObject({ color: { value: 'ABCDEF' } });
  });

  it('edits only the selected scatter point label and preserves sibling identities', async () => {
    const { slide } = emptyPresentation();
    const chart = await slide.addChart('scatter', [{
      name: 'XY', xValues: [1, 2, 3], values: [10, 20, 30],
    }]);
    const base = chart.definition!;
    await chart.replaceDefinition({ groups: [{
      type: 'scatter',
      series: base.groups[0]!.series,
      options: { series: [{ dataLabels: { pointLabels: [
        { index: 0, text: 'Alpha', fields: ['xValue', 'yValue'] },
        { index: 1, text: 'Beta' },
        { index: 2, fields: ['xValue', 'yValue'] },
      ] } }] },
    }] });
    const fragments = (xml: string) => xml.match(/<c:dLbl>[\s\S]*?<\/c:dLbl>/gu) ?? [];
    const before = fragments(chart.xml);
    const current = chart.definition!;
    const group = current.groups[0]!;
    if (group.type !== 'scatter') throw new Error('Expected scatter chart');
    const seriesOptions = group.options!.series![0]!;
    await chart.replaceDefinition({
      groups: [{
        type: 'scatter',
        ...(group.axis === undefined ? {} : { axis: group.axis }),
        series: group.series,
        options: {
          ...group.options,
          series: [{
            ...seriesOptions,
            dataLabels: {
              ...seriesOptions.dataLabels,
              pointLabels: seriesOptions.dataLabels!.pointLabels!.map((label) =>
                label.index === 1 ? { ...label, text: 'Beta edited' } : label),
            },
          }],
        },
      }],
      options: current.options,
    });
    const after = fragments(chart.xml);

    expect(after[0]).toBe(before[0]);
    expect(after[1]).not.toBe(before[1]);
    expect(after[2]).toBe(before[2]);
    expect(chart.definition?.groups[0]?.options?.series?.[0]?.dataLabels?.pointLabels?.[1])
      .toMatchObject({ index: 1, text: 'Beta edited' });
  });

  it('patches chart line caps without rewriting join, arrow, effect, or extension payload', async () => {
    const { pkg, slide } = emptyPresentation();
    const chart = await slide.addChart('line', [{
      name: 'Trend', categories: ['A', 'B'], values: [1, 2],
    }]);
    const base = chart.definition!;
    const line = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 1,
      dash: 'solid' as const,
      cap: 'round' as const,
    };
    await chart.replaceDefinition({
      groups: [{
        ...base.groups[0]!,
        options: { series: [{ line }] },
      }],
      options: {
        categoryAxis: { line },
        valueAxis: { majorGridLine: line, minorGridLine: line },
      },
    });
    updateChartXml(pkg, chart, (xml) => {
      let index = 0;
      return xml.replace(/<a:ln\b[^>]*cap="rnd"[^>]*>[\s\S]*?<\/a:ln>/gu, (fragment) => {
        index += 1;
        if (index > 4) return fragment;
        return fragment
          .replace('cap="rnd"', 'cap="rnd" cmpd="sng" algn="ctr"')
          .replace('</a:ln>', '<a:round/><a:headEnd type="triangle"/>'
            + `<a:extLst><a:ext uri="urn:keep-${index}"/></a:extLst></a:ln>`);
      });
    });
    expect(chart.definition).toBeDefined();

    const imported = chart.definition!;
    const group = imported.groups[0]!;
    const importedSeriesLine = group.options?.series?.[0]?.line;
    const importedMajorGridLine = imported.options.valueAxis?.majorGridLine;
    const importedMinorGridLine = imported.options.valueAxis?.minorGridLine;
    const importedAxisLine = imported.options.categoryAxis?.line;
    if (
      importedSeriesLine?.kind !== 'line'
      || importedMajorGridLine?.kind !== 'line'
      || importedMinorGridLine?.kind !== 'line'
      || importedAxisLine?.kind !== 'line'
    ) throw new Error('Expected imported line owners');
    await chart.replaceDefinition({
      groups: [{
        ...group,
        options: { series: [{
          ...group.options!.series![0]!,
          line: { ...importedSeriesLine, cap: 'square' },
        }] },
      }],
      options: {
        ...imported.options,
        valueAxis: {
          ...imported.options.valueAxis,
          majorGridLine: {
            ...importedMajorGridLine,
            cap: 'square',
          },
          minorGridLine: {
            ...importedMinorGridLine,
            cap: 'square',
          },
        },
        categoryAxis: {
          ...imported.options.categoryAxis,
          line: { ...importedAxisLine, cap: 'square' },
        },
      },
    });
    expect(chart.xml.match(/cmpd="sng" algn="ctr"/gu)).toHaveLength(4);
    expect(chart.xml.match(/cap="sq" cmpd="sng" algn="ctr"/gu)).toHaveLength(4);
    expect(chart.xml.match(/<a:headEnd type="triangle"\/>/gu)).toHaveLength(4);
    expect(chart.xml).toContain('uri="urn:keep-1"');
    expect(chart.xml).toContain('uri="urn:keep-2"');
    expect(chart.xml).toContain('uri="urn:keep-3"');
    expect(chart.xml).toContain('uri="urn:keep-4"');

    const squared = chart.definition!;
    const squaredGroup = squared.groups[0]!;
    const squaredSeriesLine = squaredGroup.options?.series?.[0]?.line;
    const squaredMajorGridLine = squared.options.valueAxis?.majorGridLine;
    const squaredMinorGridLine = squared.options.valueAxis?.minorGridLine;
    const squaredAxisLine = squared.options.categoryAxis?.line;
    if (
      squaredSeriesLine?.kind !== 'line'
      || squaredMajorGridLine?.kind !== 'line'
      || squaredMinorGridLine?.kind !== 'line'
      || squaredAxisLine?.kind !== 'line'
    ) throw new Error('Expected squared line owners');
    const { cap: _seriesCap, ...seriesLine } = squaredSeriesLine;
    const { cap: _gridCap, ...gridLine } = squaredMajorGridLine;
    const { cap: _minorGridCap, ...minorGridLine } = squaredMinorGridLine;
    const { cap: _axisCap, ...axisLine } = squaredAxisLine;
    await chart.replaceDefinition({
      groups: [{
        ...squaredGroup,
        options: { series: [{ ...squaredGroup.options!.series![0]!, line: seriesLine }] },
      }],
      options: {
        ...squared.options,
        valueAxis: {
          ...squared.options.valueAxis,
          majorGridLine: gridLine,
          minorGridLine,
        },
        categoryAxis: { ...squared.options.categoryAxis, line: axisLine },
      },
    });
    expect(chart.xml.match(/cmpd="sng" algn="ctr"/gu)).toHaveLength(4);
    expect(chart.xml).not.toMatch(/cap="sq" cmpd="sng" algn="ctr"/u);
    expect(chart.xml.match(/<a:headEnd type="triangle"\/>/gu)).toHaveLength(4);
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

function pptxGenJsSeriesLabels(): string {
  const points = [0, 1].map((index) => '<c:dLbl>'
    + `<c:idx val="${index}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:spPr/>`
    + '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
    + '<a:defRPr sz="1200" b="0" i="0" u="none" strike="noStrike">'
    + '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
    + '<a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
    + '<c:dLblPos val="bestFit"/><c:showLegendKey val="0"/><c:showVal val="0"/>'
    + '<c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="0"/>'
    + '<c:showBubbleSize val="0"/></c:dLbl>').join('');
  return '<c:dLbls>' + points
    + '<c:numFmt formatCode="General" sourceLinked="0"/>'
    + '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
    + '<a:defRPr sz="1800" b="0" i="0" u="none" strike="noStrike">'
    + '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
    + '<a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>'
    + '<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/>'
    + '<c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/>'
    + '<c:showBubbleSize val="0"/><c:showLeaderLines val="0"/></c:dLbls>';
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
