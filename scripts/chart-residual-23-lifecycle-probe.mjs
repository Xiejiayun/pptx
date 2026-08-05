const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

export const CHART_RESIDUAL_23_ATOMS = Object.freeze([
  ['class:PptxGenJS@property:ChartType', 'deliberate-difference'],
  ['inline:interface:IChartOpts@property:titlePos@property:titlePos.x', 'deliberate-difference'],
  ['inline:interface:IChartOpts@property:titlePos@property:titlePos.y', 'deliberate-difference'],
  ['interface:IChartOpts@property:bar3DShape', 'deliberate-difference'],
  ['interface:IChartOpts@property:dataBorder', 'deliberate-difference'],
  ['interface:IChartOpts@property:dataLabelBkgrdColors', 'deliberate-difference'],
  ['interface:IChartOpts@property:dataLabelFormatScatter', 'deliberate-difference'],
  ['interface:IChartOpts@property:dataNoEffects', 'deliberate-difference'],
  ['interface:IChartOpts@property:invertedColors', 'deliberate-difference'],
  ['interface:IChartOpts@property:lang', 'deliberate-difference'],
  ['interface:IChartOpts@property:layout', 'deliberate-difference'],
  ['interface:IChartOpts@property:lineCap', 'deliberate-difference'],
  ['interface:IChartOpts@property:shadow', 'deliberate-difference'],
  ['interface:IChartOpts@property:titleAlign', 'deliberate-difference'],
  ['interface:OptsChartGridLine@property:cap', 'deliberate-difference'],
  ['union:ChartLineCap#flat', 'supported'],
  ['union:ChartLineCap#round', 'supported'],
  ['union:ChartLineCap#square', 'supported'],
  ['union:ChartType#bubble3D', 'supported'],
  ['union:interface:IChartOpts@property:dataLabelFormatScatter#XY', 'deliberate-difference'],
  ['union:interface:IChartOpts@property:dataLabelFormatScatter#custom', 'deliberate-difference'],
  ['union:interface:IChartOpts@property:dataLabelFormatScatter#customXY', 'deliberate-difference'],
  ['union:interface:OptsChartData@property:labels#string[][]', 'deliberate-difference'],
].map(([id, status]) => Object.freeze({ id, status })));

const FORMATS = ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'];
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function packageSnapshot(document) {
  return JSON.stringify({
    parts: document.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: [...bytes],
      relationships,
    })),
    mutations: [...document.opcPackage.mutations],
  });
}

function charts(api, document) {
  return document.slides.flatMap(({ shapes }) => shapes)
    .filter((shape) => shape instanceof api.ChartModel);
}

function chartOn(api, document, slideIndex) {
  const chart = document.slides[slideIndex]?.shapes.find(
    (shape) => shape instanceof api.ChartModel,
  );
  if (!chart) throw new Error(`Missing chart on slide ${slideIndex + 1}`);
  return chart;
}

function seriesOptions(chart, seriesIndex = 0) {
  return chart.definition?.groups[0]?.options?.series?.[seriesIndex];
}

function pointLabels(chart) {
  return seriesOptions(chart)?.dataLabels?.pointLabels;
}

function pointFragments(xml, tag) {
  return xml.match(new RegExp(`<c:${tag}\\b[^>]*>[\\s\\S]*?<\\/c:${tag}>`, 'gu')) ?? [];
}

function packageGraph(document) {
  const owned = document.opcPackage.parts.filter(({ contentType }) =>
    contentType === CHART_CONTENT_TYPE || contentType === WORKBOOK_CONTENT_TYPE);
  const ownership = owned.map(({ uri }) => ({
    uri,
    incoming: document.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0,
  }));
  return {
    slides: document.slides.length,
    charts: document.opcPackage.parts.filter(
      ({ contentType }) => contentType === CHART_CONTENT_TYPE,
    ).length,
    workbooks: document.opcPackage.parts.filter(
      ({ contentType }) => contentType === WORKBOOK_CONTENT_TYPE,
    ).length,
    ownership,
    orphanCount: ownership.filter(({ incoming }) => incoming === 0).length,
  };
}

async function addChart(api, document, group, options = undefined) {
  const slideIndex = document.slides.length;
  document.addSlide();
  const chart = await document.addChart(slideIndex, [group]);
  if (options !== undefined) {
    await chart.replaceDefinition({ groups: chart.definition.groups, options });
  }
  return chart;
}

function pointLabelDefinition(chart, labels) {
  const definition = chart.definition;
  const group = definition?.groups[0];
  if (!definition || group?.type !== 'scatter') throw new Error('Expected scatter chart');
  const currentSeries = group.options?.series?.[0] ?? {};
  return {
    groups: [{
      type: 'scatter',
      ...(group.axis === undefined ? {} : { axis: group.axis }),
      series: group.series,
      options: {
        ...group.options,
        series: [{
          ...currentSeries,
          dataLabels: { ...currentSeries.dataLabels, pointLabels: labels },
        }],
      },
    }],
    options: definition.options,
  };
}

function allTrue(value) {
  return Object.values(value).every(Boolean);
}

export async function runChartResidual23LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function'
      || typeof api.ChartModel !== 'function'
      || typeof api.chartWorkbookMatches !== 'function') {
    throw new TypeError(
      'Chart Residual 23 probe requires PptxDocument, ChartModel, and chartWorkbookMatches',
    );
  }

  const document = api.PptxDocument.create();
  const bar3D = await addChart(api, document, {
    type: 'bar3D',
    series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
    options: { shape: 'coneToMax' },
  }, {
    language: 'fr-FR',
    layout: { x: 0.1, y: 0.2, width: 0.7, height: 0.6 },
    title: {
      text: 'Revenue',
      position: { x: 0.15, y: 0.1 },
      align: 'right',
    },
  });
  const bubble3D = await addChart(api, document, {
    type: 'bubble3D',
    series: [{
      name: 'Depth',
      xValues: [1, 2],
      values: [10, 20],
      sizes: [3, 6],
    }],
  });
  const nested = await addChart(api, document, {
    type: 'bar',
    series: [{
      name: 'Nested',
      categories: [['FY26', 'FY26'], ['Q1', 'Q2']],
      values: [1, 2],
    }],
  });
  const line = await addChart(api, document, {
    type: 'line',
    series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [1, 2] }],
    options: { series: [{
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        cap: 'flat',
      },
      marker: { line: {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        cap: 'round',
      } },
    }] },
  }, {
    categoryAxis: {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '778899' },
        cap: 'square',
      },
      majorGridLine: {
        kind: 'line',
        color: { kind: 'srgb', value: 'AABBCC' },
        cap: 'round',
      },
    },
    valueAxis: { minorGridLine: {
      kind: 'line',
      color: { kind: 'srgb', value: 'DDEEFF' },
      cap: 'square',
    } },
  });
  const pointEffects = await addChart(api, document, {
    type: 'pie',
    series: [{ name: 'Share', categories: ['A', 'B', 'C'], values: [1, 2, 3] }],
    options: { series: [{
      shadow: { kind: 'outer', color: { kind: 'srgb', value: '778899' } },
      points: [{
        index: 1,
        fill: { kind: 'solid', color: { kind: 'srgb', value: '00AA00' } },
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '123456' },
          width: 2,
        },
        shadow: { kind: 'inner', color: { kind: 'srgb', value: '445566' } },
      }],
      dataLabels: {
        fill: { kind: 'solid', color: { kind: 'srgb', value: '112233' } },
      },
    }] },
  });
  const noEffects = await addChart(api, document, {
    type: 'pie',
    series: [{ name: 'No effects', categories: ['A', 'B'], values: [1, 2] }],
    options: { series: [{ points: [{ index: 0, shadow: { kind: 'outer' } }] }] },
  });
  const customLabels = [
    { index: 0, text: 'Alpha' },
    { index: 1, text: ' ' },
    { index: 2, text: 'Gamma' },
  ];
  const custom = await addChart(api, document, {
    type: 'scatter',
    series: [{ name: 'Custom', xValues: [1, 2, 3], values: [10, 20, 30] }],
    options: { series: [{ dataLabels: { pointLabels: customLabels } }] },
  });
  const customXYLabels = [
    { index: 0, text: 'Alpha', fields: ['xValue', 'yValue'] },
    { index: 1, text: ' ' },
    { index: 2, text: 'Gamma', fields: ['xValue', 'yValue'] },
  ];
  const customXY = await addChart(api, document, {
    type: 'scatter',
    series: [{ name: 'Custom XY', xValues: [1, 2, 3], values: [10, 20, 30] }],
    options: { series: [{ dataLabels: { pointLabels: customXYLabels } }] },
  });
  const xy = await addChart(api, document, {
    type: 'scatter',
    series: [{ name: 'XY', xValues: [1, 2, 3], values: [10, 20, 30] }],
    options: { dataLabels: { showValue: true, showCategoryName: true, position: 'top' } },
  });

  customLabels[0].text = 'CALLER MUTATED';
  customXYLabels[0].fields.push('xValue');
  const callerDetached = pointLabels(custom)?.[0]?.text === 'Alpha'
    && JSON.stringify(pointLabels(customXY)?.[0]?.fields) === JSON.stringify(['xValue', 'yValue'])
    && Object.isFrozen(pointLabels(customXY))
    && pointLabels(customXY).every((label) => Object.isFrozen(label))
    && pointLabels(customXY).filter(({ fields }) => fields !== undefined)
      .every(({ fields }) => Object.isFrozen(fields));

  const beforeNoOp = packageSnapshot(document);
  await customXY.replaceDefinition(customXY.definition);
  const noOp = packageSnapshot(document) === beforeNoOp;

  const beforeInvalid = packageSnapshot(document);
  let invalidRejected = false;
  try {
    await customXY.replaceDefinition(pointLabelDefinition(customXY, [
      { index: 0, text: 'Duplicate A' },
      { index: 0, text: 'Duplicate B' },
    ]));
  } catch {
    invalidRejected = packageSnapshot(document) === beforeInvalid;
  }

  const beforeRollback = packageSnapshot(document);
  const beforeRollbackSlides = document.slides.length;
  let rollbackMessage;
  try {
    document.transaction(() => {
      document.addSlide();
      throw new Error('restore chart residual package');
    });
  } catch (error) {
    rollbackMessage = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackMessage === 'restore chart residual package'
    && document.slides.length === beforeRollbackSlides
    && packageSnapshot(document) === beforeRollback;

  const duplicateSlide = document.duplicateSlide(7);
  const duplicateChart = duplicateSlide.shapes.find((shape) => shape instanceof api.ChartModel);
  if (!duplicateChart) throw new Error('Missing duplicated Custom XY chart');
  const duplicateBefore = JSON.stringify(pointLabels(duplicateChart));
  const pointBefore = pointFragments(customXY.xml, 'dLbl');
  await customXY.replaceDefinition(pointLabelDefinition(
    customXY,
    pointLabels(customXY).map((label) =>
      label.index === 1 ? { ...label, text: 'Beta edited' } : label),
  ));
  const pointAfter = pointFragments(customXY.xml, 'dLbl');
  const selectedPointIsolation = pointBefore[0] === pointAfter[0]
    && pointBefore[1] !== pointAfter[1]
    && pointBefore[2] === pointAfter[2];
  const copyOnWrite = JSON.stringify(pointLabels(duplicateChart)) === duplicateBefore
    && pointLabels(duplicateChart)?.[1]?.text === ' ';

  const unrelatedBefore = pointFragments(customXY.xml, 'dLbl');
  await customXY.replaceDefinition({
    groups: customXY.definition.groups,
    options: { ...customXY.definition.options, title: { text: 'Unrelated title' } },
  });
  const byteIsolation = JSON.stringify(pointFragments(customXY.xml, 'dLbl'))
    === JSON.stringify(unrelatedBefore);

  const initialState = {
    catalog: Array.isArray(api.CHART_TYPES) && api.CHART_TYPES.includes('bubble3D'),
    titlePosition: bar3D.definition.options.title?.position?.x === 0.15
      && bar3D.definition.options.title?.position?.y === 0.1,
    bar3DShapes: bar3D.definition.groups[0]?.options?.shape === 'coneToMax',
    effectiveLanguage: bar3D.definition.options.language === 'fr-FR',
    plotLayout: bar3D.definition.options.layout?.width === 0.7,
    titleAlignment: bar3D.definition.options.title?.align === 'right',
    bubble3D: bubble3D.definition.groups[0]?.type === 'bubble3D',
    nestedCategories: Array.isArray(nested.definition.groups[0]?.series[0]?.categories?.[0]),
    seriesLineCap: seriesOptions(line)?.line?.cap === 'flat',
    axisLineCap: line.definition.options.categoryAxis?.line?.cap === 'square',
    gridLineCap: line.definition.options.categoryAxis?.majorGridLine?.cap === 'round'
      && line.definition.options.valueAxis?.minorGridLine?.cap === 'square',
    lineCapTokens: ['flat', 'round', 'square'].every((cap) =>
      line.xml.includes(`cap="${cap === 'round' ? 'rnd' : cap === 'square' ? 'sq' : cap}"`)),
    pointBorders: seriesOptions(pointEffects)?.points?.[0]?.line?.color?.value === '123456',
    dataLabelFills: seriesOptions(pointEffects)?.dataLabels?.fill?.color?.value === '112233',
    dataNoEffects: seriesOptions(noEffects)?.shadow === undefined
      && seriesOptions(noEffects)?.points?.[0]?.shadow?.kind === 'outer',
    invertedPointFills: seriesOptions(pointEffects)?.points?.[0]?.fill?.color?.value === '00AA00',
    seriesShadow: seriesOptions(pointEffects)?.shadow?.kind === 'outer',
    scatterXY: xy.definition.groups[0]?.options?.dataLabels?.showValue === true
      && xy.definition.groups[0]?.options?.dataLabels?.showCategoryName === true,
    scatterCustom: JSON.stringify(pointLabels(custom)?.map(({ index, text, fields }) => ({
      index, text: text ?? null, fields: fields ?? [],
    }))) === JSON.stringify([
      { index: 0, text: 'Alpha', fields: [] },
      { index: 1, text: ' ', fields: [] },
      { index: 2, text: 'Gamma', fields: [] },
    ]),
    scatterCustomXY: pointLabels(customXY)?.[0]?.fields?.join(',') === 'xValue,yValue'
      && pointLabels(customXY)?.[2]?.fields?.join(',') === 'xValue,yValue',
  };

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedCharts = charts(api, reopened);
  const reopenedCustomXY = chartOn(api, reopened, 7);
  const graph = packageGraph(reopened);
  const workbookMatches = await Promise.all(reopenedCharts.map((chart) =>
    api.chartWorkbookMatches(
      reopened.opcPackage.requirePart(chart.workbookPartUri).bytes,
      chart.definition,
      chart.xml,
    )));

  const formats = [];
  for (const format of FORMATS) {
    const formatted = api.PptxDocument.create({ format });
    formatted.addSlide();
    await formatted.addChart(0, [{
      type: 'bubble3D',
      series: [{ name: format, xValues: [1], values: [2], sizes: [3] }],
    }]);
    const formattedReopened = await api.PptxDocument.open(await formatted.write());
    formats.push({
      format,
      reopenedFormat: formattedReopened.format,
      bubble3D: chartOn(api, formattedReopened, 0).definition?.groups[0]?.type === 'bubble3D',
      errors: formattedReopened.diagnostics.filter(({ severity }) => severity === 'error').length,
    });
  }
  const allFormats = formats.every(({ format, reopenedFormat, bubble3D: retained, errors }) =>
    reopenedFormat === format && retained && errors === 0);

  const exactOoxml = {
    bar3D: bar3D.xml.includes('<c:shape val="coneToMax"/>')
      && bar3D.xml.includes('<c:layoutTarget val="inner"/>')
      && bar3D.xml.includes('<c:xMode val="edge"/><c:yMode val="edge"/>')
      && bar3D.xml.includes('<c:wMode val="factor"/><c:hMode val="factor"/>')
      && bar3D.xml.includes('<a:pPr algn="r"/>')
      && bar3D.xml.includes('<c:lang val="fr-FR"/>'),
    bubble3D: bubble3D.xml.includes('<c:bubbleChart>')
      && bubble3D.xml.includes('<c:bubble3D val="1"/>')
      && !bubble3D.xml.includes('<c:bubble3DChart>'),
    nestedCategories: nested.xml.includes('<c:multiLvlStrRef>')
      && nested.xml.includes('<c:multiLvlStrCache>'),
    lineCaps: line.xml.includes('<a:ln w="12700" cap="flat">')
      && line.xml.includes('cap="rnd"')
      && line.xml.includes('cap="sq"'),
    pointEffects: pointEffects.xml.includes('<c:dPt><c:idx val="1"/><c:spPr>')
      && pointEffects.xml.includes('<a:innerShdw')
      && pointEffects.xml.includes('<a:outerShdw')
      && pointEffects.xml.includes('<c:dLbls><c:spPr><a:solidFill>'),
    scatterFields: (customXY.xml.match(/type="XVALUE"/gu) ?? []).length === 2
      && (customXY.xml.match(/type="YVALUE"/gu) ?? []).length === 2
      && (customXY.xml.match(/<c16:uniqueId /gu) ?? []).length === 3,
  };
  const diagnostics = {
    createdErrors: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    createdWarnings: document.diagnostics.filter(({ severity }) => severity === 'warning').length,
    reopenedErrors: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopenedWarnings: reopened.diagnostics.filter(({ severity }) => severity === 'warning').length,
  };
  const lifecycle = {
    callerDetached,
    noOp,
    invalidIsolation: invalidRejected,
    rollback,
    copyOnWrite,
    selectedPointIsolation,
    reopen: pointLabels(reopenedCustomXY)?.[1]?.text === 'Beta edited',
    allFormats,
    packageGraph: graph.orphanCount === 0
      && graph.charts === reopenedCharts.length
      && graph.workbooks === reopenedCharts.length
      && graph.ownership.every(({ incoming }) => incoming === 1),
    workbookCacheMatch: workbookMatches.every(Boolean),
    byteIsolation,
    exactOoxml: allTrue(exactOoxml),
    diagnostics: Object.values(diagnostics).every((count) => count === 0),
  };
  const state = { ...initialState, ...lifecycle };
  const classification = {
    supported: CHART_RESIDUAL_23_ATOMS.filter(({ status }) => status === 'supported').length,
    'deliberate-difference': CHART_RESIDUAL_23_ATOMS.filter(
      ({ status }) => status === 'deliberate-difference',
    ).length,
    'defect-excluded': 0,
  };
  return {
    ok: CHART_RESIDUAL_23_ATOMS.length === 23
      && classification.supported === 4
      && classification['deliberate-difference'] === 19
      && allTrue(state),
    family: 'chart-residual-23',
    atomCount: CHART_RESIDUAL_23_ATOMS.length,
    classification,
    atoms: CHART_RESIDUAL_23_ATOMS,
    state,
    details: {
      formats,
      graph,
      workbookMatches,
      exactOoxml,
      diagnostics,
    },
    explicitOutputBytes,
    mime: typeof Blob === 'function'
      ? new Blob([explicitOutputBytes], { type: PPTX_MIME }).type
      : PPTX_MIME,
  };
}
