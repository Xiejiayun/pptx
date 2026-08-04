const chartContentType =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const workbookContentType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const expectedReadback = [
  {
    face: 'Arial',
    size: 12,
    color: { kind: 'srgb', value: '000000' },
    position: 'bestFit',
    showCategoryName: true,
    numberFormat: 'General',
  },
  {
    face: 'Arial',
    size: 12,
    color: { kind: 'srgb', value: '000000' },
    position: 'bestFit',
    showPercent: true,
    numberFormat: '0%',
  },
  {
    face: 'Arial',
    size: 12,
    color: { kind: 'srgb', value: '000000' },
    position: 'bestFit',
    numberFormat: 'General',
  },
  {
    face: 'Arial',
    size: 12,
    color: { kind: 'srgb', value: '000000' },
    showCategoryName: true,
    numberFormat: 'General',
  },
  {
    face: 'Arial',
    size: 12,
    color: { kind: 'srgb', value: '000000' },
    position: 'top',
    numberFormat: 'General',
  },
];

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(Object.entries(value)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, child]) => [key, stable(child)]));
}

function equal(left, right) {
  return JSON.stringify(stable(left)) === JSON.stringify(stable(right));
}

function seriesLabelFragment(xml) {
  return xml.match(
    /<c:ser>[\s\S]*?(<c:dLbls>[\s\S]*?<\/c:dLbls>)[\s\S]*?<\/c:ser>/u,
  )?.[1] ?? null;
}

function labelContainerCount(xml) {
  return xml.match(/<c:dLbls>/gu)?.length ?? 0;
}

function chartOnSlide(api, document, index) {
  const chart = document.slides[index]?.shapes.find(
    (shape) => shape instanceof api.ChartModel,
  );
  if (!chart) throw new Error(`Missing chart on slide ${index + 1}`);
  return chart;
}

function charts(api, document) {
  return document.slides.flatMap(({ shapes }) => shapes)
    .filter((shape) => shape instanceof api.ChartModel);
}

async function writeAndReopen(api, document) {
  const bytes = await document.write({
    outputType: 'uint8array',
    compatibility: 'powerpoint-2010',
  });
  return {
    bytes,
    document: await api.PptxDocument.open(bytes),
  };
}

function graphState(api, document) {
  const allCharts = charts(api, document);
  const chartParts = document.opcPackage.parts.filter(
    ({ contentType }) => contentType === chartContentType,
  );
  const workbookParts = document.opcPackage.parts.filter(
    ({ contentType }) => contentType === workbookContentType,
  );
  const ownership = [...chartParts, ...workbookParts].map(({ uri }) => ({
    uri,
    incoming: document.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0,
  }));
  return {
    slides: document.slides.length,
    charts: allCharts.length,
    chartParts: chartParts.length,
    workbooks: workbookParts.length,
    ownership,
    orphanCount: ownership.filter(({ incoming }) => incoming === 0).length,
  };
}

export async function runChartPresentation91LifecycleProbe(api, fixtureBytes) {
  const source = await api.PptxDocument.open(fixtureBytes);
  const sourceCharts = charts(api, source);
  const sourceGraph = graphState(api, source);
  const actualReadback = sourceCharts.map(
    (chart) => chart.definition?.groups[0]?.options?.dataLabels ?? null,
  );
  const sourceSeriesLabels = sourceCharts.map(({ xml }) => seriesLabelFragment(xml));

  const fixtureShape = sourceGraph.slides === 5
    && sourceGraph.charts === 5
    && sourceGraph.chartParts === 5
    && sourceGraph.workbooks === 5;
  const readback = equal(actualReadback, expectedReadback);
  const triStateOverrides = actualReadback[0]?.showPercent === undefined
    && actualReadback[1]?.showCategoryName === undefined
    && actualReadback[2]?.showCategoryName === undefined
    && actualReadback[2]?.showPercent === undefined;

  for (let index = 0; index < sourceCharts.length; index += 1) {
    const chart = sourceCharts[index];
    const definition = chart.definition;
    if (!definition) throw new Error(`Missing definition for chart ${index + 1}`);
    await chart.replaceDefinition({
      groups: definition.groups,
      options: {
        ...definition.options,
        title: { text: `Unrelated ${index + 1}` },
      },
    });
  }
  const unrelatedOutput = await writeAndReopen(api, source);
  const unrelatedCharts = charts(api, unrelatedOutput.document);
  const unrelatedFragments = unrelatedCharts.map(({ xml }) => seriesLabelFragment(xml));
  const unrelatedEditPreserved = equal(unrelatedFragments, sourceSeriesLabels);

  for (let index = 0; index < 4; index += 1) {
    const chart = chartOnSlide(api, unrelatedOutput.document, index);
    const definition = chart.definition;
    if (!definition) throw new Error(`Missing definition for chart ${index + 1}`);
    await chart.replaceDefinition({
      groups: [{
        ...definition.groups[0],
        options: { dataLabels: { position: 'center', showPercent: true } },
      }],
      options: definition.options,
    });
  }
  const scatter = chartOnSlide(api, unrelatedOutput.document, 4);
  const scatterDefinition = scatter.definition;
  if (!scatterDefinition) throw new Error('Missing scatter definition');
  await scatter.replaceDefinition({
    groups: [{
      ...scatterDefinition.groups[0],
      options: { dataLabels: { position: 'center', showCategoryName: true } },
    }],
    options: scatterDefinition.options,
  });

  const explicitOutput = await writeAndReopen(api, unrelatedOutput.document);
  const explicitCharts = charts(api, explicitOutput.document);
  const explicitSafeState = explicitCharts.slice(0, 4).map((chart) => ({
    labelContainerCount: labelContainerCount(chart.xml),
    hasSeriesLabels: seriesLabelFragment(chart.xml) !== null,
    dataLabels: chart.definition?.groups[0]?.options?.dataLabels ?? null,
  }));
  const explicitSafeEdit = explicitSafeState.every(({ labelContainerCount: count, hasSeriesLabels, dataLabels }) =>
    count === 1
      && !hasSeriesLabels
      && dataLabels?.position === 'center'
      && dataLabels?.showPercent === true);

  const explicitScatter = explicitCharts[4];
  if (!explicitScatter) throw new Error('Missing reopened scatter chart');
  const scatterSeriesLabels = seriesLabelFragment(explicitScatter.xml) ?? '';
  const customScatterState = {
    labelContainerCount: labelContainerCount(explicitScatter.xml),
    seriesFragmentByteIdentical: scatterSeriesLabels === sourceSeriesLabels[4],
    customTextCount: scatterSeriesLabels.match(/<c:tx>/gu)?.length ?? 0,
    shapePropertiesCount: scatterSeriesLabels.match(/<c:spPr>/gu)?.length ?? 0,
    extensionListCount: scatterSeriesLabels.match(/<c:extLst>/gu)?.length ?? 0,
    customStrings: ['A', 'B', 'C'].filter(
      (value) => scatterSeriesLabels.includes(`<a:t>${value}</a:t>`),
    ),
    dataLabels: explicitScatter.definition?.groups[0]?.options?.dataLabels ?? null,
  };
  const customScatterPreserved = customScatterState.labelContainerCount === 2
    && customScatterState.seriesFragmentByteIdentical
    && customScatterState.customTextCount === 3
    && customScatterState.shapePropertiesCount >= 3
    && customScatterState.extensionListCount >= 3
    && customScatterState.customStrings.length === 3
    && customScatterState.dataLabels?.position === 'center'
    && customScatterState.dataLabels?.showCategoryName === true;

  const finalGraph = graphState(api, explicitOutput.document);
  const workbookMatches = await Promise.all(explicitCharts.map((chart) =>
    api.chartWorkbookMatches(
      explicitOutput.document.opcPackage.requirePart(chart.workbookPartUri).bytes,
      chart.definition,
      chart.xml,
    )));
  const diagnosticErrors = explicitOutput.document.diagnostics
    .filter(({ severity }) => severity === 'error');
  const packageLifecycle = finalGraph.slides === 5
    && finalGraph.charts === 5
    && finalGraph.chartParts === 5
    && finalGraph.workbooks === 5
    && finalGraph.ownership.every(({ incoming }) => incoming === 1)
    && finalGraph.orphanCount === 0
    && workbookMatches.every(Boolean)
    && diagnosticErrors.length === 0;

  const state = {
    fixtureShape,
    readback,
    triStateOverrides,
    unrelatedEditPreserved,
    explicitSafeEdit,
    customScatterPreserved,
    packageLifecycle,
  };
  return {
    ok: Object.values(state).every(Boolean),
    state,
    actualReadback,
    expectedReadback,
    sourceGraph,
    explicitSafeState,
    customScatterState,
    finalGraph,
    workbookMatches,
    diagnosticErrors,
    unrelatedOutputBytes: unrelatedOutput.bytes,
    explicitOutputBytes: explicitOutput.bytes,
  };
}
