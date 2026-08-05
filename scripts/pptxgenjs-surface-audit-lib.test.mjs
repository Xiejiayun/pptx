import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

import { buildPptxGenJSAudit } from './pptxgenjs-surface-audit-lib.mjs';
import { PPTXGENJS_SURFACE_MANIFEST } from './pptxgenjs-surface-manifest.mjs';

const IDS = Object.freeze([
  'class:Deck#write',
  'method:Slide#addText',
  'union:OutputType#blob',
  'interface:ImageProps@property:altText',
  'property:Slide#bkgd',
  'interface:ShapeProps@property:shadow',
]);

function atom(id) {
  return Object.freeze({
    id,
    kind: id.startsWith('union:') ? 'union-member' : 'property',
    owner: id.split(/[:#@]/u)[1] ?? 'Deck',
    name: id.split(/[:#@]/u).at(-1) ?? id,
    declaredIn: 'Fixture',
    optional: false,
    readonly: false,
    deprecated: false,
    signatures: [],
    deprecatedSignatures: [],
    catalogKey: '',
    typeText: 'unknown',
  });
}

function surface(ids = IDS.slice(0, 3)) {
  return Object.freeze({
    schemaVersion: 1,
    atoms: Object.freeze(ids.map(atom)),
    roots: Object.freeze({ presentation: 'Deck', slide: 'Slide' }),
    diagnostics: Object.freeze([]),
  });
}

const runtimeProbe = Object.freeze({
  schemaVersion: 1,
  packageVersion: '4.0.1',
  declarationSha256: 'a'.repeat(64),
  runtimeEntrySha256: 'b'.repeat(64),
  runtimeMismatches: Object.freeze([]),
});

function emptyEvidence() {
  return {
    code: [],
    tests: [],
    package: [],
    ooxml: [],
    clients: [],
  };
}

function manifest(entries = [], extensions = []) {
  return {
    schemaVersion: 1,
    packageVersion: '4.0.1',
    entries,
    extensions,
  };
}

function entry(id, status, overrides = {}) {
  return {
    id,
    status,
    native: [],
    evidence: emptyEvidence(),
    note: `${id} fixture`,
    ...overrides,
  };
}

function validEvidence() {
  return {
    code: [{ path: 'src/write.ts', pattern: 'export function write', commit: 'abc1234' }],
    tests: [{ path: 'tests/write.test.ts', title: 'writes blob', commit: 'abc1234' }],
    package: [{ path: 'scripts/smoke.mjs', pattern: 'packed write', commit: 'abc1234' }],
    ooxml: [{ path: 'docs/ooxml.md', pattern: 'ppt/slides/slide1.xml', commit: 'abc1234' }],
    clients: [{ path: 'docs/client.md', pattern: 'PowerPoint pass', commit: 'abc1234' }],
  };
}

const validControl = Object.freeze({
  path: 'controls/control.mjs',
  pattern: 'PptxGenJS control',
  commit: 'abc1234',
});

async function withRepository(run) {
  const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-audit-lib-'));
  try {
    for (const path of ['src', 'tests', 'scripts', 'docs', 'controls']) {
      await mkdir(join(directory, path), { recursive: true });
    }
    await writeFile(join(directory, 'src', 'write.ts'), 'export function write() {}\n');
    await writeFile(join(directory, 'tests', 'write.test.ts'), "test('writes blob', () => {});\n");
    await writeFile(join(directory, 'scripts', 'smoke.mjs'), 'packed write\n');
    await writeFile(join(directory, 'docs', 'ooxml.md'), 'ppt/slides/slide1.xml\n');
    await writeFile(join(directory, 'docs', 'client.md'), 'PowerPoint pass\n');
    await writeFile(join(directory, 'controls', 'control.mjs'), 'PptxGenJS control\n');
    return await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

function assertDeepFrozen(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  assert.equal(Object.isFrozen(value), true);
  for (const nested of Object.values(value)) assertDeepFrozen(nested, seen);
}

test('exports an immutable evidence-backed initial manifest batch', () => {
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.schemaVersion, 1);
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.packageVersion, '4.0.1');
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.entries.length, 1734);
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries.map(({ status }) => status).sort(),
    [
      ...Array(371).fill('defect-excluded'),
      ...Array(752).fill('supported'),
      ...Array(517).fill('deliberate-difference'),
      ...Array(94).fill('deprecated-alias'),
    ].sort(),
  );
  const coreContentPrimitiveInputsIds = new Set([
    'interface:TableCell@property:options',
    'interface:TableCell@property:text',
    'interface:TextProps@property:options',
    'interface:TextProps@property:text',
    'union:interface:TableCell@property:text#TableCell[]',
    'union:interface:TableCell@property:text#string',
    'union:method:Slide#addText@path:text#TextProps[]',
    'union:method:Slide#addText@path:text#string',
    'union:Color#HexColor',
    'union:Color#ThemeColor',
    'union:Coord#${number}%',
    'union:Coord#number',
    'union:Margin#[number,number,number,number]',
    'union:Margin#number',
  ]);
  const coreContentPrimitiveInputsEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => coreContentPrimitiveInputsIds.has(id));
  assert.equal(coreContentPrimitiveInputsEntries.length, 14);
  assert.deepEqual(
    coreContentPrimitiveInputsEntries.map(({ status }) => status).sort(),
    [
      ...Array(5).fill('supported'),
      ...Array(9).fill('deliberate-difference'),
    ].sort(),
  );
  const lineFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^(?:union:)?interface:(?:ShapeLineProps@property:(?:alpha|beginArrowType|color|dashType|endArrowType|lineDash|lineHead|lineTail|pt|size|transparency|type|width)|(?:ShapeProps|TextPropsOptions)@property:(?:line|lineDash|lineHead|lineSize|lineTail))(?:#.+)?$/u
      .test(id));
  assert.equal(lineFamilyEntries.length, 105);
  assert.deepEqual(
    lineFamilyEntries.map(({ status }) => status).sort(),
    [
      ...Array(31).fill('deliberate-difference'),
      ...Array(74).fill('deprecated-alias'),
    ].sort(),
  );
  const lineFamilyById = new Map(lineFamilyEntries.map((entry) => [entry.id, entry]));
  for (const entry of lineFamilyEntries.filter(
    ({ status }) => status === 'deprecated-alias',
  )) {
    assert.equal(lineFamilyById.get(entry.canonical)?.status, 'deliberate-difference');
  }
  assert.equal(
    lineFamilyById.get('union:interface:TextPropsOptions@property:lineHead#triangle')
      ?.canonical,
    'union:interface:ShapeLineProps@property:beginArrowType#triangle',
  );
  assert.equal(
    lineFamilyById.get('interface:ShapeProps@property:lineSize')?.canonical,
    'interface:ShapeLineProps@property:width',
  );
  const fillFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^(?:union:)?interface:(?:ShapeFillProps@property:(?:alpha|color|transparency|type)|(?:ShapeProps|TextPropsOptions)@property:fill)(?:#.+)?$/u.test(id));
  assert.equal(fillFamilyEntries.length, 8);
  assert.deepEqual(
    fillFamilyEntries.map(({ status }) => status).sort(),
    [
      ...Array(7).fill('deliberate-difference'),
      'deprecated-alias',
    ].sort(),
  );
  const fillFamilyById = new Map(fillFamilyEntries.map((entry) => [entry.id, entry]));
  assert.equal(
    fillFamilyById.get('interface:ShapeFillProps@property:alpha')?.canonical,
    'interface:ShapeFillProps@property:transparency',
  );
  const tableFillEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^interface:(?:TableCellProps|TableProps|TableToSlidesProps)@property:fill$/u.test(id));
  assert.deepEqual(
    tableFillEntries.map(({ id, status }) => ({ id, status })),
    [
      {
        id: 'interface:TableCellProps@property:fill',
        status: 'deliberate-difference',
      },
      {
        id: 'interface:TableProps@property:fill',
        status: 'deliberate-difference',
      },
      {
        id: 'interface:TableToSlidesProps@property:fill',
        status: 'defect-excluded',
      },
    ],
  );
  const chartAreaFillLineEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^interface:(?:IChartAreaProps@property:(?:border|fill|roundedCorners)|IChartPropsFillLine@property:(?:border|fill)|IChartOpts@property:(?:border|chartArea|fill|plotArea))$/u.test(id));
  assert.equal(chartAreaFillLineEntries.length, 9);
  assert.deepEqual(
    chartAreaFillLineEntries.map(({ status }) => status).sort(),
    [
      ...Array(7).fill('deliberate-difference'),
      ...Array(2).fill('deprecated-alias'),
    ].sort(),
  );
  const chartAreaFillLineById = new Map(
    chartAreaFillLineEntries.map((entry) => [entry.id, entry]),
  );
  assert.equal(
    chartAreaFillLineById.get('interface:IChartOpts@property:border')?.canonical,
    'interface:IChartPropsFillLine@property:border',
  );
  assert.equal(
    chartAreaFillLineById.get('interface:IChartOpts@property:fill')?.canonical,
    'interface:IChartPropsFillLine@property:fill',
  );
  const chartTypeValues = [
    'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
  ];
  const expectedChartCreationEntries = [
    ...['CHART_NAME', 'ChartType'].flatMap((owner) =>
      chartTypeValues.map((value) => ({
        id: `union:${owner}#${value}`,
        status: 'supported',
      }))),
    {
      id: 'union:method:Slide#addChart@path:type#CHART_NAME',
      status: 'supported',
    },
    { id: 'interface:IChartMulti@property:type', status: 'supported' },
    { id: 'interface:OptsChartData@property:name', status: 'supported' },
    { id: 'interface:OptsChartData@property:sizes', status: 'supported' },
    { id: 'interface:OptsChartData@property:values', status: 'supported' },
    { id: 'method:Slide#addChart', status: 'deliberate-difference' },
    {
      id: 'union:method:Slide#addChart@path:type#IChartMulti[]',
      status: 'deliberate-difference',
    },
    { id: 'interface:IChartMulti@property:data', status: 'deliberate-difference' },
    { id: 'interface:IChartMulti@property:options', status: 'deliberate-difference' },
    { id: 'interface:OptsChartData@property:labels', status: 'deliberate-difference' },
    {
      id: 'union:interface:OptsChartData@property:labels#string[]',
      status: 'deliberate-difference',
    },
  ].sort((left, right) => left.id.localeCompare(right.id));
  const chartCreationIds = new Set(expectedChartCreationEntries.map(({ id }) => id));
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => chartCreationIds.has(id))
      .map(({ id, status }) => ({ id, status }))
      .sort((left, right) => left.id.localeCompare(right.id)),
    expectedChartCreationEntries,
  );
  assert.equal(
    PPTXGENJS_SURFACE_MANIFEST.entries.some(({ id }) => [
      'union:ChartType#bubble3D',
      'class:PptxGenJS@property:ChartType',
      'union:interface:OptsChartData@property:labels#string[][]',
    ].includes(id)),
    false,
  );
  const chartPresentationSupportedIds = [
    'interface:IChartOpts@property:altText',
    'interface:IChartOpts@property:barGrouping',
    'interface:IChartOpts@property:dataLabelFontBold',
    'interface:IChartOpts@property:dataLabelFontItalic',
    'interface:IChartOpts@property:dataLabelFontSize',
    'interface:IChartOpts@property:dataTableFontSize',
    'interface:IChartOpts@property:holeSize',
    'interface:IChartOpts@property:legendFontSize',
    'interface:IChartOpts@property:lineSmooth',
    'interface:IChartOpts@property:radarStyle',
    'interface:IChartOpts@property:showDataTable',
    'interface:IChartOpts@property:showDataTableHorzBorder',
    'interface:IChartOpts@property:showDataTableKeys',
    'interface:IChartOpts@property:showDataTableOutline',
    'interface:IChartOpts@property:showDataTableVertBorder',
    'interface:IChartOpts@property:showLeaderLines',
    'interface:IChartOpts@property:showLegend',
    'interface:IChartOpts@property:showPercent',
    'interface:IChartOpts@property:showTitle',
    'interface:IChartOpts@property:showValue',
    'interface:IChartOpts@property:title',
    'interface:IChartOpts@property:titleBold',
    'interface:IChartOpts@property:titleFontSize',
    'union:interface:IChartOpts@property:radarStyle#filled',
    'union:interface:IChartOpts@property:radarStyle#marker',
    'union:interface:IChartOpts@property:radarStyle#standard',
  ];
  const chartPresentationDifferenceIds = [
    'interface:IChartOpts@property:barDir',
    'interface:IChartOpts@property:barGapDepthPct',
    'interface:IChartOpts@property:barGapWidthPct',
    'interface:IChartOpts@property:barOverlapPct',
    'interface:IChartOpts@property:chartColors',
    'interface:IChartOpts@property:chartColorsOpacity',
    'interface:IChartOpts@property:dataLabelColor',
    'interface:IChartOpts@property:dataLabelFontFace',
    'interface:IChartOpts@property:dataLabelFormatCode',
    'interface:IChartOpts@property:dataLabelPosition',
    'interface:IChartOpts@property:dataTableFormatCode',
    'interface:IChartOpts@property:displayBlanksAs',
    'interface:IChartOpts@property:firstSliceAng',
    'interface:IChartOpts@property:h',
    'interface:IChartOpts@property:legendColor',
    'interface:IChartOpts@property:legendFontFace',
    'interface:IChartOpts@property:legendPos',
    'interface:IChartOpts@property:lineDash',
    'interface:IChartOpts@property:lineDataSymbol',
    'interface:IChartOpts@property:lineDataSymbolLineColor',
    'interface:IChartOpts@property:lineDataSymbolLineSize',
    'interface:IChartOpts@property:lineDataSymbolSize',
    'interface:IChartOpts@property:lineSize',
    'interface:IChartOpts@property:objectName',
    'interface:IChartOpts@property:showLabel',
    'interface:IChartOpts@property:showSerName',
    'interface:IChartOpts@property:titleColor',
    'interface:IChartOpts@property:titleFontFace',
    'interface:IChartOpts@property:titlePos',
    'interface:IChartOpts@property:titleRotate',
    'interface:IChartOpts@property:v3DPerspective',
    'interface:IChartOpts@property:v3DRAngAx',
    'interface:IChartOpts@property:v3DRotX',
    'interface:IChartOpts@property:v3DRotY',
    'interface:IChartOpts@property:w',
    'interface:IChartOpts@property:x',
    'interface:IChartOpts@property:y',
    'union:interface:IChartOpts@property:dataLabelPosition#b',
    'union:interface:IChartOpts@property:dataLabelPosition#bestFit',
    'union:interface:IChartOpts@property:dataLabelPosition#ctr',
    'union:interface:IChartOpts@property:dataLabelPosition#inEnd',
    'union:interface:IChartOpts@property:dataLabelPosition#l',
    'union:interface:IChartOpts@property:dataLabelPosition#outEnd',
    'union:interface:IChartOpts@property:dataLabelPosition#r',
    'union:interface:IChartOpts@property:dataLabelPosition#t',
    'union:interface:IChartOpts@property:legendPos#b',
    'union:interface:IChartOpts@property:legendPos#l',
    'union:interface:IChartOpts@property:legendPos#r',
    'union:interface:IChartOpts@property:legendPos#t',
    'union:interface:IChartOpts@property:legendPos#tr',
    'union:interface:IChartOpts@property:lineDash#dash',
    'union:interface:IChartOpts@property:lineDash#dashDot',
    'union:interface:IChartOpts@property:lineDash#lgDash',
    'union:interface:IChartOpts@property:lineDash#lgDashDot',
    'union:interface:IChartOpts@property:lineDash#lgDashDotDot',
    'union:interface:IChartOpts@property:lineDash#solid',
    'union:interface:IChartOpts@property:lineDash#sysDash',
    'union:interface:IChartOpts@property:lineDash#sysDot',
    'union:interface:IChartOpts@property:lineDataSymbol#circle',
    'union:interface:IChartOpts@property:lineDataSymbol#dash',
    'union:interface:IChartOpts@property:lineDataSymbol#diamond',
    'union:interface:IChartOpts@property:lineDataSymbol#dot',
    'union:interface:IChartOpts@property:lineDataSymbol#none',
    'union:interface:IChartOpts@property:lineDataSymbol#square',
    'union:interface:IChartOpts@property:lineDataSymbol#triangle',
  ];
  const expectedChartPresentationEntries = [
    ...chartPresentationSupportedIds.map((id) => ({ id, status: 'supported' })),
    ...chartPresentationDifferenceIds.map((id) => ({
      id, status: 'deliberate-difference',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  assert.equal(new Set(expectedChartPresentationEntries.map(({ id }) => id)).size, 91);
  assert.equal(chartPresentationSupportedIds.length, 26);
  assert.equal(chartPresentationDifferenceIds.length, 65);
  const chartPresentationIds = new Set(
    expectedChartPresentationEntries.map(({ id }) => id),
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => chartPresentationIds.has(id))
      .map(({ id, status }) => ({ id, status }))
      .sort((left, right) => left.id.localeCompare(right.id)),
    expectedChartPresentationEntries,
  );
  const chartPresentationById = new Map(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => chartPresentationIds.has(id))
      .map((entry) => [entry.id, entry]),
  );
  assert.equal(
    chartPresentationById.get('interface:IChartOpts@property:showLabel')?.status,
    'deliberate-difference',
  );
  assert.equal(
    chartPresentationById.get('interface:IChartOpts@property:showPercent')?.status,
    'supported',
  );
  assert.equal(
    chartPresentationById.get(
      'union:interface:IChartOpts@property:dataLabelPosition#bestFit',
    )?.status,
    'deliberate-difference',
  );
  const axisOwners = [
    ['IChartOpts', 'cat'],
    ['IChartOpts', 'val'],
    ['IChartPropsAxisCat', 'cat'],
    ['IChartPropsAxisVal', 'val'],
  ];
  const expectedChartAxisFoundationEntries = [
    ...axisOwners.flatMap(([owner, prefix]) => {
      const labelPositionId = `interface:${owner}@property:${prefix}AxisLabelPos`;
      return [
        { id: labelPositionId, status: 'supported' },
        ...['high', 'low', 'nextTo', 'none'].map((value) => ({
          id: `union:${labelPositionId}#${value}`,
          status: 'supported',
        })),
        ...['Color', 'Show', 'Size', 'Style'].map((suffix) => ({
          id: `interface:${owner}@property:${prefix}AxisLine${suffix}`,
          status: 'deliberate-difference',
        })),
        {
          id: `interface:${owner}@property:${prefix}GridLine`,
          status: 'deliberate-difference',
        },
        ...['dash', 'dot', 'solid'].map((value) => ({
          id: `union:interface:${owner}@property:${prefix}AxisLineStyle#${value}`,
          status: 'deliberate-difference',
        })),
        {
          id: `interface:${owner}@property:${prefix}AxisHidden`,
          status: 'deliberate-difference',
        },
        {
          id: `interface:${owner}@property:${prefix}AxisLabelRotate`,
          status: 'deliberate-difference',
        },
        ...['Major', 'Minor'].map((level) => ({
          id: `interface:${owner}@property:${prefix}Axis${level}TickMark`,
          status: 'deliberate-difference',
        })),
      ];
    }),
    ...['color', 'size', 'style'].map((property) => ({
      id: `interface:OptsChartGridLine@property:${property}`,
      status: 'deliberate-difference',
    })),
    ...['dash', 'dot', 'none', 'solid'].map((value) => ({
      id: `union:interface:OptsChartGridLine@property:style#${value}`,
      status: 'deliberate-difference',
    })),
    ...['cross', 'none'].map((value) => ({
      id: `union:ChartAxisTickMark#${value}`,
      status: 'supported',
    })),
    ...['inside', 'outside'].map((value) => ({
      id: `union:ChartAxisTickMark#${value}`,
      status: 'deliberate-difference',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const chartAxisFoundationIds = new Set(
    expectedChartAxisFoundationEntries.map(({ id }) => id),
  );
  assert.equal(chartAxisFoundationIds.size, 79);
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => chartAxisFoundationIds.has(id))
      .map(({ id, status }) => ({ id, status }))
      .sort((left, right) => left.id.localeCompare(right.id)),
    expectedChartAxisFoundationEntries,
  );
  const advancedAxisProperty = (owner, property, status) => ({
    id: `interface:${owner}@property:${property}`,
    status,
  });
  const advancedAxisUnion = (owner, property, value, status) => ({
    id: `union:interface:${owner}@property:${property}#${value}`,
    status,
  });
  const categorySupportedProperties = [
    'catAxisBaseTimeUnit',
    'catAxisLabelFontBold',
    'catAxisLabelFontItalic',
    'catAxisLabelFontSize',
    'catAxisMajorTimeUnit',
    'catAxisMajorUnit',
    'catAxisMinorTimeUnit',
    'catAxisMinorUnit',
    'catAxisMultiLevelLabels',
    'catAxisOrientation',
    'catAxisTitle',
    'catAxisTitleFontSize',
    'showCatAxisTitle',
  ];
  const categoryDifferenceProperties = [
    'catAxes',
    'catAxisCrossesAt',
    'catAxisLabelColor',
    'catAxisLabelFontFace',
    'catAxisLabelFrequency',
    'catAxisMaxVal',
    'catAxisMinVal',
    'catAxisTitleColor',
    'catAxisTitleFontFace',
    'catAxisTitleRotate',
    'catLabelFormatCode',
    'secondaryCatAxis',
  ];
  const valueSupportedProperties = [
    'showValAxisTitle',
    'valAxisDisplayUnit',
    'valAxisDisplayUnitLabel',
    'valAxisLabelFontBold',
    'valAxisLabelFontItalic',
    'valAxisLabelFontSize',
    'valAxisMajorUnit',
    'valAxisOrientation',
    'valAxisTitle',
    'valAxisTitleFontSize',
  ];
  const valueDifferenceProperties = [
    'secondaryValAxis',
    'valAxes',
    'valAxisCrossesAt',
    'valAxisLabelColor',
    'valAxisLabelFontFace',
    'valAxisLabelFormatCode',
    'valAxisLogScaleBase',
    'valAxisMaxVal',
    'valAxisMinVal',
    'valAxisTitleColor',
    'valAxisTitleFontFace',
    'valAxisTitleRotate',
    'valLabelFormatCode',
  ];
  const seriesSupportedProperties = [
    'serAxisLabelFontBold',
    'serAxisLabelFontItalic',
    'serAxisLabelFontSize',
    'serAxisLabelPos',
    'serAxisMajorUnit',
    'serAxisMinorUnit',
    'serAxisOrientation',
    'serAxisTitle',
    'serAxisTitleFontSize',
    'showSerAxisTitle',
  ];
  const seriesDifferenceProperties = [
    'serAxisHidden',
    'serAxisLabelColor',
    'serAxisLabelFontFace',
    'serAxisLabelFrequency',
    'serAxisLineColor',
    'serAxisLineShow',
    'serAxisTitleColor',
    'serAxisTitleFontFace',
    'serAxisTitleRotate',
    'serGridLine',
    'serLabelFormatCode',
  ];
  const displayUnits = [
    'billions',
    'hundredMillions',
    'hundredThousands',
    'hundreds',
    'millions',
    'tenMillions',
    'tenThousands',
    'thousands',
    'trillions',
  ];
  const expectedAdvancedAxisEntries = [
    ...['IChartOpts', 'IChartPropsAxisCat'].flatMap((owner) => [
      ...categorySupportedProperties.map((property) =>
        advancedAxisProperty(owner, property, 'supported')),
      advancedAxisUnion(owner, 'catAxisCrossesAt', 'autoZero', 'supported'),
      ...categoryDifferenceProperties.map((property) =>
        advancedAxisProperty(owner, property, 'deliberate-difference')),
      advancedAxisUnion(owner, 'catAxisCrossesAt', 'number', 'deliberate-difference'),
    ]),
    ...['IChartOpts', 'IChartPropsAxisVal'].flatMap((owner) => [
      ...valueSupportedProperties.map((property) =>
        advancedAxisProperty(owner, property, 'supported')),
      advancedAxisUnion(owner, 'valAxisCrossesAt', 'autoZero', 'supported'),
      ...displayUnits.map((value) =>
        advancedAxisUnion(owner, 'valAxisDisplayUnit', value, 'supported')),
      ...valueDifferenceProperties.map((property) =>
        advancedAxisProperty(owner, property, 'deliberate-difference')),
      advancedAxisUnion(owner, 'valAxisCrossesAt', 'number', 'deliberate-difference'),
    ]),
    ...seriesSupportedProperties.map((property) =>
      advancedAxisProperty('IChartOpts', property, 'supported')),
    advancedAxisUnion('IChartOpts', 'serAxisLabelPos', 'low', 'supported'),
    ...seriesDifferenceProperties.map((property) =>
      advancedAxisProperty('IChartOpts', property, 'deliberate-difference')),
    ...['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'].map(
      (property) => advancedAxisProperty('IChartOpts', property, 'defect-excluded'),
    ),
    ...['high', 'nextTo', 'none'].map((value) =>
      advancedAxisUnion('IChartOpts', 'serAxisLabelPos', value, 'defect-excluded')),
    advancedAxisProperty('IChartOpts', 'axisPos', 'defect-excluded'),
    ...['b', 'l', 'r', 't'].map((value) =>
      advancedAxisUnion('IChartOpts', 'axisPos', value, 'defect-excluded')),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const advancedAxisIds = new Set(expectedAdvancedAxisEntries.map(({ id }) => id));
  assert.equal(advancedAxisIds.size, 155);
  assert.deepEqual(
    expectedAdvancedAxisEntries.map(({ status }) => status).sort(),
    [
      ...Array(79).fill('supported'),
      ...Array(65).fill('deliberate-difference'),
      ...Array(11).fill('defect-excluded'),
    ].sort(),
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => advancedAxisIds.has(id))
      .map(({ id, status }) => ({ id, status }))
      .sort((left, right) => left.id.localeCompare(right.id)),
    expectedAdvancedAxisEntries,
  );
  assert.equal(
    PPTXGENJS_SURFACE_MANIFEST.entries.some(({ id }) => [
      'interface:OptsChartGridLine@property:cap',
      'union:ChartLineCap#flat',
      'union:ChartLineCap#round',
      'union:ChartLineCap#square',
    ].includes(id)),
    false,
  );
  const inertChartOptionProperties = new Set([
    'align',
    'bold',
    'breakLine',
    'bullet',
    'cap',
    'color',
    'fontFace',
    'fontSize',
    'highlight',
    'italic',
    'size',
    'softBreakBefore',
    'style',
    'tabStops',
    'textDirection',
    'transparency',
    'underline',
    'valign',
  ]);
  const inertChartOptionEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) => {
    const property = id.match(/interface:IChartOpts@property:([^@#]+)/u)?.[1];
    return property !== undefined && inertChartOptionProperties.has(property);
  });
  assert.equal(inertChartOptionEntries.length, 79);
  assert.deepEqual(
    Object.fromEntries(
      ['property', 'inline', 'union'].map((kind) => [
        kind,
        inertChartOptionEntries.filter(({ id }) => {
          if (kind === 'inline') return id.startsWith('inline:');
          if (kind === 'union') return id.startsWith('union:');
          return id.startsWith('interface:');
        }).length,
      ]),
    ),
    { property: 18, inline: 13, union: 48 },
  );
  assert.deepEqual(
    [...new Set(inertChartOptionEntries.map(({ status }) => status))],
    ['defect-excluded'],
  );
  const tabStopsFamilyOwners = new Set([
    'PlaceholderProps',
    'SlideNumberProps',
    'TableCellProps',
    'TableProps',
    'TableToSlidesProps',
    'TextPropsOptions',
  ]);
  const tabStopsFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) => {
    const owner = id.match(/interface:([^@]+)@property:tabStops/u)?.[1];
    return owner !== undefined && tabStopsFamilyOwners.has(owner);
  });
  assert.equal(tabStopsFamilyEntries.length, 42);
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'defect-excluded'].map((status) => [
        status,
        tabStopsFamilyEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 3, 'deliberate-difference': 18, 'defect-excluded': 21 },
  );
  for (const owner of tabStopsFamilyOwners) {
    const entries = tabStopsFamilyEntries.filter(({ id }) =>
      id.includes(`interface:${owner}@property:tabStops`));
    assert.equal(entries.length, 7, owner);
    assert.deepEqual(
      entries.map(({ status }) => status).sort(),
      ['PlaceholderProps', 'TableCellProps', 'TextPropsOptions'].includes(owner)
        ? ['supported', ...Array(6).fill('deliberate-difference')].sort()
        : Array(7).fill('defect-excluded'),
      owner,
    );
  }
  const underlineFamilyOwners = new Set([
    'PlaceholderProps',
    'SlideNumberProps',
    'TableCellProps',
    'TableProps',
    'TableToSlidesProps',
    'TextPropsOptions',
  ]);
  const underlineFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) => {
    const owner = id.match(/interface:([^@]+)@property:underline/u)?.[1];
    return owner !== undefined && underlineFamilyOwners.has(owner);
  });
  assert.equal(underlineFamilyEntries.length, 120);
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'defect-excluded'].map((status) => [
        status,
        underlineFamilyEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 68, 'deliberate-difference': 8, 'defect-excluded': 44 },
  );
  for (const owner of underlineFamilyOwners) {
    const entries = underlineFamilyEntries.filter(({ id }) =>
      id.includes(`interface:${owner}@property:underline`));
    assert.equal(entries.length, 20, owner);
    assert.deepEqual(
      entries.map(({ status }) => status).sort(),
      ['PlaceholderProps', 'TableCellProps', 'TableProps', 'TextPropsOptions'].includes(owner)
        ? [
            ...Array(17).fill('supported'),
            ...Array(2).fill('deliberate-difference'),
            'defect-excluded',
          ].sort()
        : Array(20).fill('defect-excluded'),
      owner,
    );
  }
  const textDirectionFamilyOwners = new Set([
    'PlaceholderProps',
    'SlideNumberProps',
    'TableCellProps',
    'TableProps',
    'TableToSlidesProps',
    'TextPropsOptions',
  ]);
  const textDirectionFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) => {
    const match = id.match(
      /interface:([^@]+)@property:(textDirection|vert)(?:#|$)/u,
    );
    return match !== null
      && textDirectionFamilyOwners.has(match[1])
      && (match[2] === 'textDirection' || match[1] === 'TextPropsOptions');
  });
  assert.equal(textDirectionFamilyEntries.length, 38);
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'defect-excluded'].map((status) => [
        status,
        textDirectionFamilyEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 18, 'defect-excluded': 20 },
  );
  for (const owner of textDirectionFamilyOwners) {
    const entries = textDirectionFamilyEntries.filter(({ id }) =>
      id.includes(`interface:${owner}@property:textDirection`));
    assert.equal(entries.length, 5, owner);
    assert.deepEqual(
      entries.map(({ status }) => status),
      ['TableCellProps', 'TableProps'].includes(owner)
        ? Array(5).fill('supported')
        : Array(5).fill('defect-excluded'),
      owner,
    );
  }
  const textVertEntries = textDirectionFamilyEntries.filter(({ id }) =>
    id.includes('interface:TextPropsOptions@property:vert'));
  assert.equal(textVertEntries.length, 8);
  assert.deepEqual(textVertEntries.map(({ status }) => status), Array(8).fill('supported'));
  const textBoxFitPropertyId = 'interface:TextPropsOptions@property:fit';
  const textBoxFitExpected = [
    { id: textBoxFitPropertyId, status: 'supported' },
    ...['none', 'resize', 'shrink'].map((value) => ({
      id: `union:${textBoxFitPropertyId}#${value}`,
      status: 'supported',
    })),
    {
      id: 'interface:TextPropsOptions@property:autoFit',
      status: 'deprecated-alias',
    },
    {
      id: 'interface:TextPropsOptions@property:shrinkText',
      status: 'deprecated-alias',
    },
  ].sort((left, right) => left.id.localeCompare(right.id));
  const textBoxFitEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => textBoxFitExpected.some((expected) => expected.id === id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    textBoxFitEntries.map(({ id, status }) => ({ id, status })),
    textBoxFitExpected,
  );
  assert.equal(textBoxFitEntries.length, 6);
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deprecated-alias'].map((status) => [
        status,
        textBoxFitEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 4, 'deprecated-alias': 2 },
  );
  assert.equal(
    textBoxFitEntries.every(({
      native,
      evidence,
      serialization,
      client,
    }) => native.length === 3 && evidence.code.length === 2 &&
      evidence.tests.length === 2 &&
      evidence.package.some(({ pattern }) =>
        pattern === 'const initialTextFit = createdText.textFit;') &&
      evidence.ooxml.some(({ pattern }) =>
        pattern === 'creates, edits, duplicates, and reopens text-box fit modes') &&
      evidence.clients.some(({ pattern }) =>
        pattern === 'const textBoxFitFamilyState = {') &&
      serialization === true && client === true),
    true,
  );
  for (const entry of textBoxFitEntries.filter(
    ({ status }) => status === 'deprecated-alias',
  )) {
    assert.equal(entry.canonical, textBoxFitPropertyId);
    assert.equal(
      entry.control.pattern,
      'imports public PptxGenJS output and continues editing in the OOXML kernel',
    );
  }
  const textParagraphLayoutExpectedStatus = {
    align: 'supported',
    inset: 'deprecated-alias',
    lineSpacing: 'deliberate-difference',
    lineSpacingMultiple: 'deliberate-difference',
    margin: 'deliberate-difference',
    paraSpaceAfter: 'deliberate-difference',
    paraSpaceBefore: 'deliberate-difference',
    rtlMode: 'supported',
    valign: 'supported',
    wrap: 'supported',
  };
  const textParagraphLayoutEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => id.startsWith('interface:TextPropsOptions@property:'))
    .filter(({ id }) => Object.hasOwn(
      textParagraphLayoutExpectedStatus,
      id.slice(id.lastIndexOf(':') + 1),
    ))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.equal(textParagraphLayoutEntries.length, 10);
  assert.equal(new Set(textParagraphLayoutEntries.map(({ id }) => id)).size, 10);
  assert.deepEqual(
    Object.fromEntries(textParagraphLayoutEntries.map(({ id, status }) => [
      id.slice(id.lastIndexOf(':') + 1),
      status,
    ])),
    textParagraphLayoutExpectedStatus,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'deprecated-alias'].map((status) => [
        status,
        textParagraphLayoutEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 4, 'deliberate-difference': 5, 'deprecated-alias': 1 },
  );
  assert.equal(
    textParagraphLayoutEntries.every(({
      native,
      evidence,
      serialization,
      client,
    }) => native.length >= 2 && evidence.code.length === 1 &&
      evidence.tests.length === 2 &&
      evidence.package.some(({ pattern }) =>
        pattern === 'const createdText = created.addSlide().addText(') &&
      evidence.ooxml.length === 1 &&
      evidence.clients.some(({ pattern }) =>
        pattern === 'const textParagraphLayoutFamilyState = {') &&
      serialization === true && client === true),
    true,
  );
  for (const entry of textParagraphLayoutEntries.filter(
    ({ status }) => status !== 'supported',
  )) {
    assert.equal(
      entry.control.pattern,
      'imports public PptxGenJS output and continues editing in the OOXML kernel',
    );
  }
  const textInsetEntry = textParagraphLayoutEntries.find(({ id }) =>
    id === 'interface:TextPropsOptions@property:inset');
  assert.equal(
    textInsetEntry.canonical,
    'interface:TextPropsOptions@property:margin',
  );
  const richTextEffectsExpectedStatus = {
    'inline:interface:TextPropsOptions@property:outline@property:outline.color':
      'deliberate-difference',
    'inline:interface:TextPropsOptions@property:outline@property:outline.size': 'supported',
    'interface:TextGlowProps@property:color': 'deliberate-difference',
    'interface:TextGlowProps@property:opacity': 'supported',
    'interface:TextGlowProps@property:size': 'supported',
    'interface:TextPropsOptions@property:baseline': 'deliberate-difference',
    'interface:TextPropsOptions@property:charSpacing': 'deliberate-difference',
    'interface:TextPropsOptions@property:glow': 'supported',
    'interface:TextPropsOptions@property:outline': 'supported',
    'interface:TextPropsOptions@property:strike': 'supported',
    'interface:TextPropsOptions@property:subscript': 'deliberate-difference',
    'interface:TextPropsOptions@property:superscript': 'deliberate-difference',
    'interface:TextPropsOptions@property:transparency': 'supported',
    'union:interface:TextPropsOptions@property:strike#boolean': 'deliberate-difference',
    'union:interface:TextPropsOptions@property:strike#dblStrike': 'supported',
    'union:interface:TextPropsOptions@property:strike#sngStrike': 'supported',
  };
  const richTextEffectsEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => Object.hasOwn(richTextEffectsExpectedStatus, id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.equal(richTextEffectsEntries.length, 16);
  assert.equal(new Set(richTextEffectsEntries.map(({ id }) => id)).size, 16);
  assert.deepEqual(
    Object.fromEntries(richTextEffectsEntries.map(({ id, status }) => [id, status])),
    richTextEffectsExpectedStatus,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference'].map((status) => [
        status,
        richTextEffectsEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 9, 'deliberate-difference': 7 },
  );
  assert.equal(
    richTextEffectsEntries.every(({
      native,
      evidence,
      serialization,
      client,
    }) => native.length >= 3 && evidence.code.length === 1 &&
      evidence.tests.length === 3 && evidence.package.length === 1 &&
      evidence.ooxml.length === 1 &&
      evidence.clients.some(({ pattern }) =>
        pattern === 'const richTextEffectsFamilyState = {') &&
      serialization === true && client === true),
    true,
  );
  for (const entry of richTextEffectsEntries.filter(
    ({ status }) => status === 'deliberate-difference',
  )) {
    assert.equal(entry.control.path, 'packages/pptxgenjs-adapter/src/index.test.ts');
  }
  const textRunScalarFamilyStatus = {
    PlaceholderProps: {
      supported: 4,
      'deliberate-difference': 3,
      'defect-excluded': 2,
    },
    SlideNumberProps: {
      supported: 2,
      'deliberate-difference': 2,
      'defect-excluded': 5,
    },
    TableCellProps: {
      supported: 5,
      'deliberate-difference': 4,
      'defect-excluded': 0,
    },
    TableProps: {
      supported: 2,
      'deliberate-difference': 2,
      'defect-excluded': 5,
    },
    TableToSlidesProps: {
      supported: 0,
      'deliberate-difference': 0,
      'defect-excluded': 9,
    },
    TextPropsOptions: {
      supported: 6,
      'deliberate-difference': 3,
      'defect-excluded': 0,
    },
  };
  const textRunScalarProperties = new Set([
    'bold',
    'breakLine',
    'color',
    'fontFace',
    'fontSize',
    'highlight',
    'italic',
    'lang',
    'softBreakBefore',
  ]);
  const textRunScalarFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) => {
    const match = id.match(/interface:([^@]+)@property:([^@#]+)$/u);
    return match !== null
      && Object.hasOwn(textRunScalarFamilyStatus, match[1])
      && textRunScalarProperties.has(match[2]);
  });
  assert.equal(textRunScalarFamilyEntries.length, 54);
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'defect-excluded'].map((status) => [
        status,
        textRunScalarFamilyEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 19, 'deliberate-difference': 14, 'defect-excluded': 21 },
  );
  for (const [owner, expected] of Object.entries(textRunScalarFamilyStatus)) {
    const entries = textRunScalarFamilyEntries.filter(({ id }) =>
      id.startsWith(`interface:${owner}@property:`));
    assert.equal(entries.length, 9, owner);
    assert.deepEqual(
      Object.fromEntries(
        ['supported', 'deliberate-difference', 'defect-excluded'].map((status) => [
          status,
          entries.filter((entry) => entry.status === status).length,
        ]),
      ),
      expected,
      owner,
    );
  }
  const tableToSlidesPropertyId = (property) =>
    `interface:TableToSlidesProps@property:${property}`;
  const tableToSlidesExpected = [
    ...[
      'autoPageCharWeight',
      'autoPageLineWeight',
      'autoPageRepeatHeader',
      'masterSlideName',
    ].map((property) => ({
      id: tableToSlidesPropertyId(property),
      status: 'supported',
    })),
    {
      id: 'class:PptxGenJS#tableToSlides',
      status: 'deliberate-difference',
    },
    ...[
      'addImage',
      'addShape',
      'addTable',
      'addText',
      'autoPageSlideStartY',
      'h',
      'slideMargin',
      'verbose',
      'w',
      'x',
      'y',
    ].map((property) => ({
      id: tableToSlidesPropertyId(property),
      status: 'deliberate-difference',
    })),
    ...[
      ['addImage', 'image'],
      ['addImage', 'options'],
      ['addShape', 'options'],
      ['addShape', 'shapeName'],
      ['addTable', 'options'],
      ['addTable', 'rows'],
      ['addText', 'options'],
      ['addText', 'text'],
    ].map(([owner, field]) => ({
      id: `inline:interface:TableToSlidesProps@property:${owner}@property:${owner}.${field}`,
      status: 'deliberate-difference',
    })),
    ...['addHeaderToEach', 'newSlideStartY'].map((property) => ({
      id: tableToSlidesPropertyId(property),
      status: 'deprecated-alias',
    })),
    ...[
      'align',
      'autoPage',
      'autoPageHeaderRows',
      'border',
      'colW',
      'margin',
      'objectName',
      'rowH',
      'transparency',
      'valign',
    ].map((property) => ({
      id: tableToSlidesPropertyId(property),
      status: 'defect-excluded',
    })),
    ...Object.entries({
      border: ['BorderProps', '[BorderProps,BorderProps,BorderProps,BorderProps]'],
      colW: ['number', 'number[]'],
      rowH: ['number', 'number[]'],
    }).flatMap(([property, tokens]) => tokens.map((token) => ({
      id: `union:${tableToSlidesPropertyId(property)}#${token}`,
      status: 'defect-excluded',
    }))),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const tableToSlidesIds = new Set(tableToSlidesExpected.map(({ id }) => id));
  assert.equal(tableToSlidesIds.size, 42);
  const tableToSlidesEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => tableToSlidesIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    tableToSlidesEntries.map(({ id, status }) => ({ id, status })),
    tableToSlidesExpected,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'deprecated-alias', 'defect-excluded']
        .map((status) => [
          status,
          tableToSlidesEntries.filter((entry) => entry.status === status).length,
        ]),
    ),
    {
      supported: 4,
      'deliberate-difference': 20,
      'deprecated-alias': 2,
      'defect-excluded': 16,
    },
  );
  const tableToSlidesById = new Map(
    tableToSlidesEntries.map((entry) => [entry.id, entry]),
  );
  assert.equal(
    tableToSlidesById.get(tableToSlidesPropertyId('addHeaderToEach'))?.canonical,
    tableToSlidesPropertyId('autoPageRepeatHeader'),
  );
  assert.equal(
    tableToSlidesById.get(tableToSlidesPropertyId('newSlideStartY'))?.canonical,
    tableToSlidesPropertyId('autoPageSlideStartY'),
  );
  const addTablePropertyId = (owner, property) =>
    `interface:${owner}@property:${property}`;
  const addTableUnionId = (owner, property, token) =>
    `union:${addTablePropertyId(owner, property)}#${token}`;
  const addTableExpected = [
    ...['align', 'colspan', 'hyperlink', 'rowspan', 'transparency', 'valign']
      .map((property) => ({
        id: addTablePropertyId('TableCellProps', property),
        status: 'supported',
      })),
    ...[
      'align',
      'autoPage',
      'autoPageCharWeight',
      'autoPageHeaderRows',
      'autoPageLineWeight',
      'autoPageRepeatHeader',
      'valign',
    ].map((property) => ({
      id: addTablePropertyId('TableProps', property),
      status: 'supported',
    })),
    { id: 'method:Slide#addTable', status: 'deliberate-difference' },
    ...['color', 'pt', 'type'].map((property) => ({
      id: addTablePropertyId('BorderProps', property),
      status: 'deliberate-difference',
    })),
    ...['dash', 'none', 'solid'].map((token) => ({
      id: addTableUnionId('BorderProps', 'type', token),
      status: 'deliberate-difference',
    })),
    ...['border', 'margin'].map((property) => ({
      id: addTablePropertyId('TableCellProps', property),
      status: 'deliberate-difference',
    })),
    ...[
      'BorderProps',
      '[BorderProps,BorderProps,BorderProps,BorderProps]',
    ].map((token) => ({
      id: addTableUnionId('TableCellProps', 'border', token),
      status: 'deliberate-difference',
    })),
    ...[
      'autoPageSlideStartY',
      'border',
      'colW',
      'h',
      'margin',
      'objectName',
      'rowH',
      'verbose',
      'w',
      'x',
      'y',
    ].map((property) => ({
      id: addTablePropertyId('TableProps', property),
      status: 'deliberate-difference',
    })),
    ...[
      ['border', 'BorderProps'],
      ['border', '[BorderProps,BorderProps,BorderProps,BorderProps]'],
      ['colW', 'number'],
      ['colW', 'number[]'],
      ['rowH', 'number'],
      ['rowH', 'number[]'],
    ].map(([property, token]) => ({
      id: addTableUnionId('TableProps', property, token),
      status: 'deliberate-difference',
    })),
    {
      id: addTablePropertyId('TableProps', 'newSlideStartY'),
      status: 'deprecated-alias',
    },
    ...[
      ['TableCellProps', 'autoPageCharWeight'],
      ['TableCellProps', 'autoPageLineWeight'],
      ['TableProps', 'transparency'],
    ].map(([owner, property]) => ({
      id: addTablePropertyId(owner, property),
      status: 'defect-excluded',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const addTableIds = new Set(addTableExpected.map(({ id }) => id));
  assert.equal(addTableIds.size, 45);
  const addTableEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => addTableIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    addTableEntries.map(({ id, status }) => ({ id, status })),
    addTableExpected,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'deprecated-alias', 'defect-excluded']
        .map((status) => [
          status,
          addTableEntries.filter((entry) => entry.status === status).length,
        ]),
    ),
    {
      supported: 13,
      'deliberate-difference': 28,
      'deprecated-alias': 1,
      'defect-excluded': 3,
    },
  );
  const addTableById = new Map(addTableEntries.map((entry) => [entry.id, entry]));
  assert.equal(
    addTableById.get(addTablePropertyId('TableProps', 'newSlideStartY'))?.canonical,
    addTablePropertyId('TableProps', 'autoPageSlideStartY'),
  );
  assert.equal(
    addTableById.get(addTablePropertyId('TableCellProps', 'transparency'))?.status,
    'supported',
  );
  const rootPropertyId = (property) => `class:PptxGenJS@property:${property}`;
  const outputReturnId = (method, token) =>
    `union:class:PptxGenJS#${method}@path:return#${token}`;
  const rootOutputExpected = [
    { id: 'class:PptxGenJS#addSlide', status: 'supported' },
    ...['author', 'company', 'revision', 'rtlMode', 'subject', 'theme', 'title']
      .map((property) => ({ id: rootPropertyId(property), status: 'supported' })),
    ...['height', 'width'].map((property) => ({
      id: addTablePropertyId('PresLayout', property),
      status: 'supported',
    })),
    ...['bodyFontFace', 'headFontFace'].map((property) => ({
      id: addTablePropertyId('ThemeProps', property),
      status: 'supported',
    })),
    {
      id: addTablePropertyId('WriteBaseProps', 'compression'),
      status: 'supported',
    },
    { id: 'union:WRITE_OUTPUT_TYPE#JSZIP_OUTPUT_TYPE', status: 'supported' },
    ...['ArrayBuffer', 'Blob', 'Uint8Array', 'string'].map((token) => ({
      id: outputReturnId('write', token),
      status: 'supported',
    })),
    ...[
      'addSection',
      'defineLayout',
      'defineSlideMaster',
      'stream',
      'write',
      'writeFile',
    ].map((method) => ({
      id: `class:PptxGenJS#${method}`,
      status: 'deliberate-difference',
    })),
    { id: rootPropertyId('layout'), status: 'deliberate-difference' },
    {
      id: addTablePropertyId('PresLayout', 'name'),
      status: 'deliberate-difference',
    },
    ...[
      ['WriteFileProps', 'compression'],
      ['WriteFileProps', 'fileName'],
      ['WriteProps', 'compression'],
      ['WriteProps', 'outputType'],
    ].map(([owner, property]) => ({
      id: addTablePropertyId(owner, property),
      status: 'deliberate-difference',
    })),
    { id: 'union:WRITE_OUTPUT_TYPE#STREAM', status: 'deliberate-difference' },
    {
      id: outputReturnId('stream', 'Uint8Array'),
      status: 'deliberate-difference',
    },
    ...['ArrayBuffer', 'Blob', 'string'].map((token) => ({
      id: outputReturnId('stream', token),
      status: 'defect-excluded',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const rootOutputIds = new Set(rootOutputExpected.map(({ id }) => id));
  assert.equal(rootOutputIds.size, 35);
  const rootOutputEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => rootOutputIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    rootOutputEntries.map(({ id, status }) => ({ id, status })),
    rootOutputExpected,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'defect-excluded'].map((status) => [
        status,
        rootOutputEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 18, 'deliberate-difference': 14, 'defect-excluded': 3 },
  );
  assert.equal(
    rootOutputEntries.some(({ id }) => id.includes('ChartType')),
    false,
  );
  assert.deepEqual(
    rootOutputEntries
      .filter(({ status }) => status === 'defect-excluded')
      .map(({ id }) => id),
    ['ArrayBuffer', 'Blob', 'string']
      .map((token) => outputReturnId('stream', token))
      .sort(),
  );
  const slidePropertyId = (property) => `property:Slide#${property}`;
  const slideSectionExpected = [
    ...['addChart', 'addImage', 'addNotes', 'addShape', 'addTable', 'addText']
      .map((property) => ({
        id: addTablePropertyId('PresSlide', property),
        status: 'supported',
      })),
    {
      id: addTablePropertyId('PresSlide', 'hidden'),
      status: 'supported',
    },
    { id: 'method:Slide#addNotes', status: 'supported' },
    { id: slidePropertyId('hidden'), status: 'supported' },
    {
      id: addTablePropertyId('SectionProps', 'title'),
      status: 'supported',
    },
    ...[
      ['AddSlideProps', 'masterName'],
      ['AddSlideProps', 'sectionTitle'],
      ['PresSlide', 'addMedia'],
      ['PresSlide', 'background'],
      ['PresSlide', 'color'],
      ['PresSlide', 'slideNumber'],
      ['SectionProps', 'order'],
    ].map(([owner, property]) => ({
      id: addTablePropertyId(owner, property),
      status: 'deliberate-difference',
    })),
    ...['addImage', 'addMedia', 'addShape', 'addText'].map((method) => ({
      id: `method:Slide#${method}`,
      status: 'deliberate-difference',
    })),
    ...['background', 'color', 'newAutoPagedSlides', 'slideNumber']
      .map((property) => ({
        id: slidePropertyId(property),
        status: 'deliberate-difference',
      })),
    { id: slidePropertyId('bkgd'), status: 'deprecated-alias' },
  ].sort((left, right) => left.id.localeCompare(right.id));
  const slideSectionIds = new Set(slideSectionExpected.map(({ id }) => id));
  assert.equal(slideSectionIds.size, 26);
  const slideSectionEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => slideSectionIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    slideSectionEntries.map(({ id, status }) => ({ id, status })),
    slideSectionExpected,
  );
  assert.deepEqual(
    Object.fromEntries(
      ['supported', 'deliberate-difference', 'deprecated-alias'].map((status) => [
        status,
        slideSectionEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    { supported: 10, 'deliberate-difference': 15, 'deprecated-alias': 1 },
  );
  assert.equal(
    slideSectionEntries.find(({ id }) => id === slidePropertyId('bkgd'))?.canonical,
    slidePropertyId('background'),
  );
  assert.deepEqual(
    slideSectionEntries
      .filter(({ id }) => /^interface:PresSlide@property:add/u.test(id))
      .map(({ id, status }) => ({ id, status })),
    [
      { id: addTablePropertyId('PresSlide', 'addChart'), status: 'supported' },
      { id: addTablePropertyId('PresSlide', 'addImage'), status: 'supported' },
      { id: addTablePropertyId('PresSlide', 'addMedia'), status: 'deliberate-difference' },
      { id: addTablePropertyId('PresSlide', 'addNotes'), status: 'supported' },
      { id: addTablePropertyId('PresSlide', 'addShape'), status: 'supported' },
      { id: addTablePropertyId('PresSlide', 'addTable'), status: 'supported' },
      { id: addTablePropertyId('PresSlide', 'addText'), status: 'supported' },
    ],
  );
  const inlineMasterObjectId = (property) =>
    `inline:interface:SlideMasterProps@property:objects@property:objects.${property}`;
  const propertyUnionId = (owner, property, token) =>
    `union:${addTablePropertyId(owner, property)}#${token}`;
  const masterBackgroundSlideNumberExpected = [
    ...[
      inlineMasterObjectId('placeholder.text'),
      addTablePropertyId('SlideMasterProps', 'title'),
      addTablePropertyId('SlideNumberProps', 'margin'),
      addTablePropertyId('SlideNumberProps', 'valign'),
    ].map((id) => ({ id, status: 'supported' })),
    ...[
      ...[
        'image',
        'line',
        'placeholder',
        'placeholder.options',
        'rect',
        'text',
      ].map(inlineMasterObjectId),
      ...['background', 'margin', 'objects', 'slideNumber']
        .map((property) => addTablePropertyId('SlideMasterProps', property)),
      ...['color', 'data', 'path', 'transparency', 'type']
        .map((property) => addTablePropertyId('BackgroundProps', property)),
      propertyUnionId('BackgroundProps', 'type', 'none'),
      propertyUnionId('BackgroundProps', 'type', 'solid'),
      ...['align', 'h', 'w', 'x', 'y']
        .map((property) => addTablePropertyId('SlideNumberProps', property)),
    ].map((id) => ({ id, status: 'deliberate-difference' })),
    ...[
      addTablePropertyId('SlideMasterProps', 'bkgd'),
      propertyUnionId('SlideMasterProps', 'bkgd', 'string'),
      addTablePropertyId('BackgroundProps', 'alpha'),
      addTablePropertyId('BackgroundProps', 'fill'),
    ].map((id) => ({ id, status: 'deprecated-alias' })),
    ...[
      inlineMasterObjectId('chart'),
      propertyUnionId('SlideMasterProps', 'bkgd', 'BackgroundProps'),
      addTablePropertyId('BackgroundProps', 'src'),
      addTablePropertyId('SlideNumberProps', 'transparency'),
    ].map((id) => ({ id, status: 'defect-excluded' })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const masterBackgroundSlideNumberIds = new Set(
    masterBackgroundSlideNumberExpected.map(({ id }) => id),
  );
  assert.equal(masterBackgroundSlideNumberIds.size, 34);
  const masterBackgroundSlideNumberEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => masterBackgroundSlideNumberIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    masterBackgroundSlideNumberEntries.map(({ id, status }) => ({ id, status })),
    masterBackgroundSlideNumberExpected,
  );
  assert.deepEqual(
    Object.fromEntries(
      [
        'supported',
        'deliberate-difference',
        'deprecated-alias',
        'defect-excluded',
      ].map((status) => [
        status,
        masterBackgroundSlideNumberEntries.filter((entry) => entry.status === status).length,
      ]),
    ),
    {
      supported: 4,
      'deliberate-difference': 22,
      'deprecated-alias': 4,
      'defect-excluded': 4,
    },
  );
  const masterBackgroundSlideNumberById = new Map(
    masterBackgroundSlideNumberEntries.map((entry) => [entry.id, entry]),
  );
  assert.equal(
    masterBackgroundSlideNumberById
      .get(addTablePropertyId('SlideMasterProps', 'bkgd'))?.canonical,
    addTablePropertyId('SlideMasterProps', 'background'),
  );
  assert.equal(
    masterBackgroundSlideNumberById
      .get(propertyUnionId('SlideMasterProps', 'bkgd', 'string'))?.canonical,
    addTablePropertyId('SlideMasterProps', 'background'),
  );
  assert.equal(
    masterBackgroundSlideNumberById
      .get(addTablePropertyId('BackgroundProps', 'alpha'))?.canonical,
    addTablePropertyId('BackgroundProps', 'transparency'),
  );
  assert.equal(
    masterBackgroundSlideNumberById
      .get(addTablePropertyId('BackgroundProps', 'fill'))?.canonical,
    addTablePropertyId('BackgroundProps', 'color'),
  );
  assert.deepEqual(
    masterBackgroundSlideNumberEntries
      .filter(({ status }) => status === 'defect-excluded')
      .map(({ id }) => id),
    [
      inlineMasterObjectId('chart'),
      propertyUnionId('SlideMasterProps', 'bkgd', 'BackgroundProps'),
      addTablePropertyId('BackgroundProps', 'src'),
      addTablePropertyId('SlideNumberProps', 'transparency'),
    ].sort(),
  );
  const shapeTextShadowExpected = [
    ...[
      'angle',
      'blur',
      'color',
      'offset',
      'opacity',
      'rotateWithShape',
      'type',
    ].map((property) => addTablePropertyId('ShadowProps', property)),
    ...['inner', 'none', 'outer']
      .map((token) => propertyUnionId('ShadowProps', 'type', token)),
    addTablePropertyId('ShapeProps', 'shadow'),
    addTablePropertyId('TextPropsOptions', 'shadow'),
  ].sort();
  const shapeTextShadowIds = new Set(shapeTextShadowExpected);
  assert.equal(shapeTextShadowIds.size, 12);
  const shapeTextShadowEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => shapeTextShadowIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    shapeTextShadowEntries.map(({ id, status }) => ({ id, status })),
    shapeTextShadowExpected.map((id) => ({
      id,
      status: 'deliberate-difference',
    })),
  );
  assert.equal(
    shapeTextShadowEntries.every(({ native, serialization, client }) =>
      native.length > 0 && serialization === true && client === true),
    true,
  );
  const shapeTextShadowById = new Map(
    shapeTextShadowEntries.map((entry) => [entry.id, entry]),
  );
  assert.deepEqual(
    shapeTextShadowById
      .get(propertyUnionId('ShadowProps', 'type', 'none'))?.native,
    ['ShapeModel.shadow'],
  );
  assert.deepEqual(
    shapeTextShadowById
      .get(addTablePropertyId('ShapeProps', 'shadow'))?.evidence.package,
    [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const shapeShadows =' }],
  );
  assert.deepEqual(
    shapeTextShadowById
      .get(addTablePropertyId('TextPropsOptions', 'shadow'))?.evidence.package,
    [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const textShapeShadows =' }],
  );
  const imageSourceSizingTransformExpected = [
    ...['data', 'path', 'rotate', 'flipH', 'flipV', 'sizing']
      .map((property) => addTablePropertyId('ImageProps', property)),
    ...['type', 'w', 'h', 'x', 'y']
      .map((property) =>
        `inline:interface:ImageProps@property:sizing@property:sizing.${property}`),
    ...['contain', 'cover', 'crop']
      .map((token) =>
        `union:interface:ImageProps@property:sizing@path:sizing.type#${token}`),
  ].sort();
  const imageSourceSizingTransformIds = new Set(imageSourceSizingTransformExpected);
  assert.equal(imageSourceSizingTransformIds.size, 14);
  const imageSourceSizingTransformEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => imageSourceSizingTransformIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    imageSourceSizingTransformEntries.map(({ id, status }) => ({ id, status })),
    imageSourceSizingTransformExpected.map((id) => ({
      id,
      status: 'deliberate-difference',
    })),
  );
  assert.equal(
    imageSourceSizingTransformEntries.every(({
      native,
      evidence,
      control,
      serialization,
      client,
    }) => native.length > 0 && evidence.code.length > 0 &&
      evidence.tests.length >= 3 && evidence.package.length > 0 &&
      evidence.ooxml.length > 0 && evidence.clients.some(
        ({ pattern }) => pattern === 'const imageSourceSizingTransformState = {',
      ) && control.pattern ===
        'locks ImageProps source, sizing, and transform divergences against PptxGenJS 4.0.1' &&
      serialization === true && client === true),
    true,
  );
  const imageSourceSizingTransformById = new Map(
    imageSourceSizingTransformEntries.map((entry) => [entry.id, entry]),
  );
  assert.deepEqual(
    imageSourceSizingTransformById
      .get(addTablePropertyId('ImageProps', 'path'))?.evidence.package,
    [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const packedSvgPath =' }],
  );
  assert.deepEqual(
    imageSourceSizingTransformEntries
      .filter(({ id }) => id.includes('@path:sizing.type#'))
      .map(({ native }) => native[0]),
    ['ImageSizing.type', 'ImageSizing.type', 'ImageSizing.type'],
  );
  assert.equal(
    [
      ...['hyperlink', 'x', 'y', 'w', 'h', 'placeholder']
        .map((property) => addTablePropertyId('ImageProps', property)),
      addTablePropertyId('DataOrPathProps', 'data'),
      addTablePropertyId('DataOrPathProps', 'path'),
    ].some((id) => imageSourceSizingTransformIds.has(id)),
    false,
  );
  const imageIdentityEffectsExpected = [
    ...['altText', 'rounding', 'transparency'].map((property) => ({
      id: addTablePropertyId('ImageProps', property),
      status: 'supported',
    })),
    ...['objectName', 'shadow'].map((property) => ({
      id: addTablePropertyId('ImageProps', property),
      status: 'deliberate-difference',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const imageIdentityEffectsIds = new Set(
    imageIdentityEffectsExpected.map(({ id }) => id),
  );
  assert.equal(imageIdentityEffectsIds.size, 5);
  const imageIdentityEffectsEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => imageIdentityEffectsIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    imageIdentityEffectsEntries.map(({ id, status }) => ({ id, status })),
    imageIdentityEffectsExpected,
  );
  assert.equal(
    imageIdentityEffectsEntries.every(({
      native,
      evidence,
      control,
      serialization,
      client,
    }) => native.length >= 3
      && evidence.code.length === 2
      && evidence.tests.length === 2
      && evidence.package.some(
        ({ pattern }) => pattern === 'const imageIdentityEffects5Probe =',
      )
      && evidence.ooxml.some(({ pattern }) => pattern === 'const exactOoxml = {')
      && evidence.clients.some(
        ({ pattern }) => pattern === 'const imageIdentityEffects5State = {',
      )
      && control.pattern ===
        'matches PptxGenJS image identity and visual effects through create edit reopen'
      && serialization === true
      && client === true),
    true,
  );
  assert.deepEqual(
    Object.fromEntries(['supported', 'deliberate-difference'].map((status) => [
      status,
      imageIdentityEffectsEntries.filter((entry) => entry.status === status).length,
    ])),
    { supported: 3, 'deliberate-difference': 2 },
  );
  assert.equal(
    PPTXGENJS_SURFACE_MANIFEST.entries.some(
      ({ id }) => id === addTablePropertyId('ImageProps', 'hyperlink'),
    ),
    false,
  );
  const shapeTextTransformIdentityExpected = [
    ...['flipH', 'flipV', 'objectName', 'rectRadius', 'rotate'].map((property) => ({
      id: addTablePropertyId('ShapeProps', property),
      status: 'deliberate-difference',
    })),
    {
      id: addTablePropertyId('ShapeProps', 'shapeName'),
      status: 'defect-excluded',
    },
    ...['flipH', 'flipV', 'objectName', 'rectRadius', 'rotate'].map((property) => ({
      id: addTablePropertyId('TextPropsOptions', property),
      status: 'deliberate-difference',
    })),
    ...['isTextBox', 'shape'].map((property) => ({
      id: addTablePropertyId('TextPropsOptions', property),
      status: 'supported',
    })),
  ].sort((left, right) => left.id.localeCompare(right.id));
  const shapeTextTransformIdentityIds = new Set(
    shapeTextTransformIdentityExpected.map(({ id }) => id),
  );
  assert.equal(shapeTextTransformIdentityIds.size, 13);
  const shapeTextTransformIdentityEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => shapeTextTransformIdentityIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    shapeTextTransformIdentityEntries.map(({ id, status }) => ({ id, status })),
    shapeTextTransformIdentityExpected,
  );
  const shapeTextTransformIdentityImplemented = shapeTextTransformIdentityEntries
    .filter(({ status }) => status !== 'defect-excluded');
  assert.equal(
    shapeTextTransformIdentityImplemented.every(({
      native,
      evidence,
      control,
      serialization,
      client,
    }) => native.length >= 2
      && evidence.code.length >= 2
      && evidence.tests.length >= 3
      && evidence.package.some(
        ({ pattern }) => pattern === 'const shapeTextTransformIdentity13Probe =',
      )
      && evidence.ooxml.some(({ pattern }) => pattern === 'const exactOoxml = {')
      && evidence.clients.some(
        ({ pattern }) => pattern === 'const shapeTextTransformIdentity13State = {',
      )
      && control.pattern ===
        'closes PptxGenJS shape and text transform identity through strict native state'
      && serialization === true
      && client === true),
    true,
  );
  const shapeNameDefect = shapeTextTransformIdentityEntries.find(
    ({ id }) => id === addTablePropertyId('ShapeProps', 'shapeName'),
  );
  assert.deepEqual({
    native: shapeNameDefect?.native,
    code: shapeNameDefect?.evidence.code,
    package: shapeNameDefect?.evidence.package,
    ooxml: shapeNameDefect?.evidence.ooxml,
    clients: shapeNameDefect?.evidence.clients,
    control: shapeNameDefect?.control.pattern,
  }, {
    native: [],
    code: [],
    package: [],
    ooxml: [],
    clients: [],
    control: 'closes PptxGenJS shape and text transform identity through strict native state',
  });
  assert.deepEqual(
    Object.fromEntries([
      'supported',
      'deliberate-difference',
      'defect-excluded',
    ].map((status) => [
      status,
      shapeTextTransformIdentityEntries.filter((entry) => entry.status === status).length,
    ])),
    { supported: 2, 'deliberate-difference': 10, 'defect-excluded': 1 },
  );
  const mediaCoreExpected = [
    ...['cover', 'data', 'extn', 'h', 'link', 'objectName', 'path', 'type', 'w', 'x', 'y']
      .map((property) => addTablePropertyId('MediaProps', property)),
    ...['audio', 'online', 'video'].map((value) => `union:MediaType#${value}`),
  ].sort();
  const mediaCoreIds = new Set(mediaCoreExpected);
  assert.equal(mediaCoreIds.size, 14);
  const mediaCoreEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => mediaCoreIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    mediaCoreEntries.map(({ id, status }) => ({ id, status })),
    mediaCoreExpected.map((id) => ({ id, status: 'deliberate-difference' })),
  );
  assert.equal(
    mediaCoreEntries.every(({ native, evidence, control, serialization, client }) =>
      native.length > 0 && evidence.code.length >= 3 && evidence.tests.length === 4 &&
      evidence.package.some(({ pattern }) => pattern === 'const packedOnlineVideoState =') &&
      evidence.ooxml.some(({ pattern }) =>
        pattern === 'creates every public media source, MIME family, poster family, and external mode') &&
      evidence.clients.some(({ pattern }) => pattern === 'const browserOnlineVideoState =') &&
      control.pattern ===
        'locks MediaProps source, type, metadata, and geometry against PptxGenJS 4.0.1' &&
      serialization === true && client === true),
    true,
  );
  const mediaCoreById = new Map(mediaCoreEntries.map((entry) => [entry.id, entry]));
  for (const id of [addTablePropertyId('MediaProps', 'link'), 'union:MediaType#online']) {
    assert.equal(mediaCoreById.get(id)?.native.includes('MediaModel.externalUrl'), true);
    assert.equal(mediaCoreById.get(id)?.native.some((name) => name.includes('addOnlineVideo')), false);
  }
  assert.equal(
    [
      addTablePropertyId('DataOrPathProps', 'data'),
      addTablePropertyId('DataOrPathProps', 'path'),
      addTablePropertyId('ImageProps', 'objectName'),
      addTablePropertyId('PlaceholderProps', 'type'),
    ].some((id) => mediaCoreIds.has(id)),
    false,
  );
  const addMediaEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    id === addTablePropertyId('PresSlide', 'addMedia') || id === 'method:Slide#addMedia');
  assert.equal(addMediaEntries.length, 2);
  assert.equal(
    addMediaEntries.every(({ native, note }) =>
      native.every((name) => !name.includes('addOnlineVideo')) &&
      !note.includes('addOnlineVideo')),
    true,
  );
  const placeholderCoreExpected = [
    ...['name', 'type', 'x', 'y', 'w', 'h']
      .map((property) => addTablePropertyId('PlaceholderProps', property)),
    addTablePropertyId('ImageProps', 'placeholder'),
    addTablePropertyId('TextPropsOptions', 'placeholder'),
  ].sort();
  const placeholderCoreIds = new Set(placeholderCoreExpected);
  assert.equal(placeholderCoreIds.size, 8);
  const placeholderCoreEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => placeholderCoreIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    placeholderCoreEntries.map(({ id, status }) => ({ id, status })),
    placeholderCoreExpected.map((id) => ({
      id,
      status: 'deliberate-difference',
    })),
  );
  assert.equal(
    placeholderCoreEntries.every(({ native, serialization, client }) =>
      native.length > 0 && serialization === true && client === true),
    true,
  );
  assert.equal(
    placeholderCoreEntries.every(({ evidence, control }) =>
      evidence.package.some(({ pattern }) => pattern === 'const masterLayoutChecks = {') &&
      evidence.clients.some(({ pattern }) => pattern === 'const masterLayoutState = {') &&
      control.pattern ===
        'locks PlaceholderProps and text/image placeholder population against PptxGenJS 4.0.1'),
    true,
  );
  assert.deepEqual(
    placeholderCoreEntries
      .filter(({ id }) => /PlaceholderProps@property:(?:w|h)$/u.test(id))
      .map(({ native }) => native[0])
      .sort(),
    ['AddPlaceholderOptions.height', 'AddPlaceholderOptions.width'],
  );
  assert.equal(
    placeholderCoreIds.has(addTablePropertyId('PlaceholderProps', 'align')) ||
      placeholderCoreIds.has(addTablePropertyId('PlaceholderProps', 'margin')) ||
      placeholderCoreIds.has(addTablePropertyId('PlaceholderProps', 'transparency')) ||
      placeholderCoreIds.has(addTablePropertyId('PlaceholderProps', 'valign')),
    false,
  );
  const shapeCustomPathExpected = [
    addTablePropertyId('ShapeProps', 'points'),
    ...[
      'close',
      'curve',
      'curve.hR',
      'curve.stAng',
      'curve.swAng',
      'curve.type',
      'curve.wR',
      'curve.x1',
      'curve.x2',
      'curve.y1',
      'curve.y2',
      'moveTo',
      'x',
      'y',
    ].map((property) =>
      `inline:interface:ShapeProps@property:points@property:points.${property}`),
  ].sort();
  const shapeCustomPathIds = new Set(shapeCustomPathExpected);
  assert.equal(shapeCustomPathIds.size, 15);
  const shapeCustomPathEntries = PPTXGENJS_SURFACE_MANIFEST.entries
    .filter(({ id }) => shapeCustomPathIds.has(id))
    .sort((left, right) => left.id.localeCompare(right.id));
  assert.deepEqual(
    shapeCustomPathEntries.map(({ id, status }) => ({ id, status })),
    shapeCustomPathExpected.map((id) => ({
      id,
      status: 'deliberate-difference',
    })),
  );
  assert.equal(
    shapeCustomPathEntries.every(({
      native,
      evidence,
      control,
      serialization,
      client,
    }) => native.length > 0 && evidence.code.length === 3 &&
      evidence.tests.length === 4 &&
      evidence.package.some(({ pattern }) => pattern === 'const customGeometryPaths =') &&
      evidence.ooxml.some(({ pattern }) =>
        pattern === 'creates and reopens styled custom geometry shapes through the public SDK') &&
      evidence.clients.some(({ pattern }) => pattern === 'const browserCustomPathState = {') &&
      control.pattern ===
        'classifies PptxGenJS custom path unit heuristics and malformed runtime output' &&
      serialization === true && client === true),
    true,
  );
  const shapeCustomPathById = new Map(
    shapeCustomPathEntries.map((entry) => [entry.id, entry]),
  );
  assert.deepEqual(
    shapeCustomPathById.get(
      'inline:interface:ShapeProps@property:points@property:points.curve.x2',
    )?.native,
    ['CustomGeometryCommand.cubicBezierTo'],
  );
  assert.equal(
    [
      addTablePropertyId('ShapeProps', 'angleRange'),
      addTablePropertyId('ShapeProps', 'arcThicknessRatio'),
      addTablePropertyId('ShapeProps', 'rectRadius'),
    ].some((id) => shapeCustomPathIds.has(id)),
    false,
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => id.endsWith('#folderCorner'))
      .map(({ id, status }) => ({ id, status })),
    [
      { id: 'union:ShapeType#folderCorner', status: 'defect-excluded' },
      { id: 'union:SHAPE_NAME#folderCorner', status: 'defect-excluded' },
    ],
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => /^union:PLACEHOLDER_TYPES?#(?:pic|tbl)$/u.test(id))
      .map(({ id, status }) => ({ id, status })),
    [
      { id: 'union:PLACEHOLDER_TYPE#pic', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPE#tbl', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPES#pic', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPES#tbl', status: 'deliberate-difference' },
    ],
  );
  assert.deepEqual(PPTXGENJS_SURFACE_MANIFEST.extensions, []);
  assertDeepFrozen(PPTXGENJS_SURFACE_MANIFEST);
});

test('defaults missing declaration entries to unverified without double-counting', async () => {
  await withRepository(async (repositoryRoot) => {
    const commits = [];
    const report = await buildPptxGenJSAudit({
      surface: surface(),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', {
          native: ['Deck.write'],
          evidence: validEvidence(),
          serialization: true,
          client: true,
        }),
      ]),
      repositoryRoot,
      gitCommitExists: async (commit) => {
        commits.push(commit);
        return commit === 'abc1234';
      },
    });

    assert.deepEqual(report.counts, {
      supported: 1,
      'deliberate-difference': 0,
      'deprecated-alias': 0,
      'defect-excluded': 0,
      unsupported: 0,
      unverified: 2,
      stale: 0,
    });
    assert.equal(report.declarationTotal, 3);
    assert.equal(report.complete, false);
    assert.deepEqual(report.incompleteIds, [IDS[1], IDS[2]]);
    assert.deepEqual(report.diagnostics, []);
    assert.deepEqual(commits, ['abc1234']);
    assertDeepFrozen(report);
  });
});

test('accepts every explicit non-stale status while open statuses still fail completion', async () => {
  await withRepository(async (repositoryRoot) => {
    const evidence = validEvidence();
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', { native: ['Deck.write'], evidence }),
        entry(IDS[1], 'deliberate-difference', {
          native: ['Slide.addText'], evidence, control: validControl,
        }),
        entry(IDS[2], 'deprecated-alias', {
          native: ['Deck.write'], evidence, canonical: IDS[0], control: validControl,
        }),
        entry(IDS[3], 'defect-excluded', {
          evidence, control: validControl,
        }),
        entry(IDS[4], 'unsupported', { control: validControl }),
        entry(IDS[5], 'unverified'),
      ]),
      repositoryRoot,
      gitCommitExists: async () => true,
    });

    assert.deepEqual(report.counts, {
      supported: 1,
      'deliberate-difference': 1,
      'deprecated-alias': 1,
      'defect-excluded': 1,
      unsupported: 1,
      unverified: 1,
      stale: 0,
    });
    assert.deepEqual(report.diagnostics, []);
    assert.deepEqual(report.incompleteIds, [IDS[4], IDS[5]].sort());
    assert.equal(report.complete, false);
  });
});

test('reports status-specific missing evidence with stable diagnostics', async () => {
  await withRepository(async (repositoryRoot) => {
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS.slice(0, 5)),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', { serialization: true, client: true }),
        entry(IDS[1], 'deliberate-difference'),
        entry(IDS[2], 'deprecated-alias', { canonical: 'class:Deck#missing' }),
        entry(IDS[3], 'defect-excluded'),
        entry(IDS[4], 'unsupported', { note: '' }),
      ]),
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    const codes = report.diagnostics.map(({ code }) => code);
    for (const code of [
      'invalid-canonical',
      'missing-client-evidence',
      'missing-code-evidence',
      'missing-control',
      'missing-native',
      'missing-note',
      'missing-ooxml-evidence',
      'missing-package-evidence',
      'missing-test-evidence',
    ]) {
      assert.equal(codes.includes(code), true, code);
    }
    assert.deepEqual(
      report.diagnostics,
      [...report.diagnostics].sort((left, right) => (
        left.id.localeCompare(right.id)
        || left.code.localeCompare(right.code)
        || left.message.localeCompare(right.message)
      )),
    );
    assert.equal(report.complete, false);
  });
});

test('rejects duplicate IDs, illegal status, unknown fields, accessors, and class data', async () => {
  await withRepository(async (repositoryRoot) => {
    const build = (candidate) => buildPptxGenJSAudit({
      surface: surface(),
      runtimeProbe,
      manifest: candidate,
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    await assert.rejects(
      build(manifest([entry(IDS[0], 'unverified'), entry(IDS[0], 'unverified')])),
      /duplicate manifest entry/u,
    );
    await assert.rejects(
      build(manifest([entry(IDS[0], 'stale')])),
      /invalid status stale/u,
    );
    await assert.rejects(
      build({ ...manifest(), schemaVersion: 2 }),
      /manifest schemaVersion must be 1/u,
    );
    await assert.rejects(
      build({ ...manifest(), packageVersion: '4.1.0' }),
      /manifest packageVersion must be 4\.0\.1/u,
    );
    await assert.rejects(
      build(manifest([{ ...entry(IDS[0], 'unverified'), surprise: true }])),
      /unknown key surprise/u,
    );
    const accessorEntry = entry(IDS[0], 'unverified');
    Object.defineProperty(accessorEntry, 'note', { enumerable: true, get: () => 'unsafe' });
    await assert.rejects(build(manifest([accessorEntry])), /note must be a data property/u);
    const accessorLink = { path: 'src/write.ts', pattern: 'write' };
    Object.defineProperty(accessorLink, 'pattern', { enumerable: true, get: () => 'unsafe' });
    await assert.rejects(
      build(manifest([entry(IDS[0], 'unverified', {
        evidence: { ...emptyEvidence(), code: [accessorLink] },
      })])),
      /pattern must be a data property/u,
    );

    class ManifestData {
      schemaVersion = 1;
      packageVersion = '4.0.1';
      entries = [];
      extensions = [];
    }
    await assert.rejects(build(new ManifestData()), /manifest must be a plain data object/u);
  });
});

test('diagnoses invalid paths, missing literals, missing files, and commit objects', async () => {
  await withRepository(async (repositoryRoot) => {
    const evidence = validEvidence();
    evidence.code = [
      { path: 'src/missing.ts', pattern: 'missing', commit: 'deadbee' },
      { path: 'src/write.ts', pattern: 'absent literal' },
      { path: resolve(repositoryRoot, 'src', 'write.ts'), pattern: 'export function write' },
    ];
    evidence.tests = [{ path: 'tests/write.test.ts', title: 'absent title' }];
    evidence.package = [{ path: '../outside.mjs', pattern: 'outside' }];
    const report = await buildPptxGenJSAudit({
      surface: surface([IDS[0]]),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', {
          native: ['Deck.write'], evidence, serialization: true, client: true,
        }),
      ]),
      repositoryRoot,
      gitCommitExists: async () => false,
    });
    const codes = new Set(report.diagnostics.map(({ code }) => code));
    for (const code of [
      'evidence-commit-missing',
      'evidence-file-missing',
      'evidence-path-invalid',
      'evidence-pattern-missing',
      'evidence-title-missing',
    ]) {
      assert.equal(codes.has(code), true, code);
    }
    assert.deepEqual(report.incompleteIds, [IDS[0]]);
    assert.equal(report.complete, false);
  });
});

test('verifies real Git commit objects with the default checker', async () => {
  const evidence = emptyEvidence();
  evidence.code = [{
    path: 'scripts/pptxgenjs-surface-audit-lib.mjs',
    pattern: 'buildPptxGenJSAudit',
    commit: '0b6094e',
  }];
  evidence.tests = [{
    path: 'scripts/pptxgenjs-surface-audit-lib.test.mjs',
    title: 'verifies real Git commit objects',
    commit: '0b6094e',
  }];
  evidence.package = [{
    path: 'package.json',
    pattern: 'typecheck',
    commit: '0b6094e',
  }];
  const report = await buildPptxGenJSAudit({
    surface: surface([IDS[0]]),
    runtimeProbe,
    manifest: manifest([
      entry(IDS[0], 'supported', { native: ['Deck.write'], evidence }),
    ]),
    repositoryRoot: resolve('.'),
  });
  assert.deepEqual(report.diagnostics, []);
  assert.equal(report.complete, true);
});

test('keeps stale entries and extensions outside the declaration denominator', async () => {
  await withRepository(async (repositoryRoot) => {
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS.slice(0, 2)),
      runtimeProbe,
      manifest: manifest(
        [entry('class:Deck#removed', 'unverified')],
        [{
          id: 'extension:PptxDocument#strictWrite',
          native: ['PptxDocument.write'],
          evidence: emptyEvidence(),
          note: 'strict native write mode',
        }],
      ),
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    assert.equal(report.declarationTotal, 2);
    assert.equal(report.counts.unverified, 2);
    assert.equal(report.counts.stale, 1);
    assert.equal(report.extensions.length, 1);
    assert.deepEqual(report.incompleteIds, [
      'class:Deck#removed',
      IDS[0],
      IDS[1],
    ].sort());

    await assert.rejects(
      buildPptxGenJSAudit({
        surface: surface([IDS[0]]),
        runtimeProbe,
        manifest: manifest([entry('extension:collision', 'unverified')], [{
          id: 'extension:collision',
          native: ['Deck.write'],
          evidence: emptyEvidence(),
          note: 'collision',
        }]),
        repositoryRoot,
        gitCommitExists: async () => true,
      }),
      /extension ID collides with manifest entry/u,
    );
  });
});
