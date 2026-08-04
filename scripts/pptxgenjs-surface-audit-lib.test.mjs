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
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.entries.length, 1221);
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries.map(({ status }) => status).sort(),
    [
      ...Array(349).fill('defect-excluded'),
      ...Array(575).fill('supported'),
      ...Array(212).fill('deliberate-difference'),
      ...Array(85).fill('deprecated-alias'),
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
  assert.equal(
    PPTXGENJS_SURFACE_MANIFEST.entries.some(({ id }) => [
      'interface:OptsChartGridLine@property:cap',
      'union:ChartLineCap#flat',
      'union:ChartLineCap#round',
      'union:ChartLineCap#square',
      'interface:IChartOpts@property:serGridLine',
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
