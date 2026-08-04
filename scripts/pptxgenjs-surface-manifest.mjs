function evidence(code, tests, packageEvidence) {
  return {
    code: [code],
    tests: [tests],
    package: [packageEvidence],
    ooxml: [],
    clients: [],
  };
}

function supported(id, native, code, tests, packageEvidence, note) {
  return {
    id,
    status: 'supported',
    native,
    evidence: evidence(code, tests, packageEvidence),
    note,
  };
}

function serializedCatalogMember(
  id,
  native,
  code,
  tests,
  packageEvidence,
  ooxml,
  clientEvidence,
  note,
) {
  return {
    id,
    status: 'supported',
    native,
    evidence: {
      code: [code],
      tests: [tests],
      package: [packageEvidence],
      ooxml: [ooxml],
      clients: [clientEvidence],
    },
    serialization: true,
    client: true,
    note,
  };
}

const SHAPE_TEXT_COORDINATE_CONTROL_TITLE =
  'matches PptxGenJS shape and text percentage coordinate output with explicit native units';
const SHAPE_TEXT_COORDINATE_OOXML_TITLE =
  'creates and reopens shape and text percentage coordinates against the current slide size';
const IMAGE_COORDINATE_CONTROL_TITLE =
  'matches PptxGenJS image percentage coordinate output with explicit native units';
const IMAGE_COORDINATE_OOXML_TITLE =
  'creates and reopens raster and SVG percentage coordinates through the source loader';

function deliberateDifference(id, native) {
  return {
    id,
    status: 'deliberate-difference',
    native,
    evidence: {
      code: [{
        path: 'packages/model/src/slide-coordinate.internal.ts',
        pattern: 'export function resolveSlideCoordinate(',
      }],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: SHAPE_TEXT_COORDINATE_CONTROL_TITLE,
      }],
      package: [{
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const shapeTextPercentageCoordinates =',
      }],
      ooxml: [{
        path: 'packages/sdk/src/index.test.ts',
        pattern: SHAPE_TEXT_COORDINATE_OOXML_TITLE,
      }],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: SHAPE_TEXT_COORDINATE_CONTROL_TITLE,
    },
    serialization: true,
    note: 'Native covers the same legal percentage geometry with width/height and explicit Emu or inches() numeric units instead of PptxGenJS w/h and implicit-inch numbers.',
  };
}

const SHAPE_TEXT_COORDINATE_ATOMS = Object.freeze([
  'interface:PositionProps@property:x',
  'interface:PositionProps@property:y',
  'interface:PositionProps@property:w',
  'interface:PositionProps@property:h',
  'interface:ShapeProps@property:x',
  'interface:ShapeProps@property:y',
  'interface:ShapeProps@property:w',
  'interface:ShapeProps@property:h',
  'interface:TextPropsOptions@property:x',
  'interface:TextPropsOptions@property:y',
  'interface:TextPropsOptions@property:w',
  'interface:TextPropsOptions@property:h',
]);

const IMAGE_COORDINATE_ATOMS = Object.freeze([
  'interface:ImageProps@property:x',
  'interface:ImageProps@property:y',
  'interface:ImageProps@property:w',
  'interface:ImageProps@property:h',
]);

const DECLARED_PRESET_SHAPE_VALUES = Object.freeze(`
  accentBorderCallout1 accentBorderCallout2 accentBorderCallout3
  accentCallout1 accentCallout2 accentCallout3
  actionButtonBackPrevious actionButtonBeginning actionButtonBlank
  actionButtonDocument actionButtonEnd actionButtonForwardNext
  actionButtonHelp actionButtonHome actionButtonInformation actionButtonMovie
  actionButtonReturn actionButtonSound arc bentArrow bentUpArrow bevel blockArc
  borderCallout1 borderCallout2 borderCallout3 bracePair bracketPair callout1
  callout2 callout3 can chartPlus chartStar chartX chevron chord circularArrow
  cloud cloudCallout corner cornerTabs cube curvedDownArrow curvedLeftArrow
  curvedRightArrow curvedUpArrow decagon diagStripe diamond dodecagon donut
  doubleWave downArrow downArrowCallout ellipse ellipseRibbon ellipseRibbon2
  flowChartAlternateProcess flowChartCollate flowChartConnector flowChartDecision
  flowChartDelay flowChartDisplay flowChartDocument flowChartExtract
  flowChartInputOutput flowChartInternalStorage flowChartMagneticDisk
  flowChartMagneticDrum flowChartMagneticTape flowChartManualInput
  flowChartManualOperation flowChartMerge flowChartMultidocument
  flowChartOfflineStorage flowChartOffpageConnector flowChartOnlineStorage
  flowChartOr flowChartPredefinedProcess flowChartPreparation flowChartProcess
  flowChartPunchedCard flowChartPunchedTape flowChartSort
  flowChartSummingJunction flowChartTerminator folderCorner frame funnel gear6
  gear9 halfFrame heart heptagon hexagon homePlate horizontalScroll
  irregularSeal1 irregularSeal2 leftArrow leftArrowCallout leftBrace leftBracket
  leftCircularArrow leftRightArrow leftRightArrowCallout leftRightCircularArrow
  leftRightRibbon leftRightUpArrow leftUpArrow lightningBolt line lineInv
  mathDivide mathEqual mathMinus mathMultiply mathNotEqual mathPlus moon
  noSmoking nonIsoscelesTrapezoid notchedRightArrow octagon parallelogram
  pentagon pie pieWedge plaque plaqueTabs plus quadArrow quadArrowCallout rect
  ribbon ribbon2 rightArrow rightArrowCallout rightBrace rightBracket round1Rect
  round2DiagRect round2SameRect roundRect rtTriangle smileyFace snip1Rect
  snip2DiagRect snip2SameRect snipRoundRect squareTabs star10 star12 star16
  star24 star32 star4 star5 star6 star7 star8 stripedRightArrow sun swooshArrow
  teardrop trapezoid triangle upArrow upArrowCallout upDownArrow
  upDownArrowCallout uturnArrow verticalScroll wave wedgeEllipseCallout
  wedgeRectCallout wedgeRoundRectCallout
`.trim().split(/\s+/u));

const SCHEME_COLOR_VALUES = Object.freeze([
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'bg1', 'bg2', 'tx1', 'tx2',
]);
const HORIZONTAL_ALIGNMENT_VALUES = Object.freeze(['center', 'justify', 'left', 'right']);
const VERTICAL_ALIGNMENT_VALUES = Object.freeze(['bottom', 'middle', 'top']);
const OUTPUT_TYPE_VALUES = Object.freeze([
  'arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array',
]);
const PLACEHOLDER_TYPE_VALUES = Object.freeze([
  'body', 'chart', 'media', 'pic', 'tbl', 'title',
]);
const LINE_DASH_VALUES = Object.freeze([
  'solid', 'dash', 'dashDot', 'lgDash',
  'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot',
]);
const LINE_ARROW_VALUES = Object.freeze([
  'none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle',
]);
const LINE_TYPE_VALUES = Object.freeze(['none', 'solid']);
const LINE_ALIAS_CONTROL_TITLE =
  'locks every deprecated line alias against its owner-specific runtime behavior';

function linePropertyId(owner, property) {
  return `interface:${owner}@property:${property}`;
}

function lineUnionId(owner, property, value) {
  return `union:interface:${owner}@property:${property}#${value}`;
}

function lineFamilyEvidence(id, deprecated = false) {
  const arrow = /(?:beginArrowType|endArrowType|lineHead|lineTail)/u.test(id);
  const textOwner = id.includes('TextPropsOptions');
  const shapeOwner = id.includes('ShapeProps');
  const sharedOwner = id.includes('ShapeLineProps');
  const controlTitle = deprecated
    ? LINE_ALIAS_CONTROL_TITLE
    : arrow
      ? textOwner
        ? 'compares text shape arrows public output and strict native divergences'
        : 'compares shape arrow public output and strict native divergences'
      : textOwner
        ? 'compares text shape line public output and strict native divergences'
        : 'compares shape line public output and strict native divergences';
  const tests = deprecated
    ? [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: controlTitle }]
    : [
        ...(!textOwner ? [{
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: arrow
            ? 'compares shape arrow public output and strict native divergences'
            : 'compares shape line public output and strict native divergences',
        }] : []),
        ...(!shapeOwner ? [{
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: arrow
            ? 'compares text shape arrows public output and strict native divergences'
            : 'compares text shape line public output and strict native divergences',
        }] : []),
      ];
  const packageEvidence = deprecated
    ? [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const deprecatedLineAliases =' }]
    : [
        ...(!textOwner ? [{
          path: 'scripts/smoke-npm-package.mjs',
          pattern: arrow ? 'const shapeArrows =' : 'const shapeLines =',
        }] : []),
        ...(!shapeOwner ? [{
          path: 'scripts/smoke-npm-package.mjs',
          pattern: arrow ? 'const textShapeArrows =' : 'const textShapeLines =',
        }] : []),
      ];
  const ooxml = [
    ...(!textOwner ? [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: arrow
        ? 'creates preset shape arrows through the public SDK surface'
        : 'creates preset shape lines through the public SDK surface',
    }] : []),
    ...(!shapeOwner ? [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: arrow
        ? 'creates text arrows across slide layout master and placeholder owners'
        : 'creates text lines across slide layout master and placeholder owners',
    }] : []),
  ];
  const clients = deprecated
    ? []
    : [
        ...(!textOwner ? [{
          path: 'scripts/playwright-browser-smoke.js',
          pattern: arrow
            ? 'const shapeArrows = JSON.stringify(shapeArrowState)'
            : 'const shapeLines = JSON.stringify(shapeLineState)',
        }] : []),
        ...(!shapeOwner ? [{
          path: 'scripts/playwright-browser-smoke.js',
          pattern: arrow
            ? 'const textShapeArrowState = {'
            : 'const textShapeLineState = {',
        }] : []),
      ];
  return {
    arrow,
    control: { path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: controlTitle },
    evidence: {
      code: [{
        path: arrow
          ? 'packages/model/src/shape-arrows.internal.ts'
          : 'packages/model/src/simple-line.internal.ts',
        pattern: arrow
          ? 'export function normalizeShapeArrows('
          : 'export function normalizeSimpleLine(',
      }],
      tests,
      package: packageEvidence,
      ooxml,
      clients,
    },
    sharedOwner,
    textOwner,
  };
}

function canonicalLineEntry(id) {
  const { arrow, control, evidence, sharedOwner, textOwner } = lineFamilyEvidence(id);
  return {
    id,
    status: 'deliberate-difference',
    native: [
      arrow ? 'ShapeArrows' : 'ShapeLine',
      arrow ? 'ShapeModel.arrows' : 'ShapeModel.line',
      ...(sharedOwner
        ? ['SlideModel.addShape', 'SlideModel.addText']
        : [textOwner ? 'SlideModel.addText' : 'SlideModel.addShape']),
    ],
    evidence,
    control,
    serialization: true,
    client: true,
    note: arrow
      ? 'Native covers the same six legal begin/end arrow tokens through strict ShapeArrows begin/end fields, lossless editing, and deterministic OOXML instead of PptxGenJS line aliases and permissive passthrough.'
      : 'Native covers the same legal none/solid line semantics, colors, transparency, width, and eight dash tokens through a strict ShapeLine contract with deterministic direct OOXML instead of PptxGenJS fallbacks and permissive aliases.',
  };
}

function deprecatedLineEntry(owner, alias, canonicalProperty, value) {
  const id = value === undefined
    ? linePropertyId(owner, alias)
    : lineUnionId(owner, alias, value);
  const canonical = value === undefined
    ? linePropertyId('ShapeLineProps', canonicalProperty)
    : lineUnionId('ShapeLineProps', canonicalProperty, value);
  const { arrow, control, evidence, sharedOwner, textOwner } = lineFamilyEvidence(id, true);
  let note;
  if (sharedOwner && alias === 'alpha') {
    note = 'PptxGenJS declares alpha as a deprecated transparency alias, but 4.0.1 applies it only to ordinary text lines and ignores it for shapes and line-shaped text; native rejects the alias and exposes the strict canonical transparency field.';
  } else if (sharedOwner) {
    note = `PptxGenJS declares ${alias} as a deprecated ${canonicalProperty} alias but 4.0.1 ignores it in nested shape and text line objects; native rejects the alias and exposes only the strict canonical field.`;
  } else if (textOwner) {
    note = `PptxGenJS declares top-level ${alias} as a deprecated nested-line alias, ignores it for ordinary text, and applies it to line-shaped text; the native type contract rejects the alias while direct JavaScript calls leave it inert, exposing only the strict canonical line/arrows field.`;
  } else {
    note = `PptxGenJS maps top-level ${alias} to the canonical nested-line field and lets it override modern nested input; native rejects the deprecated alias and exposes only the strict canonical line/arrows field.`;
  }
  return {
    id,
    status: 'deprecated-alias',
    native: [
      arrow ? 'ShapeArrows' : 'ShapeLine',
      arrow ? 'ShapeModel.arrows' : 'ShapeModel.line',
      ...(sharedOwner
        ? ['SlideModel.addShape', 'SlideModel.addText']
        : [textOwner ? 'SlideModel.addText' : 'SlideModel.addShape']),
    ],
    evidence,
    control,
    canonical,
    serialization: true,
    note,
  };
}

const CANONICAL_LINE_ATOM_IDS = Object.freeze([
  ...[
    'beginArrowType', 'color', 'dashType', 'endArrowType',
    'transparency', 'type', 'width',
  ].map((property) => linePropertyId('ShapeLineProps', property)),
  ...LINE_ARROW_VALUES.map((value) =>
    lineUnionId('ShapeLineProps', 'beginArrowType', value)),
  ...LINE_DASH_VALUES.map((value) =>
    lineUnionId('ShapeLineProps', 'dashType', value)),
  ...LINE_ARROW_VALUES.map((value) =>
    lineUnionId('ShapeLineProps', 'endArrowType', value)),
  ...LINE_TYPE_VALUES.map((value) =>
    lineUnionId('ShapeLineProps', 'type', value)),
  linePropertyId('ShapeProps', 'line'),
  linePropertyId('TextPropsOptions', 'line'),
]);

const DEPRECATED_LINE_ENTRIES = Object.freeze([
  deprecatedLineEntry('ShapeLineProps', 'alpha', 'transparency'),
  ...LINE_DASH_VALUES.map((value) =>
    deprecatedLineEntry('ShapeLineProps', 'lineDash', 'dashType', value)),
  deprecatedLineEntry('ShapeLineProps', 'lineDash', 'dashType'),
  ...LINE_ARROW_VALUES.map((value) =>
    deprecatedLineEntry('ShapeLineProps', 'lineHead', 'beginArrowType', value)),
  deprecatedLineEntry('ShapeLineProps', 'lineHead', 'beginArrowType'),
  ...LINE_ARROW_VALUES.map((value) =>
    deprecatedLineEntry('ShapeLineProps', 'lineTail', 'endArrowType', value)),
  deprecatedLineEntry('ShapeLineProps', 'lineTail', 'endArrowType'),
  deprecatedLineEntry('ShapeLineProps', 'pt', 'width'),
  deprecatedLineEntry('ShapeLineProps', 'size', 'width'),
  ...['ShapeProps', 'TextPropsOptions'].flatMap((owner) => [
    ...LINE_DASH_VALUES.map((value) =>
      deprecatedLineEntry(owner, 'lineDash', 'dashType', value)),
    deprecatedLineEntry(owner, 'lineDash', 'dashType'),
    ...LINE_ARROW_VALUES.map((value) =>
      deprecatedLineEntry(owner, 'lineHead', 'beginArrowType', value)),
    deprecatedLineEntry(owner, 'lineHead', 'beginArrowType'),
    deprecatedLineEntry(owner, 'lineSize', 'width'),
    ...LINE_ARROW_VALUES.map((value) =>
      deprecatedLineEntry(owner, 'lineTail', 'endArrowType', value)),
    deprecatedLineEntry(owner, 'lineTail', 'endArrowType'),
  ]),
]);
const FILL_CONTROL_TITLE =
  'locks fill boundaries and deprecated alpha across shape and text owners';
const SHAPE_FILL_CONTROL_TITLE =
  'compares shape fill public output and strict native divergences';
const TEXT_SHAPE_FILL_CONTROL_TITLE =
  'compares text shape fill public output and strict native divergences';

function fillFamilyEvidence(id, deprecated = false) {
  const textOwner = id.includes('TextPropsOptions');
  const shapeOwner = id.includes('ShapeProps');
  const packageEvidence = [
    ...(!textOwner ? [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const shapeFills =',
    }] : []),
    ...(!shapeOwner ? [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const textShapeFills =',
    }] : []),
  ];
  const ooxml = [
    ...(!textOwner ? [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: 'creates preset shape fills through the public SDK surface',
    }, {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: "const shapeFillDeckPath = join(directory, 'shape-fill-smoke.pptx');",
    }] : []),
    ...(!shapeOwner ? [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: 'creates text fills across slide layout master and placeholder owners',
    }, {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: "const textShapeFillDeckPath = join(directory, 'text-shape-fill-smoke.pptx');",
    }] : []),
  ];
  const clients = deprecated
    ? []
    : [
        ...(!textOwner ? [{
          path: 'scripts/playwright-browser-smoke.js',
          pattern: 'const shapeFills = JSON.stringify(shapeFillState)',
        }] : []),
        ...(!shapeOwner ? [{
          path: 'scripts/playwright-browser-smoke.js',
          pattern: 'const textShapeFillState = {',
        }] : []),
      ];
  return {
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: deprecated
        ? FILL_CONTROL_TITLE
        : textOwner
          ? TEXT_SHAPE_FILL_CONTROL_TITLE
          : SHAPE_FILL_CONTROL_TITLE,
    },
    evidence: {
      code: [
        {
          path: 'packages/model/src/simple-fill.internal.ts',
          pattern: 'export function normalizeSimpleFill(',
        },
        {
          path: 'packages/model/src/shape-fill.internal.ts',
          pattern: 'export function readShapeFill(',
        },
      ],
      tests: [
        {
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: FILL_CONTROL_TITLE,
        },
        ...(!deprecated && !textOwner ? [{
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: SHAPE_FILL_CONTROL_TITLE,
        }] : []),
        ...(!deprecated && !shapeOwner ? [{
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: TEXT_SHAPE_FILL_CONTROL_TITLE,
        }] : []),
      ],
      package: packageEvidence,
      ooxml,
      clients,
    },
    shapeOwner,
    textOwner,
  };
}

function canonicalFillEntry(id) {
  const { control, evidence, shapeOwner, textOwner } = fillFamilyEvidence(id);
  return {
    id,
    status: 'deliberate-difference',
    native: [
      'ShapeFill',
      'ShapeModel.fill',
      ...(shapeOwner
        ? ['SlideModel.addShape']
        : textOwner
          ? ['SlideModel.addText']
          : ['SlideModel.addShape', 'SlideModel.addText']),
    ],
    evidence,
    control,
    serialization: true,
    client: true,
    note: 'Native covers the legal none/solid, sRGB/scheme, and transparency domain through a strict ShapeFill kind union with explicit direct-state editing; PptxGenJS instead omits none and zero-alpha intent and permits fallback values that native rejects before mutation.',
  };
}

const CANONICAL_FILL_ATOM_IDS = Object.freeze([
  ...['color', 'transparency', 'type'].map((property) =>
    linePropertyId('ShapeFillProps', property)),
  ...['none', 'solid'].map((value) =>
    lineUnionId('ShapeFillProps', 'type', value)),
  linePropertyId('ShapeProps', 'fill'),
  linePropertyId('TextPropsOptions', 'fill'),
]);
const DEPRECATED_FILL_ENTRY = Object.freeze((() => {
  const id = linePropertyId('ShapeFillProps', 'alpha');
  const { control, evidence } = fillFamilyEvidence(id, true);
  return {
    id,
    status: 'deprecated-alias',
    native: ['ShapeFill', 'ShapeModel.fill', 'SlideModel.addShape', 'SlideModel.addText'],
    evidence,
    control,
    canonical: linePropertyId('ShapeFillProps', 'transparency'),
    serialization: true,
    note: 'PptxGenJS declares alpha as a deprecated transparency alias across shape and text fills; when both nonzero fields are present it writes duplicate alpha children, while native rejects the alias and exposes only the strict canonical transparency field.',
  };
})());
const TABLE_FILL_CONTROLS = Object.freeze({
  TableCellProps: 'imports PptxGenJS table-cell fills from direct cell properties',
  TableProps: 'matches native table-level solid fill creation to PptxGenJS final state',
});

function tableFillEntry(owner) {
  const id = linePropertyId(owner, 'fill');
  const tableLevel = owner === 'TableProps';
  const title = TABLE_FILL_CONTROLS[owner];
  return {
    id,
    status: 'deliberate-difference',
    native: tableLevel
      ? ['TableCellFill', 'AddTableOptions.fill', 'TableModel.fill', 'SlideModel.addTable']
      : [
          'TableCellFill',
          'AddTableCellOptions.fill',
          'TableCell.fill',
          'TableModel.setCellFill',
          'SlideModel.addTable',
        ],
    evidence: {
      code: [{
        path: 'packages/model/src/table-cell-fill.internal.ts',
        pattern: 'export function normalizeTableCellFill(',
      }],
      tests: [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', title }],
      package: [{
        path: 'scripts/smoke-npm-package.mjs',
        pattern: tableLevel
          ? 'const tableFillDocument = PptxDocument.create();'
          : 'const initialCellFill = table?.rows[0]?.cells[0]?.fill;',
      }],
      ooxml: [{
        path: tableLevel
          ? 'scripts/smoke-npm-package.mjs'
          : 'packages/sdk/src/index.test.ts',
        pattern: tableLevel
          ? "const tableFillDeckPath = join(directory, 'table-fill-smoke.pptx');"
          : 'edits table-cell fills through duplicate, rollback, and reopen lifecycles',
      }],
      clients: [{
        path: 'scripts/playwright-browser-smoke.js',
        pattern: 'const tableFill = JSON.stringify(tableFillState)',
      }],
    },
    control: { path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: title },
    serialization: true,
    client: true,
    note: 'Native preserves the legal none/solid, sRGB/scheme, and transparency domain through strict TableCellFill creation and editing; PptxGenJS collapses explicit none and zero-alpha intent and permits malformed or out-of-range fill values that native rejects before mutation.',
  };
}

const CHART_AREA_FILL_LINE_CONTROL_TITLE =
  'locks chart area and plot area fill, border, rounding, and deprecated aliases';

function chartAreaFillLineEvidence(id) {
  const rounded = id === linePropertyId('IChartAreaProps', 'roundedCorners');
  const chartArea = id === linePropertyId('IChartOpts', 'chartArea');
  return {
    code: [
      ...(!rounded ? [{
        path: 'packages/model/src/chart-options.internal.ts',
        pattern: 'function normalizeAreaOptions(',
      }, {
        path: 'packages/model/src/chart-render.internal.ts',
        pattern: 'function renderAreaShapeProperties(',
      }, {
        path: 'packages/model/src/chart-state.internal.ts',
        pattern: 'function readAreaOptions(',
      }] : []),
      ...(rounded || chartArea ? [{
        path: 'packages/model/src/chart-options.internal.ts',
        pattern: 'export function normalizeChartOptions(',
      }, {
        path: 'packages/model/src/chart-render.internal.ts',
        pattern: 'export function renderChartPart(',
      }, {
        path: 'packages/model/src/chart-state.internal.ts',
        pattern: 'function readRootChartOptions(',
      }] : []),
    ],
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: CHART_AREA_FILL_LINE_CONTROL_TITLE,
    }],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const chartAreaFillLineState =',
    }],
    ooxml: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const nativeChartAreaFillLineFragments = [',
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const chartAreaFillLineState =',
    }],
  };
}

function chartAreaFillLineEntry(id, native, note) {
  return {
    id,
    status: 'deliberate-difference',
    native,
    evidence: chartAreaFillLineEvidence(id),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: CHART_AREA_FILL_LINE_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note,
  };
}

const CHART_AREA_FILL_LINE_ENTRIES = Object.freeze([
  chartAreaFillLineEntry(
    linePropertyId('IChartAreaProps', 'border'),
    ['ChartAreaOptions.line', 'ChartOptions.chartArea', 'ChartModel.replaceDefinition'],
    'Native covers chart-area none/solid lines, sRGB/scheme colors, transparency, width, and eight dash styles through a strict ShapeLine; PptxGenJS ignores the declared border type, substitutes defaults for falsy values, and permits invalid widths.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartAreaProps', 'fill'),
    ['ChartAreaOptions.fill', 'ChartOptions.chartArea', 'ChartModel.replaceDefinition'],
    'Native covers chart-area none/solid fills, sRGB/scheme colors, and transparency through a strict ShapeFill; PptxGenJS collapses several explicit none and zero-alpha forms and permits malformed or out-of-range values.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartAreaProps', 'roundedCorners'),
    ['ChartOptions.roundedCorners', 'ChartModel.replaceDefinition'],
    'Native exposes roundedCorners at the chart root where OOXML stores it; PptxGenJS nests the option under chartArea, defaults it to true, and serializes false as an explicit default state.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartPropsFillLine', 'border'),
    ['ChartAreaOptions.line', 'ChartOptions.plotArea', 'ChartModel.replaceDefinition'],
    'Native covers plot-area none/solid lines, sRGB/scheme colors, transparency, width, and eight dash styles through a strict ShapeLine; PptxGenJS ignores the declared border type, substitutes defaults for falsy values, and permits invalid widths.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartPropsFillLine', 'fill'),
    ['ChartAreaOptions.fill', 'ChartOptions.plotArea', 'ChartModel.replaceDefinition'],
    'Native covers plot-area none/solid fills, sRGB/scheme colors, and transparency through a strict ShapeFill; PptxGenJS collapses several explicit none and zero-alpha forms and permits malformed or out-of-range values.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartOpts', 'chartArea'),
    ['ChartAreaOptions', 'ChartOptions.chartArea', 'ChartOptions.roundedCorners'],
    'Native separates root roundedCorners from strict chart-area fill and line state, while PptxGenJS combines them in one permissive chartArea object with runtime defaults.',
  ),
  chartAreaFillLineEntry(
    linePropertyId('IChartOpts', 'plotArea'),
    ['ChartAreaOptions', 'ChartOptions.plotArea', 'ChartModel.replaceDefinition'],
    'Native exposes strict plot-area fill and line state through ChartOptions.plotArea; PptxGenJS accepts the same legal visual domain plus permissive and deprecated override behavior that native rejects.',
  ),
]);

function deprecatedChartAreaAliasEntry(property) {
  const border = property === 'border';
  return {
    id: linePropertyId('IChartOpts', property),
    status: 'deprecated-alias',
    native: border
      ? ['ChartAreaOptions.line', 'ChartOptions.plotArea', 'ChartModel.replaceDefinition']
      : ['ChartAreaOptions.fill', 'ChartOptions.plotArea', 'ChartModel.replaceDefinition'],
    evidence: chartAreaFillLineEvidence(linePropertyId('IChartOpts', property)),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: CHART_AREA_FILL_LINE_CONTROL_TITLE,
    },
    canonical: linePropertyId('IChartPropsFillLine', property),
    serialization: true,
    client: true,
    note: border
      ? 'PptxGenJS keeps top-level border as a deprecated plotArea.border alias and lets it completely replace the nested border; native rejects the alias and exposes only strict ChartOptions.plotArea.line.'
      : 'PptxGenJS keeps top-level fill as a deprecated plotArea.fill color alias and lets it replace the nested color while retaining nested transparency; native rejects the alias and exposes only strict ChartOptions.plotArea.fill.',
  };
}

const DEPRECATED_CHART_AREA_ALIAS_ENTRIES = Object.freeze([
  deprecatedChartAreaAliasEntry('border'),
  deprecatedChartAreaAliasEntry('fill'),
]);

const CHART_CREATION_CONTROL_TITLE =
  'matches all public slide master chart types and combo workbooks to native output';
const CHART_CREATION_RETURN_CONTROL_TITLE =
  'compares PptxGenJS and native chart creation return semantics';
const CHART_CREATION_NORMALIZATION_CONTROL_TITLE =
  'normalizes and semantically replaces imported PptxGenJS chart families';
const CHART_CREATION_OOXML_TITLE =
  'creates and reopens all native chart types through the public SDK in all six formats';
const CHART_TYPE_VALUES = Object.freeze([
  'area',
  'bar',
  'bar3D',
  'bubble',
  'doughnut',
  'line',
  'pie',
  'radar',
  'scatter',
]);

function chartCreationEvidence(
  id,
  testTitle = id === 'union:method:Slide#addChart@path:type#CHART_NAME'
    ? CHART_CREATION_RETURN_CONTROL_TITLE
    : CHART_CREATION_CONTROL_TITLE,
) {
  const catalog = /^union:(?:CHART_NAME|ChartType)#/u.test(id);
  const method = id === 'method:Slide#addChart'
    || id === 'union:method:Slide#addChart@path:type#CHART_NAME';
  const group = id.includes('IChartMulti');
  const code = catalog
    ? [{
        path: 'packages/model/src/chart.ts',
        pattern: 'export const CHART_TYPES = Object.freeze([',
      }]
    : method
      ? [{
          path: 'packages/model/src/slide.ts',
          pattern: 'return this.presentation.opcPackage.transaction(() => commitPreparedChart(this, prepared));',
        }]
      : group
        ? [{
            path: 'packages/model/src/chart.ts',
            pattern: 'export type ChartGroupInput =',
          }]
        : [{
            path: 'packages/model/src/chart.ts',
            pattern: 'export interface ChartSeriesInput {',
          }];
  if (testTitle === CHART_CREATION_NORMALIZATION_CONTROL_TITLE) {
    code.push({
      path: 'packages/model/src/chart-definition.internal.ts',
      pattern: "const GROUP_KEYS = new Set(['type', 'series', 'axis', 'options']);",
    }, {
      path: 'packages/model/src/chart-definition.internal.ts',
      pattern: "const SERIES_KEYS = new Set(['name', 'categories', 'values', 'xValues', 'sizes']);",
    });
  }
  return {
    code,
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: testTitle,
    }],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const nativeCharts = reopenedNativeChartModels.length === 10',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: CHART_CREATION_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const nativeCharts = reopenedCharts.length === 10',
    }],
  };
}

function chartCreationSupportedEntry(id, native, note) {
  return {
    id,
    status: 'supported',
    native,
    evidence: chartCreationEvidence(id),
    serialization: true,
    client: true,
    note,
  };
}

function chartCreationDifferenceEntry(id, native, note, returnSemantics = false) {
  const title = returnSemantics
    ? CHART_CREATION_RETURN_CONTROL_TITLE
    : CHART_CREATION_NORMALIZATION_CONTROL_TITLE;
  return {
    id,
    status: 'deliberate-difference',
    native,
    evidence: chartCreationEvidence(id, title),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: title,
    },
    serialization: true,
    client: true,
    note,
  };
}

const CHART_CREATION_SUPPORTED_ENTRIES = Object.freeze([
  ...['CHART_NAME', 'ChartType'].flatMap((owner) =>
    CHART_TYPE_VALUES.map((value) => chartCreationSupportedEntry(
      `union:${owner}#${value}`,
      ['CHART_TYPES', 'ChartType'],
      `Native exposes and serializes the same legal ${value} chart type through the frozen CHART_TYPES catalog and strict ChartType union.`,
    ))),
  chartCreationSupportedEntry(
    'union:method:Slide#addChart@path:type#CHART_NAME',
    ['ChartType', 'PptxDocument.addChart', 'SlideModel.addChart'],
    'Native accepts the same nine legal single-chart type values through strict PptxDocument and SlideModel addChart overloads.',
  ),
  chartCreationSupportedEntry(
    'interface:IChartMulti@property:type',
    ['ChartGroupInput.type'],
    'Native represents each combination-chart member with a strict ChartGroupInput discriminated by the same nine legal chart types.',
  ),
  chartCreationSupportedEntry(
    'interface:OptsChartData@property:name',
    ['ChartSeriesInput.name'],
    'Native preserves each public chart series name through ChartSeriesInput and the embedded workbook.',
  ),
  chartCreationSupportedEntry(
    'interface:OptsChartData@property:sizes',
    ['ChartSeriesInput.sizes'],
    'Native accepts bubble sizes explicitly on ChartSeriesInput and preserves them through chart OOXML and workbook reopen.',
  ),
  chartCreationSupportedEntry(
    'interface:OptsChartData@property:values',
    ['ChartSeriesInput.values'],
    'Native requires and preserves numeric chart series values through ChartSeriesInput and the embedded workbook.',
  ),
]);

const CHART_CREATION_DIFFERENCE_ENTRIES = Object.freeze([
  chartCreationDifferenceEntry(
    'method:Slide#addChart',
    ['PptxDocument.addChart', 'SlideModel.addChart', 'ChartModel'],
    'PptxGenJS synchronously returns the chainable Slide; native chart creation is asynchronous and returns the created ChartModel after transactional workbook and relationship commits.',
    true,
  ),
  chartCreationDifferenceEntry(
    'union:method:Slide#addChart@path:type#IChartMulti[]',
    ['ChartGroupInput[]', 'PptxDocument.addChart', 'SlideModel.addChart'],
    'Native covers combination charts with readonly ChartGroupInput arrays instead of the permissive PptxGenJS IChartMulti array shape.',
  ),
  chartCreationDifferenceEntry(
    'interface:IChartMulti@property:data',
    ['ChartGroupInput.series'],
    'Native names combination-chart data series explicitly as ChartGroupInput.series and validates them before mutation.',
  ),
  chartCreationDifferenceEntry(
    'interface:IChartMulti@property:options',
    ['ChartGroupInput.options', 'ChartGroupOptions'],
    'Native maps supported per-group semantics to chart-type-specific ChartGroupOptions instead of accepting the permissive PptxGenJS options bag.',
  ),
  chartCreationDifferenceEntry(
    'interface:OptsChartData@property:labels',
    ['ChartSeriesInput.categories', 'ChartSeriesInput.xValues'],
    'Native separates categorical labels from numeric scatter and bubble x-values instead of overloading one permissive labels field.',
  ),
  chartCreationDifferenceEntry(
    'union:interface:OptsChartData@property:labels#string[]',
    ['ChartSeriesInput.categories'],
    'Native maps the legal flat string label form to explicit ChartSeriesInput.categories and preserves it in chart formulas and workbook cells.',
  ),
]);

const CHART_AXIS_FOUNDATION_CONTROL_TITLE =
  'compares chart axis line gridline visibility label and tick semantics';
const CHART_AXIS_FOUNDATION_OOXML_TITLE =
  'creates and reopens chart axis foundation options in all six formats';
const CHART_AXIS_LABEL_POSITIONS = Object.freeze(['high', 'low', 'nextTo', 'none']);
const CHART_AXIS_DIFFERENT_TICK_MARKS = Object.freeze(['inside', 'outside']);
const CHART_AXIS_LINE_STYLES = Object.freeze(['dash', 'dot', 'solid']);
const CHART_AXIS_GRID_STYLES = Object.freeze(['dash', 'dot', 'none', 'solid']);
const CHART_AXIS_OWNERS = Object.freeze([
  Object.freeze({ owner: 'IChartOpts', prefix: 'cat', nativeAxis: 'categoryAxis' }),
  Object.freeze({ owner: 'IChartOpts', prefix: 'val', nativeAxis: 'valueAxis' }),
  Object.freeze({
    owner: 'IChartPropsAxisCat', prefix: 'cat', nativeAxis: 'categoryAxis',
  }),
  Object.freeze({
    owner: 'IChartPropsAxisVal', prefix: 'val', nativeAxis: 'valueAxis',
  }),
]);

function chartAxisFoundationEvidence() {
  return {
    code: [{
      path: 'packages/model/src/chart.ts',
      pattern: 'export interface ChartAxisOptions extends ChartFontOptions {',
    }, {
      path: 'packages/model/src/chart-options.internal.ts',
      pattern: 'function normalizeAxisOptions(value: unknown, context: string): Readonly<ChartAxisOptions> {',
    }, {
      path: 'packages/model/src/chart-render.internal.ts',
      pattern: 'function renderAxisShapeProperties(',
    }, {
      path: 'packages/model/src/chart-state.internal.ts',
      pattern: 'function readAxisOptions(',
    }],
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: CHART_AXIS_FOUNDATION_CONTROL_TITLE,
    }],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const chartAxisFoundation =',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: CHART_AXIS_FOUNDATION_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const chartAxisFoundation =',
    }],
  };
}

function chartAxisSupportedEntry(id, native, note =
  'Native exposes the same four explicit label-position semantics through strict nested ChartAxisOptions; nextTo remains the canonical serialized default after reopen.') {
  return {
    id,
    status: 'supported',
    native,
    evidence: chartAxisFoundationEvidence(),
    serialization: true,
    client: true,
    note,
  };
}

function chartAxisDifferenceEntry(id, native, note) {
  return {
    id,
    status: 'deliberate-difference',
    native,
    evidence: chartAxisFoundationEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: CHART_AXIS_FOUNDATION_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note,
  };
}

const CHART_AXIS_FOUNDATION_SUPPORTED_ENTRIES = Object.freeze(
  CHART_AXIS_OWNERS.flatMap(({ owner, prefix, nativeAxis }) => {
    const property = `${prefix}AxisLabelPos`;
    const id = linePropertyId(owner, property);
    const native = [`ChartOptions.${nativeAxis}.labelPosition`, 'ChartAxisOptions.labelPosition'];
    return [
      chartAxisSupportedEntry(id, native),
      ...CHART_AXIS_LABEL_POSITIONS.map((value) =>
        chartAxisSupportedEntry(`union:${id}#${value}`, native)),
    ];
  }),
);

const CHART_AXIS_TICK_SUPPORTED_ENTRIES = Object.freeze(
  ['cross', 'none'].map((value) => chartAxisSupportedEntry(
    `union:ChartAxisTickMark#${value}`,
    ['ChartAxisOptions.majorTickMark', 'ChartAxisOptions.minorTickMark'],
    value === 'none'
      ? 'Native exposes and serializes the same legal none tick-mark value; none becomes the canonical absence after reopen.'
      : 'Native exposes, serializes, and reopens the same legal cross tick-mark value.',
  )),
);

const CHART_AXIS_LINE_GRID_DIFFERENCE_ENTRIES = Object.freeze([
  ...CHART_AXIS_OWNERS.flatMap(({ owner, prefix, nativeAxis }) => {
    const axisLine = `ChartOptions.${nativeAxis}.line`;
    const gridLine = `ChartOptions.${nativeAxis}.majorGridLine`;
    const lineProperties = ['Color', 'Show', 'Size', 'Style'].map((suffix) =>
      chartAxisDifferenceEntry(
        linePropertyId(owner, `${prefix}AxisLine${suffix}`),
        [axisLine, 'ShapeLine'],
        'Native maps the flat permissive PptxGenJS axis-line fields to a strict nested ShapeLine, preserves none/solid color and width intent, and uses sysDot as an intentional approximation of the distinct PptxGenJS dot preset.',
      ));
    const lineStyles = CHART_AXIS_LINE_STYLES.map((value) =>
      chartAxisDifferenceEntry(
        `union:${linePropertyId(owner, `${prefix}AxisLineStyle`)}#${value}`,
        [axisLine, 'ShapeLine.dash'],
        'Native maps the flat PptxGenJS solid/dash/dot style domain to strict ShapeLine dash semantics; dash and solid remain exact, while dot uses the distinct sysDot preset as an intentional approximation.',
      ));
    return [
      ...lineProperties,
      chartAxisDifferenceEntry(
        linePropertyId(owner, `${prefix}GridLine`),
        [gridLine, 'ChartAxisOptions.majorGridLine', 'ShapeLine'],
        'Native maps the flat PptxGenJS gridline object to the axis majorGridLine ShapeLine; minorGridLine and gridline cap remain outside this family.',
      ),
      ...lineStyles,
    ];
  }),
  ...['color', 'size', 'style'].map((property) => chartAxisDifferenceEntry(
    linePropertyId('OptsChartGridLine', property),
    [
      property === 'color' ? 'ShapeLine.color' : property === 'size'
        ? 'ShapeLine.width'
        : 'ShapeLine.dash',
      'ChartAxisOptions.majorGridLine',
    ],
    'Native replaces the permissive PptxGenJS gridline field with a strict nested ShapeLine; the separately declared cap field is intentionally not closed by this mapping.',
  )),
  ...CHART_AXIS_GRID_STYLES.map((value) => chartAxisDifferenceEntry(
    `union:${linePropertyId('OptsChartGridLine', 'style')}#${value}`,
    ['ShapeLine', 'ChartAxisOptions.majorGridLine'],
    'Native maps the PptxGenJS solid/dash/dot/none gridline styles to strict ShapeLine state, uses sysDot as an intentional approximation of the distinct dot preset, and represents none explicitly.',
  )),
]);

const CHART_AXIS_BEHAVIOR_DIFFERENCE_ENTRIES = Object.freeze([
  ...CHART_AXIS_OWNERS.flatMap(({ owner, prefix, nativeAxis }) => [
    chartAxisDifferenceEntry(
      linePropertyId(owner, `${prefix}AxisHidden`),
      [`ChartOptions.${nativeAxis}.visible`, 'ChartAxisOptions.visible'],
      'Native expresses the same legal boolean through positive visible semantics instead of the inverted PptxGenJS hidden flag.',
    ),
    chartAxisDifferenceEntry(
      linePropertyId(owner, `${prefix}AxisLabelRotate`),
      [`ChartOptions.${nativeAxis}.labelRotation`, 'ChartAxisOptions.labelRotation'],
      'Native accepts only finite label rotations from -90 through 90 degrees and writes exact 60000-degree units; PptxGenJS wraps or serializes out-of-range and non-finite inputs.',
    ),
    ...['Major', 'Minor'].map((level) => chartAxisDifferenceEntry(
      linePropertyId(owner, `${prefix}Axis${level}TickMark`),
      [`ChartOptions.${nativeAxis}.${level.toLowerCase()}TickMark`, 'ChartAxisOptions'],
      'Native keeps the semantic cross/inside/none/outside API but emits the canonical OOXML cross/in/none/out tokens; PptxGenJS writes inside/outside as invalid lexical values.',
    )),
  ]),
  ...CHART_AXIS_DIFFERENT_TICK_MARKS.map((value) => chartAxisDifferenceEntry(
    `union:ChartAxisTickMark#${value}`,
    ['ChartAxisOptions.majorTickMark', 'ChartAxisOptions.minorTickMark'],
    'Native keeps the semantic cross/inside/none/outside API but emits the canonical OOXML cross/in/none/out tokens; PptxGenJS writes inside/outside as invalid lexical values.',
  )),
]);

const INERT_CHART_OPTION_CONTROL_TITLE =
  'isolates inherited inert IChartOpts text and top-level gridline declarations from chart output';
const INERT_CHART_OPTION_PROPERTIES = Object.freeze([
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
const INERT_CHART_OPTION_INLINE_FIELDS = Object.freeze({
  bullet: Object.freeze([
    'characterCode',
    'code',
    'indent',
    'marginPt',
    'numberStartAt',
    'numberType',
    'startAt',
    'style',
    'type',
  ]),
  tabStops: Object.freeze(['alignment', 'position']),
  underline: Object.freeze(['color', 'style']),
});
const INERT_CHART_BULLET_NUMBER_TYPES = Object.freeze([
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaLcPeriod',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'alphaUcPeriod',
  'arabicParenBoth',
  'arabicParenR',
  'arabicPeriod',
  'arabicPlain',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcParenBoth',
  'romanUcParenR',
  'romanUcPeriod',
]);
const INERT_CHART_UNDERLINE_STYLES = Object.freeze([
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dbl',
  'dotDash',
  'dotDashHeave',
  'dotDotDash',
  'dotDotDashHeavy',
  'dotted',
  'dottedHeavy',
  'heavy',
  'none',
  'sng',
  'wavy',
  'wavyDbl',
  'wavyHeavy',
]);

function inertChartOptionInlineId(property, field) {
  return `inline:${linePropertyId('IChartOpts', property)}@property:${property}.${field}`;
}

function inertChartOptionUnionIds(property, path, values) {
  const suffix = path === undefined ? '' : `@path:${path}`;
  return values.map((value) =>
    `union:${linePropertyId('IChartOpts', property)}${suffix}#${value}`);
}

const INERT_CHART_OPTION_IDS = Object.freeze([
  ...INERT_CHART_OPTION_PROPERTIES.map((property) =>
    linePropertyId('IChartOpts', property)),
  ...Object.entries(INERT_CHART_OPTION_INLINE_FIELDS).flatMap(([property, fields]) =>
    fields.map((field) => inertChartOptionInlineId(property, field))),
  ...inertChartOptionUnionIds('bullet', undefined, ['boolean']),
  ...inertChartOptionUnionIds(
    'bullet',
    'bullet.numberType',
    INERT_CHART_BULLET_NUMBER_TYPES,
  ),
  ...inertChartOptionUnionIds('bullet', 'bullet.type', ['bullet', 'number']),
  ...inertChartOptionUnionIds('style', undefined, ['dash', 'dot', 'none', 'solid']),
  ...inertChartOptionUnionIds(
    'tabStops',
    'tabStops.alignment',
    ['ctr', 'dec', 'l', 'r'],
  ),
  ...inertChartOptionUnionIds(
    'textDirection',
    undefined,
    ['horz', 'vert', 'vert270', 'wordArtVert'],
  ),
  ...inertChartOptionUnionIds(
    'underline',
    'underline.style',
    INERT_CHART_UNDERLINE_STYLES,
  ),
].sort());

function inertChartOptionDefectEntry(id) {
  return {
    id,
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: [{
        path: 'packages/model/src/chart-options.internal.ts',
        pattern: 'const OPTION_KEYS = [',
      }],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: INERT_CHART_OPTION_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: INERT_CHART_OPTION_CONTROL_TITLE,
    },
    note: 'PptxGenJS 4.0.1 inherits this IChartOpts field through IChartPropsTitle/TextBaseProps or top-level OptsChartGridLine, but its chart writer ignores every legal value; native uses explicit nested chart options and does not copy inert declaration noise.',
  };
}

const INERT_CHART_OPTION_DEFECT_ENTRIES = Object.freeze(
  INERT_CHART_OPTION_IDS.map((id) => inertChartOptionDefectEntry(id)),
);

const BULLET_FAMILY_CONTROL_TITLE =
  'locks bullet and numbering behavior across every declared owner';
const BULLET_FAMILY_OOXML_TITLE =
  'creates and reopens bullet and numbering owners in all six formats';
const BULLET_FAMILY_FIELDS = Object.freeze([
  'characterCode',
  'code',
  'indent',
  'marginPt',
  'numberStartAt',
  'numberType',
  'startAt',
  'style',
  'type',
]);
const BULLET_FAMILY_ACTIVE_OWNERS = Object.freeze([
  'PlaceholderProps',
  'TableCellProps',
  'TextPropsOptions',
]);
const BULLET_FAMILY_INERT_OWNERS = Object.freeze([
  'SlideNumberProps',
  'TableProps',
  'TableToSlidesProps',
]);
function bulletFamilyInlineId(owner, field) {
  return `inline:${linePropertyId(owner, 'bullet')}@property:bullet.${field}`;
}

function bulletFamilyUnionId(owner, path, value) {
  const suffix = path === undefined ? '' : `@path:${path}`;
  return `union:${linePropertyId(owner, 'bullet')}${suffix}#${value}`;
}

function bulletFamilyOwnerIds(owner) {
  return [
    linePropertyId(owner, 'bullet'),
    ...BULLET_FAMILY_FIELDS.map((field) => bulletFamilyInlineId(owner, field)),
    bulletFamilyUnionId(owner, undefined, 'boolean'),
    ...INERT_CHART_BULLET_NUMBER_TYPES.map((value) =>
      bulletFamilyUnionId(owner, 'bullet.numberType', value)),
    ...['bullet', 'number'].map((value) =>
      bulletFamilyUnionId(owner, 'bullet.type', value)),
  ].sort();
}

function bulletFamilyNative(owner, id) {
  const ownerMapping = owner === 'PlaceholderProps'
    ? ['AddTextOptions.bullet', 'SlideModel.addPlaceholder']
    : owner === 'TextPropsOptions'
      ? ['AddTextOptions.bullet', 'SlideModel.addText']
      : ['RichTextParagraph.bullet', 'AddTableCell.text', 'SlideModel.addTable'];
  const fieldMapping = id.includes('characterCode') || id.includes('bullet.code')
    ? ['CharacterBullet.character']
    : id.includes('numberStartAt') || id.includes('bullet.startAt')
      ? ['NumberedBullet.startAt']
      : id.includes('numberType') || id.includes('bullet.style')
        ? ['NumberedBullet.style', 'NumberingStyle']
        : id.includes('bullet.indent') || id.includes('marginPt')
          ? ['CharacterBullet.indent', 'NumberedBullet.indent']
          : id.includes('bullet.type')
            ? ['CharacterBullet.kind', 'NumberedBullet.kind']
            : ['ParagraphBullet'];
  return [...fieldMapping, ...ownerMapping];
}

function bulletFamilyEvidence() {
  return {
    code: [{
      path: 'packages/model/src/rich-text.internal.ts',
      pattern: 'export function normalizeParagraphBullet(',
    }],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: BULLET_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: BULLET_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const bulletNumberingState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: BULLET_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const bulletNumberingState = {',
    }],
  };
}

function bulletFamilyDefectEntry(owner, id) {
  const inertOwner = BULLET_FAMILY_INERT_OWNERS.includes(owner);
  const marginAlias = id === bulletFamilyInlineId(owner, 'marginPt');
  const typedBullet = id === bulletFamilyUnionId(owner, 'bullet.type', 'bullet');
  return {
    id,
    status: 'defect-excluded',
    native: inertOwner ? [] : bulletFamilyNative(owner, id),
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: BULLET_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: BULLET_FAMILY_CONTROL_TITLE,
    },
    note: inertOwner
      ? `PptxGenJS 4.0.1 inherits bullet declarations into ${owner}, but its writer ignores every boolean, character, numbering, indentation, and alias value for that owner.`
      : marginAlias
        ? `PptxGenJS 4.0.1 declares ${owner}.bullet.marginPt but its writer never reads it and always uses bullet.indent or the default margin.`
        : typedBullet
          ? `PptxGenJS 4.0.1 declares ${owner}.bullet.type='bullet', but its object writer only handles the number branch and emits no bullet for this legal token.`
          : `PptxGenJS 4.0.1 declares ${owner}.bullet.numberType and all sixteen legal tokens but its writer ignores the field and reads deprecated bullet.style instead.`,
  };
}

function bulletFamilyDeprecatedEntry(owner, field, canonicalField) {
  const id = bulletFamilyInlineId(owner, field);
  return {
    id,
    status: 'deprecated-alias',
    native: bulletFamilyNative(owner, id),
    evidence: bulletFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: BULLET_FAMILY_CONTROL_TITLE,
    },
    canonical: bulletFamilyInlineId(owner, canonicalField),
    serialization: true,
    client: true,
    note: `PptxGenJS 4.0.1 keeps bullet.${field} as a working deprecated alias of bullet.${canonicalField}; native rejects the alias and exposes only its strict semantic bullet field.`,
  };
}

function bulletFamilyDifferenceEntry(owner, id) {
  const tableOwner = owner === 'TableCellProps';
  return {
    id,
    status: 'deliberate-difference',
    native: bulletFamilyNative(owner, id),
    evidence: bulletFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: BULLET_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note: tableOwner
      ? `Native represents ${owner} bullets in explicit cell rich-text paragraphs with a strict discriminated union; PptxGenJS uses permissive inherited cell or table defaults, hexadecimal character codes, and truthy fallbacks.`
      : `Native represents ${owner} bullets with a strict discriminated union, actual Unicode characters, bounded indentation and start values, and deterministic OOXML; PptxGenJS uses permissive inline fields and truthy fallbacks.`,
  };
}

function bulletFamilySupportedPropertyEntry(owner) {
  const id = linePropertyId(owner, 'bullet');
  return {
    id,
    status: 'supported',
    native: bulletFamilyNative(owner, id),
    evidence: bulletFamilyEvidence(),
    serialization: true,
    client: true,
    note: `Native covers the complete effective ${owner}.bullet output through its strict ParagraphBullet model and preserves standard, character, and numbered bullets through serialization and reopen.`,
  };
}

function bulletFamilyActiveEntry(owner, id) {
  if (
    id === bulletFamilyInlineId(owner, 'numberType')
    || id === bulletFamilyInlineId(owner, 'marginPt')
    || id === bulletFamilyUnionId(owner, 'bullet.type', 'bullet')
    || id.startsWith(`union:${linePropertyId(owner, 'bullet')}@path:bullet.numberType#`)
  ) return bulletFamilyDefectEntry(owner, id);
  if (id === bulletFamilyInlineId(owner, 'code')) {
    return bulletFamilyDeprecatedEntry(owner, 'code', 'characterCode');
  }
  if (id === bulletFamilyInlineId(owner, 'startAt')) {
    return bulletFamilyDeprecatedEntry(owner, 'startAt', 'numberStartAt');
  }
  if (id === linePropertyId(owner, 'bullet')) {
    return bulletFamilySupportedPropertyEntry(owner);
  }
  return bulletFamilyDifferenceEntry(owner, id);
}

const BULLET_FAMILY_ENTRIES = Object.freeze([
  ...BULLET_FAMILY_ACTIVE_OWNERS.flatMap((owner) =>
    bulletFamilyOwnerIds(owner).map((id) => bulletFamilyActiveEntry(owner, id))),
  ...BULLET_FAMILY_INERT_OWNERS.flatMap((owner) =>
    bulletFamilyOwnerIds(owner).map((id) => bulletFamilyDefectEntry(owner, id))),
  {
    id: linePropertyId('TextPropsOptions', 'indentLevel'),
    status: 'deliberate-difference',
    native: [
      'AddTextOptions.level',
      'RichTextParagraph.level',
      'SlideModel.addText',
      'SlideModel.addRichText',
    ],
    evidence: {
      code: [{
        path: 'packages/model/src/rich-text.internal.ts',
        pattern: 'export function normalizeParagraphLevel(',
      }],
      tests: [
        {
          path: 'packages/pptxgenjs-adapter/src/index.test.ts',
          title: BULLET_FAMILY_CONTROL_TITLE,
        },
        {
          path: 'packages/sdk/src/index.test.ts',
          title: BULLET_FAMILY_OOXML_TITLE,
        },
      ],
      package: [{
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const bulletNumberingState = {',
      }],
      ooxml: [{
        path: 'packages/sdk/src/index.test.ts',
        pattern: BULLET_FAMILY_OOXML_TITLE,
      }],
      clients: [{
        path: 'scripts/playwright-browser-smoke.js',
        pattern: 'const bulletNumberingState = {',
      }],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: BULLET_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note: 'Native preserves list levels as strict integers from zero through eight, including explicit level zero; PptxGenJS uses indentLevel, ignores non-positive values, and permissively emits out-of-range positive values.',
  },
]);

const TAB_STOPS_FAMILY_CONTROL_TITLE =
  'locks tab stop behavior across every declared owner';
const TAB_STOPS_FAMILY_OOXML_TITLE =
  'creates edits clears and reopens tab stops in all six formats';
const TAB_STOPS_FAMILY_ACTIVE_OWNERS = Object.freeze([
  'PlaceholderProps',
  'TableCellProps',
  'TextPropsOptions',
]);
const TAB_STOPS_FAMILY_INERT_OWNERS = Object.freeze([
  'SlideNumberProps',
  'TableProps',
  'TableToSlidesProps',
]);
const TAB_STOPS_FAMILY_FIELDS = Object.freeze(['alignment', 'position']);
const TAB_STOPS_FAMILY_ALIGNMENT_TOKENS = Object.freeze(['ctr', 'dec', 'l', 'r']);

function tabStopsFamilyInlineId(owner, field) {
  return `inline:${linePropertyId(owner, 'tabStops')}@property:tabStops.${field}`;
}

function tabStopsFamilyUnionId(owner, value) {
  return `union:${linePropertyId(owner, 'tabStops')}@path:tabStops.alignment#${value}`;
}

function tabStopsFamilyOwnerIds(owner) {
  return [
    linePropertyId(owner, 'tabStops'),
    ...TAB_STOPS_FAMILY_FIELDS.map((field) => tabStopsFamilyInlineId(owner, field)),
    ...TAB_STOPS_FAMILY_ALIGNMENT_TOKENS.map((value) =>
      tabStopsFamilyUnionId(owner, value)),
  ].sort();
}

function tabStopsFamilyNative(owner, id) {
  const ownerMapping = owner === 'PlaceholderProps'
    ? ['AddTextOptions.tabStops', 'RichTextParagraph.tabStops', 'SlideModel.addPlaceholder']
    : owner === 'TextPropsOptions'
      ? ['AddTextOptions.tabStops', 'RichTextParagraph.tabStops', 'SlideModel.addText']
      : [
          'AddTableCell.text',
          'RichTextParagraph.tabStops',
          'SlideModel.addTable',
          'TableModel.setCellRichText',
        ];
  const fieldMapping = id.includes('tabStops.position')
    ? ['ParagraphTabStop.position']
    : id.includes('tabStops.alignment')
      ? ['ParagraphTabStopAlignment']
      : ['ParagraphTabStop', 'ParagraphTabStopAlignment'];
  return [...fieldMapping, ...ownerMapping];
}

function tabStopsFamilyEvidence() {
  return {
    code: [{
      path: 'packages/model/src/rich-text.internal.ts',
      pattern: 'export function normalizeParagraphTabStops(',
    }],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TAB_STOPS_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: TAB_STOPS_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const tabStopsState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: TAB_STOPS_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const tabStopsState = {',
    }],
  };
}

function tabStopsFamilySupportedEntry(owner) {
  const id = linePropertyId(owner, 'tabStops');
  return {
    id,
    status: 'supported',
    native: tabStopsFamilyNative(owner, id),
    evidence: tabStopsFamilyEvidence(),
    serialization: true,
    client: true,
    note: `Native covers the effective ${owner}.tabStops output with strict paragraph tab stops and preserves every stop through serialization and reopen.`,
  };
}

function tabStopsFamilyDifferenceEntry(owner, id) {
  return {
    id,
    status: 'deliberate-difference',
    native: tabStopsFamilyNative(owner, id),
    evidence: tabStopsFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TAB_STOPS_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note: `Native maps ${owner}.tabStops to semantic left/center/right/decimal alignments with finite signed-32-bit positions and preserves zero and negative values; PptxGenJS exposes l/ctr/r/dec and its truthy fallback changes position zero to one inch.`,
  };
}

function tabStopsFamilyDefectEntry(owner, id) {
  return {
    id,
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TAB_STOPS_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TAB_STOPS_FAMILY_CONTROL_TITLE,
    },
    note: `PptxGenJS 4.0.1 inherits tabStops declarations into ${owner}, but its writer ignores the property, position, and all four alignment tokens for that owner.`,
  };
}

const TAB_STOPS_FAMILY_ENTRIES = Object.freeze([
  ...TAB_STOPS_FAMILY_ACTIVE_OWNERS.flatMap((owner) =>
    tabStopsFamilyOwnerIds(owner).map((id) => id === linePropertyId(owner, 'tabStops')
      ? tabStopsFamilySupportedEntry(owner)
      : tabStopsFamilyDifferenceEntry(owner, id))),
  ...TAB_STOPS_FAMILY_INERT_OWNERS.flatMap((owner) =>
    tabStopsFamilyOwnerIds(owner).map((id) => tabStopsFamilyDefectEntry(owner, id))),
]);

const UNDERLINE_FAMILY_CONTROL_TITLE =
  'locks underline behavior across every declared owner';
const UNDERLINE_FAMILY_OOXML_TITLE =
  'creates edits clears and reopens underline owners in all six formats';
const UNDERLINE_FAMILY_ACTIVE_OWNERS = Object.freeze([
  'PlaceholderProps',
  'TableCellProps',
  'TableProps',
  'TextPropsOptions',
]);
const UNDERLINE_FAMILY_INERT_OWNERS = Object.freeze([
  'SlideNumberProps',
  'TableToSlidesProps',
]);
const UNDERLINE_FAMILY_FIELDS = Object.freeze(['color', 'style']);
const UNDERLINE_FAMILY_STYLE_TOKENS = Object.freeze([
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dbl',
  'dotDash',
  'dotDashHeave',
  'dotDotDash',
  'dotDotDashHeavy',
  'dotted',
  'dottedHeavy',
  'heavy',
  'none',
  'sng',
  'wavy',
  'wavyDbl',
  'wavyHeavy',
]);

function underlineFamilyInlineId(owner, field) {
  return `inline:${linePropertyId(owner, 'underline')}@property:underline.${field}`;
}

function underlineFamilyUnionId(owner, value) {
  return `union:${linePropertyId(owner, 'underline')}@path:underline.style#${value}`;
}

function underlineFamilyOwnerIds(owner) {
  return [
    linePropertyId(owner, 'underline'),
    ...UNDERLINE_FAMILY_FIELDS.map((field) => underlineFamilyInlineId(owner, field)),
    ...UNDERLINE_FAMILY_STYLE_TOKENS.map((value) =>
      underlineFamilyUnionId(owner, value)),
  ].sort();
}

function underlineFamilyNative(owner, id) {
  const ownerMapping = owner === 'PlaceholderProps'
    ? ['ShapeModel.richText', 'SlideModel.addPlaceholder']
    : owner === 'TextPropsOptions'
      ? ['RichTextRunStyle.underline', 'SlideModel.addRichText', 'SlideModel.addText']
      : owner === 'TableCellProps'
        ? [
            'AddTableCell.text',
            'RichTextRunStyle.underline',
            'SlideModel.addTable',
            'TableModel.setCellRichText',
          ]
        : [
            'AddTableCell.text',
            'RichTextRunStyle.underline',
            'SlideModel.addTable',
            'TableModel.setCellRichText',
          ];
  const fieldMapping = id.includes('underline.color')
    ? ['RichTextColor', 'RichTextUnderline.color']
    : id.includes('underline.style')
      ? ['RichTextUnderline.style', 'RichTextUnderlineStyle']
      : ['RichTextRunStyle.underline', 'RichTextUnderline', 'RichTextUnderlineStyle'];
  const native = [...fieldMapping, ...ownerMapping];
  return [...new Set(native)];
}

function underlineFamilyEvidence() {
  return {
    code: [{
      path: 'packages/model/src/rich-text.internal.ts',
      pattern: 'function normalizeUnderline(',
    }],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: UNDERLINE_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: UNDERLINE_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const underlineFamilyState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: UNDERLINE_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const underlineFamilyState = {',
    }],
  };
}

function underlineFamilySupportedEntry(owner, id) {
  const tableOwner = owner === 'TableProps';
  return {
    id,
    status: 'supported',
    native: underlineFamilyNative(owner, id),
    evidence: underlineFamilyEvidence(),
    serialization: true,
    client: true,
    note: tableOwner
      ? 'Native covers effective TableProps underline output through explicit cell rich-text equivalents, strict same-named legal styles, deterministic OOXML, and reopen editing.'
      : `Native covers the effective ${owner} underline output through strict run styles, deterministic OOXML, and reopen editing.`,
  };
}

function underlineFamilyDifferenceEntry(owner, id) {
  const color = id === underlineFamilyInlineId(owner, 'color');
  return {
    id,
    status: 'deliberate-difference',
    native: underlineFamilyNative(owner, id),
    evidence: underlineFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: UNDERLINE_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note: color
      ? `Native represents ${owner}.underline.color as a strict sRGB or scheme RichTextColor and normalizes color-only input to a single underline; PptxGenJS accepts a permissive color string and emits only uFill for color-only input.`
      : `Native expresses ${owner}.underline.style='none' as underline false and may omit or explicitly serialize the local disable state while preserving the same final text semantics.`,
  };
}

function underlineFamilyDefectEntry(owner, id) {
  const inertOwner = UNDERLINE_FAMILY_INERT_OWNERS.includes(owner);
  return {
    id,
    status: 'defect-excluded',
    native: inertOwner ? [] : underlineFamilyNative(owner, id),
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: UNDERLINE_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: UNDERLINE_FAMILY_CONTROL_TITLE,
    },
    note: inertOwner
      ? `PptxGenJS 4.0.1 inherits underline declarations into ${owner}, but its writer ignores the property, both inline fields, and all seventeen declared style tokens for that owner.`
      : `PptxGenJS 4.0.1 declares ${owner}.underline.style='dotDashHeave' and writes that invalid token literally; native exposes the correct DrawingML token dotDashHeavy and rejects the typo before mutation.`,
  };
}

function underlineFamilyActiveEntry(owner, id) {
  if (id === underlineFamilyUnionId(owner, 'dotDashHeave')) {
    return underlineFamilyDefectEntry(owner, id);
  }
  if (
    id === underlineFamilyInlineId(owner, 'color')
    || id === underlineFamilyUnionId(owner, 'none')
  ) return underlineFamilyDifferenceEntry(owner, id);
  return underlineFamilySupportedEntry(owner, id);
}

const UNDERLINE_FAMILY_ENTRIES = Object.freeze([
  ...UNDERLINE_FAMILY_ACTIVE_OWNERS.flatMap((owner) =>
    underlineFamilyOwnerIds(owner).map((id) => underlineFamilyActiveEntry(owner, id))),
  ...UNDERLINE_FAMILY_INERT_OWNERS.flatMap((owner) =>
    underlineFamilyOwnerIds(owner).map((id) => underlineFamilyDefectEntry(owner, id))),
]);

const TEXT_DIRECTION_FAMILY_CONTROL_TITLE =
  'locks text direction behavior across every declared owner';
const TEXT_DIRECTION_FAMILY_OOXML_TITLE =
  'creates and reopens text direction owners in all six formats';
const TEXT_DIRECTION_TABLE_TOKENS = Object.freeze([
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
]);
const TEXT_DIRECTION_TEXT_TOKENS = Object.freeze([
  'eaVert',
  'horz',
  'mongolianVert',
  'vert',
  'vert270',
  'wordArtVert',
  'wordArtVertRtl',
]);
const TEXT_DIRECTION_ACTIVE_TABLE_OWNERS = Object.freeze([
  'TableCellProps',
  'TableProps',
]);
const TEXT_DIRECTION_INERT_OWNERS = Object.freeze([
  'PlaceholderProps',
  'SlideNumberProps',
  'TableToSlidesProps',
  'TextPropsOptions',
]);

function textDirectionPropertyIds(owner, property, tokens) {
  const propertyId = linePropertyId(owner, property);
  return [
    propertyId,
    ...tokens.map((value) => `union:${propertyId}#${value}`),
  ].sort();
}

function textDirectionFamilyEvidence(kind) {
  return {
    code: [{
      path: kind === 'table'
        ? 'packages/model/src/table-cell-text-direction.internal.ts'
        : 'packages/model/src/text-box-text-direction.internal.ts',
      pattern: kind === 'table'
        ? 'export function normalizeTableCellTextDirection('
        : 'export function normalizeTextBoxTextDirection(',
    }],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TEXT_DIRECTION_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: TEXT_DIRECTION_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const textDirectionFamilyState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: TEXT_DIRECTION_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const textDirectionFamilyState = {',
    }],
  };
}

function textDirectionSupportedEntry(owner, property, id) {
  const table = property === 'textDirection';
  const native = table
    ? owner === 'TableCellProps'
      ? [
          'AddTableCellOptions.textDirection',
          'TableCell.textDirection',
          'TableCellTextDirection',
          'TableModel.setCellTextDirection',
        ]
      : [
          'AddTableOptions.textDirection',
          'TableCellTextDirection',
          'TableModel.textDirection',
        ]
    : [
        'AddTextOptions.vert',
        'ShapeModel.textDirection',
        'TextBoxTextDirection',
      ];
  return {
    id,
    status: 'supported',
    native,
    evidence: textDirectionFamilyEvidence(table ? 'table' : 'text'),
    serialization: true,
    client: true,
    note: table
      ? `Native covers ${owner}.textDirection with the same four legal tokens, strict pre-mutation validation, editable table-cell OOXML, and canonical omission of the horizontal default.`
      : 'Native covers TextPropsOptions.vert with the same seven legal text-body tokens, strict pre-mutation validation, exact bodyPr OOXML, and reopen editing.',
  };
}

function textDirectionDefectEntry(owner, id) {
  return {
    id,
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TEXT_DIRECTION_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TEXT_DIRECTION_FAMILY_CONTROL_TITLE,
    },
    note: `PptxGenJS 4.0.1 inherits textDirection into ${owner}, but its writer ignores the property and all four declared tokens for that owner.`,
  };
}

const TEXT_DIRECTION_FAMILY_ENTRIES = Object.freeze([
  ...TEXT_DIRECTION_ACTIVE_TABLE_OWNERS.flatMap((owner) =>
    textDirectionPropertyIds(owner, 'textDirection', TEXT_DIRECTION_TABLE_TOKENS)
      .map((id) => textDirectionSupportedEntry(owner, 'textDirection', id))),
  ...TEXT_DIRECTION_INERT_OWNERS.flatMap((owner) =>
    textDirectionPropertyIds(owner, 'textDirection', TEXT_DIRECTION_TABLE_TOKENS)
      .map((id) => textDirectionDefectEntry(owner, id))),
  ...textDirectionPropertyIds(
    'TextPropsOptions',
    'vert',
    TEXT_DIRECTION_TEXT_TOKENS,
  ).map((id) => textDirectionSupportedEntry('TextPropsOptions', 'vert', id)),
]);

const TEXT_RUN_SCALAR_FAMILY_CONTROL_TITLE =
  'locks scalar text formatting behavior across every declared owner';
const TEXT_RUN_SCALAR_FAMILY_OOXML_TITLE =
  'creates and reopens scalar text formatting owners in all six formats';
const TEXT_RUN_SCALAR_FAMILY_PROPERTIES = Object.freeze([
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
const TEXT_RUN_SCALAR_FAMILY_STATUS = Object.freeze({
  PlaceholderProps: Object.freeze({
    bold: 'supported',
    breakLine: 'defect-excluded',
    color: 'deliberate-difference',
    fontFace: 'deliberate-difference',
    fontSize: 'supported',
    highlight: 'deliberate-difference',
    italic: 'supported',
    lang: 'supported',
    softBreakBefore: 'defect-excluded',
  }),
  SlideNumberProps: Object.freeze({
    bold: 'supported',
    breakLine: 'defect-excluded',
    color: 'deliberate-difference',
    fontFace: 'deliberate-difference',
    fontSize: 'supported',
    highlight: 'defect-excluded',
    italic: 'defect-excluded',
    lang: 'defect-excluded',
    softBreakBefore: 'defect-excluded',
  }),
  TableCellProps: Object.freeze({
    bold: 'deliberate-difference',
    breakLine: 'supported',
    color: 'deliberate-difference',
    fontFace: 'deliberate-difference',
    fontSize: 'supported',
    highlight: 'deliberate-difference',
    italic: 'supported',
    lang: 'supported',
    softBreakBefore: 'supported',
  }),
  TableProps: Object.freeze({
    bold: 'supported',
    breakLine: 'defect-excluded',
    color: 'deliberate-difference',
    fontFace: 'deliberate-difference',
    fontSize: 'supported',
    highlight: 'defect-excluded',
    italic: 'defect-excluded',
    lang: 'defect-excluded',
    softBreakBefore: 'defect-excluded',
  }),
  TableToSlidesProps: Object.freeze(Object.fromEntries(
    TEXT_RUN_SCALAR_FAMILY_PROPERTIES.map((property) => [property, 'defect-excluded']),
  )),
  TextPropsOptions: Object.freeze({
    bold: 'supported',
    breakLine: 'supported',
    color: 'deliberate-difference',
    fontFace: 'deliberate-difference',
    fontSize: 'supported',
    highlight: 'deliberate-difference',
    italic: 'supported',
    lang: 'supported',
    softBreakBefore: 'supported',
  }),
});

function textRunScalarFamilyEvidence() {
  return {
    code: [
      {
        path: 'packages/model/src/rich-text.internal.ts',
        pattern: 'function normalizeStyle(',
      },
      {
        path: 'packages/model/src/table-create.internal.ts',
        pattern: 'export function normalizeTableDefinition(',
      },
      {
        path: 'packages/codecs/src/slide-number.internal.ts',
        pattern: 'function normalizeStyle(',
      },
    ],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TEXT_RUN_SCALAR_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: TEXT_RUN_SCALAR_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const textRunScalarFamilyState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: TEXT_RUN_SCALAR_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const textRunScalarFamilyState = {',
    }],
  };
}

function textRunScalarFamilyNative(owner, property) {
  if (owner === 'TableToSlidesProps') return [];
  const nativeProperty = property === 'fontFace' ? 'fontFamily' : property;
  if (owner === 'SlideNumberProps') {
    return [`SlideNumberTextStyle.${nativeProperty}`, 'SlideNumberOptions.style'];
  }
  if (owner === 'TableProps') {
    return [`AddTableOptions.${nativeProperty}`, 'SlideModel.addTable', 'TableModel'];
  }
  const runProperty = property === 'breakLine' || property === 'softBreakBefore'
    ? `RichTextRun.${property}`
    : `RichTextRunStyle.${nativeProperty}`;
  if (owner === 'TableCellProps') {
    return [runProperty, 'AddTableCell.text', 'TableModel.setCellRichText'];
  }
  if (owner === 'PlaceholderProps') {
    return [runProperty, 'ShapeModel.richText', 'SlideModel.addPlaceholder'];
  }
  return [runProperty, 'SlideModel.addRichText', 'ShapeModel.richText'];
}

function textRunScalarFamilySupportedEntry(owner, property) {
  const lineBreak = property === 'breakLine' || property === 'softBreakBefore';
  return {
    id: linePropertyId(owner, property),
    status: 'supported',
    native: textRunScalarFamilyNative(owner, property),
    evidence: textRunScalarFamilyEvidence(),
    serialization: true,
    client: true,
    note: lineBreak
      ? `Native covers effective ${owner}.${property} with canonical paragraph or soft-break OOXML and preserves it through edit, duplicate, all six formats, and reopen.`
      : `Native covers effective ${owner}.${property} with strict detached state, exact OOXML, edit, duplicate, all six formats, and reopen.`,
  };
}

function textRunScalarFamilyDifferenceEntry(owner, property) {
  const note = property === 'fontFace'
    ? `Native names ${owner}.fontFace as fontFamily, requires a non-empty XML-safe string, and writes the same effective typeface.`
    : property === 'color'
      ? `Native represents ${owner}.color as a strict sRGB or scheme RichTextColor instead of a permissive PptxGenJS color string.`
      : property === 'highlight'
        ? `Native represents ${owner}.highlight as a strict sRGB or scheme RichTextColor and emits it independently; PptxGenJS accepts permissive strings and conditionally drops some legal highlight-only input.`
        : `Native preserves legal ${owner}.bold=false when overriding a true table default; PptxGenJS uses a truthy fallback that changes the effective value to true.`;
  return {
    id: linePropertyId(owner, property),
    status: 'deliberate-difference',
    native: textRunScalarFamilyNative(owner, property),
    evidence: textRunScalarFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TEXT_RUN_SCALAR_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note,
  };
}

function textRunScalarFamilyDefectEntry(owner, property) {
  const tableToSlides = owner === 'TableToSlidesProps';
  return {
    id: linePropertyId(owner, property),
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: tableToSlides ? [{
        path: 'packages/sdk/src/table-to-slides-css.ts',
        pattern: 'export function mapComputedCellOptions(',
      }] : [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TEXT_RUN_SCALAR_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TEXT_RUN_SCALAR_FAMILY_CONTROL_TITLE,
    },
    note: tableToSlides
      ? `PptxGenJS 4.0.1 inherits ${owner}.${property}, but tableToSlides ignores the flat field; computed CSS remains the positive styling input and native does not copy this inert alias.`
      : `PptxGenJS 4.0.1 inherits ${owner}.${property}, but the owner writer ignores the field for every legal value.`,
  };
}

function textRunScalarFamilyEntry(owner, property, status) {
  if (status === 'supported') return textRunScalarFamilySupportedEntry(owner, property);
  if (status === 'deliberate-difference') {
    return textRunScalarFamilyDifferenceEntry(owner, property);
  }
  return textRunScalarFamilyDefectEntry(owner, property);
}

const TEXT_RUN_SCALAR_FAMILY_ENTRIES = Object.freeze(
  Object.entries(TEXT_RUN_SCALAR_FAMILY_STATUS).flatMap(([owner, statuses]) =>
    TEXT_RUN_SCALAR_FAMILY_PROPERTIES.map((property) =>
      textRunScalarFamilyEntry(owner, property, statuses[property]))),
);

const TABLE_TO_SLIDES_FILL_CONTROL_TITLE =
  'isolates the ignored tableToSlides fill declaration from computed CSS backgrounds';
const TABLE_TO_SLIDES_FILL_DEFECT_ENTRY = Object.freeze({
  id: linePropertyId('TableToSlidesProps', 'fill'),
  status: 'defect-excluded',
  native: [],
  evidence: {
    code: [{
      path: 'packages/sdk/src/table-to-slides-css.ts',
      pattern: 'export function mapComputedCellOptions(',
    }],
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: TABLE_TO_SLIDES_FILL_CONTROL_TITLE,
    }],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const tableToSlidesStyles =',
    }],
    ooxml: [],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const tableToSlidesState = {',
    }],
  },
  control: {
    path: 'packages/pptxgenjs-adapter/src/index.test.ts',
    pattern: TABLE_TO_SLIDES_FILL_CONTROL_TITLE,
  },
  note: 'PptxGenJS 4.0.1 inherits TableProps.fill into TableToSlidesProps but drops the option before creating any auto-paged table; computed CSS backgrounds alone determine cell fills, which native exposes directly without copying the inert declaration.',
});

const TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE =
  'locks every declared tableToSlides option and nested addition against PptxGenJS 4.0.1';
const TABLE_TO_SLIDES_FAMILY_OOXML_TITLE =
  'creates editable styled tables in all six OOXML presentation formats';
const TABLE_TO_SLIDES_TOP_STATUS = Object.freeze({
  autoPageCharWeight: 'supported',
  autoPageLineWeight: 'supported',
  autoPageRepeatHeader: 'supported',
  masterSlideName: 'supported',
  addImage: 'deliberate-difference',
  addShape: 'deliberate-difference',
  addTable: 'deliberate-difference',
  addText: 'deliberate-difference',
  autoPageSlideStartY: 'deliberate-difference',
  h: 'deliberate-difference',
  slideMargin: 'deliberate-difference',
  verbose: 'deliberate-difference',
  w: 'deliberate-difference',
  x: 'deliberate-difference',
  y: 'deliberate-difference',
  addHeaderToEach: 'deprecated-alias',
  newSlideStartY: 'deprecated-alias',
  align: 'defect-excluded',
  autoPage: 'defect-excluded',
  autoPageHeaderRows: 'defect-excluded',
  border: 'defect-excluded',
  colW: 'defect-excluded',
  margin: 'defect-excluded',
  objectName: 'defect-excluded',
  rowH: 'defect-excluded',
  transparency: 'defect-excluded',
  valign: 'defect-excluded',
});
const TABLE_TO_SLIDES_NESTED_FIELDS = Object.freeze({
  addImage: Object.freeze({ image: 'source', options: 'options' }),
  addShape: Object.freeze({ options: 'options', shapeName: 'type' }),
  addTable: Object.freeze({ options: 'options', rows: 'rows' }),
  addText: Object.freeze({ options: 'options', text: 'text' }),
});
const TABLE_TO_SLIDES_UNION_TOKENS = Object.freeze({
  border: Object.freeze(['BorderProps', '[BorderProps,BorderProps,BorderProps,BorderProps]']),
  colW: Object.freeze(['number', 'number[]']),
  rowH: Object.freeze(['number', 'number[]']),
});

function tableToSlidesFamilyEvidence() {
  return {
    code: [
      {
        path: 'packages/sdk/src/table-to-slides.ts',
        pattern: 'export interface TableToSlidesOptions {',
      },
      {
        path: 'packages/sdk/src/index.ts',
        pattern: 'async tableToSlides(',
      },
    ],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/table-to-slides.test.ts',
        title: TABLE_TO_SLIDES_FAMILY_OOXML_TITLE,
      },
    ],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const tableToSlidesState = {',
    }],
    ooxml: [{
      path: 'packages/sdk/src/table-to-slides.test.ts',
      pattern: TABLE_TO_SLIDES_FAMILY_OOXML_TITLE,
    }],
    clients: [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const tableToSlidesState = {',
    }],
  };
}

function tableToSlidesNative(property) {
  if (property === 'w') return ['TableToSlidesOptions.width'];
  if (property === 'h') return ['TableToSlidesOptions.height'];
  if (property === 'masterSlideName') {
    return ['TableToSlidesOptions.masterSlideName', 'PptxDocument.defineSlideMaster'];
  }
  return [`TableToSlidesOptions.${property}`, 'PptxDocument.tableToSlides'];
}

function tableToSlidesSupportedEntry(property) {
  return {
    id: linePropertyId('TableToSlidesProps', property),
    status: 'supported',
    native: tableToSlidesNative(property),
    evidence: tableToSlidesFamilyEvidence(),
    serialization: true,
    client: true,
    note: `Native exposes effective TableToSlidesProps.${property} under the same name with strict detached input, editable output, packed-package use, browser execution, and reopen evidence.`,
  };
}

function tableToSlidesDifferenceEntry(
  id,
  native,
  note,
  { serialization = true, client = true } = {},
) {
  return {
    id,
    status: 'deliberate-difference',
    native,
    evidence: tableToSlidesFamilyEvidence(),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE,
    },
    serialization,
    client,
    note,
  };
}

function tableToSlidesTopDifferenceEntry(property) {
  const id = linePropertyId('TableToSlidesProps', property);
  if (property === 'verbose') {
    return tableToSlidesDifferenceEntry(
      id,
      ['PptxDocument.tableToSlides'],
      'PptxGenJS verbose produces process-global console traces without changing the deck; native deliberately keeps tableToSlides deterministic and free of global logging.',
      { serialization: false, client: false },
    );
  }
  if (['w', 'h'].includes(property)) {
    return tableToSlidesDifferenceEntry(
      id,
      tableToSlidesNative(property),
      `Native maps TableToSlidesProps.${property} to strict exact geometry with explicit EMU or inches() units, preserves legal zero values, and uses width/height names instead of PptxGenJS truthy implicit-inch behavior.`,
    );
  }
  if (['x', 'y', 'slideMargin', 'autoPageSlideStartY'].includes(property)) {
    return tableToSlidesDifferenceEntry(
      id,
      tableToSlidesNative(property),
      `Native maps TableToSlidesProps.${property} to strict exact geometry with explicit EMU or inches() units and preserves legal zero values instead of PptxGenJS truthy implicit-inch behavior.`,
    );
  }
  return tableToSlidesDifferenceEntry(
    id,
    tableToSlidesNative(property),
    `Native covers TableToSlidesProps.${property} with a strict detached addition record and transactional page creation instead of PptxGenJS caller mutation and permissive nested inputs.`,
  );
}

function tableToSlidesNestedDifferenceEntry(owner, field, nativeField) {
  return tableToSlidesDifferenceEntry(
    `inline:interface:TableToSlidesProps@property:${owner}@property:${owner}.${field}`,
    [`TableToSlides${owner.slice(0, 1).toUpperCase()}${owner.slice(1)}.${nativeField}`],
    `Native maps TableToSlidesProps.${owner}.${field} to the detached, strictly typed ${nativeField} field with native geometry and content types while preserving the same legal page addition.`,
  );
}

function tableToSlidesDeprecatedEntry(property) {
  const canonicalProperty = property === 'addHeaderToEach'
    ? 'autoPageRepeatHeader'
    : 'autoPageSlideStartY';
  return {
    id: linePropertyId('TableToSlidesProps', property),
    status: 'deprecated-alias',
    native: tableToSlidesNative(canonicalProperty),
    evidence: tableToSlidesFamilyEvidence(),
    canonical: linePropertyId('TableToSlidesProps', canonicalProperty),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE,
    },
    serialization: true,
    client: true,
    note: `PptxGenJS 4.0.1 still consumes deprecated ${property}; native exposes only its canonical ${canonicalProperty} capability.`,
  };
}

function tableToSlidesDefectEntry(id, property) {
  const notes = {
    autoPage: 'PptxGenJS tableToSlides always paginates even when autoPage is false; native implements the declared false value, but that strict correction does not make the broken upstream field a delivered capability.',
    autoPageHeaderRows: 'PptxGenJS never reads autoPageHeaderRows in tableToSlides and always repeats every thead row; native supports an explicit count as a strict correction.',
    colW: 'PptxGenJS overwrites caller colW with DOM-derived widths before output, so neither declared scalar nor vector input controls the table; native columnWidths is a strict correction.',
    margin: 'PptxGenJS margin only enters an inconsistent pagination fallback and is not forwarded as the declared final cell margin; computed CSS padding remains the effective source.',
    rowH: 'PptxGenJS never forwards tableToSlides rowH to the final table, so neither declared scalar nor vector branch affects output.',
  };
  return {
    id,
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: TABLE_TO_SLIDES_FAMILY_CONTROL_TITLE,
    },
    note: notes[property]
      ?? `PptxGenJS tableToSlides does not forward inherited ${property} to the final table; computed DOM/CSS state remains the effective input and native does not copy the inert alias.`,
  };
}

function tableToSlidesTopEntry(property, status) {
  if (status === 'supported') return tableToSlidesSupportedEntry(property);
  if (status === 'deliberate-difference') {
    return tableToSlidesTopDifferenceEntry(property);
  }
  if (status === 'deprecated-alias') return tableToSlidesDeprecatedEntry(property);
  return tableToSlidesDefectEntry(
    linePropertyId('TableToSlidesProps', property),
    property,
  );
}

const TABLE_TO_SLIDES_FAMILY_ENTRIES = Object.freeze([
  tableToSlidesDifferenceEntry(
    'class:PptxGenJS#tableToSlides',
    ['PptxDocument.tableToSlides'],
    'Native tableToSlides returns a Promise of frozen created SlideModel identities, snapshots caller input, and rolls back atomically instead of returning undefined and mutating caller options.',
  ),
  ...Object.entries(TABLE_TO_SLIDES_TOP_STATUS).map(([property, status]) =>
    tableToSlidesTopEntry(property, status)),
  ...Object.entries(TABLE_TO_SLIDES_NESTED_FIELDS).flatMap(([owner, fields]) =>
    Object.entries(fields).map(([field, nativeField]) =>
      tableToSlidesNestedDifferenceEntry(owner, field, nativeField))),
  ...Object.entries(TABLE_TO_SLIDES_UNION_TOKENS).flatMap(([property, tokens]) =>
    tokens.map((token) => tableToSlidesDefectEntry(
      `union:${linePropertyId('TableToSlidesProps', property)}#${token}`,
      property,
    ))),
]);

const ADD_TABLE_CORE_CONTROL_TITLE =
  'locks the remaining addTable core declarations against PptxGenJS 4.0.1';
const ADD_TABLE_CORE_OOXML_TITLE =
  'creates, edits, duplicates, rolls back, and reopens a basic table';
const ADD_TABLE_CORE_STATUS = Object.freeze({
  BorderProps: Object.freeze({
    color: 'deliberate-difference',
    pt: 'deliberate-difference',
    type: 'deliberate-difference',
  }),
  TableCellProps: Object.freeze({
    align: 'supported',
    autoPageCharWeight: 'defect-excluded',
    autoPageLineWeight: 'defect-excluded',
    border: 'deliberate-difference',
    colspan: 'supported',
    hyperlink: 'supported',
    margin: 'deliberate-difference',
    rowspan: 'supported',
    transparency: 'supported',
    valign: 'supported',
  }),
  TableProps: Object.freeze({
    align: 'supported',
    autoPage: 'supported',
    autoPageCharWeight: 'supported',
    autoPageHeaderRows: 'supported',
    autoPageLineWeight: 'supported',
    autoPageRepeatHeader: 'supported',
    autoPageSlideStartY: 'deliberate-difference',
    border: 'deliberate-difference',
    colW: 'deliberate-difference',
    h: 'deliberate-difference',
    margin: 'deliberate-difference',
    newSlideStartY: 'deprecated-alias',
    objectName: 'deliberate-difference',
    rowH: 'deliberate-difference',
    transparency: 'defect-excluded',
    valign: 'supported',
    verbose: 'deliberate-difference',
    w: 'deliberate-difference',
    x: 'deliberate-difference',
    y: 'deliberate-difference',
  }),
});
const ADD_TABLE_CORE_UNIONS = Object.freeze([
  Object.freeze({
    owner: 'BorderProps',
    property: 'type',
    tokens: Object.freeze(['dash', 'none', 'solid']),
  }),
  Object.freeze({
    owner: 'TableCellProps',
    property: 'border',
    tokens: Object.freeze([
      'BorderProps',
      '[BorderProps,BorderProps,BorderProps,BorderProps]',
    ]),
  }),
  Object.freeze({
    owner: 'TableProps',
    property: 'border',
    tokens: Object.freeze([
      'BorderProps',
      '[BorderProps,BorderProps,BorderProps,BorderProps]',
    ]),
  }),
  Object.freeze({
    owner: 'TableProps',
    property: 'colW',
    tokens: Object.freeze(['number', 'number[]']),
  }),
  Object.freeze({
    owner: 'TableProps',
    property: 'rowH',
    tokens: Object.freeze(['number', 'number[]']),
  }),
]);

function addTableCoreFamilyEvidence() {
  return {
    code: [
      {
        path: 'packages/model/src/slide.ts',
        pattern: 'export interface AddTableOptions {',
      },
      {
        path: 'packages/model/src/table-create.internal.ts',
        pattern: 'export function normalizeTableDefinition(',
      },
      {
        path: 'packages/model/src/table-auto-page.internal.ts',
        pattern: 'export function normalizeTableAutoPageRequest(',
      },
    ],
    tests: [
      {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: ADD_TABLE_CORE_CONTROL_TITLE,
      },
      {
        path: 'packages/sdk/src/index.test.ts',
        title: ADD_TABLE_CORE_OOXML_TITLE,
      },
    ],
    package: [
      {
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const tableCreation =',
      },
      {
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const tableAutoPageState = {',
      },
      {
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const tableBordersState = {',
      },
    ],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: ADD_TABLE_CORE_OOXML_TITLE,
    }],
    clients: [
      {
        path: 'scripts/playwright-browser-smoke.js',
        pattern: 'const tableAutoPageState = {',
      },
      {
        path: 'scripts/playwright-browser-smoke.js',
        pattern: 'const tableBordersState = {',
      },
    ],
  };
}

function addTableCoreNative(owner, property) {
  if (owner === 'Slide' && property === 'addTable') {
    return ['SlideModel.addTable', 'TableModel'];
  }
  if (owner === 'BorderProps') {
    if (property === 'color') return ['TableCellBorder.color'];
    if (property === 'pt') return ['TableCellBorder.width'];
    return ['TableCellBorder.kind', 'TableCellBorder.style'];
  }
  if (owner === 'TableCellProps') {
    const native = {
      align: ['AddTableCellOptions.align', 'TableCell.horizontalAlignment'],
      border: ['AddTableCellOptions.border', 'TableCell.borders'],
      colspan: ['AddTableCellOptions.colspan', 'TableModel.mergeRegions'],
      hyperlink: ['AddTableCellOptions.hyperlink', 'TableCell.hyperlink'],
      margin: ['AddTableCellOptions.margin', 'TableCell.margins'],
      rowspan: ['AddTableCellOptions.rowspan', 'TableModel.mergeRegions'],
      transparency: ['RichTextRunStyle.transparency', 'TableCell.richText'],
      valign: ['AddTableCellOptions.valign', 'TableCell.verticalAlignment'],
    };
    return native[property] ?? [];
  }
  const nativeProperty = {
    border: 'border',
    colW: 'columnWidths',
    h: 'height',
    margin: 'margin',
    objectName: 'name',
    rowH: 'rowHeights',
    w: 'width',
  }[property] ?? property;
  return [`AddTableOptions.${nativeProperty}`, 'SlideModel.addTable', 'TableModel'];
}

function addTableCoreControl() {
  return {
    path: 'packages/pptxgenjs-adapter/src/index.test.ts',
    pattern: ADD_TABLE_CORE_CONTROL_TITLE,
  };
}

function addTableCoreSupportedEntry(owner, property) {
  return {
    id: linePropertyId(owner, property),
    status: 'supported',
    native: addTableCoreNative(owner, property),
    evidence: addTableCoreFamilyEvidence(),
    serialization: true,
    client: true,
    note: owner === 'TableCellProps' && property === 'transparency'
      ? 'Native preserves the same legal table-cell color and alpha final state through RichTextRunStyle.transparency, editable TableCell.richText, packed output, and reopen.'
      : `Native preserves legal ${owner}.${property} table behavior under the same public name with strict detached input, editable output, packed-package use, browser execution, and reopen evidence.`,
  };
}

function addTableCoreDifferenceNote(owner, property) {
  if (owner === 'BorderProps' || property === 'border') {
    return 'Native represents table borders as a strict kind/color/width/style union with sRGB or scheme colors, explicit zero-width and none intent, and partial side maps instead of permissive PptxGenJS type/color/pt fallbacks.';
  }
  if (property === 'margin') {
    return owner === 'TableCellProps'
      ? 'PptxGenJS mixes scalar-inch and vector-point table-cell margins; native uses one strict point-based scalar, tuple, or partial-side contract and preserves the same legal final OOXML state.'
      : 'PptxGenJS table-level margins use implicit inches; native uses one strict point-based scalar, tuple, or partial-side contract and preserves the same legal final OOXML state.';
  }
  if (property === 'autoPageSlideStartY' || ['w', 'h', 'x', 'y'].includes(property)) {
    return `Native maps TableProps.${property} to exact geometry with explicit EMU or inches() units and preserves legal zero values instead of PptxGenJS truthy implicit-inch behavior.`;
  }
  if (property === 'colW' || property === 'rowH') {
    const nativeProperty = property === 'colW' ? 'columnWidths' : 'rowHeights';
    return `Native names TableProps.${property} as ${nativeProperty}, requires explicit EMU or inches() units, and normalizes scalar/vector intent deterministically instead of PptxGenJS mutation and fallback behavior.`;
  }
  if (property === 'objectName') {
    return 'Native names TableProps.objectName as AddTableOptions.name, snapshots the original XML-safe string, and returns an editable TableModel instead of entity-encoding the caller option in place.';
  }
  if (property === 'verbose') {
    return 'PptxGenJS verbose produces process-global auto-page console traces without changing table output; native deliberately keeps SlideModel.addTable deterministic and free of global logging.';
  }
  return `Native covers ${owner}.${property} through a strict detached table contract and transactional editing instead of PptxGenJS permissive input and caller mutation.`;
}

function addTableCoreDifferenceEntry(id, owner, property) {
  const verbose = owner === 'TableProps' && property === 'verbose';
  return {
    id,
    status: 'deliberate-difference',
    native: addTableCoreNative(owner, property),
    evidence: addTableCoreFamilyEvidence(),
    control: addTableCoreControl(),
    serialization: !verbose,
    client: !verbose,
    note: addTableCoreDifferenceNote(owner, property),
  };
}

function addTableCoreDeprecatedEntry() {
  return {
    id: linePropertyId('TableProps', 'newSlideStartY'),
    status: 'deprecated-alias',
    native: addTableCoreNative('TableProps', 'autoPageSlideStartY'),
    evidence: addTableCoreFamilyEvidence(),
    canonical: linePropertyId('TableProps', 'autoPageSlideStartY'),
    control: addTableCoreControl(),
    serialization: true,
    client: true,
    note: 'PptxGenJS 4.0.1 still consumes deprecated TableProps.newSlideStartY; native exposes only its canonical autoPageSlideStartY capability.',
  };
}

function addTableCoreDefectEntry(owner, property) {
  const notes = {
    autoPageCharWeight: 'PptxGenJS overwrites every cell autoPageCharWeight with the table-level value or null before measurement, so the declared cell-local value never controls output.',
    autoPageLineWeight: 'PptxGenJS never reads cell-local autoPageLineWeight; only the table-level value enters automatic row-height calculation.',
    transparency: 'PptxGenJS does not propagate TableProps.transparency to table cells or text runs, so every legal table-level value is inert.',
  };
  return {
    id: linePropertyId(owner, property),
    status: 'defect-excluded',
    native: [],
    evidence: {
      code: [],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: ADD_TABLE_CORE_CONTROL_TITLE,
      }],
      package: [],
      ooxml: [],
      clients: [],
    },
    control: addTableCoreControl(),
    note: notes[property],
  };
}

function addTableCorePropertyEntry(owner, property, status) {
  if (status === 'supported') return addTableCoreSupportedEntry(owner, property);
  if (status === 'deliberate-difference') {
    return addTableCoreDifferenceEntry(linePropertyId(owner, property), owner, property);
  }
  if (status === 'deprecated-alias') return addTableCoreDeprecatedEntry();
  return addTableCoreDefectEntry(owner, property);
}

const ADD_TABLE_CORE_FAMILY_ENTRIES = Object.freeze([
  addTableCoreDifferenceEntry(
    'method:Slide#addTable',
    'Slide',
    'addTable',
  ),
  ...Object.entries(ADD_TABLE_CORE_STATUS).flatMap(([owner, properties]) =>
    Object.entries(properties).map(([property, status]) =>
      addTableCorePropertyEntry(owner, property, status))),
  ...ADD_TABLE_CORE_UNIONS.flatMap(({ owner, property, tokens }) =>
    tokens.map((token) => addTableCoreDifferenceEntry(
      `union:${linePropertyId(owner, property)}#${token}`,
      owner,
      property,
    ))),
]);

const PRESENTATION_ROOT_OUTPUT_CONTROL_TITLE =
  'locks the presentation root, output, and theme declarations against PptxGenJS 4.0.1';
const PRESENTATION_ROOT_OUTPUT_SUPPORTED_IDS = Object.freeze([
  'class:PptxGenJS#addSlide',
  ...['author', 'company', 'revision', 'rtlMode', 'subject', 'theme', 'title']
    .map((property) => `class:PptxGenJS@property:${property}`),
  ...['height', 'width'].map((property) => linePropertyId('PresLayout', property)),
  ...['bodyFontFace', 'headFontFace']
    .map((property) => linePropertyId('ThemeProps', property)),
  linePropertyId('WriteBaseProps', 'compression'),
  'union:WRITE_OUTPUT_TYPE#JSZIP_OUTPUT_TYPE',
  ...['ArrayBuffer', 'Blob', 'Uint8Array', 'string']
    .map((token) => `union:class:PptxGenJS#write@path:return#${token}`),
]);
const PRESENTATION_ROOT_OUTPUT_DIFFERENCE_IDS = Object.freeze([
  ...[
    'addSection',
    'defineLayout',
    'defineSlideMaster',
    'stream',
    'write',
    'writeFile',
  ].map((method) => `class:PptxGenJS#${method}`),
  'class:PptxGenJS@property:layout',
  linePropertyId('PresLayout', 'name'),
  linePropertyId('WriteFileProps', 'compression'),
  linePropertyId('WriteFileProps', 'fileName'),
  linePropertyId('WriteProps', 'compression'),
  linePropertyId('WriteProps', 'outputType'),
  'union:WRITE_OUTPUT_TYPE#STREAM',
  'union:class:PptxGenJS#stream@path:return#Uint8Array',
]);
const PRESENTATION_ROOT_OUTPUT_DEFECT_IDS = Object.freeze(
  ['ArrayBuffer', 'Blob', 'string']
    .map((token) => `union:class:PptxGenJS#stream@path:return#${token}`),
);

function presentationRootOutputCategory(id) {
  if (/^class:PptxGenJS@property:(?:author|company|revision|rtlMode|subject|title)$/u.test(id)) {
    return 'metadata';
  }
  if (id.includes('ThemeProps') || id === 'class:PptxGenJS@property:theme') return 'theme';
  if (
    id.includes('PresLayout')
    || id === 'class:PptxGenJS@property:layout'
    || id === 'class:PptxGenJS#defineLayout'
  ) return 'layout';
  if (id === 'class:PptxGenJS#addSection') return 'section';
  if (id === 'class:PptxGenJS#addSlide' || id === 'class:PptxGenJS#defineSlideMaster') {
    return 'master';
  }
  if (id.includes('#stream') || id.endsWith('#STREAM')) return 'stream';
  if (id.includes('WriteFileProps') || id === 'class:PptxGenJS#writeFile') return 'file';
  return 'write';
}

function presentationRootOutputNative(id) {
  if (id === 'class:PptxGenJS#addSection') {
    return ['PresentationModel.addSection', 'PresentationSection'];
  }
  if (id === 'class:PptxGenJS#addSlide') return ['PptxDocument.addSlide', 'SlideModel'];
  if (id === 'class:PptxGenJS#defineLayout' || id.includes('PresLayout')
      || id === 'class:PptxGenJS@property:layout') {
    return ['PptxDocument.create', 'PptxDocument.presLayout'];
  }
  if (id === 'class:PptxGenJS#defineSlideMaster') {
    return ['PptxDocument.defineSlideMaster', 'SlideLayoutModel'];
  }
  if (id === 'class:PptxGenJS#stream' || id.includes('#stream@path:return')
      || id.endsWith('#STREAM')) {
    return ['PptxDocument.stream', 'PptxNodeReadableStream'];
  }
  if (id === 'class:PptxGenJS#writeFile' || id.includes('WriteFileProps')) {
    return ['PptxDocument.writeFile', 'WriteBaseOptions'];
  }
  if (id.includes('Write') || id.includes('#write@path:return')) {
    return ['PptxDocument.write', 'WriteOptions', 'OUTPUT_TYPES'];
  }
  if (id.includes('ThemeProps') || id === 'class:PptxGenJS@property:theme') {
    return ['PptxDocument.theme', 'PresentationThemeOptions'];
  }
  const property = id.split('@property:')[1];
  return [`PptxDocument.${property}`];
}

function presentationRootOutputEvidence(id) {
  const category = presentationRootOutputCategory(id);
  const aggregate = {
    path: 'packages/pptxgenjs-adapter/src/index.test.ts',
    title: PRESENTATION_ROOT_OUTPUT_CONTROL_TITLE,
  };
  if (category === 'metadata') {
    const property = id.split('@property:')[1];
    const titles = {
      author: 'imports and reopens PptxGenJS presentation author metadata from public output',
      company: 'imports and reopens PptxGenJS presentation company metadata from public output',
      revision: 'imports and reopens PptxGenJS presentation revision metadata from public output',
      rtlMode: 'imports only direct PptxGenJS presentation RTL and reopens it',
      subject: 'imports and reopens PptxGenJS presentation subject metadata from public output',
      title: 'imports and reopens PptxGenJS presentation title metadata from public output',
    };
    const ooxmlPatterns = {
      author: 'presentation author',
      company: 'presentation company',
      revision: 'presentation revision',
      rtlMode: 'presentation RTL',
      subject: 'presentation subject',
      title: 'presentation title',
    };
    return {
      code: [{ path: 'packages/model/src/presentation.ts', pattern: `get ${property}()` }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: titles[property],
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const metadata = PptxDocument.create({' }],
      ooxml: [{
        path: 'packages/sdk/src/index.test.ts',
        pattern: ooxmlPatterns[property],
      }],
      clients: [],
    };
  }
  if (category === 'theme') {
    return {
      code: [{ path: 'packages/sdk/src/index.ts', pattern: 'get theme(): PresentationTheme | undefined {' }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: 'matches public PptxGenJS presentation theme fonts and reopens a partial edit',
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const themed = PptxDocument.create({' }],
      ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'creates presentation theme fonts with explicit PptxGenJS partial fallbacks' }],
      clients: [],
    };
  }
  if (category === 'layout') {
    return {
      code: [{ path: 'packages/sdk/src/index.ts', pattern: 'get presLayout(): PresentationLayout {' }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: 'matches the public presentation layout projection and locks the custom-name boundary',
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const presentationLayoutState = {' }],
      ooxml: [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: 'matches the public presentation layout projection and locks the custom-name boundary' }],
      clients: [{ path: 'scripts/playwright-browser-smoke.js', pattern: 'const presentationLayoutState = {' }],
    };
  }
  if (category === 'section') {
    return {
      code: [{ path: 'packages/model/src/presentation.ts', pattern: 'addSection(options: AddSectionOptions): PresentationSection {' }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: 'imports and matches public PptxGenJS presentation sections',
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const sectioned = PptxDocument.create();' }],
      ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'creates, reads, and atomically edits detached presentation sections' }],
      clients: [],
    };
  }
  if (category === 'master') {
    return {
      code: [{ path: 'packages/sdk/src/index.ts', pattern: 'async defineSlideMaster(options: DefineSlideMasterOptions): Promise<SlideLayoutModel> {' }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: 'diagnoses, canonicalizes, reorders, and reopens named-master public output',
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const masterLayoutChecks = {' }],
      ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'define slide master works in all six presentation formats' }],
      clients: [{ path: 'scripts/playwright-browser-smoke.js', pattern: 'const masterLayoutState = {' }],
    };
  }
  if (category === 'stream') {
    return {
      code: [{ path: 'packages/sdk/src/index.ts', pattern: 'async stream(options: WriteBaseOptions = {}): Promise<PptxNodeReadableStream> {' }],
      tests: [aggregate, {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: 'records the public PptxGenJS stream result and provides a real Node readable',
      }],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const nodeReadableStreamState = {' }],
      ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'keeps Node path, stream, and writeFile adapters working' }],
      clients: [{ path: 'scripts/playwright-browser-smoke.js', pattern: 'const nodeReadableStreamState = {' }],
    };
  }
  if (category === 'file') {
    return {
      code: [{ path: 'packages/sdk/src/index.ts', pattern: 'async writeFile(path: string, options: WriteBaseOptions = {}): Promise<void> {' }],
      tests: [aggregate],
      package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: "typedWriteOutputDocument.writeFile('typed.pptx', { compression: true })" }],
      ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'keeps Node path, stream, and writeFile adapters working' }],
      clients: [],
    };
  }
  return {
    code: [{ path: 'packages/sdk/src/index.ts', pattern: 'async write<TOutputType extends OutputType' }],
    tests: [aggregate, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: 'matches the PptxGenJS output type runtime catalog and return kinds',
    }, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: 'matches legal compression booleans without copying the explicit-output defect',
    }],
    package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: 'const writeOutputTypeState = {' }],
    ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: 'writes every selected output representation without changing canonical bytes' }],
    clients: [{ path: 'scripts/playwright-browser-smoke.js', pattern: 'const writeOutputTypeState = {' }],
  };
}

function presentationRootOutputDifferenceNote(id) {
  if (id === 'class:PptxGenJS#addSection') {
    return 'PptxGenJS returns undefined and treats order zero as an omitted truthy value; native returns an editable PresentationSection and preserves explicit zero insertion.';
  }
  if (id === 'class:PptxGenJS#defineSlideMaster') {
    return 'PptxGenJS returns undefined, clones permissively, and permits duplicate names; native resolves an editable SlideLayoutModel and commits unique strict definitions atomically.';
  }
  if (id === 'class:PptxGenJS#defineLayout' || id === 'class:PptxGenJS@property:layout'
      || id === linePropertyId('PresLayout', 'name')) {
    return 'PptxGenJS registers and retains user layout names; native selects an explicit slide size at creation and canonicalizes reopened custom layouts to the stable name custom.';
  }
  if (id.includes('#stream') || id.endsWith('#STREAM')) {
    return 'PptxGenJS stream and STREAM resolve to an in-memory Buffer; native exposes the equivalent byte representations through write and reserves stream for a real Node Readable.';
  }
  if (id.includes('WriteFileProps') || id === 'class:PptxGenJS#writeFile') {
    return 'PptxGenJS accepts writeFile({ fileName, compression }), returns the final name, and ignores compression on the forced explicit output; native uses writeFile(path, options), returns void, and applies compression.';
  }
  if (id === linePropertyId('WriteProps', 'compression')) {
    return 'PptxGenJS applies compression only through its implicit or STREAM path and ignores it for explicit JSZip output selectors; native applies the boolean consistently.';
  }
  if (id === 'class:PptxGenJS#write' || id === linePropertyId('WriteProps', 'outputType')) {
    return 'Native matches all six JSZip representations through write but routes the misleading PptxGenJS STREAM selector to its separate real-stream API.';
  }
  return 'Native preserves the legal final-state capability through a stricter editable presentation lifecycle instead of the permissive PptxGenJS root contract.';
}

function presentationRootOutputEntry(id, status) {
  const evidence = presentationRootOutputEvidence(id);
  if (status === 'defect-excluded') {
    return {
      id,
      status,
      native: [],
      evidence: { code: [], tests: evidence.tests, package: [], ooxml: [], clients: [] },
      control: {
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        pattern: PRESENTATION_ROOT_OUTPUT_CONTROL_TITLE,
      },
      note: `PptxGenJS 4.0.1 stream always resolves to a Buffer/Uint8Array in Node; its declared ${id.split('#').at(-1)} return branch cannot be selected and never occurs.`,
    };
  }
  const entry = {
    id,
    status,
    native: presentationRootOutputNative(id),
    evidence,
    serialization: true,
    client: ['layout', 'master', 'write'].includes(presentationRootOutputCategory(id)),
    note: status === 'supported'
      ? 'Native preserves the same legal presentation-root capability with detached state, strict validation, packed-package use, serialization, and reopen evidence.'
      : presentationRootOutputDifferenceNote(id),
  };
  if (status === 'deliberate-difference') {
    entry.control = {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: PRESENTATION_ROOT_OUTPUT_CONTROL_TITLE,
    };
  }
  return entry;
}

const PRESENTATION_ROOT_OUTPUT_FAMILY_ENTRIES = Object.freeze([
  ...PRESENTATION_ROOT_OUTPUT_SUPPORTED_IDS.map((id) =>
    presentationRootOutputEntry(id, 'supported')),
  ...PRESENTATION_ROOT_OUTPUT_DIFFERENCE_IDS.map((id) =>
    presentationRootOutputEntry(id, 'deliberate-difference')),
  ...PRESENTATION_ROOT_OUTPUT_DEFECT_IDS.map((id) =>
    presentationRootOutputEntry(id, 'defect-excluded')),
]);

const SLIDE_SECTION_CONTROL_TITLE =
  'locks the slide and section declarations against PptxGenJS 4.0.1';
const SLIDE_SECTION_SUPPORTED_IDS = Object.freeze([
  ...['addChart', 'addImage', 'addNotes', 'addShape', 'addTable', 'addText']
    .map((property) => linePropertyId('PresSlide', property)),
  linePropertyId('PresSlide', 'hidden'),
  'method:Slide#addNotes',
  'property:Slide#hidden',
  linePropertyId('SectionProps', 'title'),
]);
const SLIDE_SECTION_DIFFERENCE_IDS = Object.freeze([
  linePropertyId('AddSlideProps', 'masterName'),
  linePropertyId('AddSlideProps', 'sectionTitle'),
  ...['addMedia', 'background', 'color', 'slideNumber']
    .map((property) => linePropertyId('PresSlide', property)),
  linePropertyId('SectionProps', 'order'),
  ...['addImage', 'addMedia', 'addShape', 'addText']
    .map((method) => `method:Slide#${method}`),
  ...['background', 'color', 'newAutoPagedSlides', 'slideNumber']
    .map((property) => `property:Slide#${property}`),
]);
const SLIDE_SECTION_DEPRECATED_IDS = Object.freeze(['property:Slide#bkgd']);

function slideSectionCategory(id) {
  if (id.includes('AddSlideProps@property:masterName')) return 'master';
  if (id.includes('AddSlideProps@property:sectionTitle') || id.includes('SectionProps')) {
    return 'section';
  }
  if (id.endsWith('addNotes')) return 'notes';
  if (id.endsWith('#hidden') || id.endsWith('@property:hidden')) return 'hidden';
  if (id.endsWith('addChart')) return 'chart';
  if (id.endsWith('addImage')) return 'image';
  if (id.endsWith('addMedia')) return 'media';
  if (id.endsWith('addShape')) return 'shape';
  if (id.endsWith('addTable')) return 'table';
  if (id.endsWith('addText')) return 'text';
  if (id.endsWith('#background') || id.endsWith('@property:background')
      || id.endsWith('#bkgd')) return 'background';
  if (id.endsWith('#color') || id.endsWith('@property:color')) return 'color';
  if (id.endsWith('#slideNumber') || id.endsWith('@property:slideNumber')) {
    return 'slide-number';
  }
  return 'auto-page';
}

const SLIDE_SECTION_EVIDENCE = Object.freeze({
  master: Object.freeze({
    code: Object.freeze({
      path: 'packages/sdk/src/index.ts',
      pattern: 'override addSlide(options: AddSlideOptions = {}): SlideModel {',
    }),
    test: 'locks public slide master fallbacks while native definitions reject atomically',
    packagePattern: 'const masterLayoutSlide = masterLayoutDeck.addSlide({ masterName: masterLayout.name });',
    ooxmlPattern: 'define slide master works in all six presentation formats',
    clientPattern: "const masterLayoutDocument = api.PptxDocument.create({ slideSize: 'wide' });",
  }),
  section: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/presentation.ts',
      pattern: 'addSection(options: AddSectionOptions): PresentationSection {',
    }),
    test: 'imports and matches public PptxGenJS presentation sections',
    packagePattern: "const sectioned = PptxDocument.create();",
    ooxmlPattern: 'creates, reads, and atomically edits detached presentation sections',
  }),
  notes: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'addNotes(value: string): this {',
    }),
    test: 'imports and matches public PptxGenJS speaker notes output',
    packagePattern: 'const speakerNotesDeck = PptxDocument.create();',
    ooxmlPattern: 'preserves speaker notes through lifecycle, rollback, sections, hidden state, and all formats',
  }),
  hidden: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'set hidden(value: boolean) {',
    }),
    test: 'imports and matches public PptxGenJS hidden slide output',
    packagePattern: 'const hiddenDeck = PptxDocument.create();',
    ooxmlPattern: 'preserves hidden slide state through lifecycle, rollback, and all formats',
  }),
  chart: Object.freeze({
    code: Object.freeze({ path: 'packages/model/src/slide.ts', pattern: 'async addChart(' }),
    test: 'compares PptxGenJS and native chart creation return semantics',
    packagePattern: 'const nativeCharts = reopenedNativeChartModels.length === 10',
    ooxmlPattern: 'creates and reopens all native chart types through the public SDK in all six formats',
    clientPattern: 'const chartDocument = api.PptxDocument.create();',
  }),
  image: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel {',
    }),
    test: 'matches PptxGenJS embedded raster image public output semantically',
    packagePattern: 'const embeddedRasterImages = embeddedRasterImmediate',
    ooxmlPattern: 'adds detected raster image sources with immediate live model state',
    clientPattern: 'const svgDocument = api.PptxDocument.create();',
  }),
  media: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'async addAudio(source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {',
    }),
    test: 'matches valid PptxGenJS public audio and video media output semantically',
    packagePattern: 'const embeddedMedia = packedMediaDeduplicated',
    ooxmlPattern: 'creates every public media source, MIME family, poster family, and external mode',
    clientPattern: 'const mediaDocument = api.PptxDocument.create();',
  }),
  shape: Object.freeze({
    code: Object.freeze({ path: 'packages/model/src/slide.ts', pattern: 'addShape(' }),
    test: 'matches representative preset shape public output semantically',
    packagePattern: 'const presetShapes = PRESET_SHAPE_TYPES.length === 178',
    ooxmlPattern: 'creates preset shapes with deterministic defaults, transforms, order, and identity',
    clientPattern: 'const presetShapeDocument = api.PptxDocument.create();',
  }),
  table: Object.freeze({
    code: Object.freeze({ path: 'packages/model/src/slide.ts', pattern: 'addTable(' }),
    test: 'imports PptxGenJS 4.0.1 auto-page output with repeated headers and editable tables',
    packagePattern: 'const tableAutoPageDocument = PptxDocument.create();',
    ooxmlPattern: 'exports measured table auto-page contracts through all six presentation formats',
    clientPattern: 'const tableAutoPageDocument = api.PptxDocument.create();',
  }),
  text: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'addText(value: string, options: AddTextOptions = {}): ShapeModel {',
    }),
    test: 'matches PptxGenJS shape and text percentage coordinate output with explicit native units',
    packagePattern: 'const slideDefaultColorDeck = PptxDocument.create();',
    ooxmlPattern: 'creates, edits, and round-trips a basic text shape with stable identity',
    clientPattern: 'const textShapeFillDocument = api.PptxDocument.create();',
  }),
  background: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'set background(value: SlideBackground | undefined) {',
    }),
    test: 'matches supported public PptxGenJS slide backgrounds and locks none divergences',
    packagePattern: 'const slideBackgroundDeck = PptxDocument.create();',
    ooxmlPattern: 'round-trips image, none, solid, and gradient backgrounds twice in all six formats',
    clientPattern: 'const backgroundDocument = api.PptxDocument.create();',
  }),
  color: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'set color(value: RichTextColor | undefined) {',
    }),
    test: 'matches public PptxGenJS slide default colors and locks intentional differences',
    packagePattern: 'const slideDefaultColorDeck = PptxDocument.create();',
    ooxmlPattern: 'round-trips materialized slide default colors twice in all six formats',
    clientPattern: 'const slideDefaultColorDocument = api.PptxDocument.create();',
  }),
  'slide-number': Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'set slideNumber(value: SlideNumberOptions | undefined) {',
    }),
    test: 'imports public slide-number variants and locks PptxGenJS 4.0.1 differences',
    packagePattern: 'const slideNumberDeck = PptxDocument.create({ firstSlideNumber: 5 });',
    ooxmlPattern: 'round-trips all three slide-number owners twice in all six formats',
    clientPattern: 'const slideNumberDocument = api.PptxDocument.create({ firstSlideNumber: -2 });',
  }),
  'auto-page': Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide.ts',
      pattern: 'get newAutoPagedSlides(): readonly SlideModel[] {',
    }),
    test: 'imports PptxGenJS 4.0.1 auto-page output with repeated headers and editable tables',
    packagePattern: 'const tableAutoPageGenerated = tableAutoPageSource.newAutoPagedSlides;',
    ooxmlPattern: 'exports measured table auto-page contracts through all six presentation formats',
    clientPattern: 'const tableAutoPageGenerated = tableAutoPageSource.newAutoPagedSlides;',
  }),
});

function slideSectionNative(id) {
  const category = slideSectionCategory(id);
  const mappings = {
    master: ['PptxDocument.addSlide', 'SlideModel'],
    section: ['PresentationModel.addSection', 'PresentationSection', 'PptxDocument.addSlide'],
    notes: ['SlideModel.addNotes'],
    hidden: ['SlideModel.hidden'],
    chart: ['SlideModel.addChart', 'PptxDocument.addChart', 'ChartModel'],
    image: ['SlideModel.addImage', 'PptxDocument.addImage', 'ImageModel'],
    media: ['SlideModel.addAudio', 'SlideModel.addVideo', 'SlideModel.addOnlineVideo', 'MediaModel'],
    shape: ['SlideModel.addShape', 'ShapeModel'],
    table: ['SlideModel.addTable', 'TableModel'],
    text: ['SlideModel.addText', 'ShapeModel'],
    background: ['SlideModel.background', 'SlideBackground'],
    color: ['SlideModel.color', 'RichTextColor'],
    'slide-number': ['SlideModel.slideNumber', 'SlideNumberOptions'],
    'auto-page': ['SlideModel.newAutoPagedSlides', 'SlideModel.addTable'],
  };
  return mappings[category];
}

function slideSectionEvidence(id) {
  const category = slideSectionCategory(id);
  const anchors = SLIDE_SECTION_EVIDENCE[category];
  return {
    code: [anchors.code],
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: SLIDE_SECTION_CONTROL_TITLE,
    }, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: anchors.test,
    }],
    package: [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: anchors.packagePattern,
    }],
    ooxml: [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: anchors.ooxmlPattern,
    }],
    clients: anchors.clientPattern === undefined ? [] : [{
      path: 'scripts/playwright-browser-smoke.js',
      pattern: anchors.clientPattern,
    }],
  };
}

function slideSectionDifferenceNote(id) {
  if (id === linePropertyId('AddSlideProps', 'masterName')) {
    return 'PptxGenJS silently falls back to its default layout for an unknown master name; native resolves known names and rejects an unknown reference before mutation.';
  }
  if (id === linePropertyId('AddSlideProps', 'sectionTitle')) {
    return 'PptxGenJS warns but still creates a loose slide for an unknown section title; native resolves known titles and rejects an unknown reference before mutation.';
  }
  if (id === linePropertyId('SectionProps', 'order')) {
    return 'PptxGenJS treats declared numeric order zero as omitted and appends; native preserves zero as the first insertion index while matching documented positive indices.';
  }
  if (id === linePropertyId('PresSlide', 'addMedia')) {
    return 'PptxGenJS exposes one generic callable addMedia property; native deliberately splits the capability into typed addAudio, addVideo, and addOnlineVideo operations.';
  }
  if (/^method:Slide#add(?:Image|Media|Shape|Text)$/u.test(id)) {
    return 'PptxGenJS synchronously returns the owning slide and may retain or mutate caller state; native returns a typed live model, prepares asynchronous sources where required, and snapshots validated input.';
  }
  const category = slideSectionCategory(id);
  const notes = {
    background: 'PptxGenJS retains and may mutate permissive background objects and has malformed none fallbacks; native uses strict detached background state with explicit no-fill and transactional media loading.',
    color: 'PptxGenJS uses a permissive hex string with black materialization fallbacks; native uses strict detached sRGB or scheme RichTextColor state and canonical theme inheritance.',
    'slide-number': 'PptxGenJS retains permissive caller state and has lossy defaults plus fixed shape-id collisions; native snapshots strict editable options and allocates canonical unique OOXML owners.',
    'auto-page': 'PptxGenJS starts undefined, reuses existing following slides, and mutates option and hyperlink inputs; native returns a frozen continuation list containing only transactionally created pages.',
  };
  return notes[category];
}

function slideSectionEntry(id, status) {
  const evidence = slideSectionEvidence(id);
  const entry = {
    id,
    status,
    native: slideSectionNative(id),
    evidence,
    serialization: true,
    client: evidence.clients.length > 0,
    note: status === 'supported'
      ? 'Native exposes the same legal callable or state capability with typed live models, strict validation, packed-package coverage, serialization, and editable reopen evidence.'
      : slideSectionDifferenceNote(id),
  };
  if (status === 'deliberate-difference') {
    entry.control = {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: SLIDE_SECTION_CONTROL_TITLE,
    };
  }
  if (status === 'deprecated-alias') {
    entry.canonical = 'property:Slide#background';
    entry.control = {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: SLIDE_SECTION_CONTROL_TITLE,
    };
    entry.note = 'PptxGenJS bkgd is a working deprecated setter for background.color; native exposes only the canonical strict background property.';
  }
  return entry;
}

const SLIDE_SECTION_FAMILY_ENTRIES = Object.freeze([
  ...SLIDE_SECTION_SUPPORTED_IDS.map((id) => slideSectionEntry(id, 'supported')),
  ...SLIDE_SECTION_DIFFERENCE_IDS.map((id) =>
    slideSectionEntry(id, 'deliberate-difference')),
  ...SLIDE_SECTION_DEPRECATED_IDS.map((id) =>
    slideSectionEntry(id, 'deprecated-alias')),
]);

const MASTER_BACKGROUND_SLIDE_NUMBER_CONTROL_TITLE =
  'locks master, background, and slide-number declarations against PptxGenJS 4.0.1';
const masterInlineObjectId = (property) =>
  `inline:interface:SlideMasterProps@property:objects@property:objects.${property}`;
const propertyUnionId = (owner, property, token) =>
  `union:${linePropertyId(owner, property)}#${token}`;
const MASTER_BACKGROUND_SLIDE_NUMBER_SUPPORTED_IDS = Object.freeze([
  masterInlineObjectId('placeholder.text'),
  linePropertyId('SlideMasterProps', 'title'),
  linePropertyId('SlideNumberProps', 'margin'),
  linePropertyId('SlideNumberProps', 'valign'),
]);
const MASTER_BACKGROUND_SLIDE_NUMBER_DIFFERENCE_IDS = Object.freeze([
  ...[
    'image',
    'line',
    'placeholder',
    'placeholder.options',
    'rect',
    'text',
  ].map(masterInlineObjectId),
  ...['background', 'margin', 'objects', 'slideNumber']
    .map((property) => linePropertyId('SlideMasterProps', property)),
  ...['color', 'data', 'path', 'transparency', 'type']
    .map((property) => linePropertyId('BackgroundProps', property)),
  propertyUnionId('BackgroundProps', 'type', 'none'),
  propertyUnionId('BackgroundProps', 'type', 'solid'),
  ...['align', 'h', 'w', 'x', 'y']
    .map((property) => linePropertyId('SlideNumberProps', property)),
]);
const MASTER_BACKGROUND_SLIDE_NUMBER_DEPRECATED_IDS = Object.freeze([
  linePropertyId('SlideMasterProps', 'bkgd'),
  propertyUnionId('SlideMasterProps', 'bkgd', 'string'),
  linePropertyId('BackgroundProps', 'alpha'),
  linePropertyId('BackgroundProps', 'fill'),
]);
const MASTER_BACKGROUND_SLIDE_NUMBER_DEFECT_IDS = Object.freeze([
  masterInlineObjectId('chart'),
  propertyUnionId('SlideMasterProps', 'bkgd', 'BackgroundProps'),
  linePropertyId('BackgroundProps', 'src'),
  linePropertyId('SlideNumberProps', 'transparency'),
]);

const MASTER_BACKGROUND_SLIDE_NUMBER_EVIDENCE = Object.freeze({
  master: Object.freeze({
    code: Object.freeze({
      path: 'packages/sdk/src/master-layout.ts',
      pattern: 'export interface DefineSlideMasterOptions {',
    }),
    test: 'matches public slide master objects, topology, and empty placeholder geometry',
    packagePattern:
      "const masterLayoutDeck = PptxDocument.create({ slideSize: 'wide', firstSlideNumber: 3 });",
    ooxmlPattern: 'define slide master works in all six presentation formats',
    clientPattern:
      "const masterLayoutDocument = api.PptxDocument.create({ slideSize: 'wide' });",
  }),
  background: Object.freeze({
    code: Object.freeze({
      path: 'packages/model/src/slide-background.ts',
      pattern: 'export type SlideBackground =',
    }),
    test: 'matches supported public PptxGenJS slide backgrounds and locks none divergences',
    packagePattern: "noFillBackgroundSlide.background = { kind: 'none' };",
    ooxmlPattern: 'round-trips image, none, solid, and gradient backgrounds twice in all six formats',
    clientPattern: 'const backgroundDocument = api.PptxDocument.create();',
  }),
  'background-image': Object.freeze({
    code: Object.freeze({
      path: 'packages/sdk/src/index.ts',
      pattern: 'async setSlideBackgroundImage(',
    }),
    test: 'edits and validates imported PptxGenJS backgrounds without disturbing neighbors',
    packagePattern: 'const slideBackgroundDeck = PptxDocument.create();',
    ooxmlPattern: 'round-trips image, none, solid, and gradient backgrounds twice in all six formats',
    clientPattern: 'const backgroundDocument = api.PptxDocument.create();',
  }),
  'slide-number': Object.freeze({
    code: Object.freeze({
      path: 'packages/codecs/src/slide-number.ts',
      pattern: 'export interface SlideNumberOptions {',
    }),
    test: 'imports public slide-number variants and locks PptxGenJS 4.0.1 differences',
    packagePattern: 'const slideNumberDeck = PptxDocument.create({ firstSlideNumber: 5 });',
    ooxmlPattern: 'round-trips all three slide-number owners twice in all six formats',
    clientPattern: 'const slideNumberDocument = api.PptxDocument.create({ firstSlideNumber: -2 });',
  }),
});

function masterBackgroundSlideNumberCategory(id) {
  if (id.includes('BackgroundProps')) {
    return /@property:(?:data|path|src)(?:#|$)/u.test(id)
      ? 'background-image' : 'background';
  }
  if (id.includes('SlideNumberProps')) return 'slide-number';
  return 'master';
}

function masterBackgroundSlideNumberNative(id) {
  if (MASTER_BACKGROUND_SLIDE_NUMBER_DEFECT_IDS.includes(id)) return [];
  if (id.includes('BackgroundProps')) {
    if (/@property:(?:data|path)(?:#|$)/u.test(id)) {
      return [
        'RasterImageSource',
        'PptxDocument.setSlideBackgroundImage',
        'SlideBackgroundImage',
      ];
    }
    return ['SlideBackground', 'SlideModel.background'];
  }
  if (id.includes('SlideNumberProps')) {
    const property = id.split('@property:')[1]?.split('#')[0];
    const mappings = {
      align: ['SlideNumberOptions.align', 'SlideNumber.align'],
      h: ['SlideNumberOptions.height', 'SlideNumber.height'],
      margin: ['SlideNumberOptions.margin', 'SlideNumber.margin'],
      valign: ['SlideNumberOptions.valign', 'SlideNumber.valign'],
      w: ['SlideNumberOptions.width', 'SlideNumber.width'],
      x: ['SlideNumberOptions.x', 'SlideNumber.x'],
      y: ['SlideNumberOptions.y', 'SlideNumber.y'],
    };
    return mappings[property];
  }
  if (id === linePropertyId('SlideMasterProps', 'title')) {
    return ['DefineSlideMasterOptions.title', 'PptxDocument.defineSlideMaster'];
  }
  if (id.includes('@property:background') || id.includes('@property:bkgd')) {
    return [
      'DefineSlideMasterOptions.background',
      'SlideMasterBackground',
      'SlideLayoutModel.background',
    ];
  }
  if (id === linePropertyId('SlideMasterProps', 'margin')) {
    return ['DefineSlideMasterOptions.margin', 'SlideMasterMarginInput', 'SlideLayoutModel.margin'];
  }
  if (id === linePropertyId('SlideMasterProps', 'slideNumber')) {
    return ['DefineSlideMasterOptions.slideNumber', 'SlideNumberOptions', 'SlideLayoutModel.slideNumber'];
  }
  if (id.endsWith('.image')) return ['SlideMasterObject', 'ImageModel', 'SlideLayoutModel.shapes'];
  if (id.endsWith('.placeholder') || id.includes('.placeholder.')) {
    return ['SlideMasterObject', 'SlideLayoutModel.placeholders', 'ShapeModel.placeholder'];
  }
  if (id.endsWith('.line') || id.endsWith('.rect') || id.endsWith('.text')) {
    return ['SlideMasterObject', 'ShapeModel', 'SlideLayoutModel.shapes'];
  }
  return ['DefineSlideMasterOptions.objects', 'SlideMasterObject', 'SlideLayoutModel.shapes'];
}

function masterBackgroundSlideNumberEvidence(id) {
  const category = masterBackgroundSlideNumberCategory(id);
  const anchors = MASTER_BACKGROUND_SLIDE_NUMBER_EVIDENCE[category];
  return {
    code: [anchors.code],
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: MASTER_BACKGROUND_SLIDE_NUMBER_CONTROL_TITLE,
    }, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: anchors.test,
    }],
    package: [{ path: 'scripts/smoke-npm-package.mjs', pattern: anchors.packagePattern }],
    ooxml: [{ path: 'packages/sdk/src/index.test.ts', pattern: anchors.ooxmlPattern }],
    clients: [{ path: 'scripts/playwright-browser-smoke.js', pattern: anchors.clientPattern }],
  };
}

function masterBackgroundSlideNumberDifferenceNote(id) {
  if (id === linePropertyId('SlideMasterProps', 'margin')) {
    return 'PptxGenJS stores implicit-inch master margins in its transient layout definition; native requires explicit EMU margins, exposes frozen live state, and does not claim the non-serialized input after reopen.';
  }
  if (id === linePropertyId('SlideMasterProps', 'background')) {
    return 'PptxGenJS accepts a loose BackgroundProps record; native uses a strict discriminated master background or explicit image-source contract with detached live state.';
  }
  if (id === linePropertyId('SlideMasterProps', 'slideNumber')) {
    return 'PptxGenJS retains permissive slide-number input and writes fixed owner identities; native snapshots strict editable slide-number state and allocates canonical unique owners.';
  }
  if (id === linePropertyId('SlideMasterProps', 'objects')
      || id.startsWith('inline:interface:SlideMasterProps@property:objects')) {
    return 'PptxGenJS uses loose untagged master object records with implicit-inch options; native uses strict discriminated objects, explicit sources and units, transactional preparation, and editable live models.';
  }
  if (id.includes('BackgroundProps')) {
    return 'PptxGenJS retains and mutates a permissive background record, omits legal no-fill intent, and accepts implicit image and color forms; native uses detached strict background kinds and transactional image loading.';
  }
  return 'PptxGenJS uses implicit-inch or percentage geometry, truthy width and height defaults, and degrades justify to left; native uses explicit semantic geometry, rejects zero dimensions, and preserves justify.';
}

function masterBackgroundSlideNumberAlias(id) {
  if (id === linePropertyId('BackgroundProps', 'alpha')) {
    return {
      canonical: linePropertyId('BackgroundProps', 'transparency'),
      note: 'PptxGenJS alpha is a working deprecated transparency alias and can duplicate the canonical alpha child; native exposes only canonical strict transparency.',
    };
  }
  if (id === linePropertyId('BackgroundProps', 'fill')) {
    return {
      canonical: linePropertyId('BackgroundProps', 'color'),
      note: 'PptxGenJS fill is a working deprecated color alias that mutates caller state; native exposes only canonical strict color state.',
    };
  }
  return {
    canonical: linePropertyId('SlideMasterProps', 'background'),
    note: 'PptxGenJS master bkgd and its string branch are working deprecated aliases for a solid background; native exposes only the canonical strict background property.',
  };
}

function masterBackgroundSlideNumberDefectNote(id) {
  if (id === masterInlineObjectId('chart')) {
    return 'PptxGenJS declares the master chart branch as IChartOpts, but runtime requires an undeclared internal type/data/opts record and throws for the declared shape.';
  }
  if (id === propertyUnionId('SlideMasterProps', 'bkgd', 'BackgroundProps')) {
    return 'Every legal BackgroundProps object supplied through the deprecated master bkgd union is inert; only the separate string branch produces output.';
  }
  if (id === linePropertyId('BackgroundProps', 'src')) {
    return 'PptxGenJS declares deprecated direct background src, but its writer ignores the field for every legal string and creates no image relationship.';
  }
  return 'PptxGenJS declares top-level slide-number transparency, but the owner writer ignores every legal value; native places effective transparency in strict nested text style.';
}

function masterBackgroundSlideNumberEntry(id, status) {
  const evidence = masterBackgroundSlideNumberEvidence(id);
  const entry = {
    id,
    status,
    native: masterBackgroundSlideNumberNative(id),
    evidence,
    serialization: status !== 'defect-excluded',
    client: status !== 'defect-excluded',
    note: status === 'supported'
      ? 'Native covers the same legal atomic value with strict detached state, packed and browser consumers, canonical OOXML, and editable reopen evidence.'
      : status === 'deliberate-difference'
        ? masterBackgroundSlideNumberDifferenceNote(id)
        : status === 'defect-excluded'
          ? masterBackgroundSlideNumberDefectNote(id)
          : undefined,
  };
  if (status !== 'supported') {
    entry.control = {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: MASTER_BACKGROUND_SLIDE_NUMBER_CONTROL_TITLE,
    };
  }
  if (status === 'deprecated-alias') {
    const alias = masterBackgroundSlideNumberAlias(id);
    entry.canonical = alias.canonical;
    entry.note = alias.note;
  }
  if (status === 'defect-excluded') {
    entry.evidence = {
      code: [],
      tests: evidence.tests.slice(0, 1),
      package: [],
      ooxml: [],
      clients: [],
    };
    delete entry.serialization;
    delete entry.client;
  }
  return entry;
}

const MASTER_BACKGROUND_SLIDE_NUMBER_FAMILY_ENTRIES = Object.freeze([
  ...MASTER_BACKGROUND_SLIDE_NUMBER_SUPPORTED_IDS.map((id) =>
    masterBackgroundSlideNumberEntry(id, 'supported')),
  ...MASTER_BACKGROUND_SLIDE_NUMBER_DIFFERENCE_IDS.map((id) =>
    masterBackgroundSlideNumberEntry(id, 'deliberate-difference')),
  ...MASTER_BACKGROUND_SLIDE_NUMBER_DEPRECATED_IDS.map((id) =>
    masterBackgroundSlideNumberEntry(id, 'deprecated-alias')),
  ...MASTER_BACKGROUND_SLIDE_NUMBER_DEFECT_IDS.map((id) =>
    masterBackgroundSlideNumberEntry(id, 'defect-excluded')),
]);

const SHAPE_TEXT_SHADOW_CONTROL_TITLE =
  'locks ShadowProps across shape, text, image, and chart owners against PptxGenJS 4.0.1';
const SHAPE_SHADOW_CONTROL_TITLE =
  'compares shape shadow public output and strict native divergences';
const TEXT_SHADOW_CONTROL_TITLE =
  'compares text shape shadow public output and strict native divergences';
const SHAPE_TEXT_SHADOW_IDS = Object.freeze([
  ...[
    'angle',
    'blur',
    'color',
    'offset',
    'opacity',
    'rotateWithShape',
    'type',
  ].map((property) => linePropertyId('ShadowProps', property)),
  ...['inner', 'none', 'outer']
    .map((token) => propertyUnionId('ShadowProps', 'type', token)),
  linePropertyId('ShapeProps', 'shadow'),
  linePropertyId('TextPropsOptions', 'shadow'),
]);

function shapeTextShadowControlTitle(id) {
  if (id === linePropertyId('ShapeProps', 'shadow')) return SHAPE_SHADOW_CONTROL_TITLE;
  if (id === linePropertyId('TextPropsOptions', 'shadow')) {
    return TEXT_SHADOW_CONTROL_TITLE;
  }
  return SHAPE_TEXT_SHADOW_CONTROL_TITLE;
}

function shapeTextShadowNative(id) {
  if (id === linePropertyId('ShapeProps', 'shadow')) {
    return ['AddShapeOptions.shadow', 'ShapeShadow', 'ShapeModel.shadow'];
  }
  if (id === linePropertyId('TextPropsOptions', 'shadow')) {
    return ['AddTextOptions.shadow', 'ShapeShadow', 'ShapeModel.shadow'];
  }
  if (id === propertyUnionId('ShadowProps', 'type', 'none')) {
    return ['ShapeModel.shadow'];
  }
  const property = id.split('@property:')[1]?.split('#')[0];
  const mappings = {
    angle: ['ShapeShadowBase.angle', 'ShapeModel.shadow'],
    blur: ['ShapeShadowBase.blur', 'ShapeModel.shadow'],
    color: ['ShapeShadowBase.color', 'ShapeModel.shadow'],
    offset: ['ShapeShadowBase.distance', 'ShapeModel.shadow'],
    opacity: ['ShapeShadowBase.opacity', 'ShapeModel.shadow'],
    rotateWithShape: ['ShapeShadow.rotateWithShape', 'ShapeModel.shadow'],
    type: ['ShapeShadow.kind', 'ShapeModel.shadow'],
  };
  return mappings[property];
}

function shapeTextShadowEvidence(id) {
  const shapeOwner = id === linePropertyId('ShapeProps', 'shadow');
  const textOwner = id === linePropertyId('TextPropsOptions', 'shadow');
  const shapeClient = {
    path: 'scripts/smoke-npm-package.mjs',
    pattern: 'const browserShadowChecks = {',
  };
  const textClient = {
    path: 'scripts/playwright-browser-smoke.js',
    pattern: 'const textShapeShadowState = {',
  };
  const clients = shapeOwner
    ? [shapeClient]
    : textOwner
      ? [textClient]
      : [shapeClient, textClient];
  const code = shapeOwner
    ? [{
        path: 'packages/model/src/preset-shape.ts',
        pattern: 'readonly shadow?: ShapeShadow;',
      }, {
        path: 'packages/model/src/shape-shadow.internal.ts',
        pattern: 'export function replaceShapeShadow(',
      }]
    : textOwner
      ? [{
          path: 'packages/model/src/slide.ts',
          pattern: 'readonly shadow?: ShapeShadow;',
        }, {
          path: 'packages/model/src/shape-shadow.internal.ts',
          pattern: 'export function replaceShapeShadow(',
        }]
      : [{
          path: 'packages/model/src/preset-shape.ts',
          pattern: 'export interface ShapeShadowBase {',
        }, {
          path: 'packages/model/src/preset-shape.ts',
          pattern: 'export type ShapeShadow =',
        }, {
          path: 'packages/model/src/shape-shadow.internal.ts',
          pattern: 'export function replaceShapeShadow(',
        }];
  const packageEvidence = [
    ...(textOwner ? [] : [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const shapeShadows =',
    }]),
    ...(shapeOwner ? [] : [{
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const textShapeShadows =',
    }]),
  ];
  const ooxml = [
    ...(shapeOwner ? [] : [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: 'creates text shadows across slide layout master and placeholder owners',
    }]),
    ...(textOwner ? [] : [{
      path: 'packages/sdk/src/index.test.ts',
      pattern: 'supports the shape shadow lifecycle across duplicate, move, rollback, reopen, and all formats',
    }]),
  ];
  return {
    code,
    tests: [{
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: SHAPE_TEXT_SHADOW_CONTROL_TITLE,
    }, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: SHAPE_SHADOW_CONTROL_TITLE,
    }, {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: TEXT_SHADOW_CONTROL_TITLE,
    }],
    package: packageEvidence,
    ooxml,
    clients,
  };
}

function shapeTextShadowNote(id) {
  if (id === linePropertyId('ShapeProps', 'shadow')) {
    return 'PptxGenJS shape shadows use permissive type/offset/hex input, collapse explicit zero to defaults, ignore rotate true, and produce malformed inner XML; native uses a strict detached ShapeShadow with legal outer and inner editable state.';
  }
  if (id === linePropertyId('TextPropsOptions', 'shadow')) {
    return 'PptxGenJS text shadows mutate or correct permissive input, collapse explicit zero to defaults, ignore rotate true, and produce malformed inner XML; native uses the same strict detached ShapeShadow across every text owner.';
  }
  return 'PptxGenJS shares ShadowProps across owners but varies coercion, defaults, zero handling, inner output, and rotate behavior between shape, text, image, and chart writers; native exposes one strict detached ShapeShadow contract with legal editable OOXML.';
}

function shapeTextShadowEntry(id) {
  const controlTitle = shapeTextShadowControlTitle(id);
  return {
    id,
    status: 'deliberate-difference',
    native: shapeTextShadowNative(id),
    evidence: shapeTextShadowEvidence(id),
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: controlTitle,
    },
    serialization: true,
    client: true,
    note: shapeTextShadowNote(id),
  };
}

const SHAPE_TEXT_SHADOW_FAMILY_ENTRIES = Object.freeze(
  SHAPE_TEXT_SHADOW_IDS.map((id) => shapeTextShadowEntry(id)),
);

function presetShapeCatalogEntry(owner, value) {
  const id = `union:${owner}#${value}`;
  if (value === 'folderCorner') {
    const title = 'isolates the folderCorner defect from valid preset shape public output';
    return {
      id,
      status: 'defect-excluded',
      native: [],
      evidence: {
        code: [],
        tests: [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', title }],
        package: [],
        ooxml: [],
        clients: [],
      },
      control: { path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: title },
      note: 'PptxGenJS 4.0.1 exposes folderCorner but writes an invalid preset token; native rejects it and exposes the valid OOXML foldedCorner token instead.',
    };
  }
  return serializedCatalogMember(
    id,
    ['PRESET_SHAPE_TYPES', 'PresetShapeType', 'SlideModel.addShape'],
    {
      path: 'packages/model/src/preset-shape.ts',
      pattern: 'export const PRESET_SHAPE_TYPES = Object.freeze([',
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: 'reads every legal PptxGenJS preset shape public output',
    },
    {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const typedPresetCatalog: readonly PresetShapeType[] = PRESET_SHAPE_TYPES;',
    },
    {
      path: 'packages/sdk/src/index.test.ts',
      pattern: 'creates all 178 canonical preset shapes in catalog order',
    },
    {
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const presetShapes = JSON.stringify(presetShapeState)',
    },
    'Native exposes, creates, serializes, and reopens the same legal canonical preset token.',
  );
}

function schemeColorCatalogEntry(owner, value) {
  return serializedCatalogMember(
    `union:${owner}#${value}`,
    ['SCHEME_COLORS', 'SchemeColor'],
    {
      path: 'packages/model/src/scheme-color.ts',
      pattern: 'export const SCHEME_COLORS = Object.freeze({',
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: 'matches the public PptxGenJS SchemeColor helper and legal output',
    },
    {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'const packedSchemeColorEntries = Object.entries(SCHEME_COLORS);',
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: 'matches the public PptxGenJS SchemeColor helper and legal output',
    },
    {
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const schemeColors = JSON.stringify(schemeColorState)',
    },
    'Native exposes the same scheme-color value through its frozen catalog and legal OOXML output.',
  );
}

function alignmentCatalogEntry(owner, value, vertical) {
  const title = vertical
    ? 'matches the PptxGenJS vertical alignment runtime catalog'
    : 'matches the PptxGenJS horizontal alignment runtime catalog';
  return serializedCatalogMember(
    `union:${owner}#${value}`,
    vertical
      ? ['TEXT_VERTICAL_ALIGNMENTS', 'TextBoxVerticalAlignment']
      : ['TEXT_ALIGNMENTS', 'TextAlignment'],
    {
      path: 'packages/model/src/text.ts',
      pattern: vertical
        ? 'export const TEXT_VERTICAL_ALIGNMENTS = Object.freeze(['
        : 'export const TEXT_ALIGNMENTS = Object.freeze([',
    },
    { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title },
    {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: vertical
        ? 'const packedVerticalAlignmentShapes = TEXT_VERTICAL_ALIGNMENTS.map('
        : 'const packedHorizontalAlignmentShapes = TEXT_ALIGNMENTS.map(',
    },
    {
      path: 'packages/sdk/src/index.test.ts',
      pattern: vertical
        ? 'publishes TEXT_VERTICAL_ALIGNMENTS through text and table lifecycles'
        : 'publishes TEXT_ALIGNMENTS through the SDK lifecycle',
    },
    {
      path: 'scripts/playwright-browser-smoke.js',
      pattern: vertical
        ? 'const verticalAlignments = JSON.stringify(verticalAlignmentState)'
        : 'const horizontalAlignments = JSON.stringify(horizontalAlignmentState)',
    },
    'Native exposes and serializes the same alignment value through its frozen catalog.',
  );
}

function outputTypeCatalogEntry(owner, value) {
  return serializedCatalogMember(
    `union:${owner}#${value}`,
    ['OUTPUT_TYPES', 'OutputType', 'PptxDocument.write'],
    {
      path: 'packages/sdk/src/output-type.ts',
      pattern: 'export const OUTPUT_TYPES = Object.freeze([',
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: 'matches the PptxGenJS output type runtime catalog and return kinds',
    },
    {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: 'values: [...OUTPUT_TYPES]',
    },
    {
      path: 'packages/sdk/src/write-output.test.ts',
      pattern: 'converts canonical bytes into all six Node output representations',
    },
    {
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'const writeOutputTypes = JSON.stringify(writeOutputTypeState)',
    },
    'Native exposes the same output selector and returns the matching representation.',
  );
}

function placeholderTypeCatalogEntry(owner, value) {
  const controlTitle =
    'matches public slide master objects, topology, and empty placeholder geometry';
  if (value === 'pic' || value === 'tbl') {
    return {
      id: `union:${owner}#${value}`,
      status: 'deliberate-difference',
      native: ['PLACEHOLDER_TYPES', 'PlaceholderType', 'SlideModel.addPlaceholder'],
      evidence: {
        code: [{
          path: 'packages/model/src/placeholder.ts',
          pattern: 'export const PLACEHOLDER_TYPES = [',
        }],
        tests: [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: controlTitle }],
        package: [{
          path: 'scripts/smoke-npm-package.mjs',
          pattern: "PLACEHOLDER_TYPES.join(',') === 'title,body,pic,chart,tbl,media'",
        }],
        ooxml: [{ path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: controlTitle }],
        clients: [{
          path: 'scripts/playwright-browser-smoke.js',
          pattern: 'placeholderTypes: [...api.PLACEHOLDER_TYPES]',
        }],
      },
      control: { path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: controlTitle },
      serialization: true,
      client: true,
      note: `PptxGenJS 4.0.1 normalizes an empty ${value} layout placeholder to body; native preserves the declared ${value} type and uses the correct owner geometry.`,
    };
  }
  return serializedCatalogMember(
    `union:${owner}#${value}`,
    ['PLACEHOLDER_TYPES', 'PlaceholderType', 'SlideModel.addPlaceholder'],
    {
      path: 'packages/model/src/placeholder.ts',
      pattern: 'export const PLACEHOLDER_TYPES = [',
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      title: controlTitle,
    },
    {
      path: 'scripts/smoke-npm-package.mjs',
      pattern: "PLACEHOLDER_TYPES.join(',') === 'title,body,pic,chart,tbl,media'",
    },
    {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: controlTitle,
    },
    {
      path: 'scripts/playwright-browser-smoke.js',
      pattern: 'placeholderTypes: [...api.PLACEHOLDER_TYPES]',
    },
    'Native exposes, creates, serializes, and reopens the same placeholder type.',
  );
}

function imageCoordinateDifference(id) {
  return {
    id,
    status: 'deliberate-difference',
    native: ['SlideModel.addImage', 'SlideModel.addSvgImage', 'PptxDocument.addImage'],
    evidence: {
      code: [{
        path: 'packages/model/src/image-create.internal.ts',
        pattern: 'function resolveImageCoordinate(',
      }],
      tests: [{
        path: 'packages/pptxgenjs-adapter/src/index.test.ts',
        title: IMAGE_COORDINATE_CONTROL_TITLE,
      }],
      package: [{
        path: 'scripts/smoke-npm-package.mjs',
        pattern: 'const imagePercentageCoordinates =',
      }],
      ooxml: [{
        path: 'packages/sdk/src/index.test.ts',
        pattern: IMAGE_COORDINATE_OOXML_TITLE,
      }],
      clients: [],
    },
    control: {
      path: 'packages/pptxgenjs-adapter/src/index.test.ts',
      pattern: IMAGE_COORDINATE_CONTROL_TITLE,
    },
    serialization: true,
    note: 'Native covers the same legal direct image percentage geometry with width/height and explicit Emu or inches() numeric units; nested sizing coordinates remain a separate capability family.',
  };
}

function coordinateNativeMapping(id) {
  if (id.startsWith('interface:PositionProps@')) {
    return ['SlideCoordinate', 'TransformInput'];
  }
  if (id.startsWith('interface:ShapeProps@')) {
    return ['SlideModel.addShape', 'SlideModel.addCustomShape'];
  }
  return ['SlideModel.addText'];
}

function deepFreeze(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return value;
  seen.add(value);
  for (const nested of Object.values(value)) deepFreeze(nested, seen);
  return Object.freeze(value);
}

export const PPTXGENJS_SURFACE_MANIFEST = deepFreeze({
  schemaVersion: 1,
  packageVersion: '4.0.1',
  entries: [
    supported(
      'class:PptxGenJS@property:version',
      ['PPTX_VERSION', 'PptxDocument.version'],
      { path: 'packages/sdk/src/version.ts', pattern: "export const PPTX_VERSION = '0.1.0' as const;" },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'reports each library runtime version through its public instance' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: 'const version: PptxVersion = PPTX_VERSION;' },
      'Native exposes one immutable library version through constants, instances, declarations, and packed consumers.',
    ),
    supported(
      'class:PptxGenJS@property:presLayout',
      ['PptxDocument.presLayout'],
      { path: 'packages/sdk/src/index.ts', pattern: 'get presLayout(): PresentationLayout {' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'matches the public presentation layout projection and locks the custom-name boundary' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: 'const packedDefaultLayout = PptxDocument.create().presLayout;' },
      'Native projects detached standard and custom presentation dimensions through a getter.',
    ),
    supported(
      'class:PptxGenJS@property:AlignH',
      ['TEXT_ALIGNMENTS'],
      { path: 'packages/model/src/text.ts', pattern: 'export const TEXT_ALIGNMENTS = Object.freeze([' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'matches the PptxGenJS horizontal alignment runtime catalog' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: "throw new Error('Browser horizontal alignment catalog failed');" },
      'Native publishes the same four frozen horizontal alignment values.',
    ),
    supported(
      'class:PptxGenJS@property:AlignV',
      ['TEXT_VERTICAL_ALIGNMENTS'],
      { path: 'packages/model/src/text.ts', pattern: 'export const TEXT_VERTICAL_ALIGNMENTS = Object.freeze([' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'matches the PptxGenJS vertical alignment runtime catalog' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: "throw new Error('Browser vertical alignment catalog failed');" },
      'Native publishes the same three frozen vertical alignment values.',
    ),
    supported(
      'class:PptxGenJS@property:OutputType',
      ['OUTPUT_TYPES'],
      { path: 'packages/sdk/src/output-type.ts', pattern: 'export const OUTPUT_TYPES = Object.freeze([' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'matches the PptxGenJS output type runtime catalog and return kinds' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: 'const packedOutputTypeDocument = PptxDocument.create();' },
      'Native publishes all six catalog values and returns the matching Node/browser representations.',
    ),
    supported(
      'class:PptxGenJS@property:SchemeColor',
      ['SCHEME_COLORS'],
      { path: 'packages/model/src/scheme-color.ts', pattern: 'export const SCHEME_COLORS = Object.freeze({' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'matches the public PptxGenJS SchemeColor helper and legal output' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: 'const packedSchemeColorEntries = Object.entries(SCHEME_COLORS);' },
      'Native exposes the same ten key/value mappings as an immutable shared catalog.',
    ),
    supported(
      'class:PptxGenJS@property:ShapeType',
      ['PRESET_SHAPE_TYPES'],
      { path: 'packages/model/src/preset-shape.ts', pattern: 'export const PRESET_SHAPE_TYPES = Object.freeze([' },
      { path: 'packages/pptxgenjs-adapter/src/index.test.ts', title: 'reads every legal PptxGenJS preset shape public output' },
      { path: 'scripts/smoke-npm-package.mjs', pattern: 'const typedPresetCatalog: readonly PresetShapeType[] = PRESET_SHAPE_TYPES;' },
      'Native exposes the 177 legal declared preset values plus valid foldedCorner; invalid folderCorner and runtime-only custGeom are tracked separately.',
    ),
    {
      id: 'class:PptxGenJS@property:PlaceholderType',
      status: 'defect-excluded',
      native: [],
      evidence: {
        code: [],
        tests: [{ path: 'scripts/pptxgenjs-runtime-probe.test.mjs', title: 'probes the locked public runtime and declared catalogs deterministically' }],
        package: [],
        ooxml: [],
        clients: [],
      },
      control: { path: 'scripts/pptxgenjs-runtime-probe.test.mjs', pattern: 'assert.equal(first.catalogs.PlaceholderType, null);' },
      note: 'PptxGenJS 4.0.1 declares PlaceholderType on the instance but the real runtime property is absent.',
    },
    ...SHAPE_TEXT_COORDINATE_ATOMS.map((id) =>
      deliberateDifference(id, coordinateNativeMapping(id))),
    ...IMAGE_COORDINATE_ATOMS.map((id) => imageCoordinateDifference(id)),
    ...CANONICAL_LINE_ATOM_IDS.map((id) => canonicalLineEntry(id)),
    ...DEPRECATED_LINE_ENTRIES,
    ...CANONICAL_FILL_ATOM_IDS.map((id) => canonicalFillEntry(id)),
    DEPRECATED_FILL_ENTRY,
    ...['TableCellProps', 'TableProps'].map((owner) => tableFillEntry(owner)),
    TABLE_TO_SLIDES_FILL_DEFECT_ENTRY,
    ...TABLE_TO_SLIDES_FAMILY_ENTRIES,
    ...ADD_TABLE_CORE_FAMILY_ENTRIES,
    ...PRESENTATION_ROOT_OUTPUT_FAMILY_ENTRIES,
    ...SLIDE_SECTION_FAMILY_ENTRIES,
    ...MASTER_BACKGROUND_SLIDE_NUMBER_FAMILY_ENTRIES,
    ...SHAPE_TEXT_SHADOW_FAMILY_ENTRIES,
    ...CHART_AREA_FILL_LINE_ENTRIES,
    ...DEPRECATED_CHART_AREA_ALIAS_ENTRIES,
    ...CHART_CREATION_SUPPORTED_ENTRIES,
    ...CHART_CREATION_DIFFERENCE_ENTRIES,
    ...CHART_AXIS_FOUNDATION_SUPPORTED_ENTRIES,
    ...CHART_AXIS_TICK_SUPPORTED_ENTRIES,
    ...CHART_AXIS_LINE_GRID_DIFFERENCE_ENTRIES,
    ...CHART_AXIS_BEHAVIOR_DIFFERENCE_ENTRIES,
    ...INERT_CHART_OPTION_DEFECT_ENTRIES,
    ...BULLET_FAMILY_ENTRIES,
    ...TAB_STOPS_FAMILY_ENTRIES,
    ...UNDERLINE_FAMILY_ENTRIES,
    ...TEXT_DIRECTION_FAMILY_ENTRIES,
    ...TEXT_RUN_SCALAR_FAMILY_ENTRIES,
    ...['ShapeType', 'SHAPE_NAME'].flatMap((owner) =>
      DECLARED_PRESET_SHAPE_VALUES.map((value) => presetShapeCatalogEntry(owner, value))),
    ...['SchemeColor', 'ThemeColor'].flatMap((owner) =>
      SCHEME_COLOR_VALUES.map((value) => schemeColorCatalogEntry(owner, value))),
    ...['AlignH', 'HAlign'].flatMap((owner) =>
      HORIZONTAL_ALIGNMENT_VALUES.map((value) => alignmentCatalogEntry(owner, value, false))),
    ...['AlignV', 'VAlign'].flatMap((owner) =>
      VERTICAL_ALIGNMENT_VALUES.map((value) => alignmentCatalogEntry(owner, value, true))),
    ...['OutputType', 'JSZIP_OUTPUT_TYPE'].flatMap((owner) =>
      OUTPUT_TYPE_VALUES.map((value) => outputTypeCatalogEntry(owner, value))),
    ...['PLACEHOLDER_TYPE', 'PLACEHOLDER_TYPES'].flatMap((owner) =>
      PLACEHOLDER_TYPE_VALUES.map((value) => placeholderTypeCatalogEntry(owner, value))),
  ],
  extensions: [],
});
