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
    ...CHART_AREA_FILL_LINE_ENTRIES,
    ...DEPRECATED_CHART_AREA_ALIAS_ENTRIES,
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
