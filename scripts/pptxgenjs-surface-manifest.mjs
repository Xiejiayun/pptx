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
      clients: [],
    },
    serialization: true,
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
        clients: [],
      },
      control: { path: 'packages/pptxgenjs-adapter/src/index.test.ts', pattern: controlTitle },
      serialization: true,
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
