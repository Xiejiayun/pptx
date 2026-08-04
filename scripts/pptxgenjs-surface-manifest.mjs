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
      'Native exposes and creates every one of the 178 declared canonical preset shape values.',
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
  ],
  extensions: [],
});
