import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';
import {
  degrees,
  inches,
  PRESET_SHAPE_TYPES,
  PptxDocument,
  ShapeModel,
  TableModel,
  type AddShapeOptions,
  type CustomGeometry,
  type PresetShapeType,
} from '@pptx/sdk';
import { importPptxGenJS } from './index.js';

interface BorderProps {
  readonly type?: 'none' | 'dash' | 'solid';
  readonly color?: string;
  readonly pt?: number;
}

interface PptxGenJSSlide {
  hidden: unknown;
  addNotes(notes: string): PptxGenJSSlide;
  addShape(type: string, options?: PptxGenJSShapeOptions): void;
  addText(
    text: string | readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
    options: Record<string, unknown>,
  ): void;
  addTable(
    rows: readonly (readonly {
      readonly text?: string;
      readonly options?: Record<string, unknown>;
    }[])[],
    options: Record<string, unknown>,
  ): void;
}

type PptxGenJSCoord = number | string;

type PptxGenJSCustomPoint =
  | { readonly x: PptxGenJSCoord; readonly y: PptxGenJSCoord; readonly moveTo?: boolean }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'arc';
        readonly hR: PptxGenJSCoord;
        readonly wR: PptxGenJSCoord;
        readonly stAng: number;
        readonly swAng: number;
      };
    }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'cubic';
        readonly x1: PptxGenJSCoord;
        readonly y1: PptxGenJSCoord;
        readonly x2: PptxGenJSCoord;
        readonly y2: PptxGenJSCoord;
      };
    }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'quadratic';
        readonly x1: PptxGenJSCoord;
        readonly y1: PptxGenJSCoord;
      };
    }
  | { readonly close: true };

interface PptxGenJSShapeOptions extends Record<string, unknown> {
  readonly points?: readonly PptxGenJSCustomPoint[];
}

type PptxGenJSPublicInstance = InstanceType<typeof import('pptxgenjs').default>;
type PptxGenJSPublicShapeOptions = NonNullable<
  Parameters<ReturnType<PptxGenJSPublicInstance['addSlide']>['addShape']>[1]
>;

const publicCustomShapeOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  points: [{ x: 0, y: 0 }],
};
const unsupportedPublicHandleOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no arbitrary adjustment-handle input.
  handles: [],
};
const unsupportedPublicConnectionSiteOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no arbitrary connection-site input.
  connectionSites: [],
};
void [
  publicCustomShapeOptions,
  unsupportedPublicHandleOptions,
  unsupportedPublicConnectionSiteOptions,
];

interface PptxGenJSInstance {
  readonly version: string;
  readonly ShapeType: Readonly<Record<string, string>>;
  readonly SchemeColor: {
    readonly accent1: 'accent1';
    readonly accent2: 'accent2';
  };
  author: string;
  company: string;
  layout: string;
  revision: string;
  rtlMode: unknown;
  subject: string;
  theme: {
    readonly headFontFace?: string;
    readonly bodyFontFace?: string;
  } | undefined;
  title: string;
  addSection(options: { readonly title: string; readonly order?: number }): void;
  addSlide(options?: { readonly sectionTitle?: string }): PptxGenJSSlide;
  write(options: { outputType: 'nodebuffer'; compression: boolean }): Promise<Uint8Array>;
  write(options: { outputType: 'uint8array'; compression: boolean }): Promise<Uint8Array>;
}

const require = createRequire(import.meta.url);
const PptxGenJS = require('pptxgenjs') as new () => PptxGenJSInstance;
const sectionState = (document: PptxDocument) =>
  document.sections?.map(({ title, slideIds }) => ({ title, slideIds }));

async function openPptxGenJSPublicOutput(
  presentation: PptxGenJSInstance,
): Promise<PptxDocument> {
  return PptxDocument.open(await presentation.write({
    outputType: 'nodebuffer',
    compression: true,
  }));
}

function shapeXml(document: PptxDocument, slideIndex: number, id: number): string {
  const slide = document.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex} was not found`);
  const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
  const shapes = [...xml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0]);
  const shape = shapes.find((candidate) => new RegExp(
    `<p:cNvPr\\b[^>]*\\bid="${id}"(?:\\s|/|>)`,
  ).test(candidate));
  if (!shape) throw new Error(`Shape ${id} was not found on slide ${slideIndex}`);
  return shape;
}

function directShapePaintState(xml: string): { fill: 'none'; line: 'empty' } {
  const properties = xml.match(/<p:spPr(?:\s[^>]*)?>([\s\S]*?)<\/p:spPr>/)?.[1];
  if (!properties) throw new Error('Shape properties were not found');
  const afterGeometry = properties.match(/<\/a:prstGeom>([\s\S]*)$/)?.[1];
  if (afterGeometry === undefined) throw new Error('Preset geometry was not found');
  if (!/^<a:noFill\/><a:ln(?:\/>|><\/a:ln>)/.test(afterGeometry)) {
    throw new Error('Expected direct no-fill and empty line state');
  }
  return { fill: 'none', line: 'empty' };
}

describe('importPptxGenJS', () => {
  it('matches representative preset shape public output semantically', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const cases: readonly {
      readonly type: PresetShapeType;
      readonly generatedOptions?: Record<string, unknown>;
      readonly nativeOptions?: AddShapeOptions;
      readonly expectedGeneratedName: string;
    }[] = [
      { type: 'rect', expectedGeneratedName: 'Shape 0' },
      { type: 'ellipse', generatedOptions: {}, nativeOptions: {}, expectedGeneratedName: 'Shape 0' },
      { type: 'line', expectedGeneratedName: 'Shape 0' },
      { type: 'lineInv', expectedGeneratedName: 'Shape 0' },
      { type: 'flowChartDecision', expectedGeneratedName: 'Shape 0' },
      { type: 'star5', expectedGeneratedName: 'Shape 0' },
      { type: 'actionButtonHome', expectedGeneratedName: 'Shape 0' },
      {
        type: 'roundRect',
        generatedOptions: {
          objectName: 'Public shape',
          x: 1.25,
          y: 2.5,
          w: 3.75,
          h: 4.5,
          rotate: 45,
          flipH: true,
          flipV: true,
        },
        nativeOptions: {
          name: 'Public shape',
          x: inches(1.25),
          y: inches(2.5),
          width: inches(3.75),
          height: inches(4.5),
          rotation: degrees(45),
          flipHorizontal: true,
          flipVertical: true,
        },
        expectedGeneratedName: 'Public shape',
      },
    ];

    for (const fixture of cases) {
      const publicType = generated.ShapeType[fixture.type];
      expect(publicType, fixture.type).toBe(fixture.type);
      const slide = generated.addSlide();
      if (Object.prototype.hasOwnProperty.call(fixture, 'generatedOptions')) {
        slide.addShape(publicType!, fixture.generatedOptions);
      } else {
        slide.addShape(publicType!);
      }
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const native = PptxDocument.create();
    for (const [index, fixture] of cases.entries()) {
      const importedShape = imported.slides[index]?.shapes[0];
      const nativeShape = native.addSlide().addShape(fixture.type, fixture.nativeOptions);
      expect(importedShape, fixture.type).toBeInstanceOf(ShapeModel);
      expect(importedShape?.name, fixture.type).toBe(fixture.expectedGeneratedName);
      expect((importedShape as ShapeModel).presetType, fixture.type).toBe(fixture.type);
      expect(importedShape?.transform, fixture.type).toEqual(nativeShape.transform);
      expect(directShapePaintState(shapeXml(imported, index, importedShape!.id)))
        .toEqual(directShapePaintState(shapeXml(native, index, nativeShape.id)));
    }
  });

  it('compares shape fill public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const generatedCases: readonly {
        readonly name: string;
        readonly fill?: Record<string, unknown>;
      }[] = [
        { name: 'Fill omitted' },
        { name: 'Fill none', fill: { type: 'none' } },
        { name: 'Fill sRGB', fill: { color: 'FF0000' } },
        { name: 'Fill scheme', fill: { color: generated.SchemeColor.accent2 } },
        { name: 'Fill transparency', fill: { color: '00FF00', transparency: 50 } },
        { name: 'Fill zero', fill: { color: '0000FF', transparency: 0 } },
        { name: 'Fill deprecated alpha', fill: { color: '112233', alpha: 40 } },
        { name: 'Fill empty', fill: {} },
        { name: 'Fill missing color', fill: { type: 'solid' } },
      ];
      for (const fixture of generatedCases) {
        const options: Record<string, unknown> = { objectName: fixture.name };
        if (fixture.fill !== undefined) options.fill = fixture.fill;
        generatedSlide.addShape(generated.ShapeType.rect!, options);
      }

      const imported = await openPptxGenJSPublicOutput(generated);
      const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
        shape.name,
        shape as ShapeModel,
      ]));
      expect(importedShapes.size).toBe(generatedCases.length);

      const native = PptxDocument.create();
      const nativeSlide = native.addSlide();
      const nativeShapes = new Map([
        ['Fill omitted', nativeSlide.addShape('rect', { name: 'Fill omitted' })],
        ['Fill none', nativeSlide.addShape('rect', {
          name: 'Fill none',
          fill: { kind: 'none' },
        })],
        ['Fill sRGB', nativeSlide.addShape('rect', {
          name: 'Fill sRGB',
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
        })],
        ['Fill scheme', nativeSlide.addShape('rect', {
          name: 'Fill scheme',
          fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
        })],
        ['Fill transparency', nativeSlide.addShape('rect', {
          name: 'Fill transparency',
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: '00FF00' },
            transparency: 50,
          },
        })],
        ['Fill zero', nativeSlide.addShape('rect', {
          name: 'Fill zero',
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: '0000FF' },
            transparency: 0,
          },
        })],
      ] as const);

      for (const name of [
        'Fill omitted',
        'Fill sRGB',
        'Fill scheme',
        'Fill transparency',
      ] as const) {
        const importedShape = importedShapes.get(name)!;
        const nativeShape = nativeShapes.get(name)!;
        expect(importedShape).toBeInstanceOf(ShapeModel);
        expect(importedShape.name).toBe(nativeShape.name);
        expect(importedShape.presetType).toBe(nativeShape.presetType);
        expect(importedShape.transform).toEqual(nativeShape.transform);
        expect(importedShape.fill).toEqual(nativeShape.fill);
      }

      const generatedNone = importedShapes.get('Fill none')!;
      const nativeNone = nativeShapes.get('Fill none')!;
      expect(generatedNone.fill).toBeUndefined();
      expect(nativeNone.fill).toEqual({ kind: 'none' });
      expect(shapeXml(imported, 0, generatedNone.id)).toMatch(
        /<\/a:prstGeom><a:ln(?:\/>|>)/,
      );
      expect(shapeXml(native, 0, nativeNone.id)).toContain(
        '</a:prstGeom><a:noFill/><a:ln/>',
      );

      const generatedZero = importedShapes.get('Fill zero')!;
      const nativeZero = nativeShapes.get('Fill zero')!;
      expect(generatedZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
      });
      expect(nativeZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
        transparency: 0,
      });
      expect(shapeXml(imported, 0, generatedZero.id)).not.toContain('<a:alpha');
      expect(shapeXml(native, 0, nativeZero.id)).toContain('<a:alpha val="100000"/>');
      expect(100 - (generatedZero.fill?.kind === 'solid'
        ? generatedZero.fill.transparency ?? 0
        : 100)).toBe(100);
      expect(100 - (nativeZero.fill?.kind === 'solid'
        ? nativeZero.fill.transparency ?? 0
        : 100)).toBe(100);

      expect(importedShapes.get('Fill empty')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '000000' },
      });
      expect(importedShapes.get('Fill missing color')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '000000' },
      });
      expect(importedShapes.get('Fill deprecated alpha')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 40,
      });
      expect(warning).toHaveBeenCalled();

      const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
      const invalidJournal = [...native.opcPackage.mutations];
      for (const fill of [
        {},
        { kind: 'solid' },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          alpha: 40,
        },
      ]) {
        expect(() => nativeSlide.addShape('rect', { fill } as never)).toThrow();
      }
      expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
      expect(native.opcPackage.mutations).toEqual(invalidJournal);

      const editable = importedShapes.get('Fill sRGB')!;
      editable.fill = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 25,
      };
      const reopened = await PptxDocument.open(await imported.write());
      const reopenedEditable = reopened.slides[0]!.shapes.find(
        ({ name }) => name === 'Fill sRGB',
      ) as ShapeModel;
      expect(reopenedEditable.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 25,
      });
    } finally {
      warning.mockRestore();
    }
  });

  it('compares shape line public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const dashes = [
      'solid',
      'dash',
      'dashDot',
      'lgDash',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
    ] as const;
    const generatedCases: readonly {
      readonly name: string;
      readonly line?: Record<string, unknown>;
    }[] = [
      { name: 'Line omitted' },
      { name: 'Line none', line: { type: 'none' } },
      { name: 'Line empty', line: {} },
      { name: 'Line missing color', line: { type: 'solid' } },
      { name: 'Line sRGB', line: { color: 'FF0000' } },
      { name: 'Line scheme', line: { color: generated.SchemeColor.accent2 } },
      { name: 'Line transparency', line: { color: '00FF00', transparency: 50 } },
      { name: 'Line zero transparency', line: { color: '0000FF', transparency: 0 } },
      { name: 'Line zero width', line: { color: '112233', width: 0 } },
      { name: 'Line positive width', line: { color: '223344', width: 2.5 } },
      { name: 'Line deprecated alpha', line: { color: '334455', alpha: 40 } },
      { name: 'Line deprecated dash', line: { color: '445566', lineDash: 'dash' } },
      ...dashes.map((dash) => ({
        name: `Line dash ${dash}`,
        line: { color: '556677', dashType: dash },
      })),
      {
        name: 'Line arrows',
        line: {
          color: '667788',
          width: 3,
          dashType: 'dashDot',
          beginArrowType: 'triangle',
          endArrowType: 'arrow',
        },
      },
    ];
    for (const fixture of generatedCases) {
      const options: Record<string, unknown> = { objectName: fixture.name };
      if (fixture.line !== undefined) options.line = fixture.line;
      generatedSlide.addShape(generated.ShapeType.rect!, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(generatedCases.length);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>([
      ['Line omitted', nativeSlide.addShape('rect', { name: 'Line omitted' })],
      ['Line none', nativeSlide.addShape('rect', {
        name: 'Line none',
        line: { kind: 'none' },
      })],
      ['Line sRGB', nativeSlide.addShape('rect', {
        name: 'Line sRGB',
        line: { kind: 'line', color: { kind: 'srgb', value: 'FF0000' } },
      })],
      ['Line scheme', nativeSlide.addShape('rect', {
        name: 'Line scheme',
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
      })],
      ['Line transparency', nativeSlide.addShape('rect', {
        name: 'Line transparency',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          transparency: 50,
        },
      })],
      ['Line zero transparency', nativeSlide.addShape('rect', {
        name: 'Line zero transparency',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '0000FF' },
          transparency: 0,
        },
      })],
      ['Line zero width', nativeSlide.addShape('rect', {
        name: 'Line zero width',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          width: 0,
        },
      })],
      ['Line positive width', nativeSlide.addShape('rect', {
        name: 'Line positive width',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '223344' },
          width: 2.5,
        },
      })],
      ...dashes.map((dash) => [
        `Line dash ${dash}`,
        nativeSlide.addShape('rect', {
          name: `Line dash ${dash}`,
          line: {
            kind: 'line',
            color: { kind: 'srgb', value: '556677' },
            dash,
          },
        }),
      ] as const),
      ['Line arrows', nativeSlide.addShape('rect', {
        name: 'Line arrows',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '667788' },
          width: 3,
          dash: 'dashDot',
        },
      })],
    ] as const);

    for (const name of [
      'Line omitted',
      'Line sRGB',
      'Line scheme',
      'Line transparency',
      'Line positive width',
      ...dashes.map((dash) => `Line dash ${dash}`),
      'Line arrows',
    ]) {
      const importedShape = importedShapes.get(name)!;
      const nativeShape = nativeShapes.get(name)!;
      expect(importedShape, name).toBeInstanceOf(ShapeModel);
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.line, name).toEqual(nativeShape.line);
    }

    const generatedNone = importedShapes.get('Line none')!;
    const nativeNone = nativeShapes.get('Line none')!;
    expect(generatedNone.line).toBeUndefined();
    expect(nativeNone.line).toEqual({ kind: 'none' });
    expect(shapeXml(imported, 0, generatedNone.id)).toMatch(/<a:ln><\/a:ln>|<a:ln\/>/);
    expect(shapeXml(native, 0, nativeNone.id)).toContain('<a:ln><a:noFill/></a:ln>');

    for (const name of ['Line empty', 'Line missing color']) {
      expect(importedShapes.get(name)!.line, name).toEqual({
        kind: 'line',
        color: { kind: 'srgb', value: '333333' },
        width: 1,
        dash: 'solid',
      });
    }

    const generatedZeroTransparency = importedShapes.get('Line zero transparency')!;
    const nativeZeroTransparency = nativeShapes.get('Line zero transparency')!;
    expect(generatedZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 0,
      width: 1,
      dash: 'solid',
    });
    expect(shapeXml(imported, 0, generatedZeroTransparency.id)).not.toContain('<a:alpha');
    expect(shapeXml(native, 0, nativeZeroTransparency.id))
      .toContain('<a:alpha val="100000"/>');

    const generatedZeroWidth = importedShapes.get('Line zero width')!;
    const nativeZeroWidth = nativeShapes.get('Line zero width')!;
    expect(generatedZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 0,
      dash: 'solid',
    });
    expect(shapeXml(imported, 0, generatedZeroWidth.id)).toContain('<a:ln w="12700">');
    expect(shapeXml(native, 0, nativeZeroWidth.id)).toContain('<a:ln w="0">');

    expect(importedShapes.get('Line deprecated alpha')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '334455' },
      width: 1,
      dash: 'solid',
    });
    expect(importedShapes.get('Line deprecated dash')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '445566' },
      width: 1,
      dash: 'solid',
    });

    const arrows = importedShapes.get('Line arrows')!;
    expect(shapeXml(imported, 0, arrows.id)).toContain('<a:headEnd type="triangle"');
    expect(shapeXml(imported, 0, arrows.id)).toContain('<a:tailEnd type="arrow"');
    arrows.line = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 25,
      width: 2,
      dash: 'sysDash',
    };
    const editedArrowXml = shapeXml(imported, 0, arrows.id);
    expect(editedArrowXml).toContain('<a:headEnd type="triangle"');
    expect(editedArrowXml).toContain('<a:tailEnd type="arrow"');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const line of [
      {},
      { kind: 'line' },
      { type: 'none' },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '334455' },
        alpha: 40,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        dashType: 'dash',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        lineDash: 'dash',
      },
    ]) {
      expect(() => nativeSlide.addShape('rect', { line } as never)).toThrow();
    }
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedArrows = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Line arrows',
    ) as ShapeModel;
    expect(reopenedArrows.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 25,
      width: 2,
      dash: 'sysDash',
    });
    expect(shapeXml(reopened, 0, reopenedArrows.id)).toContain('<a:headEnd type="triangle"');
    expect(shapeXml(reopened, 0, reopenedArrows.id)).toContain('<a:tailEnd type="arrow"');
  });

  it('compares shape arrow public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const arrowTypes = [
      'none',
      'arrow',
      'diamond',
      'oval',
      'stealth',
      'triangle',
    ] as const;
    const generatedCases: Record<string, Record<string, unknown>> = {};
    for (const type of arrowTypes) {
      generatedCases[`Arrow begin ${type}`] = {
        objectName: `Arrow begin ${type}`,
        line: { color: '112233', beginArrowType: type },
      };
      generatedCases[`Arrow end ${type}`] = {
        objectName: `Arrow end ${type}`,
        line: { color: '112233', endArrowType: type },
      };
    }
    Object.assign(generatedCases, {
      'Arrow both': {
        objectName: 'Arrow both',
        line: {
          color: '112233',
          beginArrowType: 'triangle',
          endArrowType: 'arrow',
        },
      },
      'Arrow only defaults': {
        objectName: 'Arrow only defaults',
        line: { beginArrowType: 'diamond' },
      },
      'Arrow none line': {
        objectName: 'Arrow none line',
        line: { type: 'none', beginArrowType: 'triangle' },
      },
      'Arrow empty ignored': {
        objectName: 'Arrow empty ignored',
        line: { color: '112233', beginArrowType: '', endArrowType: '' },
      },
      'Arrow nested aliases ignored': {
        objectName: 'Arrow nested aliases ignored',
        line: { color: '112233', lineHead: 'triangle', lineTail: 'arrow' },
      },
      'Arrow invalid passthrough': {
        objectName: 'Arrow invalid passthrough',
        line: { color: '112233', beginArrowType: 'bogus' },
      },
      'Arrow top aliases mapped': {
        objectName: 'Arrow top aliases mapped',
        line: { color: '112233' },
        lineHead: 'stealth',
        lineTail: 'oval',
      },
    });
    for (const options of Object.values(generatedCases)) {
      generatedSlide.addShape(generated.ShapeType.line!, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(Object.keys(generatedCases).length);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>();
    for (const type of arrowTypes) {
      nativeShapes.set(`Arrow begin ${type}`, nativeSlide.addShape('line', {
        name: `Arrow begin ${type}`,
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
        },
        arrows: { begin: type },
      }));
      nativeShapes.set(`Arrow end ${type}`, nativeSlide.addShape('line', {
        name: `Arrow end ${type}`,
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
        },
        arrows: { end: type },
      }));
    }
    nativeShapes.set('Arrow both', nativeSlide.addShape('line', {
      name: 'Arrow both',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
      },
      arrows: { begin: 'triangle', end: 'arrow' },
    }));
    nativeShapes.set('Arrow only defaults', nativeSlide.addShape('line', {
      name: 'Arrow only defaults',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '333333' },
      },
      arrows: { begin: 'diamond' },
    }));
    nativeShapes.set('Arrow none line', nativeSlide.addShape('line', {
      name: 'Arrow none line',
      arrows: { begin: 'triangle' },
    }));
    nativeShapes.set('Arrow top aliases mapped', nativeSlide.addShape('line', {
      name: 'Arrow top aliases mapped',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
      },
      arrows: { begin: 'stealth', end: 'oval' },
    }));

    for (const [name, nativeShape] of nativeShapes) {
      const importedShape = importedShapes.get(name)!;
      expect(importedShape, name).toBeInstanceOf(ShapeModel);
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.line, name).toEqual(nativeShape.line);
      expect(importedShape.arrows, name).toEqual(nativeShape.arrows);
      const xml = shapeXml(imported, 0, importedShape.id);
      if (importedShape.arrows?.begin !== undefined) {
        expect(xml, name).toContain(
          `<a:headEnd type="${importedShape.arrows.begin}"/>`,
        );
      }
      if (importedShape.arrows?.end !== undefined) {
        expect(xml, name).toContain(
          `<a:tailEnd type="${importedShape.arrows.end}"/>`,
        );
      }
      if (
        importedShape.arrows?.begin !== undefined
        && importedShape.arrows.end !== undefined
      ) {
        expect(xml.indexOf('<a:headEnd'), name).toBeLessThan(xml.indexOf('<a:tailEnd'));
      }
    }

    const generatedArrowOnly = importedShapes.get('Arrow only defaults')!;
    const nativeArrowOnly = nativeShapes.get('Arrow only defaults')!;
    expect(generatedArrowOnly.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '333333' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeArrowOnly.line).toEqual(generatedArrowOnly.line);
    expect(shapeXml(imported, 0, generatedArrowOnly.id)).toContain('<a:ln w="12700">');

    const generatedNoneLine = importedShapes.get('Arrow none line')!;
    const nativeNoneLine = nativeShapes.get('Arrow none line')!;
    expect(generatedNoneLine.line).toBeUndefined();
    expect(nativeNoneLine.line).toBeUndefined();
    expect(generatedNoneLine.arrows).toEqual({ begin: 'triangle' });
    expect(nativeNoneLine.arrows).toEqual({ begin: 'triangle' });
    expect(shapeXml(imported, 0, generatedNoneLine.id))
      .toContain('<a:ln><a:headEnd type="triangle"/></a:ln>');
    expect(shapeXml(native, 0, nativeNoneLine.id))
      .toContain('<a:ln><a:headEnd type="triangle"/></a:ln>');
    const nativeExplicitNone = nativeSlide.addShape('line', {
      name: 'Arrow native explicit none',
      line: { kind: 'none' },
      arrows: { begin: 'triangle' },
    });
    expect(nativeExplicitNone.line).toEqual({ kind: 'none' });
    expect(nativeExplicitNone.arrows).toEqual({ begin: 'triangle' });
    expect(shapeXml(native, 0, nativeExplicitNone.id))
      .toContain('<a:ln><a:noFill/><a:headEnd type="triangle"/></a:ln>');

    for (const name of ['Arrow empty ignored', 'Arrow nested aliases ignored']) {
      expect(importedShapes.get(name)!.arrows, name).toBeUndefined();
      expect(shapeXml(imported, 0, importedShapes.get(name)!.id), name)
        .not.toMatch(/<a:(?:headEnd|tailEnd)\b/);
    }
    expect(importedShapes.get('Arrow invalid passthrough')!.arrows).toBeUndefined();
    expect(shapeXml(imported, 0, importedShapes.get('Arrow invalid passthrough')!.id))
      .toContain('<a:headEnd type="bogus"/>');
    expect(importedShapes.get('Arrow top aliases mapped')!.arrows)
      .toEqual({ begin: 'stealth', end: 'oval' });

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const arrows of [
      { begin: '' },
      { begin: 'bogus' },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
    ]) {
      expect(() => nativeSlide.addShape('line', { arrows } as never)).toThrow(TypeError);
    }
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);

    const editable = importedShapes.get('Arrow both')!;
    const importedLine = editable.line;
    editable.arrows = { begin: 'diamond' };
    expect(editable.arrows).toEqual({ begin: 'diamond' });
    expect(editable.line).toEqual(importedLine);
    editable.arrows = undefined;
    expect(editable.arrows).toBeUndefined();
    expect(editable.line).toEqual(importedLine);
    editable.arrows = { end: 'stealth' };

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedEditable = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Arrow both',
    ) as ShapeModel;
    expect(reopenedEditable.arrows).toEqual({ end: 'stealth' });
    expect(reopenedEditable.line).toEqual(importedLine);
    const reopenedXml = shapeXml(reopened, 0, reopenedEditable.id);
    expect(reopenedXml).not.toContain('<a:headEnd');
    expect(reopenedXml).toContain('<a:tailEnd type="stealth"/>');
  });

  it('compares shape hyperlink public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const cases: readonly {
      readonly name: string;
      readonly hyperlink: unknown;
    }[] = [
      { name: 'URL', hyperlink: { url: 'https://example.com?a=1&b=2' } },
      {
        name: 'URL tooltip',
        hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      },
      { name: 'Slide', hyperlink: { slide: 2 } },
      { name: 'Slide tooltip', hyperlink: { slide: 3, tooltip: '' } },
      { name: 'Self', hyperlink: { slide: 1 } },
      { name: 'Empty', hyperlink: {} },
      { name: 'Both', hyperlink: { url: 'https://example.com', slide: 2 } },
      { name: 'Zero', hyperlink: { slide: 0 } },
      { name: 'Negative', hyperlink: { slide: -1 } },
      { name: 'Fraction', hyperlink: { slide: 1.5 } },
      { name: 'Out of range', hyperlink: { slide: 99 } },
      { name: 'Numeric URL', hyperlink: { url: 42 } },
      { name: 'String value', hyperlink: 'https://example.com' },
    ];
    const consoleOutput = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    let generatedBytes: Uint8Array;
    try {
      for (const { name, hyperlink } of cases) {
        generatedSlide.addShape(generated.ShapeType.rect!, {
          objectName: name,
          x: 1,
          y: 1,
          w: 1,
          h: 1,
          hyperlink,
        });
      }
      generated.addSlide();
      generated.addSlide();
      generatedBytes = await generated.write({
        outputType: 'nodebuffer',
        compression: true,
      });
      expect(consoleOutput).toHaveBeenCalledTimes(3);
      expect(consoleOutput.mock.calls.map(([message]) => String(message)).join('\n'))
        .toContain('hyperlink requires either');
      expect(consoleOutput.mock.calls.map(([message]) => String(message)).join('\n'))
        .toContain('should be an object');
    } finally {
      consoleOutput.mockRestore();
    }

    const imported = await PptxDocument.open(generatedBytes!);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.get('URL')!.hyperlink).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: '',
    });
    expect(importedShapes.get('URL tooltip')!.hyperlink).toEqual({
      url: 'mailto:test@example.com',
      tooltip: 'Mail & help',
    });
    expect(importedShapes.get('Slide')!.hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(importedShapes.get('Slide tooltip')!.hyperlink)
      .toEqual({ slide: 3, tooltip: '' });
    expect(importedShapes.get('Self')!.hyperlink).toEqual({ slide: 1, tooltip: '' });
    expect(importedShapes.get('Numeric URL')!.hyperlink)
      .toEqual({ url: '42', tooltip: '' });
    for (const name of [
      'Empty',
      'Both',
      'Zero',
      'Negative',
      'Fraction',
      'Out of range',
      'String value',
    ]) {
      expect(importedShapes.get(name)!.hyperlink, name).toBeUndefined();
    }

    const firstSlideXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(shapeXml(imported, 0, importedShapes.get('URL')!.id))
      .toContain('tooltip=""');
    expect(shapeXml(imported, 0, importedShapes.get('URL')!.id))
      .not.toContain('ppaction://hlinksldjump');
    expect(shapeXml(imported, 0, importedShapes.get('Slide')!.id))
      .toContain('action="ppaction://hlinksldjump"');
    expect(shapeXml(imported, 0, importedShapes.get('Both')!.id).match(/<a:hlinkClick/g))
      .toHaveLength(2);
    for (const name of ['Empty', 'Zero', 'String value']) {
      expect(shapeXml(imported, 0, importedShapes.get(name)!.id), name)
        .not.toContain('<a:hlinkClick');
    }
    expect(firstSlideXml).toContain('tooltip="Mail &amp; help"');

    const generatedRelationships = imported.slides[0]!.relationships.filter(
      ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
    );
    expect(generatedRelationships.map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    }))).toEqual([
      { type: 'hyperlink', target: 'https://example.com?a=1&b=2', targetMode: 'External' },
      { type: 'hyperlink', target: 'mailto:test@example.com', targetMode: 'External' },
      { type: 'slide', target: 'slide2.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide3.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide1.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slidehttps://example.com.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide-1.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide1.5.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide99.xml', targetMode: 'Internal' },
      { type: 'hyperlink', target: '42', targetMode: 'External' },
    ]);
    for (const target of [
      '/ppt/slides/slidehttps:/example.com.xml',
      '/ppt/slides/slide-1.xml',
      '/ppt/slides/slide1.5.xml',
      '/ppt/slides/slide99.xml',
    ]) {
      expect(imported.opcPackage.hasPart(target), target).toBe(false);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    native.addSlide();
    const nativeShapes = [
      nativeSlide.addShape('rect', {
        name: 'URL',
        hyperlink: { url: 'https://example.com?a=1&b=2' },
      }),
      nativeSlide.addShape('rect', {
        name: 'URL tooltip',
        hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      }),
      nativeSlide.addShape('rect', { name: 'Slide', hyperlink: { slide: 2 } }),
      nativeSlide.addShape('rect', {
        name: 'Slide tooltip',
        hyperlink: { slide: 3, tooltip: '' },
      }),
      nativeSlide.addShape('rect', { name: 'Self', hyperlink: { slide: 1 } }),
    ];
    expect(nativeShapes.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://example.com?a=1&b=2' },
      { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      { slide: 2 },
      { slide: 3, tooltip: '' },
      { slide: 1 },
    ]);
    expect(nativeSlide.relationships.filter(
      ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
    ).map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    }))).toEqual(generatedRelationships.slice(0, 5).map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    })));

    const beforeInvalid = {
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes: bytes.slice(),
        relationships,
      })),
      mutations: [...native.opcPackage.mutations],
      shapes: [...nativeSlide.shapes],
    };
    let accessorCalls = 0;
    const accessors = (['url', 'slide', 'tooltip'] as const).map((key) =>
      Object.defineProperty({}, key, {
        enumerable: true,
        get() {
          accessorCalls += 1;
          return key === 'slide' ? 1 : 'https://example.com';
        },
      }));
    for (const hyperlink of [
      {},
      { url: '' },
      { url: 42 },
      { url: 'https://example.com', slide: 2 },
      { slide: 0 },
      { slide: -1 },
      { slide: 1.5 },
      { slide: 99 },
      'https://example.com',
      { url: 'https://example.com', _rId: 'rId9' },
      { target: 'https://example.com' },
      { kind: 'url', url: 'https://example.com' },
      { url: 'https://example.com', [Symbol('unsafe')]: true },
      ...accessors,
    ]) {
      expect(() => nativeSlide.addShape('rect', { hyperlink } as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect({
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes,
        relationships,
      })),
      mutations: native.opcPackage.mutations,
      shapes: nativeSlide.shapes,
    }).toEqual(beforeInvalid);

    const validGenerated = new PptxGenJS();
    const validSlide = validGenerated.addSlide();
    for (const { name, hyperlink } of cases.slice(0, 5)) {
      validSlide.addShape(validGenerated.ShapeType.rect!, {
        objectName: name,
        x: 1,
        y: 1,
        w: 1,
        h: 1,
        hyperlink,
      });
    }
    validGenerated.addSlide();
    validGenerated.addSlide();
    const validImported = await openPptxGenJSPublicOutput(validGenerated);
    const duplicate = validImported.duplicateSlide(0);
    expect((duplicate.shapes[4] as ShapeModel).hyperlink).toEqual({
      slide: validImported.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    validImported.moveSlide(1, 0);
    expect((validImported.slides[1]!.shapes[2] as ShapeModel).hyperlink)
      .toEqual({ slide: 1, tooltip: '' });
    expect((validImported.slides[1]!.shapes[4] as ShapeModel).hyperlink)
      .toEqual({ slide: 2, tooltip: '' });
    const reopened = await PptxDocument.open(await validImported.write());
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).hyperlink)
      .toEqual({ url: 'https://example.com?a=1&b=2', tooltip: '' });
    expect((reopened.slides[3]!.shapes[4] as ShapeModel).hyperlink)
      .toEqual({ slide: 4, tooltip: '' });
  });

  it('compares shape shadow public output and strict native divergences', async () => {
    expect(new PptxGenJS().version).toBe('4.0.1');
    const generate = async (name: string, shadow: unknown) => {
      const generated = new PptxGenJS();
      const slide = generated.addSlide();
      const options: Record<string, unknown> = {
        objectName: name,
        x: 1,
        y: 1,
        w: 2,
        h: 1,
      };
      if (shadow !== undefined) options.shadow = shadow;
      slide.addShape(generated.ShapeType.roundRect!, options);
      const bytes = await generated.write({
        outputType: 'nodebuffer',
        compression: true,
      });
      const document = await PptxDocument.open(bytes);
      const partUri = document.slides[0]!.partUri;
      const xml = new TextDecoder().decode(document.opcPackage.requirePart(partUri).bytes);
      return { document, xml };
    };

    const omitted = await generate('Omitted', undefined);
    const none = await generate('None', { type: 'none' });
    for (const fixture of [omitted, none]) {
      expect(fixture.xml).not.toContain('<a:effectLst>');
      expect(fixture.xml).not.toMatch(/<(?:a:)?(?:inner|outer)Shdw/);
      expect((fixture.document.slides[0]!.shapes[0] as ShapeModel).shadow).toBeUndefined();
    }

    const defaults = await generate('Outer defaults', { type: 'outer' });
    expect(defaults.xml).toContain(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="0" blurRad="101600" dist="50800" dir="16200000"> ' +
      '<a:srgbClr val="000000"> <a:alpha val="75000"/></a:srgbClr> ' +
      '</a:outerShdw>',
    );
    expect((defaults.document.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });

    const custom = await generate('Outer custom', {
      type: 'outer',
      color: '123ABC',
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      offset: 5.5,
      rotateWithShape: true,
    });
    expect(custom.xml).toContain(
      'rotWithShape="0" blurRad="92075" dist="69850" dir="7404000"> ' +
      '<a:srgbClr val="123ABC"> <a:alpha val="42000"/>',
    );
    const generatedCustom = (custom.document.slides[0]!.shapes[0] as ShapeModel).shadow;
    expect(generatedCustom).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: false,
    });

    const zero = await generate('Outer zero', {
      type: 'outer',
      color: '000000',
      opacity: 0,
      blur: 0,
      angle: 0,
      offset: 0,
      rotateWithShape: false,
    });
    expect((zero.document.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });

    const malformedInner = await generate('Inner', { type: 'inner' });
    expect(malformedInner.xml).toContain('<a:innerShdw  blurRad="101600"');
    expect(malformedInner.xml).toContain('</a:outerShdw></a:effectLst>');
    expect(() => malformedInner.document.slides[0]!.shapes).toThrow();
    await expect(malformedInner.document.write()).rejects.toThrow();

    const hashColor = await generate('Hash color', {
      type: 'outer',
      color: '#ABCDEF',
    });
    expect(hashColor.xml).toContain('<a:srgbClr val="#ABCDEF">');
    expect((hashColor.document.slides[0]!.shapes[0] as ShapeModel).shadow).toBeUndefined();

    const unknownType = await generate('Unknown type', {
      type: 'bogus',
      color: 'FF0000',
    });
    expect(unknownType.xml).toContain('<a:bogusShdw  blurRad="101600"');
    expect(unknownType.xml).toContain('</a:outerShdw></a:effectLst>');
    expect(() => unknownType.document.slides[0]!.shapes).toThrow();

    const invalidRanges = await generate('Invalid ranges', {
      type: 'outer',
      color: '00FF00',
      opacity: 2,
      blur: -1,
      angle: 400,
      offset: 201,
    });
    expect(invalidRanges.xml).toContain(
      'blurRad="-12700" dist="2552700" dir="24000000"> ' +
      '<a:srgbClr val="00FF00"> <a:alpha val="200000"/>',
    );
    expect((invalidRanges.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toBeUndefined();

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeDefaults = nativeSlide.addShape('roundRect', {
      shadow: { kind: 'outer' },
    });
    const nativeCustom = nativeSlide.addShape('roundRect', {
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123ABC' },
        opacity: 0.42,
        blur: 7.25,
        angle: 123.4,
        distance: 5.5,
        rotateWithShape: true,
      },
    });
    const nativeZero = nativeSlide.addShape('roundRect', {
      shadow: {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
        rotateWithShape: false,
      },
    });
    const nativeInner = nativeSlide.addShape('roundRect', {
      shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent3' } },
    });
    expect(nativeDefaults.shadow).toEqual(
      (defaults.document.slides[0]!.shapes[0] as ShapeModel).shadow,
    );
    expect(nativeCustom.shadow).toMatchObject({
      kind: generatedCustom?.kind,
      color: generatedCustom?.color,
      opacity: generatedCustom?.opacity,
      blur: generatedCustom?.blur,
      angle: generatedCustom?.angle,
      distance: generatedCustom?.distance,
      rotateWithShape: true,
    });
    expect(generatedCustom?.kind === 'outer' && generatedCustom.rotateWithShape).toBe(false);
    expect(nativeZero.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: false,
    });
    expect(nativeInner.shadow?.kind).toBe('inner');
    expect(shapeXml(native, 0, nativeInner.id)).toContain(
      '<a:innerShdw blurRad="101600" dist="50800" dir="16200000">',
    );
    expect(shapeXml(native, 0, nativeInner.id)).toContain('</a:innerShdw>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect((reopenedNative.slides[0]!.shapes[3] as ShapeModel).shadow)
      .toEqual(nativeInner.shadow);

    const beforeInvalid = {
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes: bytes.slice(),
        relationships,
      })),
      mutations: [...native.opcPackage.mutations],
      shapes: [...nativeSlide.shapes],
    };
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'outer';
      },
    });
    for (const shadow of [
      {},
      { kind: 'none' },
      { type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: 2 },
      { kind: 'outer', blur: -1 },
      { kind: 'outer', angle: 400 },
      { kind: 'outer', distance: 201 },
      { kind: 'outer', color: { kind: 'srgb', value: '#ABC' } },
      accessor,
    ]) {
      expect(() => nativeSlide.addShape('rect', { shadow } as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect({
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes,
        relationships,
      })),
      mutations: native.opcPackage.mutations,
      shapes: nativeSlide.shapes,
    }).toEqual(beforeInvalid);
  });

  it('reads and round-trips PptxGenJS preset shape adjustment output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const fixtures = [
      {
        name: 'Round radius',
        type: 'roundRect',
        options: { rectRadius: 0.5, w: 4, h: 2 },
        expected: [{ name: 'adj', value: 25_000 }],
      },
      {
        name: 'Pie angles',
        type: 'pie',
        options: { angleRange: [270, 0] },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      },
      {
        name: 'Fractional arc angles',
        type: 'arc',
        options: { angleRange: [12.34567, 89.99999] },
        expected: [
          { name: 'adj1', value: 740_740 },
          { name: 'adj2', value: 5_399_999 },
        ],
      },
      {
        name: 'Block arc angles and thickness',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0.5 },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
          { name: 'adj3', value: 25_000 },
        ],
      },
      {
        name: 'Omitted shortcuts',
        type: 'roundRect',
        options: {},
        expected: [],
      },
      {
        name: 'Zero radius shortcut',
        type: 'roundRect',
        options: { rectRadius: 0, w: 4, h: 2 },
        expected: [],
      },
      {
        name: 'Zero thickness shortcut',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0 },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      },
    ] as const;

    for (const fixture of fixtures) {
      generatedSlide.addShape(generated.ShapeType[fixture.type]!, {
        objectName: fixture.name,
        ...fixture.options,
      });
    }

    const imported = await importPptxGenJS(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const fixture of fixtures) {
      const importedShape = importedShapes.get(fixture.name);
      expect(importedShape, fixture.name).toBeInstanceOf(ShapeModel);
      expect(importedShape?.adjustments, fixture.name).toEqual(fixture.expected);
      const nativeShape = nativeSlide.addShape(fixture.type, {
        name: fixture.name,
        adjustments: fixture.expected,
      });
      expect(nativeShape.adjustments, fixture.name).toEqual(importedShape?.adjustments);
    }

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    const expected = fixtures.map((fixture) => fixture.expected);
    expect(reopenedImported.slides[0]!.shapes.map((shape) =>
      (shape as ShapeModel).adjustments)).toEqual(expected);
    expect(reopenedNative.slides[0]!.shapes.map((shape) =>
      (shape as ShapeModel).adjustments)).toEqual(expected);
  });

  it('records PptxGenJS adjustment shortcut divergences from native lists', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const generatedCases = [
      {
        name: 'Dropped zero radius',
        type: 'roundRect',
        options: { rectRadius: 0, w: 4, h: 2 },
      },
      {
        name: 'Dropped zero thickness',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0 },
      },
      {
        name: 'String coercion',
        type: 'blockArc',
        options: { angleRange: ['12.5', '90'], arcThicknessRatio: '0.25' },
      },
      {
        name: 'Radius precedence',
        type: 'blockArc',
        options: {
          rectRadius: 0.5,
          angleRange: [270, 0],
          arcThicknessRatio: 0.5,
          w: 4,
          h: 2,
        },
      },
      {
        name: 'Thickness without angles',
        type: 'blockArc',
        options: { arcThicknessRatio: 0.5 },
      },
      {
        name: 'Out of range angles',
        type: 'arc',
        options: { angleRange: [-1, 361] },
      },
      {
        name: 'Malformed coerced angles',
        type: 'arc',
        options: { angleRange: ['not-an-angle', 'also-not-an-angle'] },
      },
      {
        name: 'Unsafe coerced angles',
        type: 'arc',
        options: { angleRange: [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER] },
      },
    ] as const;
    for (const fixture of generatedCases) {
      generatedSlide.addShape(generated.ShapeType[fixture.type]!, {
        objectName: fixture.name,
        ...fixture.options,
      });
    }

    const imported = await importPptxGenJS(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.get('Dropped zero radius')?.adjustments).toEqual([]);
    expect(importedShapes.get('Dropped zero thickness')?.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
    ]);
    expect(importedShapes.get('String coercion')?.adjustments).toEqual([
      { name: 'adj1', value: 750_000 },
      { name: 'adj2', value: 5_400_000 },
      { name: 'adj3', value: 12_500 },
    ]);
    expect(importedShapes.get('Radius precedence')?.adjustments).toEqual([
      { name: 'adj', value: 25_000 },
    ]);
    expect(importedShapes.get('Thickness without angles')?.adjustments).toEqual([]);
    expect(importedShapes.get('Out of range angles')?.adjustments).toEqual([
      { name: 'adj1', value: -60_000 },
      { name: 'adj2', value: 60_000 },
    ]);
    const malformed = importedShapes.get('Malformed coerced angles')!;
    expect(malformed.adjustments).toBeUndefined();
    const malformedXml = shapeXml(imported, 0, malformed.id);
    expect(malformedXml).toContain('fmla="val NaN"');
    const unsafe = importedShapes.get('Unsafe coerced angles')!;
    expect(unsafe.adjustments).toBeUndefined();
    const unsafeXml = shapeXml(imported, 0, unsafe.id);
    expect(unsafeXml).toContain('fmla="val 540431955284437800000"');
    const reopened = await PptxDocument.open(await imported.write());
    const reopenedMalformed = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Malformed coerced angles',
    ) as ShapeModel;
    const reopenedUnsafe = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Unsafe coerced angles',
    ) as ShapeModel;
    expect(reopenedMalformed.adjustments).toBeUndefined();
    expect(reopenedUnsafe.adjustments).toBeUndefined();
    expect(shapeXml(reopened, 0, reopenedMalformed.id)).toBe(malformedXml);
    expect(shapeXml(reopened, 0, reopenedUnsafe.id)).toBe(unsafeXml);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    expect(nativeSlide.addShape('roundRect', {
      adjustments: [{ name: 'adj', value: 0 }],
    }).adjustments).toEqual([{ name: 'adj', value: 0 }]);
    expect(() => nativeSlide.addShape('roundRect', {
      adjustments: [{ name: 'adj', value: '0' }],
    } as never)).toThrow(TypeError);
    const deliberateFinalList = [
      { name: 'adj', value: 25_000 },
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    expect(nativeSlide.addShape('blockArc', {
      adjustments: deliberateFinalList,
    }).adjustments).toEqual(deliberateFinalList);
  });

  it('keeps custom geometry connection sites at the PptxGenJS public boundary', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide().addShape(generated.ShapeType.custGeom!, {
      objectName: 'Public custom geometry',
      x: 1,
      y: 1,
      w: 4,
      h: 3,
      points: [{ x: 0, y: 0 }, { x: 4, y: 3 }, { close: true }],
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShape = imported.slides[0]!.shapes[0] as ShapeModel;
    expect(importedShape.name).toBe('Public custom geometry');
    expect(importedShape.customGeometry).toBeDefined();
    expect(Object.hasOwn(importedShape.customGeometry!, 'connectionSites')).toBe(false);
    expect(shapeXml(imported, 0, importedShape.id)).toMatch(
      /<a:cxnLst(?:\s*\/>|\s*>\s*<\/a:cxnLst>)/,
    );

    const adapterSource = await readFile(new URL('./index.ts', import.meta.url), 'utf8');
    expect(adapterSource).toContain('presentation.write({');
    expect(adapterSource).not.toMatch(/\.\s*_[A-Za-z]/);
    expect(adapterSource).not.toMatch(/\[\s*['"]_[^'"]*['"]\s*\]/);
  });

  it('imports every legal PptxGenJS custom path command as native geometry', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    expect(generated.ShapeType.custGeom).toBe('custGeom');
    const points: readonly PptxGenJSCustomPoint[] = [
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: 2, y: 1, curve: { type: 'quadratic', x1: 1.5, y1: 0 } },
      {
        x: 3,
        y: 2,
        curve: { type: 'cubic', x1: 2.25, y1: 1, x2: 2.75, y2: 2 },
      },
      {
        x: 999,
        y: 999,
        curve: { type: 'arc', wR: 1, hR: 0.5, stAng: 30, swAng: 120 },
      },
      { x: 0.5, y: 0.5, moveTo: true },
      { x: 1.25, y: 1.25 },
      { close: true },
    ];
    generated.addSlide().addShape(generated.ShapeType.custGeom!, {
      objectName: 'All custom commands',
      x: 1,
      y: 1,
      w: 4,
      h: 3,
      points,
    });

    const expected: CustomGeometry = {
      paths: [{
        width: inches(4),
        height: inches(3),
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: inches(1), y: 0 } },
          {
            kind: 'quadraticBezierTo',
            control: { x: inches(1.5), y: 0 },
            end: { x: inches(2), y: inches(1) },
          },
          {
            kind: 'cubicBezierTo',
            control1: { x: inches(2.25), y: inches(1) },
            control2: { x: inches(2.75), y: inches(2) },
            end: { x: inches(3), y: inches(2) },
          },
          {
            kind: 'arcTo',
            widthRadius: inches(1),
            heightRadius: inches(0.5),
            startAngle: degrees(30),
            sweepAngle: degrees(120),
          },
          { kind: 'moveTo', point: { x: inches(0.5), y: inches(0.5) } },
          { kind: 'lineTo', point: { x: inches(1.25), y: inches(1.25) } },
          { kind: 'close' },
        ],
      }],
    };
    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShape = imported.slides[0]!.shapes[0] as ShapeModel;
    expect(importedShape).toBeInstanceOf(ShapeModel);
    expect(importedShape.name).toBe('All custom commands');
    expect(importedShape.presetType).toBeUndefined();
    expect(importedShape.transform).toEqual({
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(importedShape.customGeometry).toEqual(expected);
    expect(Object.hasOwn(importedShape.customGeometry!, 'adjustments')).toBe(false);
    expect(Object.hasOwn(importedShape.customGeometry!, 'guides')).toBe(false);
    expect(Object.hasOwn(importedShape.customGeometry!, 'handles')).toBe(false);
    expect(shapeXml(imported, 0, importedShape.id)).not.toContain('x="999"');
    expect(shapeXml(imported, 0, importedShape.id)).not.toContain('y="999"');

    const native = PptxDocument.create();
    const nativeShape = native.addSlide().addCustomShape(expected, {
      name: 'All custom commands',
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
    });
    expect(nativeShape.customGeometry).toEqual(importedShape.customGeometry);
    const reopened = await PptxDocument.open(await imported.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).customGeometry).toEqual(expected);
  });

  it('classifies PptxGenJS custom path unit heuristics and malformed runtime output', async () => {
    const generated = new PptxGenJS();
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    const validCases: readonly {
      readonly name: string;
      readonly points: readonly PptxGenJSCustomPoint[];
      readonly expected: CustomGeometry;
    }[] = [
      {
        name: 'Custom empty',
        points: [],
        expected: { paths: [{ width: inches(4), height: inches(3), commands: [] }] },
      },
      {
        name: 'Custom direct numeric',
        points: [{ x: 100, y: 200 }, { x: 300, y: 400 }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 100, y: 200 } },
              { kind: 'lineTo', point: { x: 300, y: 400 } },
            ],
          }],
        },
      },
      {
        name: 'Custom numeric strings',
        points: [{ x: '1', y: '2' }, { x: '100', y: '200' }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: inches(1), y: inches(2) } },
              { kind: 'lineTo', point: { x: 100, y: 200 } },
            ],
          }],
        },
      },
      {
        name: 'Custom percentages',
        points: [{ x: '10%', y: '20%' }, { x: '50%', y: '60%' }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 1_219_200, y: 1_371_600 } },
              { kind: 'lineTo', point: { x: 6_096_000, y: 4_114_800 } },
            ],
          }],
        },
      },
      {
        name: 'Custom later move',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1 },
          { x: 2, y: 2, moveTo: true },
          { x: 3, y: 3 },
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 0, y: 0 } },
              { kind: 'lineTo', point: { x: inches(1), y: inches(1) } },
              { kind: 'moveTo', point: { x: inches(2), y: inches(2) } },
              { kind: 'lineTo', point: { x: inches(3), y: inches(3) } },
            ],
          }],
        },
      },
      {
        name: 'Custom invalid kind omitted',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'bogus', x1: 0, y1: 0 } } as never,
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [{ kind: 'moveTo', point: { x: 0, y: 0 } }],
          }],
        },
      },
      {
        name: 'Custom missing cubic defaults',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'cubic', x1: 0, y1: 0 } } as never,
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 0, y: 0 } },
              {
                kind: 'cubicBezierTo',
                control1: { x: 0, y: 0 },
                control2: { x: 0, y: 0 },
                end: { x: inches(1), y: inches(1) },
              },
            ],
          }],
        },
      },
    ];
    for (const fixture of validCases) {
      slide.addShape(generated.ShapeType.custGeom!, {
        objectName: fixture.name,
        x: 1,
        y: 1,
        w: 4,
        h: 3,
        points: fixture.points,
      });
    }
    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    for (const fixture of validCases) {
      const snapshot = importedShapes.get(fixture.name)?.customGeometry;
      expect(snapshot, fixture.name).toEqual(fixture.expected);
      expect(Object.hasOwn(snapshot!, 'adjustments'), fixture.name).toBe(false);
      expect(Object.hasOwn(snapshot!, 'guides'), fixture.name).toBe(false);
      expect(Object.hasOwn(snapshot!, 'handles'), fixture.name).toBe(false);
    }
    expect(imported.slideSize).toEqual({ width: 12_192_000, height: 6_858_000 });

    const unsupported = new PptxGenJS();
    unsupported.layout = 'LAYOUT_WIDE';
    const unsupportedSlide = unsupported.addSlide();
    const unsupportedCases: readonly {
      readonly name: string;
      readonly points: readonly PptxGenJSCustomPoint[];
      readonly malformedXml: string;
    }[] = [
      {
        name: 'Custom first arc',
        points: [{
          x: 0,
          y: 0,
          curve: { type: 'arc', wR: 1, hR: 1, stAng: 0, swAng: 90 },
        }],
        malformedXml: '<a:arcTo hR="914400" wR="914400"',
      },
      {
        name: 'Custom zero radius',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'arc', wR: 0, hR: 1, stAng: 0, swAng: 90 } },
        ],
        malformedXml: 'wR="0"',
      },
      {
        name: 'Custom negative radius',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'arc', wR: -1, hR: 1, stAng: 0, swAng: 90 } },
        ],
        malformedXml: 'wR="-914400"',
      },
      {
        name: 'Custom unsafe coordinate',
        points: [{ x: 0, y: 0 }, { x: Number.MAX_SAFE_INTEGER + 1, y: 1 }],
        malformedXml: 'x="9007199254740992"',
      },
    ];
    for (const fixture of unsupportedCases) {
      unsupportedSlide.addShape(unsupported.ShapeType.custGeom!, {
        objectName: fixture.name,
        x: 1,
        y: 1,
        w: 4,
        h: 3,
        points: fixture.points,
      });
    }
    const unsupportedImported = await openPptxGenJSPublicOutput(unsupported);
    const unsupportedShapes = unsupportedImported.slides[0]!.shapes as ShapeModel[];
    const unsupportedXml = unsupportedShapes.map((shape) =>
      shapeXml(unsupportedImported, 0, shape.id));
    for (const [index, fixture] of unsupportedCases.entries()) {
      expect(unsupportedShapes[index]?.name).toBe(fixture.name);
      expect(unsupportedShapes[index]?.customGeometry, fixture.name).toBeUndefined();
      expect(unsupportedXml[index], fixture.name).toContain(fixture.malformedXml);
    }
    const unsupportedReopened = await PptxDocument.open(await unsupportedImported.write());
    for (const [index, source] of unsupportedXml.entries()) {
      const reopenedShape = unsupportedReopened.slides[0]!.shapes[index] as ShapeModel;
      expect(reopenedShape.customGeometry).toBeUndefined();
      expect(shapeXml(unsupportedReopened, 0, reopenedShape.id)).toBe(source);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const directNative = nativeSlide.addCustomShape({
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 1, y: 2 } }],
      }],
    });
    expect(directNative.customGeometry?.paths[0]?.commands[0]).toEqual({
      kind: 'moveTo',
      point: { x: 1, y: 2 },
    });
    // PptxGenJS 4.0.1 has no public guide-formula input; this is a native extension.
    const formulaNative = nativeSlide.addCustomShape({
      adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
      guides: [{ name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } }],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [{ kind: 'moveTo', point: { x: 'x1', y: 0 } }],
      }],
    });
    expect(formulaNative.customGeometry).toEqual({
      adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
      guides: [{ name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } }],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [{ kind: 'moveTo', point: { x: 'x1', y: 0 } }],
      }],
    });
    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const geometry of [
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'moveTo', point: { x: '10 %', y: 0 } }],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'arcTo', widthRadius: 1, heightRadius: 1, startAngle: 0, sweepAngle: 1 }],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'arcTo', widthRadius: 0, heightRadius: 1, startAngle: 0, sweepAngle: 1 },
          ],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'unknown' },
          ],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'lineTo', point: { x: Number.MAX_SAFE_INTEGER + 1, y: 0 } },
          ],
        }],
      },
    ]) expect(() => nativeSlide.addCustomShape(geometry as never)).toThrow();
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('reads every legal PptxGenJS preset shape public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const legalPublicTypes = PRESET_SHAPE_TYPES.filter((type) => type !== 'foldedCorner');
    const slide = generated.addSlide();
    for (const type of legalPublicTypes) {
      const publicType = generated.ShapeType[type];
      expect(publicType, type).toBe(type);
      slide.addShape(publicType!);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    expect(imported.slides[0]?.shapes).toHaveLength(177);
    expect(imported.slides[0]?.shapes.map((shape) => {
      expect(shape).toBeInstanceOf(ShapeModel);
      return (shape as ShapeModel).presetType;
    })).toEqual(legalPublicTypes);
  }, 30_000);

  it('isolates the folderCorner defect from valid preset shape public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    expect(generated.ShapeType.folderCorner).toBe('folderCorner');
    expect(generated.ShapeType.foldedCorner).toBeUndefined();
    expect(generated.ShapeType.custGeom).toBe('custGeom');
    generated.addSlide().addShape(generated.ShapeType.folderCorner!);

    const imported = await openPptxGenJSPublicOutput(generated);
    const malformedShape = imported.slides[0]?.shapes[0];
    expect(malformedShape).toBeInstanceOf(ShapeModel);
    expect((malformedShape as ShapeModel).presetType).toBeUndefined();
    expect(shapeXml(imported, 0, malformedShape!.id)).toContain('prst="folderCorner"');

    expect(PRESET_SHAPE_TYPES).toContain('foldedCorner');
    expect(PRESET_SHAPE_TYPES).not.toContain('folderCorner');
    expect(PRESET_SHAPE_TYPES).not.toContain('custGeom');
    const native = PptxDocument.create();
    const foldedCorner = native.addSlide().addShape('foldedCorner');
    const reopened = await PptxDocument.open(await native.write());
    expect(foldedCorner.presetType).toBe('foldedCorner');
    expect((reopened.slides[0]?.shapes[0] as ShapeModel).presetType).toBe('foldedCorner');
    expect(() => native.slides[0]!.addShape('folderCorner' as never)).toThrow(TypeError);
  });

  it('matches native basic table creation to public PptxGenJS plain-table output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    const rows = [
      ['A1', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
    ] as const;
    generatedSlide.addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1.5, w: 6, h: 2, colW: [1, 2, 3], rowH: [0.75, 1.25] },
    );

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeOptions = {
      x: inches(1),
      y: inches(1.5),
      width: inches(6),
      height: inches(2),
      columnWidths: [inches(1), inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    };
    const nativeTable = native.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text }))),
      nativeOptions,
    );
    const stringNative = PptxDocument.create({ slideSize: 'wide' });
    stringNative.addSlide().addTable(rows, nativeOptions);

    expect(importedTable).toBeInstanceOf(TableModel);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
    );

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const stringNativeXml = new TextDecoder().decode(
      stringNative.opcPackage.requirePart(stringNative.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(nativeXml).toBe(stringNativeXml);
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).toContain(
        'uri="http://schemas.openxmlformats.org/drawingml/2006/table"',
      );
      const columnWidths = [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
        .map((match) => Number(match[1]));
      const rowHeights = [...xml.matchAll(/<a:tr h="(\d+)">/g)]
        .map((match) => Number(match[1]));
      expect(columnWidths).toHaveLength(3);
      expect(columnWidths).toEqual([inches(1), inches(2), inches(3)]);
      expect(columnWidths.reduce((sum, width) => sum + width, 0)).toBe(5_486_400);
      expect(rowHeights).toEqual([inches(0.75), inches(1.25)]);
      expect(xml.match(/<a:tc>/g)).toHaveLength(6);
      expect(xml.match(/marL="91440" marR="91440" marT="45720" marB="45720"/g))
        .toHaveLength(6);
      const properties = xml.match(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/)?.[1];
      expect(properties).toBeDefined();
      const left = properties!.indexOf('<a:lnL ');
      const right = properties!.indexOf('<a:lnR ');
      const top = properties!.indexOf('<a:lnT ');
      const bottom = properties!.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }
    expect(nativeXml).not.toContain('p14:modId');
    expect(nativeXml).toContain('<a:ext cx="5486400" cy="1828800"/>');

    expect(importedTable.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    importedTable.setColumnWidths([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(importedTable.columnWidths).toEqual([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(importedTable.transform.width).toBe(inches(6));
    expect(importedTable.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
    ]);
    importedTable.setRowHeights([inches(1), inches(1.5)]);
    expect(importedTable.rowHeights).toEqual([
      inches(1),
      inches(1.5),
    ]);
    expect(importedTable.transform.height).toBe(inches(2.5));

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNativeTable = reopenedNative.slides[0]!.shapes[0] as TableModel;
    const reopenedImportedTable = reopenedImported.slides[0]!.shapes[0] as TableModel;
    expect(reopenedNativeTable.rows).toEqual(nativeTable.rows);
    expect(reopenedImportedTable.rows).toEqual(importedTable.rows);
    expect(reopenedNativeTable.transform).toEqual(nativeTable.transform);
    expect(reopenedImportedTable.transform).toEqual(importedTable.transform);
    expect(reopenedImportedTable.columnWidths).toEqual([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(reopenedImportedTable.rowHeights).toEqual([
      inches(1),
      inches(1.5),
    ]);
  });

  it('repairs a PptxGenJS transform and column-grid mismatch through the public model', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [['A', 'B', 'C'].map((text) => ({ text, options: {} }))],
      { x: 1, y: 1, w: 5, h: 1, colW: [1, 2, 3] },
    );

    const imported = await importPptxGenJS(generated);
    const table = imported.slides[0]!.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    expect(table.transform.width).toBe(inches(5));
    const before = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const gridBefore = [...before.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => match[1]);

    table.setColumnWidths(table.columnWidths!);

    expect(table.transform.width).toBe(inches(6));
    const after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect([...after.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => match[1])).toEqual(gridBefore);
    expect(after).toContain('<a:ext cx="5486400" cy="914400"/>');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    expect(reopenedTable.transform.width).toBe(inches(6));
  });

  it('repairs PptxGenJS scalar column-width floor while preserving public intent', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const rows = [['A', 'B', 'C']] as const;
    generated.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1, h: 1, colW: 1.25 },
    );
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(rows, {
      x: inches(1),
      y: inches(1),
      height: inches(1),
      columnWidths: inches(1.25),
    });

    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const readColumnWidths = (xml: string): number[] =>
      [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((match) => Number(match[1]));
    expect(readColumnWidths(importedXml)).toEqual([
      inches(1),
      inches(1),
      inches(1),
    ]);
    expect(readColumnWidths(nativeXml)).toEqual([
      inches(1.25),
      inches(1.25),
      inches(1.25),
    ]);
    expect((imported.slides[0]!.shapes[0] as TableModel).transform.width).toBe(inches(3));
    expect(nativeTable.transform.width).toBe(inches(3.75));
    expect(nativeXml).toContain('<a:ext cx="3429000" cy="914400"/>');
  });

  it('repairs an explicit PptxGenJS row-height mismatch and preserves transform height for automatic rows', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [
        ['A', 'B'].map((text) => ({ text, options: {} })),
        ['C', 'D'].map((text) => ({ text, options: {} })),
      ],
      { x: 1, y: 1, w: 5, h: 1, rowH: [0.5, 1.5] },
    );

    const imported = await importPptxGenJS(generated);
    const table = imported.slides[0]!.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.rowHeights).toEqual([inches(0.5), inches(1.5)]);
    expect(table.transform.height).toBe(inches(1));
    const before = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const rowsBefore = [...before.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => match[1]);

    table.setRowHeights(table.rowHeights!);

    expect(table.transform.height).toBe(inches(2));
    let after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect([...after.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => match[1])).toEqual(rowsBefore);
    expect(after).toContain('cy="1828800"');

    table.setRowHeights([0, inches(1.5)]);
    expect(table.rowHeights).toEqual([0, inches(1.5)]);
    expect(table.transform.height).toBe(inches(2));
    after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(after).toContain('cy="1828800"');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.rowHeights).toEqual([0, inches(1.5)]);
    expect(reopenedTable.transform.height).toBe(inches(2));
  });

  it('repairs PptxGenJS omitted-height mismatch for explicit row heights', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const rows = [['A'], ['B'], ['C']] as const;
    generated.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1, rowH: [0.5, 1, 1.5] },
    );
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(rows, {
      x: inches(1),
      y: inches(1),
      rowHeights: [inches(0.5), inches(1), inches(1.5)],
    });

    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const readRowHeights = (xml: string): number[] =>
      [...xml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]));
    const expectedHeights = [inches(0.5), inches(1), inches(1.5)];
    expect(readRowHeights(importedXml)).toEqual(expectedHeights);
    expect(readRowHeights(nativeXml)).toEqual(expectedHeights);
    expect((imported.slides[0]!.shapes[0] as TableModel).transform.height).toBe(inches(1));
    expect(nativeTable.transform.height).toBe(inches(3));
    expect(nativeXml).toContain('<a:ext cx="914400" cy="2743200"/>');
  });

  it('imports PptxGenJS table-cell text directions with exact four-value semantics', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        { text: 'Horizontal', options: { textDirection: 'horz' } },
        { text: 'Vertical', options: { textDirection: 'vert' } },
        { text: 'Rotate 270', options: { textDirection: 'vert270' } },
        { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
        { text: 'Invalid', options: { textDirection: 'eaVert' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, textDirection: 'vert270' },
    );
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Explicit horizontal', options: { textDirection: 'horz' } },
      ]],
      { x: 0.5, y: 2, w: 12, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(2);
    expect(tables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
    expect(tables[0]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Inherited',
      'Horizontal',
      'Vertical',
      'Rotate 270',
      'Stacked',
      'Invalid',
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Explicit horizontal',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* vert="vert270"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* vert="vert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="wordArtVert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="eaVert"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:tcPr[^>]* vert="horz"/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(reopenedTables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
  });

  it('matches native table text direction creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Omitted', options: {} },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10,
      h: 1,
      colW: [2, 2, 2, 2, 2],
      rowH: 1,
      margin: 0.1,
      textDirection: 'vert270',
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Omitted' },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(10),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: 7.2,
      textDirection: 'vert270',
      valign: 'middle',
    });
    const expectedTokens = ['vert270', undefined, 'vert', 'vert270', 'wordArtVert'];
    const expectedText = ['Omitted', 'Horizontal', 'Vertical', 'Rotate 270', 'Stacked'];
    const directDirectionTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:tcPr[^>]*\svert="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directDirectionTokens(nativeXml)).toEqual(expectedTokens);
    expect(directDirectionTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(expectedTokens);
    expect(importedTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directDirectionTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directDirectionTokens(slideXml(reopenedImported))).toEqual(expectedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[{
      text: 'PptxGenJS invalid east Asian vertical',
      options: { textDirection: 'eaVert' },
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directDirectionTokens(slideXml(importedInvalid))).toEqual(['eaVert']);
    expect((importedInvalid.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.textDirection).toBeUndefined();

    const generatedInvalidTable = new PptxGenJS();
    generatedInvalidTable.addSlide().addTable([[{
      text: 'PptxGenJS invalid table direction',
      options: {},
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
      textDirection: 'eaVert',
    });
    const importedInvalidTable = await importPptxGenJS(generatedInvalidTable);
    expect(directDirectionTokens(slideXml(importedInvalidTable))).toEqual(['eaVert']);
    expect((importedInvalidTable.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.textDirection).toBeUndefined();

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([[{
      text: 'Native invalid east Asian vertical',
      options: { textDirection: 'eaVert' as never },
    }]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
    expect(() => nativeSlide.addTable([['Native invalid table direction']], {
      textDirection: 'eaVert' as never,
    })).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('imports PptxGenJS table fit-like runtime options as fit-less cells', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Fit none', options: { fit: 'none' } },
        { text: 'Fit shrink', options: { fit: 'shrink' } },
        { text: 'Fit resize', options: { fit: 'resize' } },
        { text: 'Auto fit', options: { autoFit: true } },
        { text: 'Shrink text', options: { shrinkText: true } },
        {
          text: 'Conflicting',
          options: { fit: 'resize', autoFit: true, shrinkText: true, textDirection: 'vert' },
        },
      ]],
      {
        x: 0.5,
        y: 0.5,
        w: 12,
        h: 1,
        fit: 'resize',
        autoFit: true,
        shrinkText: true,
        textDirection: 'vert270',
      },
    );

    const document = await importPptxGenJS(generated);
    const table = document.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(table?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(table?.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Auto fit',
      'Shrink text',
      'Conflicting',
    ]);
    expect(table?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).not.toMatch(/<a:(?:noAutofit|normAutofit|spAutoFit)\b/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Omitted' },
      { text: 'None', options: { fit: 'none' } },
      { text: 'Shrink', options: { fit: 'shrink' } },
      { text: 'Resize', options: { fit: 'resize', textDirection: 'vert' } },
    ]], {
      columnWidths: inches(2),
      rowHeights: inches(1),
    });
    expect(nativeTable.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
    ]);
    expect(nativeTable.rows[0]!.cells[3]!.textDirection).toBe('vert');

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    expect(nativeXml.match(/<a:normAutofit\/>/g)).toHaveLength(1);
    expect(nativeXml.match(/<a:spAutoFit\/>/g)).toHaveLength(1);
    expect(nativeXml).not.toContain('<a:noAutofit/>');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes;
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([[
      { text: 'Invalid', options: { fit: 'SHRINK' as never } },
    ]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('imports PptxGenJS table-cell vertical alignments from direct cell anchors', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited bottom', options: {} },
        { text: 'Top', options: { valign: 'top' } },
        { text: 'Middle', options: { valign: 'middle' } },
        { text: 'Bottom', options: { valign: 'bottom' } },
        { text: 'Invalid mid', options: { valign: 'mid' } },
        { text: 'Invalid distributed', options: { valign: 'distributed' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, valign: 'bottom' },
    );
    slide.addTable(
      [[{ text: 'Inherited top', options: {} }]],
      { x: 0.5, y: 2, w: 3, h: 1, valign: 'top' },
    );
    slide.addTable(
      [[{ text: 'Inherited middle', options: {} }]],
      { x: 4, y: 2, w: 3, h: 1, valign: 'middle' },
    );
    slide.addTable(
      [[{ text: 'Omitted direct alignment', options: {} }]],
      { x: 7.5, y: 2, w: 3, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(4);
    expect(tables[0]!.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(tables.slice(1).map((table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Inherited bottom',
      'Top',
      'Middle',
      'Bottom',
      'Invalid mid',
      'Invalid distributed',
      'Inherited top',
      'Inherited middle',
      'Omitted direct alignment',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* anchor="t"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="ctr"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="b"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="mid"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* anchor="distributed"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:bodyPr[^>]* anchor=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(reopenedTables.slice(1).map(
      (table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
  });

  it('matches native table-level valign creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Middle', options: { valign: 'middle' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [1.6, 1.6, 1.6, 1.6, 1.6],
      rowH: 1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(importedTable).toBeInstanceOf(TableModel);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Middle', options: { valign: 'middle' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.6),
      rowHeights: inches(1),
      valign: 'middle',
    });
    const expectedAlignments = ['middle', 'middle', 'top', 'middle', 'bottom'];
    const expectedText = ['Inherited string', 'Inherited object', 'Top', 'Middle', 'Bottom'];
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect(importedTable!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable!.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.transform).toMatchObject(importedTable!.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable!.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable!.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const xml of [nativeXml, importedXml]) {
      const properties = [...xml.matchAll(/<a:tcPr([^>]*)>/g)]
        .map((match) => match[1]!);
      expect(properties).toHaveLength(5);
      expect(properties.map((attributes) =>
        attributes.match(/\sanchor="([^"]+)"/)?.[1])).toEqual([
        'ctr',
        'ctr',
        't',
        'ctr',
        'b',
      ]);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\sanchor=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([[{
      text: 'Invalid mid',
      options: { valign: 'mid' as never },
    }]])).toThrow(TypeError);
    expect(() => nativeSlide.addTable(
      [['Invalid distributed']],
      { valign: 'distributed' as never },
    )).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('matches native table-cell horizontal alignment to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Default', options: {} },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Center', options: { align: 'center' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify this sentence', options: { align: 'justify' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10,
      h: 1,
      colW: [2, 2, 2, 2, 2],
      rowH: 1,
      margin: 0.1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Default' },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Center', options: { align: 'center' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify this sentence', options: { align: 'justify' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(10),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: 7.2,
      valign: 'middle',
    });
    const expectedTokens = [undefined, 'l', 'ctr', 'r', 'just'];
    const expectedAlignments = [undefined, 'left', 'center', 'right', 'justify'];
    const expectedText = ['Default', 'Left', 'Center', 'Right', 'Justify this sentence'];
    const directAlignmentTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directAlignmentTokens(nativeXml)).toEqual(expectedTokens);
    expect(directAlignmentTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directAlignmentTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directAlignmentTokens(slideXml(reopenedImported))).toEqual(expectedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[{
      text: 'PptxGenJS invalid distributed',
      options: { align: 'dist' },
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directAlignmentTokens(slideXml(importedInvalid))).toEqual([undefined]);
    expect((importedInvalid.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.horizontalAlignment).toBeUndefined();

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([[{
      text: 'Native invalid distributed',
      options: { align: 'dist' as never },
    }]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('matches native table horizontal alignment to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Inherited undefined', options: { align: undefined } },
      { text: 'Left override', options: { align: 'left' } },
      { text: 'Right override', options: { align: 'right' } },
      { text: 'Justify override', options: { align: 'justify' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      colW: [1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
      rowH: 1,
      align: 'center',
      margin: 0.1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Inherited undefined', options: { align: undefined } as never },
      { text: 'Left override', options: { align: 'left' } },
      { text: 'Right override', options: { align: 'right' } },
      { text: 'Justify override', options: { align: 'justify' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.5),
      rowHeights: inches(1),
      align: 'center',
      margin: 7.2,
      valign: 'middle',
    });
    const expectedTokens = ['ctr', 'ctr', 'ctr', 'l', 'r', 'just'];
    const expectedAlignments = ['center', 'center', 'center', 'left', 'right', 'justify'];
    const expectedText = [
      'Inherited string',
      'Inherited object',
      'Inherited undefined',
      'Left override',
      'Right override',
      'Justify override',
    ];
    const directAlignmentTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directAlignmentTokens(nativeXml)).toEqual(expectedTokens);
    expect(directAlignmentTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
    }

    const importedSlidePartUri = imported.slides[0]!.partUri;
    const nonSlideParts = new Map(imported.opcPackage.parts
      .filter(({ uri }) => uri !== importedSlidePartUri)
      .map(({ uri, bytes }) => [uri, bytes.slice()]));
    const importedText = importedTable.rows[0]!.cells.map(({ text }) => text);
    const importedMargins = importedTable.rows[0]!.cells.map(({ margins }) => margins);
    const importedVerticalAlignments = importedTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment);
    const importedBorders = importedTable.rows[0]!.cells.map(({ borders }) => borders);
    const importedFills = importedTable.rows[0]!.cells.map(({ fill }) => fill);
    const importedTransform = importedTable.transform;
    const importedColumnWidths = importedTable.columnWidths;
    const importedRowHeights = importedTable.rowHeights;

    importedTable.setCellHorizontalAlignment(0, 0, 'right');
    importedTable.setCellHorizontalAlignment(0, 3, undefined);
    importedTable.setCellHorizontalAlignment(0, 5, 'center');
    const editedTokens = ['r', 'ctr', 'ctr', undefined, 'r', 'ctr'];
    const editedAlignments = ['right', 'center', 'center', undefined, 'right', 'center'];
    expect(directAlignmentTokens(slideXml(imported))).toEqual(editedTokens);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(editedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(importedText);
    expect(importedTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(importedMargins);
    expect(importedTable.rows[0]!.cells.map(({ verticalAlignment }) =>
      verticalAlignment)).toEqual(importedVerticalAlignments);
    expect(importedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(importedBorders);
    expect(importedTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(importedFills);
    expect(importedTable.transform).toEqual(importedTransform);
    expect(importedTable.columnWidths).toEqual(importedColumnWidths);
    expect(importedTable.rowHeights).toEqual(importedRowHeights);
    for (const [uri, bytes] of nonSlideParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directAlignmentTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directAlignmentTokens(slideXml(reopenedImported))).toEqual(editedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(editedAlignments);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[
      { text: 'Invalid inherited', options: {} },
      { text: 'Right override', options: { align: 'right' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 4,
      h: 1,
      colW: [2, 2],
      rowH: 1,
      align: 'dist',
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directAlignmentTokens(slideXml(importedInvalid))).toEqual([
      undefined,
      'r',
    ]);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([
      ['Invalid inherited', { text: 'Right override', options: { align: 'right' } }],
    ], {
      align: 'dist' as never,
    })).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('imports PptxGenJS table-cell margins from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted defaults', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table zero', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, margin: 0 },
    );
    slide.addTable(
      [[{ text: 'Table 0.1 inch', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, margin: 0.1 },
    );
    slide.addTable(
      [[{ text: 'Table inch tuple', options: {} }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, margin: [0.05, 0.1, 0.15, 0.2] },
    );
    slide.addTable(
      [[
        { text: 'Inherited 0.1', options: {} },
        { text: 'Cell zero', options: { margin: 0 } },
        { text: 'Cell quarter inch', options: { margin: 0.25 } },
        { text: 'Cell inch tuple', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
        { text: 'Cell scalar one point', options: { margin: 1 } },
        { text: 'Cell point tuple', options: { margin: [1, 2, 3, 4] } },
        { text: 'Cell negative inch', options: { margin: -0.1 } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, margin: 0.1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins));
    expect(snapshots).toEqual([
      [{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }],
      [{ top: 0, right: 0, bottom: 0, left: 0 }],
      [{ top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 }],
      [{ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 }],
      [
        { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
        { top: 0, right: 0, bottom: 0, left: 0 },
        { top: 18, right: 18, bottom: 18, left: 18 },
        { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
        { top: 1, right: 1, bottom: 1, left: 1 },
        { top: 1, right: 2, bottom: 3, left: 4 },
        { top: -7.2, right: -7.2, bottom: -7.2, left: -7.2 },
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted defaults',
      'Table zero',
      'Table 0.1 inch',
      'Table inch tuple',
      'Inherited 0.1',
      'Cell zero',
      'Cell quarter inch',
      'Cell inch tuple',
      'Cell scalar one point',
      'Cell point tuple',
      'Cell negative inch',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* marL=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marR=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marT=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marB=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr marL="91440" marR="91440" marT="91440" marB="91440">/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr marL="182880" marR="91440" marT="45720" marB="137160">/g)).toHaveLength(2);
    expect(xml).toContain('<a:tcPr marL="12700" marR="12700" marT="12700" marB="12700">');
    expect(xml).toContain('<a:tcPr marL="50800" marR="25400" marT="12700" marB="38100">');
    expect(xml).toContain('<a:tcPr marL="-91440" marR="-91440" marT="-91440" marB="-91440">');
    expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns)=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins))).toEqual(snapshots);
  });

  it('matches native table-cell margin creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Default', options: {} },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'One point', options: { margin: 1 } },
      { text: 'Seven point two', options: { margin: 0.1 } },
      { text: 'TRBL', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
      { text: 'Negative', options: { margin: -0.1 } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10.5,
      h: 1,
      colW: [1.75, 1.75, 1.75, 1.75, 1.75, 1.75],
      rowH: 1,
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(importedTable).toBeInstanceOf(TableModel);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable([[
      { text: 'Default' },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'One point', options: { margin: 1 } },
      { text: 'Seven point two', options: { margin: 7.2 } },
      { text: 'TRBL', options: { margin: [3.6, 7.2, 10.8, 14.4] } },
      { text: 'Negative', options: { margin: -7.2 } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.75),
      rowHeights: inches(1),
    });
    const expectedMargins = [
      { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 1, bottom: 1, left: 1 },
      { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: -7.2, right: -7.2, bottom: -7.2, left: -7.2 },
    ];
    const expectedText = [
      'Default',
      'Zero',
      'One point',
      'Seven point two',
      'TRBL',
      'Negative',
    ];
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(expectedMargins);
    expect(importedTable!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(expectedMargins);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable!.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable!.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable!.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable!.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const directTokens = [
      'marL="91440" marR="91440" marT="45720" marB="45720"',
      'marL="0" marR="0" marT="0" marB="0"',
      'marL="12700" marR="12700" marT="12700" marB="12700"',
      'marL="91440" marR="91440" marT="91440" marB="91440"',
      'marL="182880" marR="91440" marT="45720" marB="137160"',
      'marL="-91440" marR="-91440" marT="-91440" marB="-91440"',
    ];
    for (const xml of [nativeXml, importedXml]) {
      for (const token of directTokens) expect(xml).toContain(token);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows).toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable!.rows);

    const nativeDifferences = PptxDocument.create({ slideSize: 'wide' });
    const nativeDifferenceTable = nativeDifferences.addSlide().addTable([[
      { text: 'Native 0.1 point', options: { margin: 0.1 } },
      { text: 'Native partial', options: { margin: { top: 2, left: -2 } } },
    ]]);
    expect(nativeDifferenceTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
      { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 },
      { top: 2, right: 7.2, bottom: 3.6, left: -2 },
    ]);
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    expect(nativeDifferencesXml).toContain(
      '<a:tcPr marL="1270" marR="1270" marT="1270" marB="1270">',
    );
    expect(nativeDifferencesXml).toContain(
      '<a:tcPr marL="-25400" marR="91440" marT="25400" marB="45720">',
    );
    expect(importedXml).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="91440" marB="91440">',
    );
  });

  it('matches native table-level margin creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Zero', options: {
        margin: 0,
        border: { type: 'solid', color: '336699', pt: 1 },
        fill: { color: 'DDEEFF' },
      } },
      { text: 'Points', options: {
        margin: [1, 2, 3, 4],
        border: { type: 'dash', color: 'CC3300', pt: 1.5 },
        fill: { color: '112233', transparency: 25 },
      } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [2, 2, 2, 2],
      rowH: [1],
      margin: [0.05, 0.1, 0.15, 0.2],
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Zero', options: {
        margin: 0,
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: '336699' },
          width: 1,
          style: 'solid',
        },
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
      } },
      { text: 'Points', options: {
        margin: [1, 2, 3, 4],
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'CC3300' },
          width: 1.5,
          style: 'dash',
        },
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 25,
        },
      } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: [3.6, 7.2, 10.8, 14.4],
      valign: 'middle',
    });
    const expectedMargins = [
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ];
    const expectedText = ['Inherited string', 'Inherited object', 'Zero', 'Points'];
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(expectedMargins);
    expect(importedTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(expectedMargins);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(Array(4).fill('middle'));
    expect(importedTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(Array(4).fill('middle'));
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const directMarginVectors = (xml: string) => [...xml.matchAll(
      /<a:tcPr marL="(-?\d+)" marR="(-?\d+)" marT="(-?\d+)" marB="(-?\d+)"/g,
    )].map((match) => match.slice(1).map(Number));
    const expectedMarginVectors = [
      [182880, 91440, 45720, 137160],
      [182880, 91440, 45720, 137160],
      [0, 0, 0, 0],
      [50800, 25400, 12700, 38100],
    ];
    expect(directMarginVectors(nativeXml)).toEqual(expectedMarginVectors);
    expect(directMarginVectors(importedXml)).toEqual(expectedMarginVectors);
    for (const xml of [nativeXml, importedXml]) {
      expect([...xml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)]
        .map((match) => match[1])).toEqual(Array(4).fill('ctr'));
      expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns|anchor)=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedDifference = new PptxGenJS();
    generatedDifference.addSlide().addTable(
      [[{ text: 'PptxGenJS 0.1 table', options: {} }]],
      { x: 1, y: 1, w: 2, h: 1, colW: [2], rowH: [1], margin: 0.1 },
    );
    const importedDifference = await importPptxGenJS(generatedDifference);
    const nativeDifference = PptxDocument.create();
    const nativeDifferenceTable = nativeDifference.addSlide().addTable(
      [['Native 0.1 table']],
      { margin: 0.1 },
    );
    expect((importedDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.margins).toEqual({
      top: 7.2,
      right: 7.2,
      bottom: 7.2,
      left: 7.2,
    });
    expect(nativeDifferenceTable.rows[0]!.cells[0]!.margins).toEqual({
      top: 0.1,
      right: 0.1,
      bottom: 0.1,
      left: 0.1,
    });

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([['Invalid']], { margin: null as never }))
      .toThrow(TypeError);
    expect(() => nativeSlide.addTable([['Invalid']], { margin: [1, 2, 3] as never }))
      .toThrow(RangeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('matches native table-level solid fill creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Cell yellow', options: {
        fill: { color: 'FFFF00', transparency: 50 },
        border: { type: 'solid', color: '336699', pt: 1 },
      } },
      { text: 'Cell full transparency', options: {
        fill: { color: '445566', transparency: 100 },
      } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [2, 2, 2, 2],
      rowH: [1],
      fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      margin: 0.1,
      valign: 'middle',
    });
    generatedSlide.addTable(
      [[{ text: 'Table sRGB opaque', options: {} }]],
      {
        x: 0.5,
        y: 2,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: 'FF0000' },
      },
    );
    generatedSlide.addTable(
      [[{ text: 'Table fractional', options: {} }]],
      {
        x: 0.5,
        y: 2.8,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: '112233', transparency: 33.333 },
      },
    );
    generatedSlide.addTable(
      [[{ text: 'Table full transparency', options: {} }]],
      {
        x: 0.5,
        y: 3.6,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: '445566', transparency: 100 },
      },
    );
    const imported = await importPptxGenJS(generated);
    const importedTables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Cell yellow', options: {
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: '336699' },
          width: 1,
          style: 'solid',
        },
      } },
      { text: 'Cell full transparency', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      } } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(8),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      margin: 7.2,
      valign: 'middle',
    });
    nativeSlide.addTable([['Table sRGB opaque']], {
      x: inches(0.5),
      y: inches(2),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
      },
    });
    nativeSlide.addTable([['Table fractional']], {
      x: inches(0.5),
      y: inches(2.8),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 33.333,
      },
    });
    nativeSlide.addTable([['Table full transparency']], {
      x: inches(0.5),
      y: inches(3.6),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      },
    });
    const nativeTables = nativeSlide.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const expectedFills = [
      [
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
      ],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } }],
      [{
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 33.333,
      }],
      [{
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      }],
    ];
    expect(nativeTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(expectedFills);
    expect(importedTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(expectedFills);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ text }) => text),
    );
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(
        ({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ borders }) => borders),
    );
    for (let index = 0; index < nativeTables.length; index += 1) {
      expect(nativeTables[index]!.transform).toMatchObject(
        importedTables[index]!.transform,
      );
      expect(nativeTables[index]!.columnWidths).toEqual(
        importedTables[index]!.columnWidths,
      );
      expect(nativeTables[index]!.rowHeights).toEqual(
        importedTables[index]!.rowHeights,
      );
    }

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml.match(
        /<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"\/><\/a:schemeClr><\/a:solidFill>/g,
      )).toHaveLength(2);
      expect(xml).toContain(
        '<a:solidFill><a:srgbClr val="FFFF00">'
        + '<a:alpha val="50000"/></a:srgbClr></a:solidFill>',
      );
      expect(xml).toContain(
        '<a:solidFill><a:srgbClr val="112233">'
        + '<a:alpha val="66667"/></a:srgbClr></a:solidFill>',
      );
      expect(xml.match(
        /<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"\/><\/a:srgbClr><\/a:solidFill>/g,
      )).toHaveLength(2);
      for (const properties of xml.matchAll(
        /<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g,
      )) {
        const direct = properties[1]!;
        const fill = Math.max(
          direct.lastIndexOf('<a:noFill/>'),
          direct.lastIndexOf('<a:solidFill>'),
        );
        expect(fill).toBeGreaterThan(direct.indexOf('</a:lnB>'));
      }
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    )).map((table) => table.rows[0]!.cells.map(({ fill }) => fill)))
      .toEqual(expectedFills);
    expect((reopenedImported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    )).map((table) => table.rows[0]!.cells.map(({ fill }) => fill)))
      .toEqual(expectedFills);

    const generatedDifferences = new PptxGenJS();
    generatedDifferences.addSlide().addTable([[
      { text: 'Collapsed none', options: {} },
      { text: 'Collapsed zero alpha', options: {
        fill: { color: '00FF00', transparency: 0 },
      } },
    ]], {
      x: 1,
      y: 1,
      w: 4,
      h: 1,
      colW: [2, 2],
      rowH: [1],
      fill: { type: 'none' },
    });
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTable =
      importedDifferences.slides[0]!.shapes[0] as TableModel;
    expect(importedDifferenceTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      undefined,
      { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceSlide = nativeDifferences.addSlide();
    const nativeDifferenceTable = nativeDifferenceSlide.addTable([[
      'Inherited none',
      { text: 'Explicit zero override', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      } } },
    ]], { fill: { kind: 'none' } });
    expect(nativeDifferenceTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    ]);
    const nativeDifferenceXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferenceSlide.partUri).bytes,
    );
    expect(nativeDifferenceXml).toContain('</a:lnB><a:noFill/></a:tcPr>');
    expect(nativeDifferenceXml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00">'
      + '<a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );
    const reopenedNativeDifference = await PptxDocument.open(
      await nativeDifferences.write(),
    );
    expect((reopenedNativeDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    ]);
    const reopenedImportedDifference = await PptxDocument.open(
      await importedDifferences.write(),
    );
    expect((reopenedImportedDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      undefined,
      { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
    ]);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([['Invalid']], { fill: {} as never }))
      .toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(4);
  });

  it('imports PptxGenJS table-cell fills from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted fill', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Type none', options: { fill: { type: 'none' } } }]],
      { x: 0.2, y: 1, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table red', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, fill: { color: 'FF0000' } },
    );
    slide.addTable(
      [[{ text: 'Table theme alpha', options: {} }]],
      {
        x: 0.2,
        y: 2.6,
        w: 2,
        h: 0.5,
        fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      },
    );
    slide.addTable(
      [[
        { text: 'Inherited blue', options: {} },
        { text: 'Cell yellow alpha', options: { fill: { color: 'FFFF00', transparency: 50 } } },
        { text: 'Explicit zero', options: { fill: { color: '00FF00', transparency: 0 } } },
        { text: 'Fractional', options: { fill: { color: '112233', transparency: 33.333 } } },
        { text: 'Full transparency', options: { fill: { color: '445566', transparency: 100 } } },
        { text: 'Deprecated alpha', options: { fill: { color: generated.SchemeColor.accent2, alpha: 25 } } },
        { text: 'Runtime negative', options: { fill: { color: '778899', transparency: -1 } } },
        { text: 'Runtime overflow', options: { fill: { color: 'AABBCC', transparency: 101 } } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, fill: { color: '0000FF' } },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) => table.rows[0]!.cells.map(({ fill }) => fill));
    expect(snapshots).toEqual([
      [undefined],
      [undefined],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } }],
      [{
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      }],
      [
        { kind: 'solid', color: { kind: 'srgb', value: '0000FF' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 33.333,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        undefined,
        undefined,
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted fill',
      'Type none',
      'Table red',
      'Table theme alpha',
      'Inherited blue',
      'Cell yellow alpha',
      'Explicit zero',
      'Fractional',
      'Full transparency',
      'Deprecated alpha',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const directFillXml = [...xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)]
      .map((match) => match[1]!.replace(/<a:ln[LRBT][\s\S]*?<\/a:ln[LRBT]>/g, ''))
      .map((properties) => properties.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0]);
    expect(directFillXml.filter(Boolean)).toHaveLength(10);
    expect(directFillXml[0]).toBeUndefined();
    expect(directFillXml[1]).toBeUndefined();
    expect(xml).toContain('<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="112233"><a:alpha val="66667"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="778899"><a:alpha val="101000"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="AABBCC"><a:alpha val="-1000"/></a:srgbClr></a:solidFill>');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(snapshots);

    const conformanceGenerated = new PptxGenJS();
    conformanceGenerated.layout = 'LAYOUT_WIDE';
    conformanceGenerated.addSlide().addTable(
      [[
        { text: 'Opaque', options: { fill: { color: 'FF0000' } } },
        { text: 'Theme alpha', options: {
          fill: { color: conformanceGenerated.SchemeColor.accent1, transparency: 25 },
        } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const conformanceImported = await importPptxGenJS(conformanceGenerated);
    const importedConformanceTable = conformanceImported.slides[0]!.shapes[0] as TableModel;
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(
      [[
        { text: 'Opaque', options: { fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
        } } },
        { text: 'Theme alpha', options: { fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        } } },
      ]],
      {
        x: inches(1),
        y: inches(1),
        width: inches(4),
        height: inches(1),
        columnWidths: [inches(2), inches(2)],
        rowHeights: [inches(1)],
      },
    );
    expect(nativeTable.transform).toMatchObject(importedConformanceTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedConformanceTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedConformanceTable.rowHeights);
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ fill }) => fill))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ fill }) => fill)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    expect(nativeXml).toContain(
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
    );
    expect(nativeXml).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill>',
    );
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedConformance = await PptxDocument.open(await conformanceImported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows).toEqual(nativeTable.rows);
    expect((reopenedConformance.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedConformanceTable.rows);

    const nativeDifferences = PptxDocument.create({ slideSize: 'wide' });
    nativeDifferences.addSlide().addTable([[
      { text: 'Explicit zero', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      } } },
      { text: 'Direct none', options: { fill: { kind: 'none' } } },
    ]]);
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    expect(nativeDifferencesXml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );
    expect(nativeDifferencesXml).toContain('</a:lnB><a:noFill/></a:tcPr>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
    expect(directFillXml[1]).toBeUndefined();
  });

  it('imports PptxGenJS table-cell borders from materialized direct lines', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    const tableTuple = [
      { type: 'none' },
      { type: 'dash', color: '00FF00', pt: 1.5 },
      { type: 'solid', color: '0000FF', pt: 0 },
      { type: 'solid' },
    ] satisfies [BorderProps, BorderProps, BorderProps, BorderProps];
    const partialCellTuple = [
      undefined,
      { type: 'dash' },
      undefined,
      { type: 'solid' },
    ] as unknown as [BorderProps, BorderProps, BorderProps, BorderProps];

    slide.addTable(
      [[{ text: 'Omitted border', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table scalar', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, border: { type: 'solid', color: 'FF0000', pt: 2 } },
    );
    slide.addTable(
      [[{ text: 'Table tuple', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, border: tableTuple },
    );
    slide.addTable(
      [[{
        text: 'Cell scalar zero',
        options: { border: { type: 'solid', color: 'FFFF00', pt: 0 } },
      }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, border: { type: 'solid', color: 'AAAAAA', pt: 3 } },
    );
    slide.addTable(
      [[{ text: 'Cell partial tuple', options: { border: partialCellTuple } }]],
      { x: 0.2, y: 3.4, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Default border values', options: {} }]],
      { x: 0.2, y: 4.2, w: 2, h: 0.5, border: {} },
    );
    slide.addTable(
      [[
        { text: 'Runtime negative', options: { border: { type: 'solid', color: '778899', pt: -1 } } },
        { text: 'Runtime overflow', options: { border: { type: 'dash', color: 'AABBCC', pt: 2000 } } },
      ]],
      { x: 0.2, y: 5, w: 4, h: 0.5 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(7);
    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({ top: value, right: value, bottom: value, left: value });
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders));
    expect(snapshots).toEqual([
      [four(none)],
      [four(line('FF0000', 2))],
      [{
        top: none,
        right: line('00FF00', 1.5, 'dash'),
        bottom: line('0000FF', 1),
        left: line('666666', 1),
      }],
      [four(line('FFFF00', 0))],
      [{
        top: none,
        right: line('666666', 1, 'dash'),
        bottom: none,
        left: line('666666', 1),
      }],
      [four(line('666666', 1))],
      [undefined, undefined],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted border',
      'Table scalar',
      'Table tuple',
      'Cell scalar zero',
      'Cell partial tuple',
      'Default border values',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).toContain(
      '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>',
    );
    expect(xml).toContain('<a:prstDash val="solid"/>');
    expect(xml).toContain('<a:prstDash val="sysDash"/>');
    expect(xml).toContain('<a:lnB w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>');
    expect(xml).toContain('<a:lnL w="-12700" cap="flat" cmpd="sng" algn="ctr">');
    expect(xml).toContain('<a:lnL w="25400000" cap="flat" cmpd="sng" algn="ctr">');
    for (const properties of xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)) {
      const direct = properties[1]!;
      const left = direct.indexOf('<a:lnL ');
      const right = direct.indexOf('<a:lnR ');
      const top = direct.indexOf('<a:lnT ');
      const bottom = direct.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders))).toEqual(snapshots);
  });

  it('matches native table-cell border creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [[
        { text: 'Scalar', options: { border: {
          type: 'solid',
          color: 'FF0000',
          pt: 2,
        } } },
        { text: 'Tuple', options: { border: [
          { type: 'none' },
          { type: 'dash', color: '00FF00', pt: 1.5 },
          { type: 'solid', color: '0000FF', pt: 0 },
          { type: 'solid', color: '666666', pt: 1 },
        ] } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(
      [[
        { text: 'Scalar', options: { border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 2,
          style: 'solid',
        } } },
        { text: 'Tuple', options: { border: [
          { kind: 'none' },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '00FF00' },
            width: 1.5,
            style: 'dash',
          },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '0000FF' },
            width: 0,
            style: 'solid',
          },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '666666' },
            width: 1,
            style: 'solid',
          },
        ] } },
      ]],
      {
        x: inches(1),
        y: inches(1),
        width: inches(4),
        height: inches(1),
        columnWidths: [inches(2), inches(2)],
        rowHeights: [inches(1)],
      },
    );

    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({ top: value, right: value, bottom: value, left: value });
    const expectedBorders = [
      four(line('FF0000', 2)),
      {
        top: none,
        right: line('00FF00', 1.5, 'dash'),
        bottom: line('0000FF', 0),
        left: line('666666', 1),
      },
    ];
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(expectedBorders);
    expect(importedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(expectedBorders);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const tableXml of [nativeXml, importedXml]) {
      for (const properties of tableXml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)) {
        const direct = properties[1]!;
        const left = direct.indexOf('<a:lnL ');
        const right = direct.indexOf('<a:lnR ');
        const top = direct.indexOf('<a:lnT ');
        const bottom = direct.indexOf('<a:lnB ');
        expect(left).toBeGreaterThanOrEqual(0);
        expect(left).toBeLessThan(right);
        expect(right).toBeLessThan(top);
        expect(top).toBeLessThan(bottom);
      }
      expect(tableXml).toContain('<a:lnR w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="sysDash"/>');
      expect(tableXml).toContain('<a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="solid"/>');
    }
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedDifferences = new PptxGenJS();
    generatedDifferences.addSlide().addTable(
      [[
        { text: 'PptxGenJS empty', options: { border: {} } },
        { text: 'PptxGenJS omitted type', options: { border: {
          color: '112233',
          pt: 2,
        } } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTable = importedDifferences.slides[0]!.shapes[0] as TableModel;
    expect(importedDifferenceTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual([
      four(line('666666', 1)),
      four(line('112233', 2)),
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceTable = nativeDifferences.addSlide().addTable([[
      { text: 'Native empty', options: { border: {} } },
      { text: 'Native omitted style', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
      } } },
      { text: 'Native named theme', options: { border: {
        top: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1,
          style: 'dash',
        },
        left: { kind: 'none' },
      } } },
      { text: 'Native zero', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '0000FF' },
        width: 0,
        style: 'solid',
      } } },
    ]]);
    expect(nativeDifferenceTable.rows[0]!.cells[0]!.borders).toEqual(four(none));
    const omittedStyle = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 2,
    };
    expect(nativeDifferenceTable.rows[0]!.cells[1]!.borders).toEqual(four(omittedStyle));
    expect(nativeDifferenceTable.rows[0]!.cells[2]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1,
        style: 'dash',
      },
      right: none,
      bottom: none,
      left: none,
    });
    expect(nativeDifferenceTable.rows[0]!.cells[3]!.borders).toEqual(
      four(line('0000FF', 0)),
    );
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    const cells = [...nativeDifferencesXml.matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)]
      .map((match) => match[1]!);
    expect(cells[0]!.match(/<a:noFill\/>/g)).toHaveLength(4);
    expect(cells[1]).not.toContain('<a:prstDash');
    expect(cells[2]).toContain('<a:schemeClr val="accent1"/>');
    expect(cells[2]).toContain('<a:prstDash val="sysDash"/>');
    expect(cells[3]).toContain('<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr">');
  });

  it('matches native table-level border creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Cell partial', options: { border: [
        { type: 'none' },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 3 },
        { type: 'none' },
      ] } },
      { text: 'Cell none', options: { border: { type: 'none' } } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 1,
      colW: [3, 3, 3, 3],
      rowH: [1],
      border: { type: 'dash', color: '4472C4', pt: 1.5 },
      fill: { color: 'D9EAF7' },
      margin: 0.1,
      valign: 'middle',
    });
    generatedSlide.addTable([[{ text: 'TRBL', options: {} }]], {
      x: 0.5,
      y: 2,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: [
        { type: 'solid', color: 'FF0000', pt: 1 },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 2 },
        { type: 'solid', color: '4472C4', pt: 3 },
      ],
    });
    generatedSlide.addTable([[{ text: 'Scalar zero', options: {} }]], {
      x: 0.5,
      y: 3,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: { type: 'solid', color: 'FF0000', pt: 0 },
    });
    generatedSlide.addTable([[{ text: 'Scalar none', options: {} }]], {
      x: 0.5,
      y: 4,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: { type: 'none' },
    });
    const imported = await importPptxGenJS(generated);
    const importedTables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object', options: {} },
      { text: 'Cell partial', options: { border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'dash',
        },
      } } },
      { text: 'Cell none', options: { border: { kind: 'none' } } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(12),
      height: inches(1),
      columnWidths: inches(3),
      rowHeights: inches(1),
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 1.5,
        style: 'dash',
      },
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
      },
      margin: 7.2,
      valign: 'middle',
    });
    nativeSlide.addTable([['TRBL']], {
      x: inches(0.5),
      y: inches(2),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 1,
          style: 'solid',
        },
        { kind: 'none' },
        {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 2,
          style: 'dash',
        },
        {
          kind: 'line',
          color: { kind: 'srgb', value: '4472C4' },
          width: 3,
          style: 'solid',
        },
      ],
    });
    nativeSlide.addTable([['Scalar zero']], {
      x: inches(0.5),
      y: inches(3),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 0,
        style: 'solid',
      },
    });
    nativeSlide.addTable([['Scalar none']], {
      x: inches(0.5),
      y: inches(4),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: { kind: 'none' },
    });
    const nativeTables = nativeSlide.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({
      top: value,
      right: value,
      bottom: value,
      left: value,
    });
    const inherited = four(line('4472C4', 1.5, 'dash'));
    const expectedBorders = [
      [
        inherited,
        inherited,
        {
          top: none,
          right: none,
          bottom: line('70AD47', 3, 'dash'),
          left: none,
        },
        four(none),
      ],
      [{
        top: line('FF0000', 1),
        right: none,
        bottom: line('70AD47', 2, 'dash'),
        left: line('4472C4', 3),
      }],
      [four(line('FF0000', 0))],
      [four(none)],
    ];
    const borderMatrices = (tables: readonly TableModel[]) => tables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders));
    expect(borderMatrices(nativeTables)).toEqual(expectedBorders);
    expect(borderMatrices(importedTables)).toEqual(expectedBorders);
    expect(nativeTables[0]!.rows[0]!.cells.map(({ text }) => text)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ text }) => text),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(
        ({ verticalAlignment }) => verticalAlignment),
    );
    for (let index = 0; index < nativeTables.length; index += 1) {
      expect(nativeTables[index]!.transform).toMatchObject(
        importedTables[index]!.transform,
      );
      expect(nativeTables[index]!.columnWidths).toEqual(
        importedTables[index]!.columnWidths,
      );
      expect(nativeTables[index]!.rowHeights).toEqual(
        importedTables[index]!.rowHeights,
      );
    }

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const tableXml of [nativeXml, importedXml]) {
      for (const properties of tableXml.matchAll(
        /<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g,
      )) {
        const direct = properties[1]!;
        const left = direct.indexOf('<a:lnL ');
        const right = direct.indexOf('<a:lnR ');
        const top = direct.indexOf('<a:lnT ');
        const bottom = direct.indexOf('<a:lnB ');
        expect(left).toBeGreaterThanOrEqual(0);
        expect(left).toBeLessThan(right);
        expect(right).toBeLessThan(top);
        expect(top).toBeLessThan(bottom);
      }
      expect(tableXml).toContain(
        '<a:lnB w="38100" cap="flat" cmpd="sng" algn="ctr">'
        + '<a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>'
        + '<a:prstDash val="sysDash"/>',
      );
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(borderMatrices(reopenedNative.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    ))).toEqual(expectedBorders);
    expect(borderMatrices(reopenedImported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    ))).toEqual(expectedBorders);

    const generatedDifferences = new PptxGenJS();
    const generatedDifferenceSlide = generatedDifferences.addSlide();
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS empty table', options: {} }]],
      { x: 0.5, y: 0.5, w: 2, h: 0.5, border: {} },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS empty cell', options: { border: {} } }]],
      {
        x: 0.5,
        y: 1.5,
        w: 2,
        h: 0.5,
        border: { type: 'solid', color: 'FF0000', pt: 2 },
      },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS short tuple', options: {} }]],
      {
        x: 0.5,
        y: 2.5,
        w: 2,
        h: 0.5,
        border: [{ type: 'solid', color: '00FF00', pt: 0 }],
      },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS omitted type', options: {} }]],
      { x: 0.5, y: 3.5, w: 2, h: 0.5, border: { color: '112233', pt: 2 } },
    );
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTables = importedDifferences.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(borderMatrices(importedDifferenceTables)).toEqual([
      [four(line('666666', 1))],
      [four(line('666666', 1))],
      [{
        top: line('00FF00', 1),
        right: none,
        bottom: none,
        left: none,
      }],
      [four(line('112233', 2))],
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceSlide = nativeDifferences.addSlide();
    const nativeEmptyTable = nativeDifferenceSlide.addTable(
      [['Native empty table']],
      { border: {} },
    );
    const nativeEmptyCell = nativeDifferenceSlide.addTable(
      [[{ text: 'Native empty cell', options: { border: {} } }]],
      {
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 2,
          style: 'solid',
        },
      },
    );
    const nativeTupleZero = nativeDifferenceSlide.addTable(
      [['Native tuple zero']],
      { border: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 0,
          style: 'solid',
        },
        undefined,
        undefined,
        undefined,
      ] },
    );
    const nativeOmittedStyle = nativeDifferenceSlide.addTable(
      [['Native omitted style']],
      { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
      } },
    );
    expect(nativeEmptyTable.rows[0]!.cells[0]!.borders).toEqual(four(none));
    expect(nativeEmptyCell.rows[0]!.cells[0]!.borders).toEqual(
      four(line('FF0000', 2)),
    );
    expect(nativeTupleZero.rows[0]!.cells[0]!.borders).toEqual({
      top: line('00FF00', 0),
      right: none,
      bottom: none,
      left: none,
    });
    const omittedStyleLine = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 2,
    };
    expect(nativeOmittedStyle.rows[0]!.cells[0]!.borders).toEqual(
      four(omittedStyleLine),
    );
    const nativeDifferenceXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferenceSlide.partUri).bytes,
    );
    const nativeDifferenceCells = [...nativeDifferenceXml.matchAll(
      /<a:tc>([\s\S]*?)<\/a:tc>/g,
    )].map((match) => match[1]!);
    expect(nativeDifferenceCells[0]!.match(/<a:noFill\/>/g)).toHaveLength(4);
    expect(nativeDifferenceCells[1]).toContain('<a:srgbClr val="FF0000"/>');
    expect(nativeDifferenceCells[2]).toContain(
      '<a:lnT w="0" cap="flat" cmpd="sng" algn="ctr">'
      + '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
    );
    expect(nativeDifferenceCells[3]).not.toContain('<a:prstDash');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable(
      [['Invalid short tuple']],
      { border: [{ kind: 'none' }] as never },
    )).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(4);
  });

  it('imports public PptxGenJS output and continues editing in the OOXML kernel', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addText('Created by PptxGenJS', { x: 1, y: 1, w: 7, h: 1, align: 'center' });
    generatedSlide.addText(
      [
        {
          text: 'Bold red',
          options: { bold: true, fontFace: 'Aptos', fontSize: 24, color: 'ff0000' },
        },
        {
          text: 'italic',
          options: { italic: true, fontSize: 14, color: '4472C4', softBreakBefore: true },
        },
      ],
      { x: 1, y: 2, w: 7, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Left', options: { align: 'left' } },
        { text: 'Center', options: { align: 'center' } },
        { text: 'Right', options: { align: 'right' } },
        { text: 'Justify', options: { align: 'justify' } },
      ],
      { x: 1, y: 3, w: 7, h: 2, align: 'left' },
    );
    generatedSlide.addText('Standard\nSecond', { x: 1, y: 5, w: 3, h: 1, bullet: true });
    generatedSlide.addText('Custom', {
      x: 4,
      y: 5,
      w: 3,
      h: 1,
      bullet: { characterCode: '25BA', indent: 18 },
    });
    generatedSlide.addText('Public numberType', {
      x: 7,
      y: 5,
      w: 3,
      h: 1,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3, indent: 22 },
    });
    generatedSlide.addText('Deprecated style', {
      x: 10,
      y: 5,
      w: 2,
      h: 1,
      bullet: { type: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
    });
    generatedSlide.addText('Exact first\nExact second', {
      x: 1,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 28,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 6.25,
      paraSpaceAfter: 8.5,
    });
    generatedSlide.addText('Multiple', {
      x: 4,
      y: 6,
      w: 3,
      h: 1,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 4.25,
      paraSpaceAfter: 7.75,
    });
    generatedSlide.addText('Zero spacing', {
      x: 7,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 0,
      lineSpacingMultiple: 0,
      paraSpaceBefore: 0,
      paraSpaceAfter: 0,
    });
    generatedSlide.addText('Level one', {
      x: 10,
      y: 6,
      w: 2,
      h: 0.5,
      bullet: true,
      indentLevel: 1,
    });
    generatedSlide.addText('Custom level two', {
      x: 10,
      y: 6.5,
      w: 2,
      h: 0.5,
      bullet: { characterCode: '25BA', indent: 18 },
      indentLevel: 2,
    });
    generatedSlide.addText('Number level three', {
      x: 10,
      y: 7,
      w: 2,
      h: 0.5,
      bullet: { type: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      indentLevel: 3,
    });
    generatedSlide.addText('No bullet level two', {
      x: 10,
      y: 7.5,
      w: 2,
      h: 0.5,
      indentLevel: 2,
    });
    generatedSlide.addText('Left\tCenter\tRight\tDecimal', {
      x: 1,
      y: 7.25,
      w: 8,
      h: 0.5,
      tabStops: [
        { position: 1 },
        { position: 2.25, alignment: 'ctr' },
        { position: 3.5, alignment: 'r' },
        { position: 4.75, alignment: 'dec' },
      ],
    });
    generatedSlide.addText('Empty tabs', {
      x: 1,
      y: 7.75,
      w: 3,
      h: 0.5,
      tabStops: [],
    });
    generatedSlide.addText(
      [
        { text: 'First\tA', options: { breakLine: true, tabStops: [{ position: 1.5, alignment: 'r' }] } },
        { text: 'Second\tB', options: { tabStops: [{ position: 2.5, alignment: 'ctr' }] } },
      ],
      { x: 5, y: 7.5, w: 4, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Single', options: { underline: true } },
        { text: ' Double', options: { underline: { style: 'dbl', color: 'ff0000' } } },
        { text: ' Wavy', options: { underline: { style: 'wavyDbl' } } },
        { text: ' None', options: { underline: { style: 'none' } } },
        { text: ' Dot dash', options: { underline: { style: 'dotDashHeavy' } } },
      ],
      { x: 9, y: 3, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'True', options: { strike: true } },
        { text: ' False', options: { strike: false } },
        { text: ' Single', options: { strike: 'sngStrike' } },
        { text: ' Double', options: { strike: 'dblStrike' } },
        { text: ' None', options: { strike: 'noStrike' } },
      ],
      { x: 9, y: 4, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Yellow', options: { highlight: 'ffff00' } },
        { text: ' Theme', options: { highlight: 'accent2' } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 5, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { outline: { color: 'ff0000', size: 1.5 } } },
        { text: ' Theme', options: { outline: { color: 'accent1', size: 2 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 6, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { glow: { color: 'ff0000', size: 8, opacity: 0.5 } } },
        { text: ' Theme', options: { glow: { color: 'accent1', size: 2.5, opacity: 1 } } },
        { text: ' Default', options: { glow: { size: 0, opacity: 0 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 7, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Super', options: { superscript: true } },
        { text: ' Sub', options: { subscript: true } },
        { text: ' Custom+', options: { baseline: 600 } },
        { text: ' Custom-', options: { baseline: -800 } },
        { text: ' Fraction', options: { baseline: 1.5 } },
        { text: ' Zero', options: { baseline: 0 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 8, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Positive', options: { charSpacing: 2.5 } },
        { text: ' Negative', options: { charSpacing: -1.25 } },
        { text: ' Fraction', options: { charSpacing: 0.004 } },
        { text: ' Zero', options: { charSpacing: 0 } },
        { text: ' Combined', options: { charSpacing: 3, baseline: 600 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 9, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Inherited', options: {} },
        { text: ' German', options: { lang: 'de-DE' } },
        { text: ' Explicit default', options: { lang: 'en-US' } },
        { text: ' Empty inherits', options: { lang: '' } },
      ],
      { x: 9, y: 10, w: 3, h: 1, lang: 'fr-CA', objectName: 'Language outer' },
    );
    generatedSlide.addText('Omitted margin', {
      x: 0,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Margin omitted',
    });
    generatedSlide.addText('Zero margin', {
      x: 0,
      y: 0.5,
      w: 2,
      h: 0.5,
      margin: 0,
      objectName: 'Margin zero',
    });
    generatedSlide.addText('Scalar margin', {
      x: 0,
      y: 1,
      w: 2,
      h: 0.5,
      margin: 10,
      objectName: 'Margin scalar',
    });
    generatedSlide.addText('Tuple margin', {
      x: 0,
      y: 1.5,
      w: 2,
      h: 0.5,
      margin: [4, 8, 8, 4],
      objectName: 'Margin tuple',
    });
    generatedSlide.addText('Fractional margin', {
      x: 0,
      y: 2,
      w: 2,
      h: 0.5,
      margin: 0.125,
      objectName: 'Margin fractional',
    });
    generatedSlide.addText('Negative margin', {
      x: 0,
      y: 2.5,
      w: 2,
      h: 0.5,
      margin: -0.5,
      objectName: 'Margin negative',
    });
    generatedSlide.addText('Asymmetric probe', {
      x: 0,
      y: 3,
      w: 2,
      h: 0.5,
      margin: [1, 2, 3, 4],
      objectName: 'Margin asymmetric probe',
    });
    generatedSlide.addText('Omitted vertical alignment', {
      x: 2,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Vertical omitted',
    });
    generatedSlide.addText('Top vertical alignment', {
      x: 2,
      y: 0.5,
      w: 2,
      h: 0.5,
      valign: 'top',
      objectName: 'Vertical top',
    });
    generatedSlide.addText('Middle vertical alignment', {
      x: 2,
      y: 1,
      w: 2,
      h: 0.5,
      valign: 'middle',
      objectName: 'Vertical middle',
    });
    generatedSlide.addText('Bottom vertical alignment', {
      x: 2,
      y: 1.5,
      w: 2,
      h: 0.5,
      valign: 'bottom',
      objectName: 'Vertical bottom',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run vertical alignment', options: { valign: 'bottom' } }],
      {
        x: 2,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Vertical run ignored',
      },
    );
    generatedSlide.addText('Omitted text wrapping', {
      x: 4,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Wrap omitted',
    });
    generatedSlide.addText('Enabled text wrapping', {
      x: 4,
      y: 0.5,
      w: 2,
      h: 0.5,
      wrap: true,
      objectName: 'Wrap true',
    });
    generatedSlide.addText('Disabled text wrapping', {
      x: 4,
      y: 1,
      w: 2,
      h: 0.5,
      wrap: false,
      objectName: 'Wrap false',
    });
    generatedSlide.addText('Invalid text wrapping', {
      x: 4,
      y: 1.5,
      w: 2,
      h: 0.5,
      wrap: 'false',
      objectName: 'Wrap invalid fallback',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text wrapping', options: { wrap: false } }],
      {
        x: 4,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Wrap run ignored',
      },
    );
    const textDirections = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    generatedSlide.addText('Omitted text direction', {
      x: 6,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Direction omitted',
    });
    for (const [index, direction] of textDirections.entries()) {
      generatedSlide.addText(direction, {
        x: 6,
        y: (index + 1) * 0.5,
        w: 2,
        h: 0.5,
        vert: direction,
        objectName: `Direction ${direction}`,
      });
    }
    generatedSlide.addText('Invalid text direction', {
      x: 6,
      y: 4,
      w: 2,
      h: 0.5,
      vert: 'vertical',
      objectName: 'Direction invalid passthrough',
    });
    generatedSlide.addText('Ignored textDirection alias', {
      x: 6,
      y: 4.5,
      w: 2,
      h: 0.5,
      textDirection: 'vert',
      objectName: 'Direction alias ignored',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text direction', options: { vert: 'vert270', textDirection: 'vert' } }],
      {
        x: 6,
        y: 5,
        w: 2,
        h: 0.5,
        objectName: 'Direction run ignored',
      },
    );
    generatedSlide.addText('Omitted text fit', {
      x: 8,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Fit omitted',
    });
    generatedSlide.addText('No text fit', {
      x: 8,
      y: 0.5,
      w: 2,
      h: 0.5,
      fit: 'none',
      objectName: 'Fit none',
    });
    generatedSlide.addText('Shrink text fit', {
      x: 8,
      y: 1,
      w: 2,
      h: 0.5,
      fit: 'shrink',
      objectName: 'Fit shrink',
    });
    generatedSlide.addText('Resize text fit', {
      x: 8,
      y: 1.5,
      w: 2,
      h: 0.5,
      fit: 'resize',
      objectName: 'Fit resize',
    });
    generatedSlide.addText('Invalid text fit', {
      x: 8,
      y: 2,
      w: 2,
      h: 0.5,
      fit: 'SHRINK',
      objectName: 'Fit invalid ignored',
    });
    generatedSlide.addText('Legacy shrink text fit', {
      x: 8,
      y: 2.5,
      w: 2,
      h: 0.5,
      shrinkText: true,
      objectName: 'Fit legacy shrink',
    });
    generatedSlide.addText('Legacy resize text fit', {
      x: 8,
      y: 3,
      w: 2,
      h: 0.5,
      autoFit: true,
      objectName: 'Fit legacy resize',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text fit', options: { fit: 'shrink', shrinkText: true, autoFit: true } }],
      {
        x: 8,
        y: 3.5,
        w: 2,
        h: 0.5,
        objectName: 'Fit run ignored',
      },
    );
    generatedSlide.addText('مرحبا\nالعالم', {
      x: 0,
      y: 4,
      w: 2,
      h: 1,
      objectName: 'RTL true',
      rtlMode: true,
    });
    generatedSlide.addText('Explicit false', {
      x: 0,
      y: 5,
      w: 2,
      h: 0.5,
      objectName: 'RTL false',
      rtlMode: false,
    });
    generatedSlide.addText('Omitted', {
      x: 0,
      y: 5.5,
      w: 2,
      h: 0.5,
      objectName: 'RTL omitted',
    });
    generatedSlide.addText(
      [
        { text: 'Run one', options: { rtlMode: true } },
        { text: ' Run two', options: { rtlMode: true } },
      ],
      {
        x: 0,
        y: 6,
        w: 2,
        h: 0.5,
        objectName: 'RTL run probe',
      },
    );
    const document = await importPptxGenJS(generated);
    expect(document.slides[0]?.title.text).toBe('Created by PptxGenJS');
    expect((document.slides[0]!.shapes[0] as ShapeModel).richText[0]!.align).toBe('center');
    const rich = document.slides[0]!.shapes[1] as ShapeModel;
    expect(rich.text).toBe('Bold red\nitalic');
    expect(rich.richText[0]!.runs).toEqual([
      {
        text: 'Bold red',
        style: {
          fontFamily: 'Aptos',
          fontSize: 24,
          lang: 'en-US',
          bold: true,
          color: { kind: 'srgb', value: 'FF0000' },
        },
      },
      {
        text: 'italic',
        softBreakBefore: true,
        style: {
          fontSize: 14,
          lang: 'en-US',
          italic: true,
          color: { kind: 'srgb', value: '4472C4' },
        },
      },
    ]);
    const aligned = document.slides[0]!.shapes[2] as ShapeModel;
    expect(aligned.richText.map(({ align }) => align)).toEqual(['left', 'center', 'right', 'justify']);
    aligned.richText = aligned.richText.map((paragraph, index) => ({
      runs: paragraph.runs,
      ...(index === 3
        ? { align: 'center' as const }
        : paragraph.align
          ? { align: paragraph.align }
          : {}),
    }));
    expect((document.slides[0]!.shapes[3] as ShapeModel).richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect((document.slides[0]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((document.slides[0]!.shapes[5] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 3,
      indent: 22,
    });
    expect((document.slides[0]!.shapes[6] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
      indent: 24,
    });
    expect((document.slides[0]!.shapes[7] as ShapeModel).richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    ]);
    expect((document.slides[0]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((document.slides[0]!.shapes[9] as ShapeModel).richText[0]!.spacing).toBeUndefined();
    expect((document.slides[0]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '•', indent: 27 },
      level: 1,
    });
    expect((document.slides[0]!.shapes[11] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '►', indent: 18 },
      level: 2,
    });
    expect((document.slides[0]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      level: 3,
    });
    const noBulletLevel = (document.slides[0]!.shapes[13] as ShapeModel).richText[0]!;
    expect(noBulletLevel.level).toBe(2);
    expect(noBulletLevel.bullet).toBeUndefined();
    expect((document.slides[0]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((document.slides[0]!.shapes[15] as ShapeModel).richText[0]!.tabStops).toEqual([]);
    expect((document.slides[0]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((document.slides[0]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((document.slides[0]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((document.slides[0]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const spaced = (document.slides[0]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(spaced.map(({ style }) => style?.characterSpacing)).toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(spaced[4]!.style!.baseline).toBe('superscript');
    const languages = (document.slides[0]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(languages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const shapeByName = (name: string): ShapeModel => {
      const shape = document.slides[0]!.shapes.find((candidate) => candidate.name === name);
      expect(shape).toBeInstanceOf(ShapeModel);
      return shape as ShapeModel;
    };
    expect(shapeByName('RTL true').richText.map(({ rtl }) => rtl)).toEqual([true, true]);
    expect(shapeByName('RTL false').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    expect(shapeByName('RTL omitted').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    const rtlRunProbe = shapeByName('RTL run probe');
    expect(rtlRunProbe.text).toBe('Run one Run two');
    expect(rtlRunProbe.richText[0]!.runs.map(({ style }) =>
      (style as Record<string, unknown> | undefined)?.rtlMode)).toEqual([undefined, undefined]);
    expect(shapeByName('Margin omitted').textMargins).toBeUndefined();
    expect(shapeByName('Margin zero').textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(shapeByName('Margin scalar').textMargins).toEqual({
      top: 10,
      right: 10,
      bottom: 10,
      left: 10,
    });
    expect(shapeByName('Margin tuple').textMargins).toEqual({ top: 4, right: 8, bottom: 8, left: 4 });
    expect(shapeByName('Margin fractional').textMargins).toEqual({
      top: 1_588 / 12_700,
      right: 1_588 / 12_700,
      bottom: 1_588 / 12_700,
      left: 1_588 / 12_700,
    });
    expect(shapeByName('Margin negative').textMargins).toEqual({
      top: -0.5,
      right: -0.5,
      bottom: -0.5,
      left: -0.5,
    });
    expect(shapeByName('Margin asymmetric probe').textMargins).toEqual({
      top: 4,
      right: 2,
      bottom: 3,
      left: 1,
    });
    expect([
      'Vertical omitted',
      'Vertical top',
      'Vertical middle',
      'Vertical bottom',
      'Vertical run ignored',
    ].map((name) => shapeByName(name).verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
      'middle',
    ]);
    expect([
      'Wrap omitted',
      'Wrap true',
      'Wrap false',
      'Wrap invalid fallback',
      'Wrap run ignored',
    ].map((name) => shapeByName(name).textWrap)).toEqual([
      true,
      true,
      false,
      true,
      true,
    ]);
    expect([
      'Direction omitted',
      ...textDirections.map((direction) => `Direction ${direction}`),
      'Direction invalid passthrough',
      'Direction alias ignored',
      'Direction run ignored',
    ].map((name) => shapeByName(name).textDirection)).toEqual([
      undefined,
      ...textDirections,
      undefined,
      undefined,
      undefined,
    ]);
    expect([
      'Fit omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Fit invalid ignored',
      'Fit legacy shrink',
      'Fit legacy resize',
      'Fit run ignored',
    ].map((name) => shapeByName(name).textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
      undefined,
      'shrink',
      'resize',
      undefined,
    ]);
    const importedXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(importedXml).toContain('lIns="12700" tIns="50800" rIns="25400" bIns="38100"');
    expect(importedXml).toMatch(/name="Vertical omitted"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical top"[\s\S]*?<a:bodyPr[^>]*anchor="t"/);
    expect(importedXml).toMatch(/name="Vertical middle"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical bottom"[\s\S]*?<a:bodyPr[^>]*anchor="b"/);
    expect(importedXml).toMatch(/name="Vertical run ignored"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Wrap omitted"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap true"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap false"[\s\S]*?<a:bodyPr[^>]*wrap="none"/);
    expect(importedXml).toMatch(/name="Wrap invalid fallback"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap run ignored"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(
      /name="Direction omitted"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    for (const direction of textDirections) {
      expect(importedXml).toMatch(
        new RegExp(`name="Direction ${direction}"[\\s\\S]*?<a:bodyPr[^>]* vert="${direction}"`),
      );
    }
    expect(importedXml).toMatch(
      /name="Direction invalid passthrough"[\s\S]*?<a:bodyPr[^>]* vert="vertical"/,
    );
    expect(importedXml).toMatch(
      /name="Direction alias ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Direction run ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Fit omitted"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit none"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit invalid ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit run ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="fr-CA" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="de-DE" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:endParaRPr lang="fr-CA" dirty="0"\/>/,
    );
    expect(importedXml).toMatch(
      /name="RTL true"[\s\S]*?<a:p><a:pPr rtl="1"[^>]*>[\s\S]*?<a:p><a:pPr rtl="1"/,
    );
    expect(importedXml).toMatch(
      /name="RTL false"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="RTL omitted"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    const rtlRunStart = importedXml.indexOf('name="RTL run probe"');
    const rtlRunEnd = importedXml.indexOf('</p:sp>', rtlRunStart);
    expect(importedXml.slice(rtlRunStart, rtlRunEnd).match(/<a:pPr rtl="1"/g)).toHaveLength(2);
    document.slides[0]!.title.text = 'Edited by the OOXML kernel';
    document.duplicateSlide(0);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ title }) => title.text)).toEqual([
      'Edited by the OOXML kernel',
      'Edited by the OOXML kernel',
    ]);
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).richText[0]!.runs[1]).toMatchObject({
      text: 'italic',
      softBreakBefore: true,
      style: { italic: true, color: { kind: 'srgb', value: '4472C4' } },
    });
    expect((reopened.slides[1]!.shapes[2] as ShapeModel).richText.map(({ align }) => align)).toEqual([
      'left',
      'center',
      'right',
      'center',
    ]);
    expect((reopened.slides[1]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((reopened.slides[1]!.shapes[6] as ShapeModel).richText[0]!.bullet).toMatchObject({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
    });
    expect((reopened.slides[1]!.shapes[7] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 6.25,
      after: 8.5,
      line: { kind: 'exact', points: 28 },
    });
    expect((reopened.slides[1]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((reopened.slides[1]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', indent: 27 },
      level: 1,
    });
    expect((reopened.slides[1]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', indent: 22 },
      level: 3,
    });
    expect((reopened.slides[1]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((reopened.slides[1]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((reopened.slides[1]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((reopened.slides[1]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((reopened.slides[1]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const reopenedSpaced = (reopened.slides[1]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(reopenedSpaced.map(({ style }) => style?.characterSpacing))
      .toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(reopenedSpaced[4]!.style!.baseline).toBe('superscript');
    const reopenedLanguages = (reopened.slides[1]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(reopenedLanguages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const reopenedMargins = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Margin '))
      .map(({ name, textMargins }) => [name, textMargins]);
    expect(reopenedMargins).toEqual([
      ['Margin omitted', undefined],
      ['Margin zero', { top: 0, right: 0, bottom: 0, left: 0 }],
      ['Margin scalar', { top: 10, right: 10, bottom: 10, left: 10 }],
      ['Margin tuple', { top: 4, right: 8, bottom: 8, left: 4 }],
      ['Margin fractional', {
        top: 1_588 / 12_700,
        right: 1_588 / 12_700,
        bottom: 1_588 / 12_700,
        left: 1_588 / 12_700,
      }],
      ['Margin negative', { top: -0.5, right: -0.5, bottom: -0.5, left: -0.5 }],
      ['Margin asymmetric probe', { top: 4, right: 2, bottom: 3, left: 1 }],
    ]);
    const reopenedVerticalAlignment = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Vertical '))
      .map(({ name, verticalAlignment }) => [name, verticalAlignment]);
    expect(reopenedVerticalAlignment).toEqual([
      ['Vertical omitted', 'middle'],
      ['Vertical top', 'top'],
      ['Vertical middle', 'middle'],
      ['Vertical bottom', 'bottom'],
      ['Vertical run ignored', 'middle'],
    ]);
    const reopenedWrapping = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Wrap '))
      .map(({ name, textWrap }) => [name, textWrap]);
    expect(reopenedWrapping).toEqual([
      ['Wrap omitted', true],
      ['Wrap true', true],
      ['Wrap false', false],
      ['Wrap invalid fallback', true],
      ['Wrap run ignored', true],
    ]);
    const reopenedDirections = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Direction '))
      .map(({ name, textDirection }) => [name, textDirection]);
    expect(reopenedDirections).toEqual([
      ['Direction omitted', undefined],
      ...textDirections.map((direction) => [`Direction ${direction}`, direction]),
      ['Direction invalid passthrough', undefined],
      ['Direction alias ignored', undefined],
      ['Direction run ignored', undefined],
    ]);
    const reopenedFit = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Fit '))
      .map(({ name, textFit }) => [name, textFit]);
    expect(reopenedFit).toEqual([
      ['Fit omitted', undefined],
      ['Fit none', undefined],
      ['Fit shrink', 'shrink'],
      ['Fit resize', 'resize'],
      ['Fit invalid ignored', undefined],
      ['Fit legacy shrink', 'shrink'],
      ['Fit legacy resize', 'resize'],
      ['Fit run ignored', undefined],
    ]);
    const reopenedRtl = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('RTL '))
      .map(({ name, richText }) => [name, richText.map(({ rtl }) => rtl)]);
    expect(reopenedRtl).toEqual([
      ['RTL true', [true, true]],
      ['RTL false', [undefined]],
      ['RTL omitted', [undefined]],
      ['RTL run probe', [true]],
    ]);
  }, 180_000);

  it('imports and reopens PptxGenJS rich text transparency from real output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const slide = generated.addSlide();
    slide.addText(
      [
        { text: 'Omitted', options: { color: 'FF0000' } },
        { text: ' Zero', options: { color: '00FF00', transparency: 0 } },
        { text: ' Quarter', options: { color: '0000FF', transparency: 25 } },
        { text: ' Fractional', options: { color: '112233', transparency: 50.5555 } },
        { text: ' Invisible', options: { color: '445566', transparency: 100 } },
        { text: ' Theme', options: { color: 'accent1', transparency: 40 } },
        { text: ' Default', options: { transparency: 60 } },
      ],
      { x: 1, y: 1, w: 10, h: 1, objectName: 'Transparency probe' },
    );

    const document = await importPptxGenJS(generated);
    const shape = document.slides[0]!.shapes.find(({ name }) => name === 'Transparency probe');
    expect(shape).toBeInstanceOf(ShapeModel);
    const runs = (shape as ShapeModel).richText[0]!.runs;
    expect(runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
    expect(runs.map(({ style }) => style?.color)).toEqual([
      { kind: 'srgb', value: 'FF0000' },
      { kind: 'srgb', value: '00FF00' },
      { kind: 'srgb', value: '0000FF' },
      { kind: 'srgb', value: '112233' },
      { kind: 'srgb', value: '445566' },
      { kind: 'scheme', value: 'accent1' },
      { kind: 'srgb', value: '000000' },
    ]);
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/<a:alpha val="\d+"\/>/g)).toEqual([
      '<a:alpha val="75000"/>',
      '<a:alpha val="49445"/>',
      '<a:alpha val="0"/>',
      '<a:alpha val="60000"/>',
      '<a:alpha val="40000"/>',
    ]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Transparency probe',
    ) as ShapeModel;
    expect(reopenedShape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
  }, 30_000);

  it('imports only direct PptxGenJS presentation RTL and reopens it', async () => {
    const cases: readonly [string, boolean, unknown, boolean | undefined][] = [
      ['omitted', false, undefined, undefined],
      ['true', true, true, true],
      ['false', true, false, undefined],
      ['truthy', true, 'yes', true],
    ];
    for (const [name, assign, value, expected] of cases) {
      const generated = new PptxGenJS();
      if (assign) generated.rtlMode = value;
      generated.addSlide();
      const document = await importPptxGenJS(generated);
      const journal = [...document.opcPackage.mutations];

      expect(document.rtlMode, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);
      const presentationXml = new TextDecoder().decode(
        document.opcPackage.requirePart(document.presentationPartUri).bytes,
      );
      if (expected === true) {
        expect(presentationXml, name).toMatch(/<p:presentation\b[^>]*\srtl="1"/);
      } else {
        expect(presentationXml, name).not.toMatch(/<p:presentation\b[^>]*\srtl=/);
      }
      expect(presentationXml, name).toMatch(/<a:lvl1pPr\b[^>]*\srtl="0"/);

      if (name === 'true') {
        const reopened = await PptxDocument.open(await document.write());
        expect(reopened.rtlMode).toBe(true);
      }
    }
  }, 20_000);

  it('imports and matches public PptxGenJS hidden slide output', async () => {
    const cases: readonly (readonly [string, boolean, unknown, boolean])[] = [
      ['omitted', false, undefined, false],
      ['false', true, false, false],
      ['true', true, true, true],
      ['truthy invalid', true, 'yes', true],
    ];
    const generated = cases.map(([, assign, value]) => {
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      if (assign) slide.hidden = value;
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(4).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    expect(imported.map(({ slides }) => slides[0]?.hidden)).toEqual(
      cases.map(([, , , expected]) => expected),
    );
    for (const [index, document] of imported.entries()) {
      const [name, , , expected] = cases[index]!;
      const journal = [...document.opcPackage.mutations];
      const slide = document.slides[0]!;
      const slideXml = new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      expect(document.slides[0]?.hidden, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);
      if (expected) {
        expect(slideXml, name).toMatch(/<p:sld\b[^>]*\sshow="0"/);
      } else {
        expect(slideXml, name).not.toMatch(/<p:sld\b[^>]*\sshow=/);
      }
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.slides[0]?.hidden, name).toBe(expected);
    }

    const native = PptxDocument.create();
    const nativeVisible = native.addSlide();
    const nativeHidden = native.addSlide();
    nativeHidden.hidden = true;
    expect([nativeVisible.hidden, nativeHidden.hidden]).toEqual([
      imported[1]!.slides[0]!.hidden,
      imported[2]!.slides[0]!.hidden,
    ]);
    const nativeXml = native.slides.map(({ partUri }) =>
      new TextDecoder().decode(native.opcPackage.requirePart(partUri).bytes));
    expect(nativeXml[0]).not.toMatch(/<p:sld\b[^>]*\sshow=/);
    expect(nativeXml[1]).toMatch(/<p:sld\b[^>]*\sshow="0"/);
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.slides.map(({ hidden }) => hidden)).toEqual([false, true]);

    const beforeInvalid = native.opcPackage.requirePart(nativeVisible.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => {
      (nativeVisible as unknown as { hidden: unknown }).hidden = 'yes';
    }).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeVisible.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  }, 20_000);

  it('imports and matches public PptxGenJS speaker notes output', async () => {
    const cases = [
      { name: 'omitted', input: undefined, expected: '' },
      { name: 'empty', input: '', expected: '' },
      { name: 'plain', input: 'Speaker & <notes>', expected: 'Speaker & <notes>' },
      {
        name: 'multiline',
        input: 'Line 1\nLine 2\r\nLine 3',
        expected: 'Line 1\nLine 2\nLine 3',
      },
    ] as const;
    const generated = cases.map(({ input }) => {
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      if (input !== undefined) slide.addNotes(input);
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(4).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    expect(imported.map(({ slides }) => slides[0]?.notes)).toEqual(
      cases.map(({ expected }) => expected),
    );
    for (const [index, document] of imported.entries()) {
      const { name, expected } = cases[index]!;
      const slide = document.slides[0]!;
      const journal = [...document.opcPackage.mutations];
      expect(slide.notes, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);

      const notesRelationships = slide.relationships.filter(
        ({ type }) => type.endsWith('/notesSlide'),
      );
      expect(notesRelationships, name).toHaveLength(1);
      const notesUri = notesRelationships[0]!.resolvedTarget!;
      expect(document.opcPackage.requirePart(notesUri).contentType, name).toBe(
        'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
      );
      const slideBacklinks = document.opcPackage.relationships(notesUri).filter(
        ({ type }) => type.endsWith('/slide'),
      );
      const masterRelationships = document.opcPackage.relationships(notesUri).filter(
        ({ type }) => type.endsWith('/notesMaster'),
      );
      expect(slideBacklinks, name).toHaveLength(1);
      expect(slideBacklinks[0]!.resolvedTarget, name).toBe(slide.partUri);
      expect(masterRelationships, name).toHaveLength(1);
      const masterUri = masterRelationships[0]!.resolvedTarget!;
      expect(document.opcPackage.requirePart(masterUri).contentType, name).toBe(
        'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
      );
      const presentationMasterRelationships = document.opcPackage
        .relationships(document.presentationPartUri)
        .filter(({ type }) => type.endsWith('/notesMaster'));
      expect(presentationMasterRelationships, name).toHaveLength(1);
      expect(presentationMasterRelationships[0]!.resolvedTarget, name).toBe(masterUri);

      const notesXml = new TextDecoder().decode(
        document.opcPackage.requirePart(notesUri).bytes,
      );
      expect(notesXml.match(/<p:ph\b[^>]*\btype="body"/g), name).toHaveLength(1);
      expect(notesXml.match(/<p:txBody>/g), name).toHaveLength(2);
      if (name === 'plain') {
        expect(notesXml).toContain('<a:t>Speaker &amp; &lt;notes&gt;</a:t>');
      }

      const reopened = await PptxDocument.open(await document.write());
      expect(document.diagnostics.filter(({ severity }) => severity === 'error'), name)
        .toEqual([]);
      expect(reopened.slides[0]?.notes, name).toBe(expected);
    }

    const importedPlain = imported[2]!;
    const importedSource = importedPlain.slides[0]!;
    const importedSourceNotesUri = importedSource.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    const importedDuplicate = importedPlain.duplicateSlide(0);
    const importedDuplicateNotesUri = importedDuplicate.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    expect(importedDuplicate.notes).toBe('Speaker & <notes>');
    expect(importedDuplicateNotesUri).not.toBe(importedSourceNotesUri);
    expect(importedPlain.opcPackage.relationships(importedDuplicateNotesUri).find(
      ({ type }) => type.endsWith('/slide'),
    )?.resolvedTarget).toBe(importedDuplicate.partUri);
    importedDuplicate.notes = 'Edited duplicate';
    expect([importedSource.notes, importedDuplicate.notes]).toEqual([
      'Speaker & <notes>',
      'Edited duplicate',
    ]);
    const reopenedDuplicate = await PptxDocument.open(await importedPlain.write());
    expect(reopenedDuplicate.slides.map(({ notes }) => notes)).toEqual([
      'Speaker & <notes>',
      'Edited duplicate',
    ]);

    const native = [
      PptxDocument.create(),
      PptxDocument.create(),
      PptxDocument.create(),
      PptxDocument.create(),
    ];
    native[0]!.addSlide();
    native[1]!.addSlide().notes = '';
    native[2]!.addSlide().addNotes('Speaker & <notes>');
    native[3]!.addSlide().addNotes('Line 1\nLine 2\r\nLine 3');
    expect(native.map(({ slides }) => slides[0]?.notes)).toEqual([
      undefined,
      '',
      'Speaker & <notes>',
      'Line 1\nLine 2\nLine 3',
    ]);
    expect(native.slice(1).map(({ slides }) => slides[0]?.notes)).toEqual(
      imported.slice(1).map(({ slides }) => slides[0]?.notes),
    );
    const nativePlain = native[2]!.slides[0]!;
    const nativeNotesUri = nativePlain.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    expect(new TextDecoder().decode(
      native[2]!.opcPackage.requirePart(nativeNotesUri).bytes,
    )).toContain('<a:t xml:space="preserve">Speaker &amp; &lt;notes&gt;</a:t>');

    const invalidDocument = PptxDocument.create();
    const invalidSlide = invalidDocument.addSlide();
    const beforeParts = invalidDocument.opcPackage.parts.map(
      ({ uri, bytes }) => ({ uri, bytes: bytes.slice() }),
    );
    const beforeJournal = [...invalidDocument.opcPackage.mutations];
    for (const invalid of [7, {}, 'A\u0001B']) {
      expect(() => {
        (invalidSlide as unknown as { notes: unknown }).notes = invalid;
      }, String(invalid)).toThrow(TypeError);
      expect(() => invalidSlide.addNotes(invalid as never), String(invalid)).toThrow(TypeError);
    }
    expect(invalidDocument.opcPackage.parts.map(({ uri, bytes }) => ({ uri, bytes })))
      .toEqual(beforeParts);
    expect(invalidDocument.opcPackage.mutations).toEqual(beforeJournal);
  }, 20_000);

  it('imports and matches public PptxGenJS presentation sections', async () => {
    const none = new PptxGenJS();
    none.addSlide();
    none.addSlide();

    const explicit = new PptxGenJS();
    explicit.addSection({ title: 'Data & <One>' });
    explicit.addSlide({ sectionTitle: 'Data & <One>' });
    explicit.addSlide({ sectionTitle: 'Data & <One>' });

    const empty = new PptxGenJS();
    empty.addSlide();
    empty.addSection({ title: 'Empty' });

    const ordered = new PptxGenJS();
    ordered.addSection({ title: 'A' });
    ordered.addSection({ title: 'C' });
    ordered.addSection({ title: 'B', order: 1 });

    const defaults = new PptxGenJS();
    defaults.addSection({ title: 'Intro' });
    defaults.addSlide();
    defaults.addSlide();

    const looseBefore = new PptxGenJS();
    looseBefore.addSlide();
    looseBefore.addSection({ title: 'Later' });
    looseBefore.addSlide({ sectionTitle: 'Later' });

    const unknown = new PptxGenJS();
    unknown.addSection({ title: 'Known' });
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      unknown.addSlide({ sectionTitle: 'Missing' });
      expect(warning).toHaveBeenCalledTimes(1);
    } finally {
      warning.mockRestore();
    }

    const orderZero = new PptxGenJS();
    orderZero.addSection({ title: 'A' });
    orderZero.addSection({ title: 'Zero', order: 0 });

    const generated = [
      none,
      explicit,
      empty,
      ordered,
      defaults,
      looseBefore,
      unknown,
      orderZero,
    ];
    expect(generated.map(({ version }) => version)).toEqual(Array(8).fill('4.0.1'));
    const imported: PptxDocument[] = [];
    for (const presentation of generated) imported.push(await importPptxGenJS(presentation));

    expect(imported.map(sectionState)).toEqual([
      [],
      [{ title: 'Data & <One>', slideIds: [256, 257] }],
      [{ title: 'Empty', slideIds: [] }],
      [
        { title: 'A', slideIds: [] },
        { title: 'B', slideIds: [] },
        { title: 'C', slideIds: [] },
      ],
      [
        { title: 'Intro', slideIds: [] },
        { title: 'Default-1', slideIds: [256, 257] },
      ],
      [{ title: 'Later', slideIds: [257] }],
      [{ title: 'Known', slideIds: [] }],
      [
        { title: 'A', slideIds: [] },
        { title: 'Zero', slideIds: [] },
      ],
    ]);
    expect(imported.map(({ slides }) => slides.length)).toEqual([2, 2, 1, 0, 2, 2, 1, 0]);
    for (const document of imported) {
      for (const { id } of document.sections ?? []) {
        expect(id).toMatch(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/i);
      }
      const reopened = await PptxDocument.open(await document.write());
      expect(sectionState(reopened)).toEqual(sectionState(document));
    }

    const nativeExplicit = PptxDocument.create();
    nativeExplicit.addSection({ title: 'Data & <One>' });
    nativeExplicit.addSlide({ sectionTitle: 'Data & <One>' });
    nativeExplicit.addSlide({ sectionTitle: 'Data & <One>' });
    expect(sectionState(nativeExplicit)).toEqual(sectionState(imported[1]!));

    const nativeOrdered = PptxDocument.create();
    nativeOrdered.addSection({ title: 'A' });
    nativeOrdered.addSection({ title: 'C' });
    nativeOrdered.addSection({ title: 'B', order: 1 });
    expect(sectionState(nativeOrdered)).toEqual(sectionState(imported[3]!));

    const nativeDefaults = PptxDocument.create();
    nativeDefaults.addSection({ title: 'Intro' });
    nativeDefaults.addSlide();
    nativeDefaults.addSlide();
    expect(sectionState(nativeDefaults)).toEqual(sectionState(imported[4]!));

    const nativeLooseBefore = PptxDocument.create();
    nativeLooseBefore.addSlide();
    nativeLooseBefore.addSection({ title: 'Later' });
    nativeLooseBefore.addSlide({ sectionTitle: 'Later' });
    expect(sectionState(nativeLooseBefore)).toEqual(sectionState(imported[5]!));

    const nativeUnknown = PptxDocument.create();
    nativeUnknown.addSection({ title: 'Known' });
    expect(() => nativeUnknown.addSlide({ sectionTitle: 'Missing' })).toThrow(RangeError);
    expect(nativeUnknown.slides).toHaveLength(0);
    expect(imported[6]!.slides).toHaveLength(1);

    const nativeOrderZero = PptxDocument.create();
    nativeOrderZero.addSection({ title: 'A' });
    nativeOrderZero.addSection({ title: 'Zero', order: 0 });
    expect(sectionState(nativeOrderZero)?.map(({ title }) => title)).toEqual(['Zero', 'A']);
    expect(sectionState(imported[7]!)?.map(({ title }) => title)).toEqual(['A', 'Zero']);
  }, 30_000);

  it('matches public PptxGenJS presentation theme fonts and reopens a partial edit', async () => {
    const cases = [
      {
        name: 'default',
        input: undefined,
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Calibri' },
      },
      {
        name: 'empty',
        input: {},
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Calibri' },
      },
      {
        name: 'head only',
        input: { headFontFace: 'Aptos Display' },
        expected: { headFontFace: 'Aptos Display', bodyFontFace: 'Calibri' },
      },
      {
        name: 'body only',
        input: { bodyFontFace: 'Aptos' },
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Aptos' },
      },
      {
        name: 'custom',
        input: { headFontFace: 'Noto Sans Display', bodyFontFace: 'Noto Sans' },
        expected: { headFontFace: 'Noto Sans Display', bodyFontFace: 'Noto Sans' },
      },
    ] as const;
    const generated = cases.map(({ input }) => {
      const presentation = new PptxGenJS();
      if (input !== undefined) presentation.theme = input;
      presentation.addSlide();
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(5).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.theme, cases[index]!.name).toEqual(cases[index]!.expected);
      expect(document.opcPackage.mutations, cases[index]!.name).toEqual(journal);

      const native = PptxDocument.create({ theme: cases[index]!.input ?? {} });
      expect(native.theme, cases[index]!.name).toEqual(document.theme);
    }

    const edited = imported[4]!;
    const theme = edited.masterLayoutTheme.presentationTheme!;
    const untouchedParts = new Map(
      edited.opcPackage.parts
        .filter(({ uri }) => uri !== theme.partUri)
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    theme.setFonts({ minorLatin: 'Noto Sans Edited' });
    expect(edited.theme).toEqual({
      headFontFace: 'Noto Sans Display',
      bodyFontFace: 'Noto Sans Edited',
    });
    for (const [uri, bytes] of untouchedParts) {
      expect(edited.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopened = await PptxDocument.open(await edited.write());
    expect(reopened.theme).toEqual(edited.theme);
  }, 30_000);

  it('imports, edits, and reopens PptxGenJS presentation created-at metadata from public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide();
    const bytes = await generated.write({ outputType: 'uint8array', compression: true });
    const imported = await PptxDocument.open(bytes);
    const journal = [...imported.opcPackage.mutations];

    expect(imported.createdAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.opcPackage.mutations).toEqual(journal);
    const createdAt = imported.createdAt!;
    const coreBefore = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreBefore).toContain(
      'xmlns:dcterms="http://purl.org/dc/terms/"',
    );
    expect(coreBefore).toContain(
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
    );
    expect(coreBefore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">([^<]+)<\/dcterms:created>/,
    )?.[1]).toBe(createdAt);
    const modifiedBefore = coreBefore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0];
    expect(modifiedBefore).toMatch(
      /^<dcterms:modified xsi:type="dcterms:W3CDTF">[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z<\/dcterms:modified>$/,
    );
    const nonCoreParts = new Map(
      imported.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes: partBytes }) => [uri, partBytes]),
    );

    const native = PptxDocument.create({ createdAt });
    expect(native.createdAt).toBe(createdAt);
    expect(PptxDocument.create().createdAt).toBeUndefined();

    imported.createdAt = '2024-02-29T12:34:56.123+05:30';
    expect(imported.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
    const coreAfter = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreAfter).not.toBe(coreBefore);
    expect(coreAfter.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0]).toBe(modifiedBefore);
    for (const [uri, partBytes] of nonCoreParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(partBytes);
    }

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
    const reopenedCore = new TextDecoder().decode(
      reopened.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(reopenedCore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0]).toBe(modifiedBefore);
  }, 20_000);

  it('imports, edits, and reopens PptxGenJS presentation modified-at metadata from public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide().addText('PptxGenJS modified-at', {
      x: 1,
      y: 1,
      w: 5,
      h: 1,
    });
    const imported = await importPptxGenJS(generated);
    const journal = [...imported.opcPackage.mutations];

    expect(imported.modifiedAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.createdAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.opcPackage.mutations).toEqual(journal);
    const modifiedAt = imported.modifiedAt!;
    const createdAt = imported.createdAt!;
    const coreBefore = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreBefore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">([^<]+)<\/dcterms:modified>/,
    )?.[1]).toBe(modifiedAt);
    const createdBefore = coreBefore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0];
    expect(createdBefore).toBe(
      `<dcterms:created xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:created>`,
    );
    const nonCoreParts = new Map(
      imported.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );

    const native = PptxDocument.create({ modifiedAt });
    expect(native.modifiedAt).toBe(modifiedAt);
    expect(native.createdAt).toBeUndefined();
    expect(PptxDocument.create().modifiedAt).toBeUndefined();

    imported.modifiedAt = '2024-03-01T01:02:03.456+08:00';
    expect(imported.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
    expect(imported.createdAt).toBe(createdAt);
    const coreAfter = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreAfter).not.toBe(coreBefore);
    expect(coreAfter).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-03-01T01:02:03.456+08:00</dcterms:modified>',
    );
    expect(coreAfter.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0]).toBe(createdBefore);
    for (const [uri, bytes] of nonCoreParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
    expect(reopened.createdAt).toBe(createdAt);
    expect(reopened.slides[0]?.title.text).toBe('PptxGenJS modified-at');
    const reopenedCore = new TextDecoder().decode(
      reopened.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(reopenedCore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0]).toBe(createdBefore);
  }, 20_000);

  it('imports and reopens PptxGenJS presentation title metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.title = 'Quarterly & <Review>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.title = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedTitles = [
      'PptxGenJS Presentation',
      'Quarterly & <Review>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ title }) => title)).toEqual(expectedTitles);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.title).toBe(expectedTitles[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain(
      '<dc:title>PptxGenJS Presentation</dc:title>',
    );
    expect(coreXml[1]).toContain(
      '<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>',
    );
    expect(coreXml[2]).toContain('<dc:title></dc:title>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ title }) => title)).toEqual(expectedTitles);

    const native = PptxDocument.create({ title: 'Quarterly & <Review>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.title).toBe(imported[1]!.title);
    expect(nativeOmitted.title).toBeUndefined();
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.title).toBe('Quarterly & <Review>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation author metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.author = 'Alice & <Bob>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.author = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedAuthors = [
      'PptxGenJS',
      'Alice & <Bob>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ author }) => author)).toEqual(expectedAuthors);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.author).toBe(expectedAuthors[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<dc:creator>PptxGenJS</dc:creator>');
    expect(coreXml[0]).toContain('<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>');
    expect(coreXml[1]).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(coreXml[1]).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );
    expect(coreXml[2]).toContain('<dc:creator></dc:creator>');
    expect(coreXml[2]).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ author }) => author)).toEqual(expectedAuthors);

    const native = PptxDocument.create({ author: 'Alice & <Bob>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.author).toBe(imported[1]!.author);
    expect(nativeOmitted.author).toBe('@jiayunxie/pptx');
    const nativeCore = new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(nativeCore).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(nativeCore).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.author).toBe('Alice & <Bob>');
    expect(new TextDecoder().decode(
      reopenedNative.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation last modified by metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.author = 'Alice & <Bob>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.author = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedEditors = ['PptxGenJS', 'Alice & <Bob>', ''] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ lastModifiedBy }) => lastModifiedBy)).toEqual(expectedEditors);
    expect(imported.map(({ author }) => author)).toEqual(expectedEditors);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.lastModifiedBy).toBe(expectedEditors[index]);
      expect(document.author).toBe(expectedEditors[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<dc:creator>PptxGenJS</dc:creator>');
    expect(coreXml[0]).toContain('<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>');
    expect(coreXml[1]).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(coreXml[1]).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );
    expect(coreXml[2]).toContain('<dc:creator></dc:creator>');
    expect(coreXml[2]).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ lastModifiedBy }) => lastModifiedBy)).toEqual(expectedEditors);
    expect(reopened.map(({ author }) => author)).toEqual(expectedEditors);

    const nativeDefault = PptxDocument.create();
    expect(nativeDefault.author).toBe('@jiayunxie/pptx');
    expect(nativeDefault.lastModifiedBy).toBe('@jiayunxie/pptx');
    expect(nativeDefault.lastModifiedBy).not.toBe(imported[0]!.lastModifiedBy);

    const nativeMirror = PptxDocument.create({
      author: 'Alice & <Bob>',
      lastModifiedBy: 'Alice & <Bob>',
    });
    expect(nativeMirror.author).toBe(imported[1]!.author);
    expect(nativeMirror.lastModifiedBy).toBe(imported[1]!.lastModifiedBy);
    const nativeCore = new TextDecoder().decode(
      nativeMirror.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(nativeCore).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(nativeCore).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );

    nativeMirror.author = 'Creator only';
    expect(nativeMirror.author).toBe('Creator only');
    expect(nativeMirror.lastModifiedBy).toBe('Alice & <Bob>');
    nativeMirror.lastModifiedBy = 'Editor only';
    expect(nativeMirror.author).toBe('Creator only');
    expect(nativeMirror.lastModifiedBy).toBe('Editor only');
    const reopenedNative = await PptxDocument.open(await nativeMirror.write());
    expect(reopenedNative.author).toBe('Creator only');
    expect(reopenedNative.lastModifiedBy).toBe('Editor only');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation subject metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.subject = 'Revenue & <Forecast>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.subject = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedSubjects = [
      'PptxGenJS Presentation',
      'Revenue & <Forecast>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ subject }) => subject)).toEqual(expectedSubjects);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.subject).toBe(expectedSubjects[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain(
      '<dc:subject>PptxGenJS Presentation</dc:subject>',
    );
    expect(coreXml[1]).toContain(
      '<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>',
    );
    expect(coreXml[2]).toContain('<dc:subject></dc:subject>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ subject }) => subject)).toEqual(expectedSubjects);

    const native = PptxDocument.create({ subject: 'Revenue & <Forecast>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.subject).toBe(imported[1]!.subject);
    expect(nativeOmitted.subject).toBeUndefined();
    expect(new TextDecoder().decode(
      nativeOmitted.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).not.toContain('<dc:subject');
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.subject).toBe('Revenue & <Forecast>');
    expect(new TextDecoder().decode(
      reopenedNative.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation revision metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const zero = new PptxGenJS();
    zero.revision = '0';
    zero.addSlide();
    const custom = new PptxGenJS();
    custom.revision = '42';
    custom.addSlide();
    const leading = new PptxGenJS();
    leading.revision = '007';
    leading.addSlide();
    expect([baseline.version, zero.version, custom.version, leading.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedRevisions = ['1', '0', '42', '007'] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(zero),
      importPptxGenJS(custom),
      importPptxGenJS(leading),
    ]);
    expect(imported.map(({ revision }) => revision)).toEqual(expectedRevisions);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.revision).toBe(expectedRevisions[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<cp:revision>1</cp:revision>');
    expect(coreXml[1]).toContain('<cp:revision>0</cp:revision>');
    expect(coreXml[2]).toContain('<cp:revision>42</cp:revision>');
    expect(coreXml[3]).toContain('<cp:revision>007</cp:revision>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ revision }) => revision)).toEqual(expectedRevisions);

    const native = PptxDocument.create({ revision: '42' });
    const nativeOmitted = PptxDocument.create();
    expect(native.revision).toBe(imported[2]!.revision);
    expect(nativeOmitted.revision).toBe(imported[0]!.revision);
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<cp:revision>42</cp:revision>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.revision).toBe('42');

    const invalidValues = ['', '-1', '1.5', 'abc'] as const;
    for (const value of invalidValues) {
      const generated = new PptxGenJS();
      generated.revision = value;
      generated.addSlide();
      const document = await importPptxGenJS(generated);
      const before = document.opcPackage.requirePart('/docProps/core.xml').bytes;
      expect(document.revision).toBeUndefined();
      expect(new TextDecoder().decode(before)).toContain(
        `<cp:revision>${value}</cp:revision>`,
      );
      const invalidReopened = await PptxDocument.open(await document.write());
      expect(invalidReopened.revision).toBeUndefined();
      expect(invalidReopened.opcPackage.requirePart('/docProps/core.xml').bytes)
        .toEqual(before);
      expect(() => PptxDocument.create({ revision: value })).toThrow(TypeError);
    }
  }, 20_000);

  it('imports and reopens PptxGenJS presentation company metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.company = 'Acme 国际';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.company = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedCompanies = [
      'PptxGenJS',
      'Acme 国际',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ company }) => company)).toEqual(expectedCompanies);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.company).toBe(expectedCompanies[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const appXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/app.xml').bytes,
    ));
    expect(appXml[0]).toContain('<Company>PptxGenJS</Company>');
    expect(appXml[1]).toContain('<Company>Acme 国际</Company>');
    expect(appXml[2]).toContain('<Company></Company>');
    for (const xml of appXml) {
      expect(xml).toContain('<Application>Microsoft Office PowerPoint</Application>');
      expect(xml).toContain('<AppVersion>16.0000</AppVersion>');
    }

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ company }) => company)).toEqual(expectedCompanies);

    const native = PptxDocument.create({ company: 'Acme & <Partners>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.company).toBe('Acme & <Partners>');
    expect(nativeOmitted.company).toBeUndefined();
    const nativeApp = new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/app.xml').bytes,
    );
    expect(nativeApp).toContain(
      '<Company>Acme &amp; &lt;Partners&gt;</Company>',
    );
    expect(nativeApp).toContain('<Application>@jiayunxie/pptx</Application>');
    expect(nativeApp).toContain('<AppVersion>1.0</AppVersion>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.company).toBe('Acme & <Partners>');

    const unsafe = new PptxGenJS();
    unsafe.company = 'A & <B>';
    unsafe.addSlide();
    const importedUnsafe = await importPptxGenJS(unsafe);
    const unsafeApp = new TextDecoder().decode(
      importedUnsafe.opcPackage.requirePart('/docProps/app.xml').bytes,
    );
    expect(unsafeApp).toContain('<Company>A & <B></Company>');
    expect(importedUnsafe.company).toBeUndefined();
  }, 20_000);

  it('imports PptxGenJS non-list zero margins and indents without aliasing bullet indentation', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    slide.addText('Plain', { name: 'Margin plain', x: 1, y: 1, w: 3, h: 0.5 });
    slide.addText([{ text: 'Rich' }], { name: 'Margin rich', x: 1, y: 2, w: 3, h: 0.5 });
    slide.addText('Bullet', {
      name: 'Margin bullet',
      x: 1,
      y: 3,
      w: 3,
      h: 0.5,
      bullet: true,
    });
    slide.addText('Number', {
      name: 'Margin number',
      x: 1,
      y: 4,
      w: 3,
      h: 0.5,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 1, indent: 22 },
    });
    const document = await importPptxGenJS(generated);
    const shapes = document.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);

    expect(shapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes[2]!.richText[0]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(shapes[3]!.richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 1,
      indent: 22,
    });
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/indent="0" marL="0"/g)).toHaveLength(2);
    expect(slideXml).toMatch(/marL="342900" indent="-342900"/);
    expect(slideXml).toMatch(/marL="279400" indent="-279400"/);
    expect(slideXml).not.toContain('marR=');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShapes = reopened.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
  }, 20_000);

  it('keeps pptxgenjs out of every non-adapter package dependency list', async () => {
    const packagesDirectory = fileURLToPath(new URL('../..', import.meta.url));
    const packageNames = ['lossless-xml', 'model', 'opc', 'sdk', 'validator'];
    for (const packageName of packageNames) {
      const manifest = JSON.parse(await readFile(`${packagesDirectory}/${packageName}/package.json`, 'utf8')) as {
        dependencies?: Record<string, string>;
      };
      expect(manifest.dependencies?.pptxgenjs, packageName).toBeUndefined();
    }
  });
});
