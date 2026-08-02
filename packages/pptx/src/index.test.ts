import { describe, expect, it } from 'vitest';
import {
  OUTPUT_TYPES as SDK_OUTPUT_TYPES,
  SCHEME_COLORS as SDK_SCHEME_COLORS,
  TEXT_ALIGNMENTS as SDK_TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS as SDK_TEXT_VERTICAL_ALIGNMENTS,
} from '@pptx/sdk';
import {
  CHART_TYPES,
  ChartModel,
  MediaModel,
  OUTPUT_TYPES,
  PLACEHOLDER_TYPES,
  PPTX_VERSION,
  PptxDocument,
  SCHEME_COLORS,
  ShapeModel,
  SlideLayoutModel,
  SlideMasterModel,
  chartWorkbookMatches,
  inches,
  slideNumberDiagnostics,
  TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS,
  type AddTextOptions,
  type ReplaceMediaPosterOptions,
  type ReplaceMediaSourceOptions,
  type RichTextColor,
  type RichTextParagraph,
  type RichTextRun,
  type RichTextRunStyle,
  type SchemeColor,
  type DefineSlideMasterOptions,
  type Emu,
  type Hyperlink,
  type OutputType,
  type PlaceholderSelector,
  type PlaceholderType,
  type PresentationLayout,
  type PresentationLayoutName,
  type PresetShapeType,
  type PptxNodeReadableStream,
  type PptxVersion,
  type SlideMasterBackground,
  type SlideMasterMargin,
  type SlideMasterObject,
  type SlideNumberOptions,
  type ShapeArrows,
  type ShapeFill,
  type ShapeLine,
  type ShapeShadow,
  type TextAlignment,
  type TextBoxVerticalAlignment,
  type WriteBaseOptions,
  type WriteOptions,
  type WriteOutput,
} from './index.js';

describe('@jiayunxie/pptx stable exports', () => {
  it('exports the read-only runtime version from the root package', async () => {
    const current: PptxVersion = PPTX_VERSION;
    const document = PptxDocument.create();

    expect(current).toBe('0.1.0');
    expect(document.version).toBe(current);
    expect((await PptxDocument.open(await document.write())).version).toBe(current);

    if (false) {
      // @ts-expect-error document version is read-only
      document.version = '9.9.9';
    }
  });

  it('exports the presentation layout projection from the root package', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const layout: PresentationLayout = document.presLayout;
    const name: PresentationLayoutName = layout.name;

    expect({ ...layout, name }).toEqual({
      name: 'custom',
      width: 12_192_000,
      height: 6_858_000,
    });
    document.slideSize = { width: inches(10), height: inches(5.625) };
    const edited: PresentationLayout = document.presLayout;
    expect(edited).toEqual({
      name: 'screen16x9',
      width: inches(10),
      height: inches(5.625),
    });
    expect((await PptxDocument.open(await document.write())).presLayout).toEqual(edited);

    if (false) {
      // @ts-expect-error presLayout is getter-only
      document.presLayout = edited;
      // @ts-expect-error presentation layout fields are read-only
      layout.width = inches(1);
    }
  });

  it('exports the frozen SCHEME_COLORS helper from the root package', async () => {
    expect(SCHEME_COLORS).toBe(SDK_SCHEME_COLORS);
    expect(Object.entries(SCHEME_COLORS)).toEqual([
      ['text1', 'tx1'],
      ['text2', 'tx2'],
      ['background1', 'bg1'],
      ['background2', 'bg2'],
      ['accent1', 'accent1'],
      ['accent2', 'accent2'],
      ['accent3', 'accent3'],
      ['accent4', 'accent4'],
      ['accent5', 'accent5'],
      ['accent6', 'accent6'],
    ]);
    expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);

    const isolated = PptxDocument.create();
    const journal = JSON.stringify(isolated.opcPackage.mutations);
    Object.values(SCHEME_COLORS);
    expect(JSON.stringify(isolated.opcPackage.mutations)).toBe(journal);

    const document = PptxDocument.create();
    document.addSlide().addRichText([{
      runs: [{
        text: 'Scheme helper',
        style: { color: { kind: 'scheme', value: SCHEME_COLORS.text1 } },
      }],
    }], {
      fill: { kind: 'solid', color: { kind: 'scheme', value: SCHEME_COLORS.accent1 } },
    });
    const reopened = await PptxDocument.open(await document.write());
    const shape = reopened.slides[0]?.shapes[0];
    expect(shape).toBeInstanceOf(ShapeModel);
    if (!(shape instanceof ShapeModel)) throw new TypeError('Expected a text shape');
    expect(shape.richText[0]?.runs[0]?.style?.color)
      .toEqual({ kind: 'scheme', value: 'tx1' });
    expect(shape.fill)
      .toEqual({ kind: 'solid', color: { kind: 'scheme', value: 'accent1' } });

    if (false) {
      const text: SchemeColor = SCHEME_COLORS.text1;
      const accent: SchemeColor = SCHEME_COLORS.accent6;
      // @ts-expect-error SchemeColor excludes key labels
      const invalid: SchemeColor = 'background1';
      // @ts-expect-error the runtime catalog is readonly
      SCHEME_COLORS.accent1 = 'accent2';
      void [text, accent, invalid];
    }
  });

  it('exports the frozen TEXT_ALIGNMENTS catalog from the root package', () => {
    const values: readonly TextAlignment[] = TEXT_ALIGNMENTS;

    expect(TEXT_ALIGNMENTS).toBe(SDK_TEXT_ALIGNMENTS);
    expect(values).toBe(TEXT_ALIGNMENTS);
    expect([...values]).toEqual(['left', 'center', 'right', 'justify']);
    expect(Object.isFrozen(TEXT_ALIGNMENTS)).toBe(true);
  });

  it('exports the frozen TEXT_VERTICAL_ALIGNMENTS catalog from the root package', () => {
    const values: readonly TextBoxVerticalAlignment[] = TEXT_VERTICAL_ALIGNMENTS;

    expect(TEXT_VERTICAL_ALIGNMENTS).toBe(SDK_TEXT_VERTICAL_ALIGNMENTS);
    expect(values).toBe(TEXT_VERTICAL_ALIGNMENTS);
    expect([...values]).toEqual(['top', 'middle', 'bottom']);
    expect(Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS)).toBe(true);

    if (false) {
      // @ts-expect-error unknown vertical alignment is not supported
      const invalid: TextBoxVerticalAlignment = 'distributed';
      void invalid;
    }
  });

  it('exports the frozen OUTPUT_TYPES catalog from the root package', () => {
    const values: readonly OutputType[] = OUTPUT_TYPES;

    expect(OUTPUT_TYPES).toBe(SDK_OUTPUT_TYPES);
    expect(values).toBe(OUTPUT_TYPES);
    expect([...values]).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
    expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);

    if (false) {
      // @ts-expect-error runtime output type catalog is readonly
      OUTPUT_TYPES.push('uint8array');
      // @ts-expect-error output type catalog indexes are readonly
      OUTPUT_TYPES[0] = 'uint8array';
      // @ts-expect-error STREAM is handled by the separate stream API
      const stream: OutputType = 'STREAM';
      // @ts-expect-error unknown output type is not supported
      const invalid: OutputType = 'buffer';
      const document = PptxDocument.create();
      const baseOptions: WriteBaseOptions = {
        compatibility: 'powerpoint-current',
        compression: true,
      };
      const blobOptions: WriteOptions<'blob'> = {
        outputType: 'blob',
        compression: false,
      };
      const dynamicOptions: WriteOptions<OutputType> = { outputType: OUTPUT_TYPES[0] };
      document.write() satisfies Promise<Uint8Array>;
      document.write(baseOptions) satisfies Promise<Uint8Array>;
      document.write(blobOptions) satisfies Promise<Blob>;
      document.write(dynamicOptions) satisfies Promise<WriteOutput<OutputType>>;
      document.write({ outputType: 'arraybuffer' }) satisfies Promise<ArrayBuffer>;
      document.write({ outputType: 'base64' }) satisfies Promise<string>;
      document.write({ outputType: 'binarystring' }) satisfies Promise<string>;
      document.write({ outputType: 'nodebuffer' }) satisfies Promise<Uint8Array>;
      document.write({ outputType: 'uint8array' }) satisfies Promise<Uint8Array>;
      document.writeBlob(baseOptions) satisfies Promise<Blob>;
      document.writeFile('output.pptx', { compression: true });
      document.download('output.pptx', { compression: false });
      // @ts-expect-error compression is boolean-only
      document.write({ compression: 'true' });
      // @ts-expect-error file compression is boolean-only
      document.writeFile('output.pptx', { compression: 1 });
      // @ts-expect-error download compression is boolean-only
      document.download('output.pptx', { compression: null });
      // @ts-expect-error convenience blob output does not accept a selector
      document.writeBlob({ outputType: 'blob' });
      void [stream, invalid];
    }
  });

  it('exports the Node readable stream contract from the root package', () => {
    const document = PptxDocument.create();
    if (false) {
      document.stream() satisfies Promise<PptxNodeReadableStream>;
      document.stream({ mode: 'permissive' }) satisfies Promise<PptxNodeReadableStream>;
      document.stream({ compatibility: 'powerpoint-current' }) satisfies Promise<PptxNodeReadableStream>;
      document.stream({ compression: true }) satisfies Promise<PptxNodeReadableStream>;
      void document.stream().then((readable) => {
        readable satisfies AsyncIterable<Uint8Array>;
        const destination = { tag: 'destination' } as const;
        readable.pipe(destination) satisfies typeof destination;
        readable.pause().resume().destroy();
      });
      // @ts-expect-error stream does not consume write output selectors
      document.stream({ outputType: 'uint8array' });
      // @ts-expect-error stream compression is boolean-only
      document.stream({ compression: 'DEFLATE' });
    }
  });

  it('exports semantic master layout models from the root package', () => {
    const document = PptxDocument.create();
    const layout: SlideLayoutModel = document.layouts[0]!;
    const master: SlideMasterModel = document.masters[0]!;
    expect(layout).toBeInstanceOf(SlideLayoutModel);
    expect(master).toBeInstanceOf(SlideMasterModel);
    expect(master.layouts[0]).toBe(layout);
    expect(layout.addText('Root layout text')).toBe(layout.shapes[0]);
    expect(master.addShape('rect')).toBe(master.shapes[0]);
  });

  it('exports define slide master types and runtime from the root package', async () => {
    const margin: SlideMasterMargin = {
      top: inches(0.1),
      right: inches(0.2),
      bottom: inches(0.3),
      left: inches(0.4),
    };
    const objects: readonly SlideMasterObject[] = [
      { kind: 'rect' },
      { kind: 'text', text: 'Root brand' },
      {
        kind: 'placeholder',
        options: { name: 'root_title', type: 'title' },
      },
    ];
    const definition: DefineSlideMasterOptions = {
      title: 'ROOT-BRAND',
      margin: [margin.top, margin.right, margin.bottom, margin.left],
      objects,
    };
    const asyncBackground: SlideMasterBackground = {
      kind: 'image-source',
      source: new Uint8Array(),
      contentType: 'image/png',
    };
    const asyncObjects: readonly SlideMasterObject[] = [
      { kind: 'image', source: new Uint8Array() },
      {
        kind: 'chart',
        groups: [{
          type: 'bar',
          series: [{ name: 'Root', categories: ['Q1'], values: [1] }],
        }],
      },
    ];
    const document = PptxDocument.create();
    const layout = await document.defineSlideMaster(definition);
    expect(layout.name).toBe('ROOT-BRAND');
    expect(layout.margin).toEqual(margin);
    expect(asyncBackground.kind).toBe('image-source');
    expect(asyncObjects.map(({ kind }) => kind)).toEqual(['image', 'chart']);
  });

  it('creates, selects, populates, and reopens placeholders through the root package', async () => {
    const placeholderType: PlaceholderType = 'chart';
    const pictureSelector: PlaceholderSelector = { type: 'pic', index: 102 };
    const png = Uint8Array.from([
      137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
      0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
      0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
      39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ]);
    const document = PptxDocument.create();
    const layout = await document.defineSlideMaster({
      title: 'ROOT-PLACEHOLDERS',
      objects: [
        {
          kind: 'placeholder',
          text: 'Title prompt',
          options: { name: 'root_title', type: 'title', index: 101 },
        },
        {
          kind: 'placeholder',
          text: 'Picture prompt',
          options: { name: 'root_picture', type: 'pic', index: 102 },
        },
        {
          kind: 'placeholder',
          text: 'Chart prompt',
          options: { name: 'root_chart', type: placeholderType, index: 103 },
        },
      ],
    });
    const slide = document.addSlide({ masterName: layout.name });
    slide.addText('Root title', { placeholder: 'root_title' });
    await document.addImage(0, png, {
      contentType: 'image/png',
      placeholder: pictureSelector,
    });
    await document.addChart(0, 'bar', [{
      name: 'Revenue',
      categories: ['Q1', 'Q2'],
      values: [10, 20],
    }], { placeholder: 'root_chart' });

    expect(PLACEHOLDER_TYPES).toEqual(['title', 'body', 'pic', 'chart', 'tbl', 'media']);
    expect(slide.relationships.find(({ type }) => type.endsWith('/slideLayout')))
      .toMatchObject({ resolvedTarget: layout.partUri });
    expect(slide.shapes.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'title', index: 101 },
      { type: 'pic', index: 102 },
      { type: 'chart', index: 103 },
    ]);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.layouts.find(({ name }) => name === layout.name)).toBeInstanceOf(SlideLayoutModel);
    expect(reopened.masters[0]).toBeInstanceOf(SlideMasterModel);
    expect(reopened.slides[0]?.shapes.map(({ kind, placeholder }) => ({ kind, placeholder })))
      .toEqual([
        { kind: 'text', placeholder: { type: 'title', index: 101 } },
        { kind: 'image', placeholder: { type: 'pic', index: 102 } },
        { kind: 'chart', placeholder: { type: 'chart', index: 103 } },
      ]);
    expect((reopened.slides[0]?.shapes[2] as ChartModel).definition?.groups[0]?.type)
      .toBe('bar');
  });

  it('exports transient slide default colors and materializes them through the root', async () => {
    const defaultColor: RichTextColor = { kind: 'scheme', value: 'accent1' };
    const document = PptxDocument.create();
    const source = document.addSlide();
    source.color = defaultColor;
    source.addText('Inherited plain');
    source.addRichText([{
      runs: [
        { text: 'Inherited rich' },
        { text: 'Override', style: { color: { kind: 'srgb', value: '00AA00' } } },
      ],
    }]);
    const duplicate = document.duplicateSlide(0);
    expect(duplicate.color).toBe(source.color);
    duplicate.addText('Duplicate inherited');

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ color }) => color)).toEqual([undefined, undefined]);
    const colors = reopened.slides.map((slide) => slide.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .map(({ richText }) => richText.flatMap(({ runs }) =>
        runs.map(({ style }) => style?.color))));
    expect(colors).toEqual([
      [
        [{ kind: 'scheme', value: 'accent1' }],
        [
          { kind: 'scheme', value: 'accent1' },
          { kind: 'srgb', value: '00AA00' },
        ],
      ],
      [
        [{ kind: 'scheme', value: 'accent1' }],
        [
          { kind: 'scheme', value: 'accent1' },
          { kind: 'srgb', value: '00AA00' },
        ],
        [{ kind: 'scheme', value: 'accent1' }],
      ],
    ]);
  });

  it('exports text shape fill creation types and runtime from the root package', async () => {
    const none: ShapeFill = { kind: 'none' };
    const srgb: ShapeFill = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'A1B2C3' },
      transparency: 25,
    };
    const scheme: ShapeFill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
    };
    const options: AddTextOptions = { name: 'root_text_fill', fill: srgb };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root plain fill', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rich fill' }] }], {
      name: 'root_rich_fill',
      fill: scheme,
    });
    const layoutText = document.layouts[0]!.addText('Root layout fill', {
      name: 'root_layout_fill',
      fill: none,
    });
    const masterText = document.masters[0]!.addText('Root master fill', {
      name: 'root_master_fill',
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent3' } },
    });

    expect([plain.fill, rich.fill, layoutText.fill, masterText.fill]).toEqual([
      srgb,
      scheme,
      none,
      { kind: 'solid', color: { kind: 'scheme', value: 'accent3' } },
    ]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_fill',
    ) as ShapeModel).fill).toEqual(srgb);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_fill',
    ) as ShapeModel).fill).toEqual(scheme);
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_fill',
    ) as ShapeModel).fill).toEqual(none);
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_fill',
    ) as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
    });

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error PptxGenJS-style fill objects are intentionally unsupported
        fill: { color: 'FF0000' },
      },
      {
        // @ts-expect-error fill kind must be none or solid
        fill: { kind: 'gradient' },
      },
      {
        // @ts-expect-error solid fills require a color
        fill: { kind: 'solid' },
      },
      {
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
          // @ts-expect-error transparency must be numeric
          transparency: '25',
        },
      },
      {
        fill: {
          kind: 'none',
          // @ts-expect-error none fills do not accept extra properties
          transparency: 0,
        },
      },
    ];
    expect(invalid).toHaveLength(5);
  });

  it('exports text shape line creation types and runtime from the root package', async () => {
    const none: ShapeLine = { kind: 'none' };
    const srgb: ShapeLine = {
      kind: 'line',
      color: { kind: 'srgb', value: 'A1B2C3' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    };
    const scheme: ShapeLine = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
      width: 0,
      dash: 'sysDot',
    };
    const options: AddTextOptions = { name: 'root_text_line', line: srgb };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root plain line', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rich line' }] }], {
      name: 'root_rich_line',
      line: scheme,
    });
    const layoutText = document.layouts[0]!.addText('Root layout line', {
      name: 'root_layout_line',
      line: none,
    });
    const masterText = document.masters[0]!.addText('Root master line', {
      name: 'root_master_line',
      line: { kind: 'line', color: { kind: 'scheme', value: 'accent3' } },
    });

    expect([plain.line, rich.line, layoutText.line, masterText.line]).toEqual([
      srgb,
      scheme,
      none,
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 1,
        dash: 'solid',
      },
    ]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_line',
    ) as ShapeModel).line).toEqual(srgb);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_line',
    ) as ShapeModel).line).toEqual(scheme);
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_line',
    ) as ShapeModel).line).toEqual(none);
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_line',
    ) as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 1,
      dash: 'solid',
    });

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error PptxGenJS-style line objects are intentionally unsupported
        line: { color: 'FF0000' },
      },
      {
        // @ts-expect-error line kind must be none or line
        line: { kind: 'solid' },
      },
      {
        // @ts-expect-error solid lines require a color
        line: { kind: 'line' },
      },
      {
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          // @ts-expect-error width must be numeric
          width: '2',
        },
      },
      {
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          // @ts-expect-error dash must be a supported preset
          dash: 'dot',
        },
      },
      {
        line: {
          kind: 'none',
          // @ts-expect-error none lines do not accept extra properties
          transparency: 0,
        },
      },
    ];
    expect(invalid).toHaveLength(6);
  });

  it('exports text shape arrow creation types and runtime from the root package', async () => {
    const both: ShapeArrows = { begin: 'triangle', end: 'arrow' };
    const explicitNone: ShapeArrows = { begin: 'none', end: 'stealth' };
    const empty: ShapeArrows = {};
    const options: AddTextOptions = { name: 'root_text_arrows', arrows: both };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root plain arrows', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rich arrows' }] }], {
      name: 'root_rich_arrows',
      arrows: explicitNone,
    });
    const layoutText = document.layouts[0]!.addText('Root layout arrows', {
      name: 'root_layout_arrows',
      arrows: empty,
    });
    const masterText = document.masters[0]!.addText('Root master arrows', {
      name: 'root_master_arrows',
      line: { kind: 'line', color: { kind: 'scheme', value: 'accent3' } },
      arrows: { end: 'diamond' },
    });

    expect([plain.arrows, rich.arrows, layoutText.arrows, masterText.arrows]).toEqual([
      both,
      explicitNone,
      undefined,
      { end: 'diamond' },
    ]);
    expect(masterText.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 1,
      dash: 'solid',
    });

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_arrows',
    ) as ShapeModel).arrows).toEqual(both);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_arrows',
    ) as ShapeModel).arrows).toEqual(explicitNone);
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_arrows',
    ) as ShapeModel).arrows).toBeUndefined();
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_arrows',
    ) as ShapeModel).arrows).toEqual({ end: 'diamond' });

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error PptxGenJS-style arrow aliases are intentionally unsupported
        arrows: { beginArrowType: 'arrow' },
      },
      {
        // @ts-expect-error deprecated arrow aliases are intentionally unsupported
        arrows: { lineHead: 'triangle' },
      },
      {
        // @ts-expect-error begin arrow tokens are a closed union
        arrows: { begin: 'bogus' },
      },
      {
        // @ts-expect-error end arrow tokens are a closed union
        arrows: { end: '' },
      },
      {
        arrows: {
          begin: 'triangle',
          // @ts-expect-error arrows reject unknown properties
          extra: true,
        },
      },
    ];
    expect(invalid).toHaveLength(5);
  });

  it('exports text shape shadow creation types and runtime from the root package', async () => {
    const outer: ShapeShadow = {
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0.4,
      blur: 2.5,
      angle: 45,
      distance: 3,
      rotateWithShape: true,
    };
    const inner: ShapeShadow = {
      kind: 'inner',
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    };
    const options: AddTextOptions = { name: 'root_text_shadow', shadow: outer };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root plain shadow', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rich shadow' }] }], {
      name: 'root_rich_shadow',
      shadow: inner,
    });
    const layoutText = document.layouts[0]!.addText('Root layout shadow', {
      name: 'root_layout_shadow',
      shadow: { kind: 'outer' },
    });
    const masterText = document.masters[0]!.addText('Root master shadow', {
      name: 'root_master_shadow',
      line: { kind: 'line', color: { kind: 'scheme', value: 'accent3' } },
      arrows: { end: 'diamond' },
      shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent4' } },
    });

    expect(plain.shadow).toEqual(outer);
    expect(rich.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect(layoutText.shadow).toMatchObject({ kind: 'outer', rotateWithShape: false });
    expect(masterText.shadow).toMatchObject({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent4' },
    });
    expect(masterText.line).toMatchObject({ kind: 'line', width: 1, dash: 'solid' });
    expect(masterText.arrows).toEqual({ end: 'diamond' });

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_shadow',
    ) as ShapeModel).shadow).toEqual(outer);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_shadow',
    ) as ShapeModel).shadow).toEqual({
      kind: 'inner',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_shadow',
    ) as ShapeModel).shadow).toMatchObject({ kind: 'outer' });
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_shadow',
    ) as ShapeModel).shadow).toMatchObject({ kind: 'inner' });

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error PptxGenJS-style shadow aliases are intentionally unsupported
        shadow: { type: 'outer' },
      },
      {
        // @ts-expect-error shadow kind must be outer or inner
        shadow: { kind: 'none' },
      },
      {
        shadow: {
          kind: 'outer',
          // @ts-expect-error PptxGenJS offset alias is intentionally unsupported
          offset: 4,
        },
      },
      {
        shadow: {
          kind: 'inner',
          // @ts-expect-error inner shadow cannot rotate with the shape
          rotateWithShape: true,
        },
      },
      {
        shadow: {
          kind: 'outer',
          // @ts-expect-error opacity must be numeric
          opacity: '0.5',
        },
      },
    ];
    expect(invalid).toHaveLength(5);
  });

  it('exports text shape hyperlink creation types and runtime from the root package', async () => {
    const url: Hyperlink = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const internal: Hyperlink = { slide: 2, tooltip: '' };
    const options: AddTextOptions = { name: 'root_text_hyperlink', hyperlink: url };
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const plain = source.addText('Root plain hyperlink', options);
    const rich = source.addRichText([{ runs: [{ text: 'Root rich hyperlink' }] }], {
      name: 'root_rich_hyperlink',
      hyperlink: internal,
    });
    const layoutText = document.layouts[0]!.addText('Root layout hyperlink', {
      name: 'root_layout_hyperlink',
      hyperlink: { url: 'https://layout.example' },
    });
    const masterText = document.masters[0]!.addText('Root master hyperlink', {
      name: 'root_master_hyperlink',
      hyperlink: { slide: document.slides.indexOf(target) + 1 },
    });

    expect(plain.hyperlink).toEqual(url);
    expect(rich.hyperlink).toEqual(internal);
    expect(layoutText.hyperlink).toEqual({ url: 'https://layout.example' });
    expect(masterText.hyperlink).toEqual({ slide: 2 });

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_hyperlink',
    ) as ShapeModel).hyperlink).toEqual(url);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_hyperlink',
    ) as ShapeModel).hyperlink).toEqual(internal);
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({ url: 'https://layout.example' });
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({ slide: 2 });

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error text hyperlink requires exactly one target
        hyperlink: {},
      },
      {
        // @ts-expect-error text hyperlink target branches are mutually exclusive
        hyperlink: { url: 'https://example.com', slide: 2 },
      },
      {
        // @ts-expect-error text hyperlink URL must be a string
        hyperlink: { url: 42 },
      },
      {
        // @ts-expect-error text hyperlink slide must be numeric
        hyperlink: { slide: '2' },
      },
      {
        hyperlink: {
          url: 'https://example.com',
          // @ts-expect-error relationship IDs are intentionally unsupported
          _rId: 'rId9',
        },
      },
      {
        // @ts-expect-error text hyperlink tooltip must be a string
        hyperlink: { slide: 2, tooltip: 7 },
      },
    ];
    expect(invalid).toHaveLength(6);
  });

  it('exports text shape preset geometry types and runtime from the root package', async () => {
    const preset: PresetShapeType = 'ellipse';
    const options: AddTextOptions = {
      name: 'root_text_preset_geometry',
      shape: preset,
    };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root plain geometry', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rich geometry' }] }], {
      name: 'root_rich_preset_geometry',
      shape: 'star5',
    });
    const layoutText = document.layouts[0]!.addText('Root layout geometry', {
      name: 'root_layout_preset_geometry',
      shape: 'roundRect',
    });
    const masterText = document.masters[0]!.addText('Root master geometry', {
      name: 'root_master_preset_geometry',
      shape: 'foldedCorner',
    });
    expect([plain.presetType, rich.presetType, layoutText.presetType, masterText.presetType])
      .toEqual(['ellipse', 'star5', 'roundRect', 'foldedCorner']);
    plain.presetType = 'hexagon';

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_preset_geometry',
    ) as ShapeModel).presetType).toBe('hexagon');
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_preset_geometry',
    ) as ShapeModel).presetType).toBe('star5');
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'root_layout_preset_geometry',
    ) as ShapeModel).presetType).toBe('roundRect');
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'root_master_preset_geometry',
    ) as ShapeModel).presetType).toBe('foldedCorner');

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error malformed upstream folded-corner spelling is excluded
        shape: 'folderCorner',
      },
      {
        // @ts-expect-error custom geometry is not a preset text shape token
        shape: 'custGeom',
      },
      {
        // @ts-expect-error unknown preset geometry is excluded
        shape: 'unknown',
      },
      {
        // @ts-expect-error preset geometry must be a string token
        shape: 1,
      },
    ];
    expect(invalid).toHaveLength(4);
  });

  it('exports text shape rectangle radius types and runtime from the root package', async () => {
    const radius: Emu = inches(0.5);
    const options: AddTextOptions = {
      name: 'root_text_rect_radius',
      shape: 'roundRect',
      rectRadius: radius,
      width: inches(4),
      height: inches(2),
    };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root rounded text', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root rounded rich text' }] }], {
      name: 'root_rich_rect_radius',
      shape: 'roundRect',
      rectRadius: inches(0),
      width: inches(2),
      height: inches(1),
    });
    const placeholder = slide.addPlaceholder('Root rounded prompt', {
      name: 'root_rect_radius_prompt',
      type: 'title',
      shape: 'roundRect',
      rectRadius: inches(0.25),
      width: inches(2),
      height: inches(1),
    });
    expect(plain.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(rich.adjustments).toEqual([{ name: 'adj', value: 0 }]);
    expect(placeholder.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    plain.adjustments = [{ name: 'adj', value: 12_500 }];
    plain.setTransform({ width: inches(8), height: inches(4) });
    expect(plain.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_rect_radius',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 12_500 }]);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_rich_rect_radius',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 0 }]);

    const invalid: readonly AddTextOptions[] = [
      {
        shape: 'roundRect',
        // @ts-expect-error radius uses branded EMU, not an implicit inch number
        rectRadius: 0.5,
      },
      {
        shape: 'roundRect',
        // @ts-expect-error radius does not accept string coercion
        rectRadius: '0.5',
      },
      {
        shape: 'roundRect',
        // @ts-expect-error radius must be numeric EMU
        rectRadius: false,
      },
      {
        shape: 'roundRect',
        // @ts-expect-error radius must not be an object
        rectRadius: {},
      },
    ];
    expect(invalid).toHaveLength(4);
  });

  it('exports text box state creation and editing from the root package', async () => {
    const options: AddTextOptions = {
      name: 'root_text_box_state',
      isTextBox: true,
    };
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Root text box', options);
    const rich = slide.addRichText([{ runs: [{ text: 'Root shape rich text' }] }], {
      name: 'root_shape_text_state',
      isTextBox: false,
    });
    const placeholder = slide.addPlaceholder('Root text box prompt', {
      name: 'root_text_box_prompt',
      type: 'title',
      isTextBox: true,
    });
    const readable: boolean | undefined = plain.isTextBox;
    expect(readable).toBe(true);
    expect(rich.isTextBox).toBe(false);
    expect(placeholder.isTextBox).toBe(true);
    plain.isTextBox = false;
    rich.isTextBox = true;
    expect([plain.isTextBox, rich.isTextBox]).toEqual([false, true]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_text_box_state',
    ) as ShapeModel).isTextBox).toBe(false);
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_shape_text_state',
    ) as ShapeModel).isTextBox).toBe(true);

    const invalid: readonly AddTextOptions[] = [
      {
        // @ts-expect-error isTextBox does not accept string truthiness
        isTextBox: 'true',
      },
      {
        // @ts-expect-error isTextBox does not accept numeric truthiness
        isTextBox: 1,
      },
      {
        // @ts-expect-error isTextBox does not accept null
        isTextBox: null,
      },
      {
        // @ts-expect-error isTextBox does not accept objects
        isTextBox: {},
      },
    ];
    expect(invalid).toHaveLength(4);
    if (false) {
      // @ts-expect-error live isTextBox setter only accepts boolean
      plain.isTextBox = 'true';
      // @ts-expect-error undefined is not a writable text box state
      plain.isTextBox = undefined;
    }
  });

  it('exports rich text line break types and runtime from the root package', async () => {
    const first: RichTextRun = { text: 'First', breakLine: true };
    const empty: RichTextRun = { text: '', breakLine: true };
    const last: RichTextRun = { text: 'Last', softBreakBefore: true, breakLine: true };
    const paragraphs: readonly RichTextParagraph[] = [{
      align: 'center',
      runs: [first, empty, last],
    }];
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText(paragraphs, { name: 'root_break_line' });

    expect(shape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
      ['First'],
      [],
      ['Last'],
    ]);
    expect(shape.richText[2]!.runs[0]!.softBreakBefore).toBe(true);
    expect(shape.richText.flatMap(({ runs }) => runs).some((run) =>
      Object.hasOwn(run, 'breakLine'))).toBe(false);
    shape.richText = [{
      runs: [
        { text: 'Edited first', breakLine: true },
        { text: 'Edited last' },
      ],
    }];

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'root_break_line',
    ) as ShapeModel).richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
      ['Edited first'],
      ['Edited last'],
    ]);

    const invalidRuns: readonly RichTextRun[] = [
      {
        text: 'String',
        // @ts-expect-error breakLine accepts only primitive boolean values
        breakLine: 'true',
      },
      {
        text: 'Number',
        // @ts-expect-error numeric truthiness is intentionally unsupported
        breakLine: 1,
      },
      {
        text: 'Null',
        // @ts-expect-error null is not a line-break state
        breakLine: null,
      },
      {
        text: 'Object',
        // @ts-expect-error objects are not line-break states
        breakLine: {},
      },
    ];
    const invalidOuter: AddTextOptions = {
      // @ts-expect-error breakLine belongs to RichTextRun, not outer text options
      breakLine: true,
    };
    expect(invalidRuns).toHaveLength(4);
    expect(invalidOuter).toEqual({ breakLine: true });
  });

  it('exports rich text run hyperlink types and runtime from the root package', async () => {
    const urlStyle: RichTextRunStyle = {
      hyperlink: { url: 'https://root-run.example', tooltip: '' },
    };
    const suppressed: RichTextRunStyle = { hyperlink: false };
    const document = PptxDocument.create();
    const source = document.addSlide();
    document.addSlide();
    const shape = source.addRichText([{
      runs: [
        { text: 'URL', style: urlStyle },
        { text: 'Slide', style: { hyperlink: { slide: 2 } } },
        { text: 'Suppressed', style: suppressed },
      ],
    }], { hyperlink: { url: 'https://outer-root.example' } });

    expect(shape.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://root-run.example', tooltip: '' },
      { slide: 2 },
      undefined,
    ]);
    shape.richText = [{
      runs: [{
        text: 'Edited',
        style: { hyperlink: { url: 'https://root-run-edited.example' } },
      }],
    }];
    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel)
      .richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://root-run-edited.example' });

    const invalid: readonly RichTextRunStyle[] = [
      {
        // @ts-expect-error true is not a valid suppression sentinel
        hyperlink: true,
      },
      {
        // @ts-expect-error run hyperlinks require exactly one target branch
        hyperlink: { url: 'https://example.com', slide: 2 },
      },
      {
        // @ts-expect-error target is not a supported hyperlink alias
        hyperlink: { target: 'https://example.com' },
      },
      {
        // @ts-expect-error tooltip must be a string
        hyperlink: { url: 'https://example.com', tooltip: 7 },
      },
    ];
    expect(invalid).toHaveLength(4);
  });

  it('exports slide-number creation, editing, and compatibility diagnostics from the root', async () => {
    const options: SlideNumberOptions = {
      align: 'justify',
      rtl: true,
      style: { italic: true, transparency: 25 },
    };
    const document = PptxDocument.create({ firstSlideNumber: 8 });
    const slide = document.addSlide();
    slide.slideNumber = options;
    document.layouts[0]!.slideNumber = { align: 'center' };
    document.masters[0]!.slideNumber = { align: 'right' };

    expect(slideNumberDiagnostics(
      document.opcPackage,
      slide.partUri,
      'slide',
      '8',
      'powerpoint-current',
    )).toEqual([]);
    await document.write();
    expect(document.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')))
      .toEqual([]);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.slideNumber).toMatchObject(options);
    expect(reopened.layouts[0]?.slideNumber?.align).toBe('center');
    expect(reopened.masters[0]?.slideNumber?.align).toBe('right');
  });

  it('runs the complete native chart lifecycle through the root package', async () => {
    const document = PptxDocument.create();
    const charts: ChartModel[] = [];
    for (const type of CHART_TYPES) {
      const slide = document.addSlide();
      const series = type === 'scatter'
        ? [{ name: 'Points', xValues: [1, 2, 3], values: [4, 6, 5] }]
        : type === 'bubble'
          ? [{ name: 'Bubbles', xValues: [1, 2, 3], values: [4, 6, 5], sizes: [8, 12, 10] }]
          : [{ name: 'Revenue', categories: ['North', 'South', 'West'], values: [120, 150, 135] }];
      charts.push(await slide.addChart(type, series, {
        name: `Root ${type} chart`,
        altText: `${type} chart lifecycle evidence`,
        x: inches(0.5),
        y: inches(0.5),
        width: inches(9),
        height: inches(6.5),
      }));
    }
    const comboSlide = document.addSlide();
    const combo = await comboSlide.addChart([
      {
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Margin', categories: ['Q1', 'Q2'], values: [25, 30] }],
      },
    ], { name: 'Root combination chart' });
    expect(combo.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['bar', 'primary'],
      ['line', 'secondary'],
    ]);

    await charts[0]!.replaceSeries([{
      name: 'Revenue edited',
      categories: ['North', 'South', 'West'],
      values: [125, 155, 140],
    }]);
    await charts[1]!.replaceDefinition({
      groups: [{
        type: 'line',
        series: [{ name: 'Converted', categories: ['Q1', 'Q2'], values: [11, 22] }],
      }],
    });

    const duplicate = document.duplicateSlide(document.slides.length - 1);
    const duplicateChart = duplicate.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!;
    expect(duplicateChart.chartPartUri).not.toBe(combo.chartPartUri);
    const duplicatePartUri = duplicateChart.chartPartUri!;
    duplicateChart.remove();
    expect(document.opcPackage.hasPart(duplicatePartUri)).toBe(false);
    expect(document.opcPackage.hasPart(combo.chartPartUri!)).toBe(true);

    const reopened = await PptxDocument.open(await document.write());
    expect(document.diagnostics.filter(({ code }) => code.startsWith('CHART_'))).toEqual([]);
    const reopenedCharts = reopened.slides.flatMap(({ shapes }) => shapes).filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(reopenedCharts).toHaveLength(10);
    expect(new Set(reopenedCharts.flatMap(({ definition }) =>
      definition?.groups.map(({ type }) => type) ?? []))).toEqual(new Set(CHART_TYPES));
    for (const chart of reopenedCharts) {
      expect(chart.workbookPartUri).toBeDefined();
      expect(await chartWorkbookMatches(
        reopened.opcPackage.requirePart(chart.workbookPartUri!).bytes,
        chart.definition!,
      )).toBe(true);
      expect(await chart.diagnostics()).toEqual([]);
    }
    for (const slide of reopened.slides) {
      const ids = slide.shapes.map(({ id }) => id);
      expect(new Set(ids).size).toBe(ids.length);
    }
    expect(reopened.opcPackage.parts
      .filter(({ contentType }) =>
        contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
        || contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      .every(({ uri }) =>
        (reopened.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) > 0))
      .toBe(true);
  });

  it('runs the complete live media lifecycle through the root package', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const audio = await document.addAudio(0, Uint8Array.of(1, 2, 3), {
      name: 'Root audio',
      altText: 'Root package narration',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(4),
      posterContentType: 'image/png',
    });
    const video = await document.addVideo(0, 'https://example.com/root-video.mp4');

    expect(audio).toBeInstanceOf(MediaModel);
    expect(document.media(0)[0]).toBe(audio);
    expect(slide.media[0]).toBe(audio);
    expect(slide.shapes[0]).toBe(audio);
    expect(audio.shapeId).toBe(audio.id);
    expect(audio.slidePartUri).toBe(slide.partUri);

    audio.name = 'Root audio edited';
    audio.altText = undefined;
    audio.settings = { play: 'auto', loop: true, volume: 0.5 };
    audio.setTransform({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(1),
    });
    const sourceOptions: ReplaceMediaSourceOptions = { contentType: 'audio/wav' };
    expect(await audio.replaceSource(Uint8Array.of(5, 6), sourceOptions)).toBe(audio);
    expect(await audio.replaceSource('https://example.com/root-audio.wav')).toBe(audio);
    expect(await audio.replaceSource(Uint8Array.of(5, 6), sourceOptions)).toBe(audio);
    const posterOptions: ReplaceMediaPosterOptions = { contentType: 'image/gif' };
    expect(await audio.replacePoster(Uint8Array.of(7), posterOptions)).toBe(audio);

    const duplicate = document.duplicateSlide(0);
    const duplicateAudio = duplicate.media[0]!;
    const duplicateVideo = duplicate.media[1]!;
    expect(duplicateAudio).not.toBe(audio);
    expect(duplicateAudio.mediaPartUri).toBe(audio.mediaPartUri);
    await duplicateAudio.replaceSource(Uint8Array.of(8), { contentType: 'audio/ogg' });
    await duplicateAudio.replacePoster(Uint8Array.of(9), { contentType: 'image/jpeg' });
    duplicateVideo.remove();
    document.moveSlide(1, 0);
    expect(document.slides[0]).toBe(duplicate);
    document.moveSlide(0, 1);
    video.remove();
    expect(document.media(0)).toEqual([audio]);
    expect(document.media(1)).toEqual([duplicateAudio]);

    await document.write();
    audio.name = 'Root audio after write';
    const reopened = await PptxDocument.open(await document.write());
    const reopenedAudio = reopened.media(0)[0]!;
    expect(reopenedAudio).toBeInstanceOf(MediaModel);
    expect(reopenedAudio.name).toBe('Root audio after write');
    expect(reopenedAudio.mediaPartUri).not.toBe(reopened.media(1)[0]!.mediaPartUri);
    expect(reopenedAudio.posterPartUri).not.toBe(reopened.media(1)[0]!.posterPartUri);
    await reopenedAudio.replacePoster();
    reopenedAudio.settings = undefined;
    expect(reopenedAudio.settings).toEqual({});
    expect(new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    )).not.toContain('<p:timing>');
    reopenedAudio.settings = { play: 'click', volume: 1 };
    const second = await PptxDocument.open(await reopened.write());
    expect(second.media(0)[0]!.posterPartUri).toMatch(/\.png$/);
    expect(second.media(0)[0]!.settings).toEqual({
      play: 'click',
      loop: false,
      hideWhenStopped: false,
      volume: 1,
    });
    for (const [slideIndex, expected] of second.slides.entries()) {
      const source = new TextDecoder().decode(second.opcPackage.requirePart(expected.partUri).bytes);
      const ids = [...source.matchAll(/<p:cTn\b[^>]*\bid="([0-9]+)"/g)]
        .map((match) => Number(match[1]));
      const targets = [...source.matchAll(/<p:spTgt\b[^>]*\bspid="([0-9]+)"/g)]
        .map((match) => Number(match[1]));
      expect(source).toContain('cmd="playFrom(0.0)"');
      expect(source).toContain('<p:audio><p:cMediaNode');
      expect(new Set(ids).size).toBe(ids.length);
      expect(new Set(targets)).toEqual(new Set(second.media(slideIndex).map(({ shapeId }) => shapeId)));
    }
    await second.write({ mode: 'permissive' });
    expect(second.diagnostics.filter(({ code }) => code.startsWith('MEDIA_TIMING_'))).toEqual([]);

    if (false) {
      // @ts-expect-error media sources exclude scalar numbers
      await audio.replaceSource(1);
      // @ts-expect-error source replacement excludes placement options
      await audio.replaceSource(Uint8Array.of(1), { x: inches(1) });
      // @ts-expect-error poster replacement excludes transcoders
      await audio.replacePoster(Uint8Array.of(1), { transcode: async () => undefined });
      // @ts-expect-error playback mode excludes hover
      audio.settings = { play: 'hover' };
      // @ts-expect-error media names must be strings
      audio.name = 1;
      // @ts-expect-error shape ids must be numbers
      slide.deleteMedia('2');
    }
  });
});
