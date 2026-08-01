import { describe, expect, it } from 'vitest';
import {
  CHART_TYPES,
  ChartModel,
  MediaModel,
  PptxDocument,
  ShapeModel,
  SlideLayoutModel,
  SlideMasterModel,
  chartWorkbookMatches,
  inches,
  slideNumberDiagnostics,
  type ReplaceMediaPosterOptions,
  type ReplaceMediaSourceOptions,
  type RichTextColor,
  type SlideNumberOptions,
} from './index.js';

describe('@jiayunxie/pptx stable exports', () => {
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
