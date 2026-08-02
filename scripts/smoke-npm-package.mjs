import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const tarball = resolve(process.argv[2] ?? '');
if (!tarball.endsWith('.tgz')) throw new Error('Usage: node scripts/smoke-npm-package.mjs <package.tgz>');
const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const placeholderTypes = ['title', 'body', 'pic', 'chart', 'tbl', 'media'];

const directory = await mkdtemp(join(tmpdir(), 'jiayunxie-pptx-smoke-'));
try {
  await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  run('npm', [
    'install',
    '--ignore-scripts',
    '--no-audit',
    '--no-fund',
    '--package-lock=false',
    '--cache',
    join(directory, '.npm-cache'),
    tarball,
  ], directory);

  const installed = join(directory, 'node_modules', '@jiayunxie', 'pptx');
  const manifest = JSON.parse(await readFile(join(installed, 'package.json'), 'utf8'));
  if (manifest.name !== '@jiayunxie/pptx' || manifest.version !== '0.1.0') {
    throw new Error(`Unexpected package identity: ${manifest.name}@${manifest.version}`);
  }
  if (JSON.stringify(manifest).includes('workspace:')) throw new Error('Packed manifest contains workspace protocol');
  if (manifest.exports?.['.']?.browser !== './dist/browser.js') {
    throw new Error('Packed manifest is missing the browser conditional export');
  }
  for (const section of ['dependencies', 'optionalDependencies', 'peerDependencies']) {
    for (const name of Object.keys(manifest[section] ?? {})) {
      if (name.startsWith('@pptx/')) throw new Error(`Packed manifest contains internal runtime dependency: ${name}`);
    }
  }

  const distFiles = (await listFiles(join(installed, 'dist'))).filter((name) => /\.(?:js|d\.ts)$/.test(name));
  for (const file of distFiles) {
    const source = await readFile(file, 'utf8');
    if (/(?:\bfrom\s*|\bimport\s*(?:\(\s*)?|\brequire\s*\(\s*)["']@pptx\//.test(source)) {
      throw new Error(`Bundled output contains internal import: ${file}`);
    }
  }
  const browserSource = await readFile(join(installed, 'dist/browser.js'), 'utf8');
  if (/\b(?:from|import)\s*['"]node:/.test(browserSource)) {
    throw new Error('Browser bundle contains a static node: import');
  }

  await writeFile(
    join(directory, 'smoke.mjs'),
    `import { CHART_TYPES, ChartModel, calculateImageSizing, chartWorkbookMatches, CustomGeometryEvaluationError, evaluateCustomGeometry, ImageModel, inches, inspectImage, inspectRasterImage, inspectSvgImage, MediaCodec, MediaModel, PLACEHOLDER_TYPES, PRESET_SHAPE_TYPES, PptxDocument, ShapeModel, SlideLayoutModel, SlideMasterModel, TableModel, GradientCodec, importPptxGenJS, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const created = PptxDocument.create({ rtlMode: true });
const slideNumberDeck = PptxDocument.create({ firstSlideNumber: 5 });
const packedNumberSource = slideNumberDeck.addSlide();
packedNumberSource.slideNumber = {
  x: 0,
  y: 0,
  width: inches(1),
  height: inches(0.3),
  align: 'justify',
  rtl: true,
  valign: 'middle',
  margin: [1, 2, 3, 4],
  style: {
    fontFamily: 'Aptos',
    fontSize: 18,
    lang: 'zh-CN',
    bold: true,
    italic: true,
    color: { kind: 'srgb', value: 'FF3399' },
    transparency: 20,
  },
};
const packedNumberSecond = slideNumberDeck.addSlide();
packedNumberSecond.slideNumber = {
  align: 'center',
  style: { color: { kind: 'scheme', value: 'accent1' } },
};
slideNumberDeck.layouts[0].slideNumber = { x: 200, align: 'center' };
slideNumberDeck.masters[0].slideNumber = { x: 300, align: 'right' };
slideNumberDeck.layouts[0].slideNumber = undefined;
slideNumberDeck.masters[0].slideNumber = undefined;
slideNumberDeck.layouts[0].slideNumber = { x: 200, align: 'center' };
slideNumberDeck.masters[0].slideNumber = { x: 300, align: 'right' };
const packedNumberDuplicate = slideNumberDeck.duplicateSlide(0);
slideNumberDeck.moveSlide(slideNumberDeck.slides.indexOf(packedNumberDuplicate), 0);
slideNumberDeck.deleteSlide(slideNumberDeck.slides.findIndex(
  ({ partUri }) => partUri === packedNumberSource.partUri,
));
const reopenedSlideNumberDeck = await PptxDocument.open(await slideNumberDeck.write());
await reopenedSlideNumberDeck.write({ compatibility: 'powerpoint-2010' });
const packedSlideNumberXml = (partUri) => new TextDecoder().decode(
  reopenedSlideNumberDeck.opcPackage.requirePart(partUri).bytes,
);
const packedSlideNumberCache = (partUri) => {
  const xml = packedSlideNumberXml(partUri);
  const fieldStart = xml.indexOf('type="slidenum"');
  const textStart = xml.indexOf('<a:t>', fieldStart);
  const textEnd = xml.indexOf('</a:t>', textStart);
  return fieldStart < 0 || textStart < 0 || textEnd < 0
    ? undefined
    : xml.slice(textStart + 5, textEnd);
};
const packedNumberValues = reopenedSlideNumberDeck.slides.map(({ slideNumber }) => slideNumber);
const packedLayoutNumber = reopenedSlideNumberDeck.layouts[0].slideNumber;
const packedMasterNumber = reopenedSlideNumberDeck.masters[0].slideNumber;
const packedMasterXml = packedSlideNumberXml(reopenedSlideNumberDeck.masters[0].partUri);
const packedOwnerCount = (partUri) => packedSlideNumberXml(partUri).split('type="sldNum"').length - 1;
const slideNumbers = reopenedSlideNumberDeck.firstSlideNumber === 5 &&
  reopenedSlideNumberDeck.slides.length === 2 &&
  packedNumberValues[0]?.align === 'justify' && packedNumberValues[0].rtl === true &&
  packedNumberValues[0].valign === 'middle' && packedNumberValues[0].margin?.left === 4 &&
  packedNumberValues[0].style.fontFamily === 'Aptos' &&
  packedNumberValues[0].style.fontSize === 18 && packedNumberValues[0].style.lang === 'zh-CN' &&
  packedNumberValues[0].style.bold === true && packedNumberValues[0].style.italic === true &&
  packedNumberValues[0].style.color?.kind === 'srgb' &&
  packedNumberValues[0].style.color.value === 'FF3399' &&
  packedNumberValues[0].style.transparency === 20 &&
  packedNumberValues[1]?.style.color?.kind === 'scheme' &&
  packedNumberValues[1].style.color.value === 'accent1' &&
  packedLayoutNumber?.x === 200 && packedLayoutNumber.align === 'center' &&
  packedMasterNumber?.x === 300 && packedMasterNumber.align === 'right' &&
  reopenedSlideNumberDeck.slides.map(({ partUri }) => packedSlideNumberCache(partUri)).join(',') === '5,6' &&
  packedSlideNumberCache(reopenedSlideNumberDeck.layouts[0].partUri) === '‹#›' &&
  packedSlideNumberCache(reopenedSlideNumberDeck.masters[0].partUri) === '‹#›' &&
  reopenedSlideNumberDeck.slides.every(({ partUri }) => packedOwnerCount(partUri) === 1) &&
  packedOwnerCount(reopenedSlideNumberDeck.layouts[0].partUri) === 1 &&
  packedOwnerCount(reopenedSlideNumberDeck.masters[0].partUri) === 1 &&
  packedMasterXml.includes('sldNum="1"') &&
  reopenedSlideNumberDeck.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')).length === 0;
await reopenedSlideNumberDeck.writeFile('slide-number-smoke.pptx');
const slideDefaultColorDeck = PptxDocument.create();
const packedDefaultColorSource = slideDefaultColorDeck.addSlide();
packedDefaultColorSource.color = { kind: 'srgb', value: 'ff3399' };
packedDefaultColorSource.addText('Packed sRGB');
packedDefaultColorSource.color = { kind: 'scheme', value: 'accent1' };
packedDefaultColorSource.addRichText([{
  runs: [
    { text: 'Packed inherited' },
    { text: 'Packed override', style: { color: { kind: 'srgb', value: '00AA00' } } },
    { text: 'Packed alpha', style: { transparency: 25 } },
  ],
}]);
const packedDefaultColorDuplicate = slideDefaultColorDeck.duplicateSlide(0);
const packedDefaultIdentity = packedDefaultColorDuplicate.color === packedDefaultColorSource.color;
packedDefaultColorDuplicate.addText('Packed duplicate inherited');
const packedDefaultLiveValues = slideDefaultColorDeck.slides.map(({ color }) => color);
const reopenedSlideDefaultColorDeck = await PptxDocument.open(await slideDefaultColorDeck.write());
await reopenedSlideDefaultColorDeck.write({ compatibility: 'powerpoint-2010' });
const packedDefaultRunState = reopenedSlideDefaultColorDeck.slides.map((slide) => slide.shapes
  .filter((shape) => shape instanceof ShapeModel)
  .map(({ richText }) => richText.flatMap(({ runs }) => runs.map(({ style }) => ({
    color: style?.color,
    transparency: style?.transparency,
  })))));
const slideDefaultColor = packedDefaultIdentity &&
  packedDefaultLiveValues.every((color) =>
    color?.kind === 'scheme' && color.value === 'accent1') &&
  reopenedSlideDefaultColorDeck.slides.every(({ color }) => color === undefined) &&
  JSON.stringify(packedDefaultRunState) === JSON.stringify([
    [
      [{ color: { kind: 'srgb', value: 'FF3399' } }],
      [
        { color: { kind: 'scheme', value: 'accent1' } },
        { color: { kind: 'srgb', value: '00AA00' } },
        { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
      ],
    ],
    [
      [{ color: { kind: 'srgb', value: 'FF3399' } }],
      [
        { color: { kind: 'scheme', value: 'accent1' } },
        { color: { kind: 'srgb', value: '00AA00' } },
        { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
      ],
      [{ color: { kind: 'scheme', value: 'accent1' } }],
    ],
  ]) &&
  reopenedSlideDefaultColorDeck.diagnostics.filter(({ severity }) => severity === 'error').length === 0;
await reopenedSlideDefaultColorDeck.writeFile('slide-default-color-smoke.pptx');
const embeddedRasterDeck = PptxDocument.create();
const embeddedRasterSlide = embeddedRasterDeck.addSlide();
const embeddedRasterInputs = [
  { contentType: 'image/png', bytes: new Uint8Array([1, 2, 3]), name: 'Packed PNG' },
  { contentType: 'image/jpeg', bytes: new Uint8Array([4, 5, 6]), name: 'Packed JPEG' },
  { contentType: 'image/gif', bytes: new Uint8Array([7, 8, 9]), name: 'Packed GIF' },
];
const embeddedRasterCreated = embeddedRasterInputs.map((input, index) =>
  embeddedRasterSlide.addImage(input.bytes, {
    contentType: input.contentType,
    name: input.name,
    altText: input.name + ' alt',
    x: inches(index + 1),
    y: inches(index + 2),
    width: inches(2),
    height: inches(1),
  }));
const packedBytesEqual = (left, right) => left.length === right.length &&
  left.every((value, index) => value === right[index]);
const embeddedRasterImmediate = embeddedRasterCreated.every((image, index) =>
  image instanceof ImageModel &&
  embeddedRasterSlide.shapes[index] === image &&
  image.name === embeddedRasterInputs[index].name &&
  embeddedRasterDeck.opcPackage.requirePart(image.sourcePartUri).contentType ===
    embeddedRasterInputs[index].contentType &&
  packedBytesEqual(
    embeddedRasterDeck.opcPackage.requirePart(image.sourcePartUri).bytes,
    embeddedRasterInputs[index].bytes,
  ));
const embeddedRasterPngUri = embeddedRasterCreated[0].sourcePartUri;
embeddedRasterCreated[0].setTransform({ x: inches(4), rotation: 2_700_000 });
embeddedRasterCreated[0].replaceData(new Uint8Array([10, 11]), 'image/png');
const embeddedRasterExclusive = embeddedRasterCreated[0].sourcePartUri === embeddedRasterPngUri &&
  packedBytesEqual(
    embeddedRasterDeck.opcPackage.requirePart(embeddedRasterPngUri).bytes,
    new Uint8Array([10, 11]),
  );
const embeddedRasterDuplicateSlide = embeddedRasterDeck.duplicateSlide(0);
const embeddedRasterDuplicateImages = embeddedRasterDuplicateSlide.shapes.filter(
  (shape) => shape instanceof ImageModel,
);
const embeddedRasterShared = embeddedRasterDuplicateImages.every(
  (image, index) => image.sourcePartUri === embeddedRasterCreated[index].sourcePartUri,
);
embeddedRasterDuplicateImages[0].replaceData(new Uint8Array([12, 13]), 'image/png');
const embeddedRasterDuplicateUri = embeddedRasterDuplicateImages[0].sourcePartUri;
const embeddedRasterCloneOnWrite = embeddedRasterDuplicateUri !== embeddedRasterPngUri &&
  packedBytesEqual(
    embeddedRasterDeck.opcPackage.requirePart(embeddedRasterPngUri).bytes,
    new Uint8Array([10, 11]),
  ) &&
  packedBytesEqual(
    embeddedRasterDeck.opcPackage.requirePart(embeddedRasterDuplicateUri).bytes,
    new Uint8Array([12, 13]),
  );
const reopenedEmbeddedRasterDeck = await PptxDocument.open(await embeddedRasterDeck.write());
const reopenedEmbeddedRasterSource = reopenedEmbeddedRasterDeck.slides[0].shapes.filter(
  (shape) => shape instanceof ImageModel,
);
const reopenedEmbeddedRasterDuplicate = reopenedEmbeddedRasterDeck.slides[1].shapes.filter(
  (shape) => shape instanceof ImageModel,
);
const reopenedEmbeddedRasterSourceBytes = [
  new Uint8Array([10, 11]),
  embeddedRasterInputs[1].bytes,
  embeddedRasterInputs[2].bytes,
];
const reopenedEmbeddedRasterDuplicateBytes = [
  new Uint8Array([12, 13]),
  embeddedRasterInputs[1].bytes,
  embeddedRasterInputs[2].bytes,
];
const embeddedRasterReopened = reopenedEmbeddedRasterSource.length === 3 &&
  reopenedEmbeddedRasterDuplicate.length === 3 &&
  reopenedEmbeddedRasterDeck.slides[0].shapes[0] === reopenedEmbeddedRasterSource[0] &&
  reopenedEmbeddedRasterDeck.slides[1].shapes[0] === reopenedEmbeddedRasterDuplicate[0] &&
  reopenedEmbeddedRasterSource[0].sourcePartUri === embeddedRasterPngUri &&
  reopenedEmbeddedRasterDuplicate[0].sourcePartUri === embeddedRasterDuplicateUri &&
  reopenedEmbeddedRasterSource[0].transform.x === inches(4) &&
  reopenedEmbeddedRasterSource[0].transform.rotation === 2_700_000 &&
  embeddedRasterInputs.every((input, index) => {
    const sourcePart = reopenedEmbeddedRasterDeck.opcPackage
      .requirePart(reopenedEmbeddedRasterSource[index].sourcePartUri);
    const duplicatePart = reopenedEmbeddedRasterDeck.opcPackage
      .requirePart(reopenedEmbeddedRasterDuplicate[index].sourcePartUri);
    return sourcePart.contentType === input.contentType &&
      duplicatePart.contentType === input.contentType &&
      packedBytesEqual(sourcePart.bytes, reopenedEmbeddedRasterSourceBytes[index]) &&
      packedBytesEqual(duplicatePart.bytes, reopenedEmbeddedRasterDuplicateBytes[index]);
  }) &&
  reopenedEmbeddedRasterDeck.slides.every((slide) =>
    slide.shapes.filter((shape) => shape instanceof ImageModel).every((image) =>
      slide.relationships.some(({ type, targetMode, resolvedTarget }) =>
        type.endsWith('/image') &&
        targetMode === 'Internal' &&
        resolvedTarget === image.sourcePartUri)));
const embeddedRasterImages = embeddedRasterImmediate && embeddedRasterExclusive &&
  embeddedRasterShared && embeddedRasterCloneOnWrite && embeddedRasterReopened;
const packedSvgDeck = PptxDocument.create();
const packedSvgSlide = packedSvgDeck.addSlide();
const packedSvgBytes = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360">' +
  '<rect width="640" height="360" fill="#4472C4"/></svg>',
);
const packedFallbackPng = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
));
const packedSvgInfo = inspectSvgImage(packedSvgBytes);
const packedGenericSvgInfo = inspectImage(packedSvgBytes);
const packedSvgSizing = calculateImageSizing(packedSvgInfo, {
  type: 'cover',
  width: inches(4),
  height: inches(3),
});
const packedExplicitSvg = await packedSvgDeck.addImage(0, packedSvgBytes, {
  fallback: packedFallbackPng,
  sizing: {
    type: 'cover',
    width: inches(4),
    height: inches(3),
  },
  name: 'Packed explicit SVG',
  altText: 'Explicit PNG fallback',
  x: inches(1),
  y: inches(1),
  rotation: 900_000,
  flipHorizontal: true,
});
const packedSvgDataUri = 'data:image/svg+xml;base64,' + Buffer.from(packedSvgBytes).toString('base64');
const packedDefaultSvg = await packedSvgDeck.addImage(0, packedSvgDataUri, {
  name: 'Packed default SVG',
  altText: 'Automatic PNG fallback',
  x: inches(5),
  y: inches(1),
  width: inches(3),
  height: inches(2),
});
const packedSvgPairIsValid = (deck, slide, image) => {
  if (!(image instanceof ImageModel) || !image.isSvg ||
      image.fallbackPartUri === undefined || image.svgPartUri === undefined ||
      image.sourcePartUri !== image.fallbackPartUri) return false;
  const fallbackPart = deck.opcPackage.requirePart(image.fallbackPartUri);
  const svgPart = deck.opcPackage.requirePart(image.svgPartUri);
  const targetUris = [image.fallbackPartUri, image.svgPartUri];
  return fallbackPart.contentType === 'image/png' &&
    inspectRasterImage(fallbackPart.bytes).contentType === 'image/png' &&
    svgPart.contentType === 'image/svg+xml' &&
    targetUris.every((target) => slide.relationships.some(
      ({ type, targetMode, resolvedTarget }) => type.endsWith('/image') &&
        targetMode === 'Internal' && resolvedTarget === target,
    ));
};
const packedSvgImmediate = packedSvgSlide.shapes[0] === packedExplicitSvg &&
  packedSvgSlide.shapes[1] === packedDefaultSvg &&
  packedSvgInfo.contentType === 'image/svg+xml' && packedSvgInfo.width === 640 &&
  packedSvgInfo.height === 360 && packedGenericSvgInfo.contentType === 'image/svg+xml' &&
  packedSvgSizing.sourceRectangle.left === 12.5 &&
  packedSvgPairIsValid(packedSvgDeck, packedSvgSlide, packedExplicitSvg) &&
  packedSvgPairIsValid(packedSvgDeck, packedSvgSlide, packedDefaultSvg);
let packedLowLevelSvgRejected = false;
try {
  packedSvgSlide.addImage(packedSvgBytes, { contentType: 'image/svg+xml' });
} catch {
  packedLowLevelSvgRejected = true;
}
let packedInvalidFallbackRejected = false;
try {
  await packedSvgDeck.addImage(0, packedSvgBytes, {
    fallback: Uint8Array.from([71, 73, 70, 56, 57, 97, 1, 0, 1, 0]),
  });
} catch {
  packedInvalidFallbackRejected = true;
}
const packedSvgUnchangedAfterFailures = packedSvgSlide.shapes.length === 2;
const packedSvgDuplicateSlide = packedSvgDeck.duplicateSlide(0);
const packedSvgDuplicateImages = packedSvgDuplicateSlide.shapes.filter(
  (shape) => shape instanceof ImageModel,
);
const packedSvgShared = packedSvgDuplicateImages.length === 2 &&
  packedSvgDuplicateImages.every((image, index) =>
    image.fallbackPartUri === packedSvgSlide.shapes[index].fallbackPartUri &&
    image.svgPartUri === packedSvgSlide.shapes[index].svgPartUri);
const packedSourceFallbackUri = packedExplicitSvg.fallbackPartUri;
const packedSourceSvgUri = packedExplicitSvg.svgPartUri;
const packedReplacementSvg = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"/>',
);
packedSvgDuplicateImages[0].replaceSvgData(packedReplacementSvg, packedFallbackPng);
const packedSvgCloneOnWrite = packedSvgDuplicateImages[0].fallbackPartUri !== packedSourceFallbackUri &&
  packedSvgDuplicateImages[0].svgPartUri !== packedSourceSvgUri &&
  packedBytesEqual(packedSvgDeck.opcPackage.requirePart(packedSourceSvgUri).bytes, packedSvgBytes) &&
  packedBytesEqual(
    packedSvgDeck.opcPackage.requirePart(packedSvgDuplicateImages[0].svgPartUri).bytes,
    packedReplacementSvg,
  ) && packedSvgPairIsValid(packedSvgDeck, packedSvgDuplicateSlide, packedSvgDuplicateImages[0]);
const reopenedPackedSvgDeck = await PptxDocument.open(await packedSvgDeck.write());
const reopenedPackedSvgImages = reopenedPackedSvgDeck.slides.flatMap((slide) =>
  slide.shapes.filter((shape) => shape instanceof ImageModel));
const packedSvgReopened = reopenedPackedSvgImages.length === 4 &&
  reopenedPackedSvgDeck.slides.every((slide) => slide.shapes.every((shape) =>
    !(shape instanceof ImageModel) || packedSvgPairIsValid(reopenedPackedSvgDeck, slide, shape))) &&
  packedBytesEqual(
    reopenedPackedSvgDeck.opcPackage.requirePart(
      reopenedPackedSvgDeck.slides[1].shapes[0].svgPartUri,
    ).bytes,
    packedReplacementSvg,
  );
await packedSvgDeck.writeFile('svg-smoke.pptx');
const svgImages = packedSvgImmediate && packedLowLevelSvgRejected &&
  packedInvalidFallbackRejected && packedSvgUnchangedAfterFailures && packedSvgShared &&
  packedSvgCloneOnWrite && packedSvgReopened;
const mediaDeck = PptxDocument.create();
const mediaSlide = mediaDeck.addSlide();
const mediaPath = (await import('node:url')).fileURLToPath(
  new URL('./packed-path-audio.mp3', import.meta.url),
);
const sharedAudioBytes = Uint8Array.of(1, 2, 3, 4);
await import('node:fs/promises').then(({ writeFile }) => writeFile(mediaPath, sharedAudioBytes));
const mediaPngPoster = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
));
const mediaJpegPoster = Uint8Array.of(255, 216, 255, 224, 0, 16, 74, 70, 73, 70, 255, 217);
const deletedSharedAudio = await mediaDeck.addAudio(
  0,
  'data:audio/mpeg;base64,AQIDBA==',
  {
    name: 'Packed audio & data',
    altText: 'Deleted shared narration',
    poster: mediaPngPoster,
    posterContentType: 'image/png',
  },
);
const pathAudio = await mediaDeck.addAudio(0, mediaPath, {
  name: 'Packed path audio',
  altText: 'Shared path narration',
  poster: mediaPngPoster,
  posterContentType: 'image/png',
  x: inches(1),
  y: inches(1.5),
  width: inches(2),
  height: inches(1),
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.25,
});
const bytesAudio = await mediaDeck.addAudio(0, Uint8Array.of(5, 6, 7), {
  name: 'Packed bytes audio',
  altText: '',
  contentType: 'audio/wav',
  fileName: 'packed-narration.wav',
  poster: mediaJpegPoster,
  posterContentType: 'image/jpeg',
  x: inches(3.5),
  y: inches(1.5),
  width: inches(2),
  height: inches(1),
});
const blobAudio = await mediaDeck.addAudio(
  0,
  new Blob([Uint8Array.of(8, 9, 10)], { type: 'audio/mp4' }),
  {
    name: 'Packed Blob audio',
    contentType: 'audio/mp4',
    fileName: 'packed-blob.m4a',
    x: inches(6),
    y: inches(1.5),
    width: inches(2),
    height: inches(1),
  },
);
const streamAudio = await mediaDeck.addAudio(
  0,
  new ReadableStream({
    start(controller) {
      controller.enqueue(Uint8Array.of(11, 12));
      controller.enqueue(Uint8Array.of(13));
      controller.close();
    },
  }),
  {
    name: 'Packed stream audio',
    contentType: 'audio/mpeg',
    fileName: 'packed-stream.mp3',
    x: inches(8.5),
    y: inches(1.5),
    width: inches(2),
    height: inches(1),
  },
);
const dataVideo = await mediaDeck.addVideo(0, 'data:video/mp4;base64,Dg8Q', {
  name: 'Packed data video',
  altText: 'Video from a data URI',
  poster: 'data:image/png;base64,' + Buffer.from(mediaPngPoster).toString('base64'),
  x: inches(1),
  y: inches(3),
  width: inches(4),
  height: inches(2.25),
  play: 'click',
  volume: 0.75,
});
const bytesVideo = await mediaDeck.addVideo(0, Uint8Array.of(17, 18, 19), {
  name: 'Packed bytes video',
  altText: 'MOV bytes with JPEG poster',
  contentType: 'video/quicktime',
  fileName: 'packed-bytes.mov',
  poster: mediaJpegPoster,
  posterContentType: 'image/jpeg',
  x: inches(5.5),
  y: inches(3),
  width: inches(4),
  height: inches(2.25),
});
const packedMediaDeduplicated = deletedSharedAudio.mediaPartUri === pathAudio.mediaPartUri &&
  deletedSharedAudio.posterPartUri === pathAudio.posterPartUri;
new MediaCodec(mediaDeck.opcPackage).delete(mediaSlide.partUri, deletedSharedAudio.shapeId);
const packedMediaSharedPreserved = mediaDeck.opcPackage.hasPart(pathAudio.mediaPartUri) &&
  mediaDeck.opcPackage.hasPart(pathAudio.posterPartUri) && mediaDeck.media(0).length === 6;
const packedMediaExpected = [
  [pathAudio, 'audio', 'audio/mpeg', '.mp3', 'image/png', '.png'],
  [bytesAudio, 'audio', 'audio/wav', '.wav', 'image/jpeg', '.jpg'],
  [blobAudio, 'audio', 'audio/mp4', '.m4a', 'image/png', '.png'],
  [streamAudio, 'audio', 'audio/mpeg', '.mp3', 'image/png', '.png'],
  [dataVideo, 'video', 'video/mp4', '.mp4', 'image/png', '.png'],
  [bytesVideo, 'video', 'video/quicktime', '.mov', 'image/jpeg', '.jpg'],
];
const packedMediaRelationshipsValid = (deck, slide, model) => {
  const relationships = slide.relationships;
  const mediaRelationships = relationships.filter(
    ({ resolvedTarget }) => resolvedTarget === model.mediaPartUri,
  );
  const posterRelationships = relationships.filter(
    ({ resolvedTarget }) => resolvedTarget === model.posterPartUri,
  );
  return mediaRelationships.some(({ type, targetMode }) =>
    type.endsWith('/' + model.kind) && targetMode === 'Internal') &&
    mediaRelationships.some(({ type, targetMode }) =>
      type === 'http://schemas.microsoft.com/office/2007/relationships/media' &&
      targetMode === 'Internal') &&
    posterRelationships.some(({ type, targetMode }) =>
      type.endsWith('/image') && targetMode === 'Internal') &&
    deck.opcPackage.hasPart(model.mediaPartUri) && deck.opcPackage.hasPart(model.posterPartUri);
};
const packedMediaImmediate = packedMediaExpected.every(
  ([model, kind, mediaType, mediaExtension, posterType, posterExtension]) =>
    model.kind === kind && model.mediaPartUri.endsWith(mediaExtension) &&
    mediaDeck.opcPackage.requirePart(model.mediaPartUri).contentType === mediaType &&
    model.posterPartUri.endsWith(posterExtension) &&
    mediaDeck.opcPackage.requirePart(model.posterPartUri).contentType === posterType &&
    packedMediaRelationshipsValid(mediaDeck, mediaSlide, model),
);
const mediaXml = new TextDecoder().decode(mediaDeck.opcPackage.requirePart(mediaSlide.partUri).bytes);
const packedMediaXmlValid = (mediaXml.match(/<a:audioFile\\b/g) ?? []).length === 4 &&
  (mediaXml.match(/<a:videoFile\\b/g) ?? []).length === 2 &&
  mediaXml.includes('name="Packed path audio"') &&
  mediaXml.includes('descr="Shared path narration"') &&
  mediaXml.includes('x="914400" y="1371600"') &&
  mediaXml.includes('cx="1828800" cy="914400"');
await mediaDeck.writeFile('media-smoke.pptx');
const reopenedMediaDeck = await PptxDocument.open(await mediaDeck.write());
const reopenedMedia = [...reopenedMediaDeck.media(0)].sort((left, right) => left.shapeId - right.shapeId);
const reopenedMediaXml = new TextDecoder().decode(
  reopenedMediaDeck.opcPackage.requirePart(reopenedMediaDeck.slides[0].partUri).bytes,
);
const packedMediaReopened = reopenedMedia.length === 6 && reopenedMedia.every((model, index) => {
  const expected = packedMediaExpected[index];
  return model.kind === expected[1] && model.mediaPartUri.endsWith(expected[3]) &&
    reopenedMediaDeck.opcPackage.requirePart(model.mediaPartUri).contentType === expected[2] &&
    model.posterPartUri.endsWith(expected[5]) &&
    reopenedMediaDeck.opcPackage.requirePart(model.posterPartUri).contentType === expected[4] &&
    packedMediaRelationshipsValid(reopenedMediaDeck, reopenedMediaDeck.slides[0], model);
}) && reopenedMedia[0].settings.play === 'auto' && reopenedMedia[0].settings.loop === true &&
  reopenedMedia[0].settings.hideWhenStopped === true && reopenedMedia[0].settings.volume === 0.25 &&
  reopenedMedia[4].settings.play === 'click' && reopenedMedia[4].settings.volume === 0.75 &&
  (reopenedMediaXml.match(/<a:audioFile\\b/g) ?? []).length === 4 &&
  (reopenedMediaXml.match(/<a:videoFile\\b/g) ?? []).length === 2;
const embeddedMedia = packedMediaDeduplicated && packedMediaSharedPreserved &&
  packedMediaImmediate && packedMediaXmlValid && packedMediaReopened;
if (!embeddedMedia) {
  throw new Error(JSON.stringify({
    packedMediaDeduplicated,
    packedMediaSharedPreserved,
    packedMediaImmediate,
    packedMediaXmlValid,
    packedMediaReopened,
    mediaXml,
    reopenedMediaXml,
    reopenedMedia,
  }));
}
const stableMediaDeck = PptxDocument.create();
const stableMediaSlide = stableMediaDeck.addSlide();
const stableAudio = await stableMediaDeck.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
  name: 'Stable packed audio',
  altText: 'Stable packed narration',
  poster: mediaPngPoster,
  posterContentType: 'image/png',
});
const stableVideo = await stableMediaDeck.addVideo(
  0,
  new Blob([Uint8Array.of(5, 6, 7, 8)], { type: 'video/mp4' }),
  {
    name: 'Stable packed video',
    contentType: 'video/mp4',
    poster: mediaJpegPoster,
    posterContentType: 'image/jpeg',
  },
);
stableVideo.settings = { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25 };
stableVideo.settings = undefined;
const stableClearedVideoXml = new TextDecoder().decode(
  stableMediaDeck.opcPackage.requirePart(stableMediaSlide.partUri).bytes,
);
const stableClearedVideoPicture = [...stableClearedVideoXml.matchAll(
  /<p:pic(?:\\s[^>]*)?>[\\s\\S]*?<\\/p:pic>/g,
)].find((match) => match[0].includes('id="' + stableVideo.shapeId + '"'))?.[0] ?? '';
const stableVideoSettingsCleared = Object.keys(stableVideo.settings).length === 0 &&
  !stableClearedVideoPicture.includes('<px:playback') &&
  ![...stableClearedVideoXml.matchAll(/<p:video>[\\s\\S]*?<\\/p:video>/g)]
    .some((match) => match[0].includes('spid="' + stableVideo.shapeId + '"'));
const stableIdentity = stableAudio instanceof MediaModel && stableVideo instanceof MediaModel &&
  stableMediaDeck.media(0)[0] === stableAudio && stableMediaSlide.media[0] === stableAudio &&
  stableMediaSlide.shapes[0] === stableAudio;
const initialStableAudioTarget = stableAudio.mediaPartUri;
stableAudio.name = 'Stable packed audio edited';
stableAudio.altText = undefined;
stableAudio.settings = { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.5 };
stableAudio.setTransform({
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(1),
});
await stableAudio.replaceSource('https://example.com/stable-audio.wav');
const stableExternalTransition = stableAudio.externalUrl ===
  'https://example.com/stable-audio.wav' && !stableMediaDeck.opcPackage.hasPart(initialStableAudioTarget);
await stableAudio.replaceSource(
  new ReadableStream({
    start(controller) {
      controller.enqueue(Uint8Array.of(21));
      controller.enqueue(Uint8Array.of(22));
      controller.close();
    },
  }),
  { contentType: 'audio/wav' },
);
await stableAudio.replacePoster(
  new Blob([Uint8Array.of(71, 73, 70, 56, 57, 97)], { type: 'image/gif' }),
  { contentType: 'image/gif' },
);
const replacedStablePoster = stableAudio.posterPartUri;
const stablePosterReplacement = stableMediaDeck.opcPackage
  .requirePart(replacedStablePoster).contentType === 'image/gif';
await stableAudio.replacePoster();
const stablePosterReset = stableAudio.posterPartUri.endsWith('.png') &&
  stableMediaDeck.opcPackage.requirePart(stableAudio.posterPartUri).contentType === 'image/png' &&
  !stableMediaDeck.opcPackage.hasPart(replacedStablePoster);
const stableDuplicate = stableMediaDeck.duplicateSlide(0);
const stableDuplicateAudio = stableDuplicate.media[0];
const stableDuplicateVideo = stableDuplicate.media[1];
const stableSharedBeforeWrite = stableDuplicateAudio.mediaPartUri === stableAudio.mediaPartUri &&
  stableDuplicateAudio.posterPartUri === stableAudio.posterPartUri &&
  stableDuplicateVideo.mediaPartUri === stableVideo.mediaPartUri;
const stableSourceBeforeDuplicateSettings = new TextDecoder().decode(
  stableMediaDeck.opcPackage.requirePart(stableMediaSlide.partUri).bytes,
);
stableDuplicateAudio.settings = {
  play: 'click',
  loop: false,
  hideWhenStopped: false,
  volume: 0.25,
};
const stableDuplicateTimingIsolation = new TextDecoder().decode(
  stableMediaDeck.opcPackage.requirePart(stableMediaSlide.partUri).bytes,
) === stableSourceBeforeDuplicateSettings && stableDuplicateAudio.settings.play === 'click' &&
  stableDuplicateAudio.settings.volume === 0.25;
await stableDuplicateAudio.replaceSource(mediaPath);
await stableDuplicateAudio.replacePoster(mediaJpegPoster, { contentType: 'image/jpeg' });
const stableDuplicateMediaTarget = stableDuplicateAudio.mediaPartUri;
const stableDuplicatePosterTarget = stableDuplicateAudio.posterPartUri;
const stableCloneOnWrite = stableDuplicateMediaTarget !== stableAudio.mediaPartUri &&
  stableDuplicatePosterTarget !== stableAudio.posterPartUri;
const stablePosterDedup = stableDuplicatePosterTarget === stableVideo.posterPartUri;
const sharedVideoTarget = stableVideo.mediaPartUri;
stableDuplicateVideo.remove();
const stableObjectIsolation = stableMediaDeck.opcPackage.hasPart(sharedVideoTarget) &&
  stableMediaDeck.media(0)[1] === stableVideo && stableDuplicate.media.length === 1;
stableMediaDeck.moveSlide(1, 0);
const stableMoveIdentity = stableMediaDeck.slides[0] === stableDuplicate &&
  stableDuplicate.media[0] === stableDuplicateAudio;
stableMediaDeck.moveSlide(0, 1);
stableMediaDeck.deleteSlide(1);
const stableSlideGc = !stableMediaDeck.opcPackage.hasPart(stableDuplicateMediaTarget) &&
  stableMediaDeck.opcPackage.hasPart(stableDuplicatePosterTarget);
const stableVideoPosterTarget = stableVideo.posterPartUri;
stableVideo.remove();
const stableObjectGc = !stableMediaDeck.opcPackage.hasPart(sharedVideoTarget) &&
  !stableMediaDeck.opcPackage.hasPart(stableVideoPosterTarget);
await stableMediaDeck.writeFile('stable-media-smoke.pptx');
const reopenedStableMediaDeck = await PptxDocument.open(await stableMediaDeck.write());
const reopenedStableAudio = reopenedStableMediaDeck.media(0)[0];
await reopenedStableMediaDeck.write({ mode: 'permissive', compatibility: 'powerpoint-2010' });
const reopenedStableXml = new TextDecoder().decode(
  reopenedStableMediaDeck.opcPackage.requirePart(reopenedStableMediaDeck.slides[0].partUri).bytes,
);
const reopenedStableTimingIds = [...reopenedStableXml.matchAll(
  /<p:cTn\\b[^>]*\\bid="([0-9]+)"/g,
)].map((match) => Number(match[1]));
const reopenedStableTargets = [...reopenedStableXml.matchAll(
  /<p:spTgt\\b[^>]*\\bspid="([0-9]+)"/g,
)].map((match) => Number(match[1]));
const nativeMediaTiming = stableVideoSettingsCleared && stableDuplicateTimingIsolation &&
  reopenedStableXml.includes('<px:playback xmlns:px="urn:pptx-ooxml:media" ' +
    'play="auto" loop="1" hideWhenStopped="1" volume="50000"') &&
  reopenedStableXml.includes('<p:cMediaNode vol="50000" showWhenStopped="0">') &&
  reopenedStableXml.includes('repeatCount="indefinite"') &&
  reopenedStableXml.includes('cmd="playFrom(0.0)"') &&
  new Set(reopenedStableTimingIds).size === reopenedStableTimingIds.length &&
  reopenedStableTargets.length > 0 && reopenedStableTargets.every(
    (target) => target === reopenedStableAudio.shapeId,
  ) && reopenedStableMediaDeck.diagnostics.every(
    ({ code }) => !code.startsWith('MEDIA_TIMING_'),
  );
const stableNoOrphans = reopenedStableMediaDeck.opcPackage.parts
  .filter(({ uri }) => uri.startsWith('/ppt/media/'))
  .every(({ uri }) =>
    (reopenedStableMediaDeck.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) > 0);
const stableMediaLifecycle = stableIdentity && stableExternalTransition && stablePosterReplacement &&
  stablePosterReset && stableVideoSettingsCleared && stableSharedBeforeWrite &&
  stableDuplicateTimingIsolation && stableCloneOnWrite && stablePosterDedup &&
  stableObjectIsolation &&
  stableMoveIdentity && stableSlideGc && stableObjectGc && stableNoOrphans &&
  reopenedStableMediaDeck.slides.length === 1 && reopenedStableMediaDeck.media(0).length === 1 &&
  reopenedStableAudio instanceof MediaModel && reopenedStableAudio.name === 'Stable packed audio edited' &&
  reopenedStableAudio.altText === undefined && reopenedStableAudio.mediaPartUri.endsWith('.wav') &&
  reopenedStableMediaDeck.opcPackage.requirePart(reopenedStableAudio.mediaPartUri).contentType ===
    'audio/wav' && reopenedStableAudio.posterPartUri.endsWith('.png') &&
  reopenedStableAudio.settings.play === 'auto' && reopenedStableAudio.settings.loop === true &&
  reopenedStableAudio.settings.hideWhenStopped === true && reopenedStableAudio.settings.volume === 0.5 &&
  reopenedStableAudio.transform.x === inches(1) && reopenedStableAudio.transform.y === inches(2) &&
  reopenedStableMediaDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!stableMediaLifecycle) {
  throw new Error(JSON.stringify({
    stableIdentity,
    stableExternalTransition,
    stablePosterReplacement,
    stablePosterReset,
    stableVideoSettingsCleared,
    stableSharedBeforeWrite,
    stableDuplicateTimingIsolation,
    stableCloneOnWrite,
    stablePosterDedup,
    stableObjectIsolation,
    stableMoveIdentity,
    stableSlideGc,
    stableObjectGc,
    stableNoOrphans,
    reopenedSlideCount: reopenedStableMediaDeck.slides.length,
    reopenedMediaCount: reopenedStableMediaDeck.media(0).length,
    reopenedKind: reopenedStableAudio?.kind,
    reopenedName: reopenedStableAudio?.name,
    reopenedAltText: reopenedStableAudio?.altText,
    reopenedMediaPartUri: reopenedStableAudio?.mediaPartUri,
    reopenedPosterPartUri: reopenedStableAudio?.posterPartUri,
    reopenedSettings: reopenedStableAudio?.settings,
    reopenedTransform: reopenedStableAudio?.transform,
    diagnostics: reopenedStableMediaDeck.diagnostics,
    nativeMediaTiming,
  }));
}
const shapeDeck = PptxDocument.create();
const shapeSlide = shapeDeck.addSlide();
const defaultShape = shapeSlide.addShape('rect');
const customShape = shapeSlide.addShape('foldedCorner', {
  name: 'Packed folded corner',
  x: inches(2),
  y: inches(3),
  width: inches(4),
  height: inches(2),
  rotation: 2_700_000,
  flipHorizontal: true,
  flipVertical: true,
});
const initialDefaultPreset = defaultShape.presetType;
const initialCustomPreset = customShape.presetType;
const shapeIdentity = shapeSlide.shapes[0] === defaultShape && shapeSlide.shapes[1] === customShape;
const shapePartCountBeforeEdit = shapeDeck.opcPackage.parts.length;
const shapeRelationshipCountBeforeEdit = shapeSlide.relationships.length;
defaultShape.presetType = 'ellipse';
customShape.presetType = 'roundRect';
const shapeEditIsolation = shapeDeck.opcPackage.parts.length === shapePartCountBeforeEdit &&
  shapeSlide.relationships.length === shapeRelationshipCountBeforeEdit;
const duplicateShapeSlide = shapeDeck.duplicateSlide(0);
const duplicateDefaultShape = duplicateShapeSlide.shapes[0];
if (!(duplicateDefaultShape instanceof ShapeModel)) throw new Error('Packed duplicate shape failed');
duplicateDefaultShape.presetType = 'star5';
const reopenedShapeDeck = await PptxDocument.open(await shapeDeck.write());
const reopenedShapeTypes = reopenedShapeDeck.slides.map((slide) =>
  slide.shapes.map((shape) => shape instanceof ShapeModel ? shape.presetType : undefined));
const reopenedShapeNames = reopenedShapeDeck.slides.map((slide) =>
  slide.shapes.map(({ name }) => name));
const presetShapes = PRESET_SHAPE_TYPES.length === 178 &&
  Object.isFrozen(PRESET_SHAPE_TYPES) &&
  PRESET_SHAPE_TYPES.includes('foldedCorner') &&
  !PRESET_SHAPE_TYPES.includes('folderCorner') &&
  defaultShape instanceof ShapeModel &&
  customShape instanceof ShapeModel &&
  initialDefaultPreset === 'rect' &&
  initialCustomPreset === 'foldedCorner' &&
  shapeIdentity &&
  shapeEditIsolation &&
  defaultShape.transform.x === inches(1) &&
  defaultShape.transform.y === inches(1) &&
  defaultShape.transform.width === inches(1) &&
  defaultShape.transform.height === inches(1) &&
  defaultShape.presetType === 'ellipse' &&
  customShape.name === 'Packed folded corner' &&
  customShape.transform.x === inches(2) &&
  customShape.transform.y === inches(3) &&
  customShape.transform.width === inches(4) &&
  customShape.transform.height === inches(2) &&
  customShape.transform.rotation === 2_700_000 &&
  customShape.transform.flipHorizontal === true &&
  customShape.transform.flipVertical === true &&
  customShape.presetType === 'roundRect' &&
  duplicateDefaultShape.presetType === 'star5' &&
  JSON.stringify(reopenedShapeTypes) === JSON.stringify([
    ['ellipse', 'roundRect'],
    ['star5', 'roundRect'],
  ]) &&
  JSON.stringify(reopenedShapeNames) === JSON.stringify([
    ['Shape 2', 'Packed folded corner'],
    ['Shape 2', 'Packed folded corner'],
  ]);
const allCommandGeometry = {
  paths: [{
    width: 400,
    height: 300,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 100, y: 0 } },
      {
        kind: 'arcTo',
        widthRadius: 50,
        heightRadius: 25,
        startAngle: 1_800_000,
        sweepAngle: 7_200_000,
      },
      {
        kind: 'quadraticBezierTo',
        control: { x: 150, y: 50 },
        end: { x: 200, y: 100 },
      },
      {
        kind: 'cubicBezierTo',
        control1: { x: 225, y: 125 },
        control2: { x: 275, y: 175 },
        end: { x: 300, y: 200 },
      },
      { kind: 'close' },
    ],
    fill: 'none',
    stroke: false,
    extrusionOk: false,
  }],
};
const multiPathGeometry = {
  paths: [
    { width: 20, height: 10, commands: [] },
    {
      width: 40,
      height: 30,
      commands: [{ kind: 'moveTo', point: { x: 5, y: 6 } }],
      fill: 'lightenLess',
    },
  ],
};
const editedGeometry = {
  paths: [{
    width: 40,
    height: 30,
    commands: [
      { kind: 'moveTo', point: { x: 1, y: 2 } },
      { kind: 'lineTo', point: { x: 3, y: 4 } },
    ],
    fill: 'darkenLess',
  }],
};
const customGeometryDeck = PptxDocument.create();
const customGeometrySlide = customGeometryDeck.addSlide();
const customGeometryShape = customGeometrySlide.addCustomShape(allCommandGeometry, {
  name: 'Packed custom geometry',
  x: inches(1),
  y: inches(2),
  width: inches(4),
  height: inches(3),
});
const multiPathShape = customGeometrySlide.addCustomShape(multiPathGeometry, {
  name: 'Packed multi-path geometry',
});
const initialCustomGeometry = customGeometryShape.customGeometry;
const initialMultiPathGeometry = multiPathShape.customGeometry;
customGeometryShape.customGeometry = editedGeometry;
const editedCustomGeometry = customGeometryShape.customGeometry;
customGeometryShape.presetType = 'diamond';
const convertedPreset = customGeometryShape.presetType;
const clearedCustomGeometry = customGeometryShape.customGeometry;
customGeometryShape.customGeometry = allCommandGeometry;
const reopenedCustomGeometryDeck = await PptxDocument.open(await customGeometryDeck.write());
const reopenedCustomGeometryShape = reopenedCustomGeometryDeck.slides[0].shapes[0];
const reopenedMultiPathShape = reopenedCustomGeometryDeck.slides[0].shapes[1];
const customGeometryPaths =
  customGeometryShape instanceof ShapeModel &&
  Object.isFrozen(initialCustomGeometry) &&
  Object.isFrozen(initialCustomGeometry?.paths) &&
  Object.isFrozen(initialCustomGeometry?.paths[0]?.commands) &&
  JSON.stringify(initialCustomGeometry) === JSON.stringify(allCommandGeometry) &&
  Object.isFrozen(initialMultiPathGeometry) &&
  Object.isFrozen(initialMultiPathGeometry?.paths) &&
  JSON.stringify(initialMultiPathGeometry) === JSON.stringify(multiPathGeometry) &&
  JSON.stringify(editedCustomGeometry) === JSON.stringify(editedGeometry) &&
  convertedPreset === 'diamond' &&
  clearedCustomGeometry === undefined &&
  customGeometryShape.presetType === undefined &&
  reopenedCustomGeometryShape instanceof ShapeModel &&
  reopenedCustomGeometryShape.name === 'Packed custom geometry' &&
  reopenedCustomGeometryShape.presetType === undefined &&
  JSON.stringify(reopenedCustomGeometryShape.customGeometry) === JSON.stringify(allCommandGeometry) &&
  reopenedMultiPathShape instanceof ShapeModel &&
  reopenedMultiPathShape.name === 'Packed multi-path geometry' &&
  JSON.stringify(reopenedMultiPathShape.customGeometry) === JSON.stringify(multiPathGeometry);
const formulaGeometrySource = {
  adjustments: [
    { name: 'adj1', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adj2', formula: { operator: 'pin', operands: [0, 75_000, 100_000] } },
  ],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
    { name: 'y1', formula: { operator: '+-', operands: ['h', 0, 'x1'] } },
    { name: 'a1', formula: { operator: 'at2', operands: ['y1', 'x1'] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 0 } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 'a1',
        sweepAngle: 'cd2',
      },
      { kind: 'close' },
    ],
  }],
};
const formulaGeometryExpected = structuredClone(formulaGeometrySource);
const formulaGeometryReplacement = {
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [50_000] } }],
  guides: [
    { name: 'x1', formula: { operator: 'pin', operands: [0, 'adj1', 100_000] } },
    { name: 'y1', formula: { operator: 'min', operands: ['h', 'x1'] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 'y1' } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 0,
        sweepAngle: 'cd2',
      },
      { kind: 'close' },
    ],
  }],
};
const formulaGeometryDeck = PptxDocument.create();
const formulaGeometrySlide = formulaGeometryDeck.addSlide();
const formulaGeometryShape = formulaGeometrySlide.addCustomShape(formulaGeometrySource, {
  name: 'Packed guide formulas',
});
formulaGeometrySource.adjustments[0].name = 'changed';
formulaGeometrySource.adjustments[0].formula.operands[0] = 1;
formulaGeometrySource.guides.splice(0);
formulaGeometrySource.paths[0].commands.splice(0);
const initialFormulaGeometry = formulaGeometryShape.customGeometry;
const formulaNoOpBytes = formulaGeometryDeck.opcPackage
  .requirePart(formulaGeometrySlide.partUri).bytes.slice();
const formulaNoOpJournal = formulaGeometryDeck.opcPackage.mutations.length;
formulaGeometryShape.customGeometry = structuredClone(formulaGeometryExpected);
const formulaNoOpCurrent = formulaGeometryDeck.opcPackage
  .requirePart(formulaGeometrySlide.partUri).bytes;
const formulaNoOp = formulaNoOpJournal === formulaGeometryDeck.opcPackage.mutations.length &&
  formulaNoOpBytes.length === formulaNoOpCurrent.length &&
  formulaNoOpBytes.every((value, index) => value === formulaNoOpCurrent[index]);
formulaGeometryShape.customGeometry = formulaGeometryReplacement;
const editedFormulaGeometry = formulaGeometryShape.customGeometry;
formulaGeometryShape.presetType = 'diamond';
const formulaConvertedPreset = formulaGeometryShape.presetType;
const formulaConvertedCustom = formulaGeometryShape.customGeometry;
formulaGeometryShape.customGeometry = formulaGeometryReplacement;
const formulaXml = new TextDecoder().decode(
  formulaGeometryDeck.opcPackage.requirePart(formulaGeometrySlide.partUri).bytes,
);
const reopenedFormulaGeometryShape = (await PptxDocument.open(
  await formulaGeometryDeck.write(),
)).slides[0].shapes[0];
const customGeometryGuideFormulas =
  formulaGeometryShape instanceof ShapeModel &&
  Object.isFrozen(initialFormulaGeometry) &&
  Object.isFrozen(initialFormulaGeometry?.adjustments) &&
  Object.isFrozen(initialFormulaGeometry?.adjustments?.[0]) &&
  Object.isFrozen(initialFormulaGeometry?.adjustments?.[0]?.formula) &&
  Object.isFrozen(initialFormulaGeometry?.adjustments?.[0]?.formula.operands) &&
  Object.isFrozen(initialFormulaGeometry?.guides) &&
  Object.isFrozen(initialFormulaGeometry?.guides?.[0]?.formula.operands) &&
  Object.isFrozen(initialFormulaGeometry?.paths) &&
  Object.isFrozen(initialFormulaGeometry?.paths[0]?.commands) &&
  JSON.stringify(initialFormulaGeometry) === JSON.stringify(formulaGeometryExpected) &&
  formulaNoOp &&
  JSON.stringify(editedFormulaGeometry) === JSON.stringify(formulaGeometryReplacement) &&
  formulaConvertedPreset === 'diamond' &&
  formulaConvertedCustom === undefined &&
  formulaGeometryShape.presetType === undefined &&
  formulaXml.includes('<a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst>') &&
  formulaXml.includes('<a:gdLst><a:gd name="x1" fmla="pin 0 adj1 100000"/><a:gd name="y1" fmla="min h x1"/></a:gdLst>') &&
  formulaXml.includes('<a:arcTo wR="x1" hR="hd2" stAng="0" swAng="cd2"/>') &&
  reopenedFormulaGeometryShape instanceof ShapeModel &&
  reopenedFormulaGeometryShape.name === 'Packed guide formulas' &&
  JSON.stringify(reopenedFormulaGeometryShape.customGeometry) ===
    JSON.stringify(formulaGeometryReplacement);
const handleGeometrySource = {
  adjustments: [
    { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adjY', formula: { operator: 'val', operands: [50_000] } },
    { name: 'adjR', formula: { operator: 'val', operands: [30_000] } },
    { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
  ],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adjR', 100_000] } },
    { name: 'y1', formula: { operator: '*/', operands: ['h', 'adjR', 100_000] } },
  ],
  handles: [
    {
      kind: 'xy',
      position: { x: 'adjX', y: 'adjY' },
      xGuide: 'adjX',
      minX: 0,
      maxX: 100_000,
      yGuide: 'adjY',
      minY: 't',
      maxY: 'b',
    },
    {
      kind: 'polar',
      position: { x: 'x1', y: 'y1' },
      radiusGuide: 'adjR',
      minRadius: 0,
      maxRadius: 'ss',
      angleGuide: 'adjAng',
      minAngle: 0,
      maxAngle: 'cd',
    },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'adjX', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'adjY' } },
      { kind: 'close' },
    ],
  }],
};
const handleGeometryExpected = structuredClone(handleGeometrySource);
const handleGeometryReplacement = {
  ...handleGeometryExpected,
  handles: [
    {
      kind: 'polar',
      position: { x: 'hc', y: 'vc' },
      radiusGuide: 'adjR',
      minRadius: 1,
      maxRadius: 'ss',
      angleGuide: 'adjAng',
      maxAngle: '3cd4',
    },
    {
      kind: 'xy',
      position: { x: 'x1', y: 'adjY' },
      xGuide: 'adjX',
      maxX: 90_000,
      yGuide: 'adjY',
      minY: 't',
    },
  ],
};
const handleGeometryDeck = PptxDocument.create();
const handleGeometrySlide = handleGeometryDeck.addSlide();
const handleGeometryShape = handleGeometrySlide.addCustomShape(handleGeometrySource, {
  name: 'Packed adjustment handles',
});
handleGeometrySource.handles[0].position.x = 'changed';
handleGeometrySource.handles[0].maxX = 1;
handleGeometrySource.handles[1].maxAngle = 1;
handleGeometrySource.handles.reverse();
const initialHandleGeometry = handleGeometryShape.customGeometry;
const handleNoOpBytes = handleGeometryDeck.opcPackage
  .requirePart(handleGeometrySlide.partUri).bytes.slice();
const handleNoOpJournal = handleGeometryDeck.opcPackage.mutations.length;
handleGeometryShape.customGeometry = structuredClone(handleGeometryExpected);
const handleNoOpCurrent = handleGeometryDeck.opcPackage
  .requirePart(handleGeometrySlide.partUri).bytes;
const handleNoOp = handleNoOpJournal === handleGeometryDeck.opcPackage.mutations.length &&
  handleNoOpBytes.length === handleNoOpCurrent.length &&
  handleNoOpBytes.every((value, index) => value === handleNoOpCurrent[index]);
handleGeometryShape.customGeometry = handleGeometryReplacement;
const editedHandleGeometry = handleGeometryShape.customGeometry;
handleGeometryShape.presetType = 'diamond';
const handleConvertedPreset = handleGeometryShape.presetType;
const handleConvertedCustom = handleGeometryShape.customGeometry;
handleGeometryShape.customGeometry = handleGeometryReplacement;
const handleXml = new TextDecoder().decode(
  handleGeometryDeck.opcPackage.requirePart(handleGeometrySlide.partUri).bytes,
);
const reopenedHandleGeometryShape = (await PptxDocument.open(
  await handleGeometryDeck.write(),
)).slides[0].shapes[0];
const customGeometryAdjustmentHandles =
  handleGeometryShape instanceof ShapeModel &&
  Object.isFrozen(initialHandleGeometry) &&
  Object.isFrozen(initialHandleGeometry?.handles) &&
  initialHandleGeometry?.handles?.every((handle) =>
    Object.isFrozen(handle) && Object.isFrozen(handle.position)) &&
  JSON.stringify(initialHandleGeometry) === JSON.stringify(handleGeometryExpected) &&
  handleNoOp &&
  JSON.stringify(editedHandleGeometry) === JSON.stringify(handleGeometryReplacement) &&
  handleConvertedPreset === 'diamond' &&
  handleConvertedCustom === undefined &&
  handleGeometryShape.presetType === undefined &&
  handleXml.includes(
    '<a:ahLst><a:ahPolar gdRefR="adjR" minR="1" maxR="ss" ' +
    'gdRefAng="adjAng" maxAng="3cd4"><a:pos x="hc" y="vc"/></a:ahPolar>' +
    '<a:ahXY gdRefX="adjX" maxX="90000" gdRefY="adjY" minY="t">' +
    '<a:pos x="x1" y="adjY"/></a:ahXY></a:ahLst>',
  ) &&
  reopenedHandleGeometryShape instanceof ShapeModel &&
  reopenedHandleGeometryShape.name === 'Packed adjustment handles' &&
  JSON.stringify(reopenedHandleGeometryShape.customGeometry) ===
    JSON.stringify(handleGeometryReplacement) &&
  handleGeometryDeck.diagnostics.every(({ severity }) => severity !== 'error');
const connectionGeometrySource = {
  adjustments: [{ name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } }],
  connectionSites: [
    { position: { x: 'hc', y: 't' }, angle: 0 },
    { position: { x: 'r', y: 'vc' }, angle: 'adjAng' },
    { position: { x: 25_000, y: 100_000 }, angle: -5_400_000 },
    { position: { x: 'hc', y: 't' }, angle: 0 },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'b' } },
    ],
  }],
};
const connectionGeometryExpected = structuredClone(connectionGeometrySource);
const connectionGeometryReplacement = {
  ...connectionGeometryExpected,
  connectionSites: [
    { position: { x: 25_000, y: 100_000 }, angle: -5_400_000 },
    { position: { x: 'hc', y: 't' }, angle: 0 },
    { position: { x: 'l', y: 'vc' }, angle: 'adjAng' },
    { position: { x: 'hc', y: 't' }, angle: 0 },
  ],
};
const connectionGeometryDeck = PptxDocument.create();
const connectionGeometrySlide = connectionGeometryDeck.addSlide();
const connectionGeometryShape = connectionGeometrySlide.addCustomShape(
  connectionGeometrySource,
  { name: 'Packed connection sites' },
);
connectionGeometrySource.connectionSites[0].angle = 1;
connectionGeometrySource.connectionSites[0].position.x = 'changed';
connectionGeometrySource.connectionSites.reverse();
const initialConnectionGeometry = connectionGeometryShape.customGeometry;
const connectionNoOpBytes = connectionGeometryDeck.opcPackage
  .requirePart(connectionGeometrySlide.partUri).bytes.slice();
const connectionNoOpJournal = connectionGeometryDeck.opcPackage.mutations.length;
connectionGeometryShape.customGeometry = structuredClone(connectionGeometryExpected);
const connectionNoOpCurrent = connectionGeometryDeck.opcPackage
  .requirePart(connectionGeometrySlide.partUri).bytes;
const connectionNoOp =
  connectionNoOpJournal === connectionGeometryDeck.opcPackage.mutations.length &&
  connectionNoOpBytes.length === connectionNoOpCurrent.length &&
  connectionNoOpBytes.every((value, index) => value === connectionNoOpCurrent[index]);
connectionGeometryShape.customGeometry = connectionGeometryReplacement;
const editedConnectionGeometry = connectionGeometryShape.customGeometry;
connectionGeometryShape.presetType = 'diamond';
const connectionConvertedPreset = connectionGeometryShape.presetType;
const connectionConvertedCustom = connectionGeometryShape.customGeometry;
connectionGeometryShape.customGeometry = connectionGeometryReplacement;
const connectionXml = new TextDecoder().decode(
  connectionGeometryDeck.opcPackage.requirePart(connectionGeometrySlide.partUri).bytes,
);
const reopenedConnectionGeometryShape = (await PptxDocument.open(
  await connectionGeometryDeck.write(),
)).slides[0].shapes[0];
const customGeometryConnectionSites =
  connectionGeometryShape instanceof ShapeModel &&
  Object.isFrozen(initialConnectionGeometry) &&
  Object.isFrozen(initialConnectionGeometry?.connectionSites) &&
  initialConnectionGeometry?.connectionSites?.every((site) =>
    Object.isFrozen(site) && Object.isFrozen(site.position)) &&
  JSON.stringify(initialConnectionGeometry) === JSON.stringify(connectionGeometryExpected) &&
  connectionNoOp &&
  JSON.stringify(editedConnectionGeometry) === JSON.stringify(connectionGeometryReplacement) &&
  connectionConvertedPreset === 'diamond' &&
  connectionConvertedCustom === undefined &&
  connectionGeometryShape.presetType === undefined &&
  connectionXml.includes(
    '<a:cxnLst><a:cxn ang="-5400000"><a:pos x="25000" y="100000"/></a:cxn>' +
    '<a:cxn ang="0"><a:pos x="hc" y="t"/></a:cxn>' +
    '<a:cxn ang="adjAng"><a:pos x="l" y="vc"/></a:cxn>' +
    '<a:cxn ang="0"><a:pos x="hc" y="t"/></a:cxn></a:cxnLst>',
  ) &&
  reopenedConnectionGeometryShape instanceof ShapeModel &&
  reopenedConnectionGeometryShape.name === 'Packed connection sites' &&
  JSON.stringify(reopenedConnectionGeometryShape.customGeometry) ===
    JSON.stringify(connectionGeometryReplacement) &&
  connectionGeometryDeck.diagnostics.every(({ severity }) => severity !== 'error');
const textRectangleGeometrySource = {
  guides: [
    { name: 'textLeft', formula: { operator: 'val', operands: [20_000] } },
    { name: 'textRight', formula: { operator: 'val', operands: [80_000] } },
  ],
  connectionSites: [{ position: { x: 'hc', y: 't' }, angle: 0 }],
  textRectangle: {
    left: 'textLeft',
    top: 12_500,
    right: 'textRight',
    bottom: 87_500,
  },
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'b' } },
    ],
  }],
};
const textRectangleGeometryExpected = structuredClone(textRectangleGeometrySource);
const textRectangleGeometryReplacement = {
  ...textRectangleGeometryExpected,
  textRectangle: {
    left: 0,
    top: 't',
    right: 90_000,
    bottom: 'b',
  },
};
const textRectangleGeometryDeck = PptxDocument.create();
const textRectangleGeometrySlide = textRectangleGeometryDeck.addSlide();
const textRectangleGeometryShape = textRectangleGeometrySlide.addCustomShape(
  textRectangleGeometrySource,
  { name: 'Packed text rectangle' },
);
textRectangleGeometrySource.textRectangle.left = 'changed';
textRectangleGeometrySource.textRectangle.top = 1;
textRectangleGeometrySource.textRectangle.right = 2;
textRectangleGeometrySource.textRectangle.bottom = 3;
const initialTextRectangleGeometry = textRectangleGeometryShape.customGeometry;
const textRectangleNoOpBytes = textRectangleGeometryDeck.opcPackage
  .requirePart(textRectangleGeometrySlide.partUri).bytes.slice();
const textRectangleNoOpJournal = textRectangleGeometryDeck.opcPackage.mutations.length;
textRectangleGeometryShape.customGeometry = structuredClone(textRectangleGeometryExpected);
const textRectangleNoOpCurrent = textRectangleGeometryDeck.opcPackage
  .requirePart(textRectangleGeometrySlide.partUri).bytes;
const textRectangleNoOp =
  textRectangleNoOpJournal === textRectangleGeometryDeck.opcPackage.mutations.length &&
  textRectangleNoOpBytes.length === textRectangleNoOpCurrent.length &&
  textRectangleNoOpBytes.every(
    (value, index) => value === textRectangleNoOpCurrent[index],
  );
textRectangleGeometryShape.customGeometry = textRectangleGeometryReplacement;
const editedTextRectangleGeometry = textRectangleGeometryShape.customGeometry;
const { textRectangle: ignoredTextRectangle, ...defaultTextRectangleGeometry } =
  textRectangleGeometryReplacement;
textRectangleGeometryShape.customGeometry = defaultTextRectangleGeometry;
const resetTextRectangleGeometry = textRectangleGeometryShape.customGeometry;
textRectangleGeometryShape.customGeometry = textRectangleGeometryReplacement;
textRectangleGeometryShape.presetType = 'diamond';
const textRectangleConvertedPreset = textRectangleGeometryShape.presetType;
const textRectangleConvertedCustom = textRectangleGeometryShape.customGeometry;
textRectangleGeometryShape.customGeometry = textRectangleGeometryReplacement;
const textRectangleXml = new TextDecoder().decode(
  textRectangleGeometryDeck.opcPackage.requirePart(textRectangleGeometrySlide.partUri).bytes,
);
const reopenedTextRectangleGeometryShape = (await PptxDocument.open(
  await textRectangleGeometryDeck.write(),
)).slides[0].shapes[0];
const customGeometryTextRectangles =
  textRectangleGeometryShape instanceof ShapeModel &&
  Object.isFrozen(initialTextRectangleGeometry) &&
  Object.isFrozen(initialTextRectangleGeometry?.textRectangle) &&
  JSON.stringify(initialTextRectangleGeometry) ===
    JSON.stringify(textRectangleGeometryExpected) &&
  textRectangleNoOp &&
  JSON.stringify(editedTextRectangleGeometry) ===
    JSON.stringify(textRectangleGeometryReplacement) &&
  !Object.hasOwn(resetTextRectangleGeometry, 'textRectangle') &&
  ignoredTextRectangle !== undefined &&
  textRectangleConvertedPreset === 'diamond' &&
  textRectangleConvertedCustom === undefined &&
  textRectangleGeometryShape.presetType === undefined &&
  textRectangleXml.includes('<a:rect l="0" t="t" r="90000" b="b"/>') &&
  reopenedTextRectangleGeometryShape instanceof ShapeModel &&
  reopenedTextRectangleGeometryShape.name === 'Packed text rectangle' &&
  Object.isFrozen(reopenedTextRectangleGeometryShape.customGeometry?.textRectangle) &&
  JSON.stringify(reopenedTextRectangleGeometryShape.customGeometry) ===
    JSON.stringify(textRectangleGeometryReplacement) &&
  textRectangleGeometryDeck.diagnostics.every(({ severity }) => severity !== 'error');
const evaluatorGeometry = {
  adjustments: [{ name: 'adj', formula: { operator: 'val', operands: [25_000] } }],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
    { name: 'y1', formula: { operator: '*/', operands: ['h', 'adj', 100_000] } },
    { name: 'rad', formula: { operator: 'max', operands: ['x1', 1] } },
  ],
  handles: [{
    kind: 'xy',
    position: { x: 'x1', y: 'y1' },
    xGuide: 'adj',
    minX: 'l',
    maxX: 'r',
  }],
  connectionSites: [{ angle: 'cd4', position: { x: 'x1', y: 'y1' } }],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 'y1' } },
      {
        kind: 'arcTo',
        widthRadius: 'rad',
        heightRadius: 'hd2',
        startAngle: 0,
        sweepAngle: 'cd4',
      },
      { kind: 'close' },
    ],
  }],
};
const evaluatorContext = { width: inches(2), height: inches(1) };
const pureEvaluatorResult = evaluateCustomGeometry(evaluatorGeometry, evaluatorContext);
const evaluatorDeck = PptxDocument.create();
const evaluatorSlide = evaluatorDeck.addSlide();
const evaluatorShape = evaluatorSlide.addCustomShape(evaluatorGeometry, {
  name: 'Packed custom geometry evaluator',
  width: inches(2),
  height: inches(1),
});
const evaluatorBefore = evaluatorDeck.opcPackage.requirePart(evaluatorSlide.partUri).bytes.slice();
const evaluatorJournal = evaluatorDeck.opcPackage.mutations.length;
const liveEvaluatorResult = evaluatorShape.evaluateCustomGeometry();
const evaluatorCurrent = evaluatorDeck.opcPackage.requirePart(evaluatorSlide.partUri).bytes;
const evaluatorReadOnly = evaluatorJournal === evaluatorDeck.opcPackage.mutations.length &&
  evaluatorBefore.length === evaluatorCurrent.length &&
  evaluatorBefore.every((value, index) => value === evaluatorCurrent[index]);
evaluatorShape.setTransform({ width: inches(3), height: inches(1.5) });
const resizedEvaluatorResult = evaluatorShape.evaluateCustomGeometry();
let typedEvaluatorFailure = false;
try {
  evaluateCustomGeometry({
    paths: [{
      width: 1,
      height: 1,
      commands: [{ kind: 'moveTo', point: { x: 'missing', y: 0 } }],
    }],
  }, { width: 1, height: 1 });
} catch (error) {
  typedEvaluatorFailure = error instanceof CustomGeometryEvaluationError &&
    error.code === 'unknown-token' && error.token === 'missing';
}
const evaluatorBytes = await evaluatorDeck.write();
const reopenedEvaluatorShape = (await PptxDocument.open(evaluatorBytes)).slides[0].shapes[0];
const reopenedEvaluatorResult = reopenedEvaluatorShape instanceof ShapeModel
  ? reopenedEvaluatorShape.evaluateCustomGeometry()
  : undefined;
const customGeometryEvaluator =
  evaluatorShape instanceof ShapeModel &&
  JSON.stringify(pureEvaluatorResult.guides) === JSON.stringify([
    { name: 'x1', value: inches(0.5) },
    { name: 'y1', value: inches(0.25) },
    { name: 'rad', value: inches(0.5) },
  ]) &&
  JSON.stringify(pureEvaluatorResult.textRectangle) === JSON.stringify({
    left: 0,
    top: 0,
    right: inches(2),
    bottom: inches(1),
  }) &&
  JSON.stringify(liveEvaluatorResult) === JSON.stringify(pureEvaluatorResult) &&
  Object.isFrozen(pureEvaluatorResult) &&
  Object.isFrozen(pureEvaluatorResult.context) &&
  Object.isFrozen(pureEvaluatorResult.handles?.[0]?.position) &&
  Object.isFrozen(pureEvaluatorResult.connectionSites?.[0]?.position) &&
  Object.isFrozen(pureEvaluatorResult.textRectangle) &&
  Object.isFrozen(pureEvaluatorResult.paths[0]?.commands[1]) &&
  evaluatorReadOnly &&
  resizedEvaluatorResult?.context.width === inches(3) &&
  resizedEvaluatorResult.context.height === inches(1.5) &&
  resizedEvaluatorResult.guides?.[0]?.value === inches(0.75) &&
  resizedEvaluatorResult.guides?.[1]?.value === inches(0.375) &&
  typedEvaluatorFailure &&
  reopenedEvaluatorShape instanceof ShapeModel &&
  JSON.stringify(reopenedEvaluatorResult) === JSON.stringify(resizedEvaluatorResult) &&
  evaluatorDeck.diagnostics.every(({ severity }) => severity !== 'error');
const shapeAdjustmentDeck = PptxDocument.create();
const shapeAdjustmentSlide = shapeAdjustmentDeck.addSlide();
const shapeAdjustmentInput = [
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
];
const packedAdjustedShape = shapeAdjustmentSlide.addShape('blockArc', {
  name: 'Packed adjusted block arc',
  adjustments: shapeAdjustmentInput,
});
const initialPackedAdjustments = packedAdjustedShape.adjustments;
const initialPackedAdjustmentsAgain = packedAdjustedShape.adjustments;
shapeAdjustmentInput[0].value = 0;
const adjustmentNoOpBytes = shapeAdjustmentDeck.opcPackage
  .requirePart(shapeAdjustmentSlide.partUri).bytes.slice();
const adjustmentNoOpJournal = shapeAdjustmentDeck.opcPackage.mutations.length;
packedAdjustedShape.adjustments = [
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
];
const adjustmentNoOpCurrent = shapeAdjustmentDeck.opcPackage
  .requirePart(shapeAdjustmentSlide.partUri).bytes;
const shapeAdjustmentNoOp =
  adjustmentNoOpJournal === shapeAdjustmentDeck.opcPackage.mutations.length &&
  adjustmentNoOpBytes.length === adjustmentNoOpCurrent.length &&
  adjustmentNoOpBytes.every((value, index) => value === adjustmentNoOpCurrent[index]);
packedAdjustedShape.adjustments = [
  { name: 'adj1', value: 10_800_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 20_000 },
];
const editedPackedAdjustments = packedAdjustedShape.adjustments;
packedAdjustedShape.adjustments = [];
const clearedPackedAdjustments = packedAdjustedShape.adjustments;
packedAdjustedShape.adjustments = [
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
];
const reopenedShapeAdjustmentDeck = await PptxDocument.open(await shapeAdjustmentDeck.write());
const reopenedPackedAdjustedShape = reopenedShapeAdjustmentDeck.slides[0].shapes[0];
const reopenedPackedAdjustments = reopenedPackedAdjustedShape.adjustments;
const shapeAdjustments =
  packedAdjustedShape instanceof ShapeModel &&
  Array.isArray(initialPackedAdjustments) &&
  Object.isFrozen(initialPackedAdjustments) &&
  initialPackedAdjustments.every((adjustment) => Object.isFrozen(adjustment)) &&
  initialPackedAdjustments !== initialPackedAdjustmentsAgain &&
  JSON.stringify(initialPackedAdjustments) === JSON.stringify([
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ]) &&
  shapeAdjustmentNoOp &&
  JSON.stringify(editedPackedAdjustments) === JSON.stringify([
    { name: 'adj1', value: 10_800_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 20_000 },
  ]) &&
  Array.isArray(clearedPackedAdjustments) &&
  Object.isFrozen(clearedPackedAdjustments) &&
  clearedPackedAdjustments.length === 0 &&
  reopenedPackedAdjustedShape instanceof ShapeModel &&
  Array.isArray(reopenedPackedAdjustments) &&
  Object.isFrozen(reopenedPackedAdjustments) &&
  reopenedPackedAdjustments.every((adjustment) => Object.isFrozen(adjustment)) &&
  JSON.stringify(reopenedPackedAdjustments) === JSON.stringify(initialPackedAdjustments);
const shapeHyperlinkDeck = PptxDocument.create();
const shapeHyperlinkSource = shapeHyperlinkDeck.addSlide();
const shapeHyperlinkTarget = shapeHyperlinkDeck.addSlide();
shapeHyperlinkDeck.addSlide();
const shapeHyperlinkInput = {
  url: 'https://example.com?a=1&b=2',
  tooltip: 'Visit & learn',
};
const packedUrlHyperlink = shapeHyperlinkSource.addShape('rect', {
  name: 'Packed URL hyperlink',
  hyperlink: shapeHyperlinkInput,
});
const packedInternalHyperlink = shapeHyperlinkSource.addShape('actionButtonForwardNext', {
  name: 'Packed internal hyperlink',
  hyperlink: { slide: 2 },
});
const packedSelfHyperlink = shapeHyperlinkSource.addShape('actionButtonHome', {
  name: 'Packed self hyperlink',
  hyperlink: { slide: 1, tooltip: '' },
});
const packedSharedA = shapeHyperlinkSource.addShape('rect', {
  name: 'Packed shared A',
  hyperlink: { url: 'https://shared.example' },
});
const packedSharedB = shapeHyperlinkSource.addShape('ellipse', {
  name: 'Packed shared B',
  hyperlink: { url: 'https://temporary.example' },
});
const initialPackedUrlHyperlink = packedUrlHyperlink.hyperlink;
const initialPackedUrlHyperlinkAgain = packedUrlHyperlink.hyperlink;
shapeHyperlinkInput.url = 'https://changed.example';
shapeHyperlinkInput.tooltip = 'Changed';
const detachedPackedUrlHyperlink = packedUrlHyperlink.hyperlink;
const hyperlinkNoOpBytes = shapeHyperlinkDeck.opcPackage
  .requirePart(shapeHyperlinkSource.partUri).bytes.slice();
const hyperlinkNoOpJournal = shapeHyperlinkDeck.opcPackage.mutations.length;
packedUrlHyperlink.hyperlink = {
  url: 'https://example.com?a=1&b=2',
  tooltip: 'Visit & learn',
};
const hyperlinkNoOpCurrent = shapeHyperlinkDeck.opcPackage
  .requirePart(shapeHyperlinkSource.partUri).bytes;
const shapeHyperlinkNoOp = hyperlinkNoOpJournal === shapeHyperlinkDeck.opcPackage.mutations.length &&
  hyperlinkNoOpBytes.length === hyperlinkNoOpCurrent.length &&
  hyperlinkNoOpBytes.every((value, index) => value === hyperlinkNoOpCurrent[index]);
const sharedRelationship = shapeHyperlinkSource.relationships.find(
  ({ type, target }) => type.endsWith('/hyperlink') && target === 'https://shared.example',
);
const temporaryRelationship = shapeHyperlinkSource.relationships.find(
  ({ type, target }) => type.endsWith('/hyperlink') && target === 'https://temporary.example',
);
if (!sharedRelationship || !temporaryRelationship) {
  throw new Error('Packed hyperlink sharing fixture failed');
}
const shapeHyperlinkPart = shapeHyperlinkDeck.opcPackage.requirePart(shapeHyperlinkSource.partUri);
const sharedShapeHyperlinkXml = new TextDecoder().decode(shapeHyperlinkPart.bytes)
  .replace('r:id="' + temporaryRelationship.id + '"', 'r:id="' + sharedRelationship.id + '"');
shapeHyperlinkDeck.opcPackage.setPart(
  shapeHyperlinkSource.partUri,
  sharedShapeHyperlinkXml,
  shapeHyperlinkPart.contentType,
);
shapeHyperlinkDeck.opcPackage.removeRelationship(
  shapeHyperlinkSource.partUri,
  temporaryRelationship.id,
);
packedSharedA.hyperlink = { url: 'https://clone.example', tooltip: 'Clone' };
const shapeHyperlinkCloneOnWrite = JSON.stringify(packedSharedA.hyperlink) ===
  JSON.stringify({ url: 'https://clone.example', tooltip: 'Clone' }) &&
  JSON.stringify(packedSharedB.hyperlink) === JSON.stringify({ url: 'https://shared.example' }) &&
  shapeHyperlinkSource.relationships.filter(({ type }) => type.endsWith('/hyperlink')).length === 3;
packedSharedA.hyperlink = undefined;
const shapeHyperlinkCloneGc = shapeHyperlinkSource.relationships.every(
  ({ target }) => target !== 'https://clone.example',
);
packedUrlHyperlink.hyperlink = {
  url: 'mailto:test@example.com',
  tooltip: '',
};
packedInternalHyperlink.hyperlink = { slide: 3, tooltip: '' };
packedInternalHyperlink.hyperlink = { slide: 2 };
const duplicateShapeHyperlinkSlide = shapeHyperlinkDeck.duplicateSlide(0);
const duplicatePackedSelfHyperlink = duplicateShapeHyperlinkSlide.shapes[2];
const shapeHyperlinkDuplicateSelf = duplicatePackedSelfHyperlink instanceof ShapeModel &&
  JSON.stringify(duplicatePackedSelfHyperlink.hyperlink) === JSON.stringify({ slide: 4, tooltip: '' });
shapeHyperlinkDeck.moveSlide(shapeHyperlinkDeck.slides.indexOf(shapeHyperlinkTarget), 0);
const shapeHyperlinkMovedTarget = JSON.stringify(packedInternalHyperlink.hyperlink) ===
  JSON.stringify({ slide: 1 });
shapeHyperlinkDeck.deleteSlide(shapeHyperlinkDeck.slides.indexOf(shapeHyperlinkTarget));
const shapeHyperlinkTargetCleanup = packedInternalHyperlink.hyperlink === undefined &&
  duplicateShapeHyperlinkSlide.shapes[1] instanceof ShapeModel &&
  duplicateShapeHyperlinkSlide.shapes[1].hyperlink === undefined;
const reopenedShapeHyperlinkDeck = await PptxDocument.open(await shapeHyperlinkDeck.write());
const reopenedShapeHyperlinkSource = reopenedShapeHyperlinkDeck.slides[0];
const reopenedShapeHyperlinkDuplicate = reopenedShapeHyperlinkDeck.slides[2];
const reopenedSourceHyperlinks = reopenedShapeHyperlinkSource.shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.hyperlink : undefined);
const reopenedDuplicateHyperlinks = reopenedShapeHyperlinkDuplicate.shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.hyperlink : undefined);
const packedHyperlinkRelationships = reopenedShapeHyperlinkDeck.slides.flatMap(
  ({ relationships }) => relationships,
);
const packedHyperlinkDangling = packedHyperlinkRelationships.some(
  ({ targetMode, resolvedTarget }) => targetMode === 'Internal' &&
    resolvedTarget !== undefined &&
    !reopenedShapeHyperlinkDeck.opcPackage.hasPart(resolvedTarget),
);
const shapeHyperlinks = packedUrlHyperlink instanceof ShapeModel &&
  packedInternalHyperlink instanceof ShapeModel &&
  packedSelfHyperlink instanceof ShapeModel &&
  initialPackedUrlHyperlink !== initialPackedUrlHyperlinkAgain &&
  Object.isFrozen(initialPackedUrlHyperlink) &&
  JSON.stringify(initialPackedUrlHyperlink) === JSON.stringify({
    url: 'https://example.com?a=1&b=2',
    tooltip: 'Visit & learn',
  }) &&
  JSON.stringify(detachedPackedUrlHyperlink) === JSON.stringify(initialPackedUrlHyperlink) &&
  shapeHyperlinkNoOp &&
  shapeHyperlinkCloneOnWrite &&
  shapeHyperlinkCloneGc &&
  shapeHyperlinkDuplicateSelf &&
  shapeHyperlinkMovedTarget &&
  shapeHyperlinkTargetCleanup &&
  JSON.stringify(reopenedSourceHyperlinks) === JSON.stringify([
    { url: 'mailto:test@example.com', tooltip: '' },
    undefined,
    { slide: 1, tooltip: '' },
    undefined,
    { url: 'https://shared.example' },
  ]) &&
  JSON.stringify(reopenedDuplicateHyperlinks) === JSON.stringify([
    { url: 'mailto:test@example.com', tooltip: '' },
    undefined,
    { slide: 3, tooltip: '' },
    undefined,
    { url: 'https://shared.example' },
  ]) &&
  packedHyperlinkRelationships.filter(({ type }) => type.endsWith('/hyperlink')).length === 4 &&
  packedHyperlinkRelationships.filter(({ type }) => type.endsWith('/slide')).length === 2 &&
  !packedHyperlinkDangling;
if (!shapeHyperlinks) {
  throw new Error('Packed shape hyperlinks failed: ' + JSON.stringify({
    initialPackedUrlHyperlink,
    detachedPackedUrlHyperlink,
    reopenedSourceHyperlinks,
    reopenedDuplicateHyperlinks,
    packedHyperlinkRelationships,
  }));
}
const shapeShadowDeck = PptxDocument.create();
const shapeShadowSlide = shapeShadowDeck.addSlide();
const shapeShadowColor = { kind: 'srgb', value: '#123abc' };
const shapeShadowInput = {
  kind: 'outer',
  color: shapeShadowColor,
  opacity: 0.42,
  blur: 7.25,
  angle: 123.4,
  distance: 5.5,
  rotateWithShape: true,
};
const packedOuterShadow = shapeShadowSlide.addShape('roundRect', {
  name: 'Packed outer shadow',
  shadow: shapeShadowInput,
});
const packedDefaultShadow = shapeShadowSlide.addShape('rect', {
  name: 'Packed default shadow',
  shadow: { kind: 'outer' },
});
const packedInnerShadow = shapeShadowSlide.addShape('ellipse', {
  name: 'Packed inner shadow',
  shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent3' } },
});
const packedZeroThemeShadow = shapeShadowSlide.addShape('star5', {
  name: 'Packed zero theme shadow',
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
const initialPackedShadow = packedOuterShadow.shadow;
const initialPackedShadowAgain = packedOuterShadow.shadow;
shapeShadowColor.value = 'FFFFFF';
shapeShadowInput.opacity = 1;
shapeShadowInput.blur = 20;
shapeShadowInput.angle = 90;
shapeShadowInput.distance = 10;
shapeShadowInput.rotateWithShape = false;
const detachedPackedShadow = packedOuterShadow.shadow;
const shapeShadowNoOpBytes = shapeShadowDeck.opcPackage
  .requirePart(shapeShadowSlide.partUri).bytes.slice();
const shapeShadowNoOpJournal = shapeShadowDeck.opcPackage.mutations.length;
packedOuterShadow.shadow = {
  kind: 'outer',
  color: { kind: 'srgb', value: '123ABC' },
  opacity: 0.42,
  blur: 7.25,
  angle: 123.4,
  distance: 5.5,
  rotateWithShape: true,
};
const shapeShadowNoOpCurrent = shapeShadowDeck.opcPackage
  .requirePart(shapeShadowSlide.partUri).bytes;
const shapeShadowNoOp = shapeShadowNoOpJournal === shapeShadowDeck.opcPackage.mutations.length &&
  shapeShadowNoOpBytes.length === shapeShadowNoOpCurrent.length &&
  shapeShadowNoOpBytes.every((value, index) => value === shapeShadowNoOpCurrent[index]);
const shapeShadowPart = shapeShadowDeck.opcPackage.requirePart(shapeShadowSlide.partUri);
const shapeShadowXml = new TextDecoder().decode(shapeShadowPart.bytes);
const shapeShadowWithGlow = shapeShadowXml.replace(
  '<a:effectLst><a:outerShdw',
  '<a:effectLst><a:glow rad="12700"><a:srgbClr val="00FF00"/>' +
    '</a:glow><a:outerShdw',
);
if (shapeShadowWithGlow === shapeShadowXml) {
  throw new Error('Packed shape shadow glow fixture failed');
}
shapeShadowDeck.opcPackage.setPart(
  shapeShadowSlide.partUri,
  shapeShadowWithGlow,
  shapeShadowPart.contentType,
);
packedOuterShadow.shadow = {
  kind: 'outer',
  color: { kind: 'scheme', value: 'accent5' },
  opacity: 0.6,
  blur: 3,
  angle: 45,
  distance: 2,
  rotateWithShape: false,
};
packedDefaultShadow.shadow = {
  kind: 'inner',
  color: { kind: 'srgb', value: '445566' },
  opacity: 0.5,
  blur: 2,
  angle: 30,
  distance: 1,
};
packedInnerShadow.shadow = undefined;
const duplicateShapeShadowSlide = shapeShadowDeck.duplicateSlide(0);
const duplicatePackedShadow = duplicateShapeShadowSlide.shapes[0];
if (!(duplicatePackedShadow instanceof ShapeModel)) {
  throw new Error('Packed duplicate shadow shape failed');
}
duplicatePackedShadow.shadow = undefined;
const writtenShapeShadowDeck = await shapeShadowDeck.write();
const reopenedShapeShadowDeck = await PptxDocument.open(writtenShapeShadowDeck);
const reopenedSourceShadows = reopenedShapeShadowDeck.slides[0].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.shadow : undefined);
const reopenedDuplicateShadows = reopenedShapeShadowDeck.slides[1].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.shadow : undefined);
const reopenedShapeShadowXml = reopenedShapeShadowDeck.slides.map(({ partUri }) =>
  new TextDecoder().decode(reopenedShapeShadowDeck.opcPackage.requirePart(partUri).bytes));
const shapeShadows = packedOuterShadow instanceof ShapeModel &&
  packedDefaultShadow instanceof ShapeModel &&
  packedInnerShadow instanceof ShapeModel &&
  packedZeroThemeShadow instanceof ShapeModel &&
  initialPackedShadow !== initialPackedShadowAgain &&
  initialPackedShadow?.color !== initialPackedShadowAgain?.color &&
  Object.isFrozen(initialPackedShadow) &&
  Object.isFrozen(initialPackedShadow?.color) &&
  JSON.stringify(initialPackedShadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'srgb', value: '123ABC' },
    opacity: 0.42,
    blur: 7.25,
    angle: 123.4,
    distance: 5.5,
    rotateWithShape: true,
  }) &&
  JSON.stringify(detachedPackedShadow) === JSON.stringify(initialPackedShadow) &&
  shapeShadowNoOp &&
  packedInnerShadow.shadow === undefined &&
  JSON.stringify(packedZeroThemeShadow.shadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent2' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
    rotateWithShape: false,
  }) &&
  reopenedShapeShadowXml[0].includes(
    '<a:glow rad="12700"><a:srgbClr val="00FF00"/></a:glow>',
  ) &&
  reopenedShapeShadowXml[1].includes(
    '<a:glow rad="12700"><a:srgbClr val="00FF00"/></a:glow>',
  ) &&
  JSON.stringify(reopenedSourceShadows) === JSON.stringify([
    {
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent5' },
      opacity: 0.6,
      blur: 3,
      angle: 45,
      distance: 2,
      rotateWithShape: false,
    },
    {
      kind: 'inner',
      color: { kind: 'srgb', value: '445566' },
      opacity: 0.5,
      blur: 2,
      angle: 30,
      distance: 1,
    },
    undefined,
    {
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: false,
    },
  ]) &&
  reopenedDuplicateShadows[0] === undefined &&
  JSON.stringify(reopenedDuplicateShadows.slice(1)) ===
    JSON.stringify(reopenedSourceShadows.slice(1)) &&
  shapeShadowDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!shapeShadows) {
  throw new Error('Packed shape shadows failed: ' + JSON.stringify({
    initialPackedShadow,
    detachedPackedShadow,
    reopenedSourceShadows,
    reopenedDuplicateShadows,
    diagnostics: shapeShadowDeck.diagnostics,
  }));
}
const shapeFillDeck = PptxDocument.create();
const shapeFillSlide = shapeFillDeck.addSlide();
const shapeFillSourceColor = { kind: 'srgb', value: '#AA0000' };
const shapeFillSource = {
  kind: 'solid',
  color: shapeFillSourceColor,
  transparency: 33.3334,
};
const packedSrgbFill = shapeFillSlide.addShape('rect', {
  name: 'Packed sRGB fill',
  fill: shapeFillSource,
});
const packedSchemeFill = shapeFillSlide.addShape('ellipse', {
  name: 'Packed scheme fill',
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
  },
});
const packedNoneFill = shapeFillSlide.addShape('star5', {
  name: 'Packed none fill',
  fill: { kind: 'none' },
});
const initialPackedSrgbFill = packedSrgbFill.fill;
const initialPackedSrgbFillAgain = packedSrgbFill.fill;
shapeFillSourceColor.value = '000000';
shapeFillSource.transparency = 1;
const detachedPackedSrgbFill = packedSrgbFill.fill;
const shapeFillPartCountBeforeEdit = shapeFillDeck.opcPackage.parts.length;
const shapeFillRelationshipCountBeforeEdit = shapeFillSlide.relationships.length;
const shapeFillNoOpBytes = shapeFillDeck.opcPackage.requirePart(shapeFillSlide.partUri).bytes.slice();
const shapeFillNoOpJournalLength = shapeFillDeck.opcPackage.mutations.length;
packedSrgbFill.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'AA0000' },
  transparency: 33.333,
};
const currentShapeFillNoOpBytes = shapeFillDeck.opcPackage.requirePart(shapeFillSlide.partUri).bytes;
const shapeFillNoOp = shapeFillNoOpJournalLength === shapeFillDeck.opcPackage.mutations.length &&
  shapeFillNoOpBytes.length === currentShapeFillNoOpBytes.length &&
  shapeFillNoOpBytes.every((value, index) => value === currentShapeFillNoOpBytes[index]);
packedSrgbFill.fill = { kind: 'none' };
const packedNoneReplacement = packedSrgbFill.fill;
packedSrgbFill.fill = undefined;
const packedClearedFill = packedSrgbFill.fill;
packedSrgbFill.fill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent4' },
  transparency: 40,
};
const shapeFillEditIsolation = shapeFillDeck.opcPackage.parts.length ===
  shapeFillPartCountBeforeEdit &&
  shapeFillSlide.relationships.length === shapeFillRelationshipCountBeforeEdit;
const duplicateShapeFillSlide = shapeFillDeck.duplicateSlide(0);
const duplicatePackedFill = duplicateShapeFillSlide.shapes[0];
if (!(duplicatePackedFill instanceof ShapeModel)) throw new Error('Packed duplicate fill shape failed');
duplicatePackedFill.fill = { kind: 'none' };
const reopenedShapeFillDeck = await PptxDocument.open(await shapeFillDeck.write());
const reopenedSourceFills = reopenedShapeFillDeck.slides[0].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.fill : undefined);
const reopenedDuplicateFills = reopenedShapeFillDeck.slides[1].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.fill : undefined);
const shapeFills = packedSrgbFill instanceof ShapeModel &&
  packedSchemeFill instanceof ShapeModel &&
  packedNoneFill instanceof ShapeModel &&
  initialPackedSrgbFill !== initialPackedSrgbFillAgain &&
  initialPackedSrgbFill?.kind === 'solid' &&
  initialPackedSrgbFillAgain?.kind === 'solid' &&
  initialPackedSrgbFill.color !== initialPackedSrgbFillAgain.color &&
  JSON.stringify(initialPackedSrgbFill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'srgb', value: 'AA0000' },
    transparency: 33.333,
  }) &&
  JSON.stringify(detachedPackedSrgbFill) === JSON.stringify(initialPackedSrgbFill) &&
  packedNoneReplacement?.kind === 'none' &&
  packedClearedFill === undefined &&
  shapeFillNoOp &&
  shapeFillEditIsolation &&
  JSON.stringify(packedSrgbFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 40,
  }) &&
  JSON.stringify(packedSchemeFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
  }) &&
  packedNoneFill.fill?.kind === 'none' &&
  duplicatePackedFill.fill?.kind === 'none' &&
  JSON.stringify(reopenedSourceFills) === JSON.stringify([
    { kind: 'solid', color: { kind: 'scheme', value: 'accent4' }, transparency: 40 },
    { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
    { kind: 'none' },
  ]) &&
  JSON.stringify(reopenedDuplicateFills) === JSON.stringify([
    { kind: 'none' },
    { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
    { kind: 'none' },
  ]);
const textShapeFillDeck = PptxDocument.create();
const textShapeFillLayout = textShapeFillDeck.layouts[0];
const textShapeFillMaster = textShapeFillDeck.masters[0];
const packedLayoutTextFill = textShapeFillLayout.addText('Packed layout text fill', {
  name: 'packed_layout_text_fill',
  fill: { kind: 'none' },
});
const packedMasterTextFill = textShapeFillMaster.addRichText([{
  runs: [{ text: 'Packed master text fill' }],
}], {
  name: 'packed_master_text_fill',
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent6' },
    transparency: 100,
  },
});
const packedLayoutPlaceholderFill = textShapeFillLayout.addPlaceholder('Packed title prompt', {
  name: 'packed_title_fill',
  type: 'title',
  index: 190,
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 50,
  },
});
const textShapeFillSource = {
  kind: 'solid',
  color: { kind: 'srgb', value: '#AB12CD' },
  transparency: 25,
};
const textShapeFillSlide = textShapeFillDeck.addSlide({ masterName: textShapeFillLayout.name });
const packedPlainTextFill = textShapeFillSlide.addText('Packed plain text fill', {
  name: 'packed_plain_text_fill',
  fill: textShapeFillSource,
});
const packedRichTextFill = textShapeFillSlide.addRichText([{
  runs: [{ text: 'Packed rich text fill' }],
}], {
  name: 'packed_rich_text_fill',
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const packedPopulatedTextFill = textShapeFillSlide.addText('Packed populated text fill', {
  placeholder: 'packed_title_fill',
  fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: '112233' },
    transparency: 25,
  },
});
const initialPackedPlainTextFill = packedPlainTextFill.fill;
textShapeFillSource.color.value = 'FFFFFF';
textShapeFillSource.transparency = 90;
const detachedPackedPlainTextFill = packedPlainTextFill.fill;
const packedDeclarativeTextFillLayout = await textShapeFillDeck.defineSlideMaster({
  title: 'PACKED-TEXT-FILLS',
  objects: [
    {
      kind: 'text',
      text: 'Packed declarative text fill',
      options: { name: 'packed_declarative_text_fill', fill: { kind: 'none' } },
    },
    {
      kind: 'placeholder',
      text: 'Packed declarative prompt',
      options: {
        name: 'packed_declarative_title_fill',
        type: 'title',
        index: 191,
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 40,
        },
      },
    },
  ],
});
const packedDeclarativeTextFill = packedDeclarativeTextFillLayout.shapes.find(
  ({ name }) => name === 'packed_declarative_text_fill',
);
const packedDeclarativePlaceholderFill = packedDeclarativeTextFillLayout.placeholders.find(
  ({ name }) => name === 'packed_declarative_title_fill',
);
const packedDeclarativeTextFillSlide = textShapeFillDeck.addSlide({
  masterName: packedDeclarativeTextFillLayout.name,
});
const packedDeclarativePopulatedFill = packedDeclarativeTextFillSlide.addText(
  'Packed declarative populated fill',
  {
    placeholder: 'packed_declarative_title_fill',
    fill: {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 50,
    },
  },
);
const duplicateTextShapeFillSlide = textShapeFillDeck.duplicateSlide(0);
const duplicatePlainTextFill = duplicateTextShapeFillSlide.shapes.find(
  ({ name }) => name === 'packed_plain_text_fill',
);
const duplicateRichTextFill = duplicateTextShapeFillSlide.shapes.find(
  ({ name }) => name === 'packed_rich_text_fill',
);
if (!(duplicatePlainTextFill instanceof ShapeModel) ||
    !(duplicateRichTextFill instanceof ShapeModel)) {
  throw new Error('Packed duplicate text shape fill failed');
}
duplicatePlainTextFill.fill = { kind: 'none' };
duplicateRichTextFill.fill = undefined;
const reopenedTextShapeFillDeck = await PptxDocument.open(await textShapeFillDeck.write());
await reopenedTextShapeFillDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedTextFillSourceSlide = reopenedTextShapeFillDeck.slides[0];
const reopenedTextFillDeclarativeSlide = reopenedTextShapeFillDeck.slides[1];
const reopenedTextFillDuplicateSlide = reopenedTextShapeFillDeck.slides[2];
const reopenedTextFillLayout = reopenedTextShapeFillDeck.layouts.find(
  ({ name }) => name === textShapeFillLayout.name,
);
const reopenedDeclarativeTextFillLayout = reopenedTextShapeFillDeck.layouts.find(
  ({ name }) => name === 'PACKED-TEXT-FILLS',
);
const reopenedTextFillByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const textShapeFills =
  packedLayoutTextFill instanceof ShapeModel &&
  packedMasterTextFill instanceof ShapeModel &&
  packedLayoutPlaceholderFill instanceof ShapeModel &&
  packedPlainTextFill instanceof ShapeModel &&
  packedRichTextFill instanceof ShapeModel &&
  packedPopulatedTextFill instanceof ShapeModel &&
  packedDeclarativeTextFill instanceof ShapeModel &&
  packedDeclarativePlaceholderFill instanceof ShapeModel &&
  packedDeclarativePopulatedFill instanceof ShapeModel &&
  JSON.stringify(initialPackedPlainTextFill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'srgb', value: 'AB12CD' },
    transparency: 25,
  }) &&
  JSON.stringify(detachedPackedPlainTextFill) === JSON.stringify(initialPackedPlainTextFill) &&
  packedLayoutTextFill.fill?.kind === 'none' &&
  JSON.stringify(packedMasterTextFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent6' },
    transparency: 100,
  }) &&
  JSON.stringify(packedLayoutPlaceholderFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 50,
  }) &&
  JSON.stringify(packedRichTextFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  }) &&
  JSON.stringify(packedPopulatedTextFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'srgb', value: '112233' },
    transparency: 25,
  }) &&
  packedDeclarativeTextFill.fill?.kind === 'none' &&
  JSON.stringify(packedDeclarativePlaceholderFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 40,
  }) &&
  JSON.stringify(packedDeclarativePopulatedFill.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent3' },
    transparency: 50,
  }) &&
  JSON.stringify(reopenedTextFillByName(
    reopenedTextFillSourceSlide,
    'packed_plain_text_fill',
  )?.fill) === JSON.stringify(initialPackedPlainTextFill) &&
  JSON.stringify(reopenedTextFillByName(
    reopenedTextFillSourceSlide,
    'packed_rich_text_fill',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  }) &&
  JSON.stringify(reopenedTextFillByName(
    reopenedTextFillDuplicateSlide,
    'packed_plain_text_fill',
  )?.fill) === JSON.stringify({ kind: 'none' }) &&
  reopenedTextFillByName(
    reopenedTextFillDuplicateSlide,
    'packed_rich_text_fill',
  )?.fill === undefined &&
  reopenedTextFillByName(
    reopenedTextFillLayout,
    'packed_layout_text_fill',
  )?.fill?.kind === 'none' &&
  JSON.stringify(reopenedTextFillByName(
    reopenedTextShapeFillDeck.masters[0],
    'packed_master_text_fill',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent6' },
    transparency: 100,
  }) &&
  reopenedTextFillByName(
    reopenedDeclarativeTextFillLayout,
    'packed_declarative_text_fill',
  )?.fill?.kind === 'none' &&
  JSON.stringify(reopenedTextFillByName(
    reopenedTextFillDeclarativeSlide,
    'packed_declarative_title_fill',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent3' },
    transparency: 50,
  }) &&
  reopenedTextShapeFillDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!textShapeFills) {
  throw new Error('Packed text shape fills failed');
}
await reopenedTextShapeFillDeck.writeFile('text-shape-fill-smoke.pptx');
const textShapeLineDeck = PptxDocument.create();
const textShapeLineLayout = textShapeLineDeck.layouts[0];
const textShapeLineMaster = textShapeLineDeck.masters[0];
const packedLayoutTextLine = textShapeLineLayout.addText('Packed layout text line', {
  name: 'packed_layout_text_line',
  line: { kind: 'none' },
});
const packedMasterTextLine = textShapeLineMaster.addRichText([{
  runs: [{ text: 'Packed master text line' }],
}], {
  name: 'packed_master_text_line',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent6' },
    transparency: 100,
    width: 0,
    dash: 'sysDot',
  },
});
const packedLayoutPlaceholderLine = textShapeLineLayout.addPlaceholder(
  'Packed text line prompt',
  {
    name: 'packed_title_line',
    type: 'title',
    index: 192,
    line: {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 50,
    },
  },
);
const textShapeLineSource = {
  kind: 'line',
  color: { kind: 'srgb', value: '#AB12CD' },
  transparency: 25,
  width: 2.5,
  dash: 'dashDot',
};
const textShapeLineSlide = textShapeLineDeck.addSlide({
  masterName: textShapeLineLayout.name,
});
const packedPlainTextLine = textShapeLineSlide.addText('Packed plain text line', {
  name: 'packed_plain_text_line',
  line: textShapeLineSource,
});
const packedRichTextLine = textShapeLineSlide.addRichText([{
  runs: [{ text: 'Packed rich text line' }],
}], {
  name: 'packed_rich_text_line',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const packedPopulatedTextLine = textShapeLineSlide.addText('Packed populated text line', {
  placeholder: 'packed_title_line',
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '112233' },
    transparency: 25,
    width: 3,
    dash: 'lgDashDot',
  },
});
const initialPackedPlainTextLine = packedPlainTextLine.line;
textShapeLineSource.color.value = 'FFFFFF';
textShapeLineSource.transparency = 90;
textShapeLineSource.width = 9;
textShapeLineSource.dash = 'solid';
const detachedPackedPlainTextLine = packedPlainTextLine.line;
const packedDeclarativeTextLineLayout = await textShapeLineDeck.defineSlideMaster({
  title: 'PACKED-TEXT-LINES',
  objects: [
    {
      kind: 'text',
      text: 'Packed declarative text line',
      options: { name: 'packed_declarative_text_line', line: { kind: 'none' } },
    },
    {
      kind: 'placeholder',
      text: 'Packed declarative line prompt',
      options: {
        name: 'packed_declarative_title_line',
        type: 'title',
        index: 193,
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 40,
          width: 1.5,
          dash: 'sysDash',
        },
      },
    },
  ],
});
const packedDeclarativeTextLine = packedDeclarativeTextLineLayout.shapes.find(
  ({ name }) => name === 'packed_declarative_text_line',
);
const packedDeclarativePlaceholderLine = packedDeclarativeTextLineLayout.placeholders.find(
  ({ name }) => name === 'packed_declarative_title_line',
);
const packedDeclarativeTextLineSlide = textShapeLineDeck.addSlide({
  masterName: packedDeclarativeTextLineLayout.name,
});
const packedDeclarativePopulatedLine = packedDeclarativeTextLineSlide.addText(
  'Packed declarative populated line',
  {
    placeholder: 'packed_declarative_title_line',
    line: {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 50,
      width: 2,
      dash: 'lgDashDotDot',
    },
  },
);
const duplicateTextShapeLineSlide = textShapeLineDeck.duplicateSlide(0);
const duplicatePlainTextLine = duplicateTextShapeLineSlide.shapes.find(
  ({ name }) => name === 'packed_plain_text_line',
);
const duplicateRichTextLine = duplicateTextShapeLineSlide.shapes.find(
  ({ name }) => name === 'packed_rich_text_line',
);
if (!(duplicatePlainTextLine instanceof ShapeModel) ||
    !(duplicateRichTextLine instanceof ShapeModel)) {
  throw new Error('Packed duplicate text shape line failed');
}
duplicatePlainTextLine.line = { kind: 'none' };
duplicateRichTextLine.line = undefined;
const reopenedTextShapeLineDeck = await PptxDocument.open(await textShapeLineDeck.write());
await reopenedTextShapeLineDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedTextLineSourceSlide = reopenedTextShapeLineDeck.slides[0];
const reopenedTextLineDeclarativeSlide = reopenedTextShapeLineDeck.slides[1];
const reopenedTextLineDuplicateSlide = reopenedTextShapeLineDeck.slides[2];
const reopenedTextLineLayout = reopenedTextShapeLineDeck.layouts.find(
  ({ name }) => name === textShapeLineLayout.name,
);
const reopenedDeclarativeTextLineLayout = reopenedTextShapeLineDeck.layouts.find(
  ({ name }) => name === 'PACKED-TEXT-LINES',
);
const reopenedTextLineByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const textShapeLines =
  packedLayoutTextLine instanceof ShapeModel &&
  packedMasterTextLine instanceof ShapeModel &&
  packedLayoutPlaceholderLine instanceof ShapeModel &&
  packedPlainTextLine instanceof ShapeModel &&
  packedRichTextLine instanceof ShapeModel &&
  packedPopulatedTextLine instanceof ShapeModel &&
  packedDeclarativeTextLine instanceof ShapeModel &&
  packedDeclarativePlaceholderLine instanceof ShapeModel &&
  packedDeclarativePopulatedLine instanceof ShapeModel &&
  JSON.stringify(initialPackedPlainTextLine) === JSON.stringify({
    kind: 'line',
    color: { kind: 'srgb', value: 'AB12CD' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  }) &&
  JSON.stringify(detachedPackedPlainTextLine) === JSON.stringify(initialPackedPlainTextLine) &&
  packedLayoutTextLine.line?.kind === 'none' &&
  JSON.stringify(packedMasterTextLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent6' },
    transparency: 100,
    width: 0,
    dash: 'sysDot',
  }) &&
  JSON.stringify(packedLayoutPlaceholderLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 50,
    width: 1,
    dash: 'solid',
  }) &&
  JSON.stringify(packedRichTextLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
    width: 1,
    dash: 'solid',
  }) &&
  JSON.stringify(packedPopulatedTextLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'srgb', value: '112233' },
    transparency: 25,
    width: 3,
    dash: 'lgDashDot',
  }) &&
  packedDeclarativeTextLine.line?.kind === 'none' &&
  JSON.stringify(packedDeclarativePlaceholderLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 40,
    width: 1.5,
    dash: 'sysDash',
  }) &&
  JSON.stringify(packedDeclarativePopulatedLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent3' },
    transparency: 50,
    width: 2,
    dash: 'lgDashDotDot',
  }) &&
  JSON.stringify(reopenedTextLineByName(
    reopenedTextLineSourceSlide,
    'packed_plain_text_line',
  )?.line) === JSON.stringify(initialPackedPlainTextLine) &&
  JSON.stringify(reopenedTextLineByName(
    reopenedTextLineSourceSlide,
    'packed_rich_text_line',
  )?.line) === JSON.stringify(packedRichTextLine.line) &&
  reopenedTextLineByName(
    reopenedTextLineDuplicateSlide,
    'packed_plain_text_line',
  )?.line?.kind === 'none' &&
  reopenedTextLineByName(
    reopenedTextLineDuplicateSlide,
    'packed_rich_text_line',
  )?.line === undefined &&
  reopenedTextLineByName(
    reopenedTextLineLayout,
    'packed_layout_text_line',
  )?.line?.kind === 'none' &&
  JSON.stringify(reopenedTextLineByName(
    reopenedTextShapeLineDeck.masters[0],
    'packed_master_text_line',
  )?.line) === JSON.stringify(packedMasterTextLine.line) &&
  reopenedTextLineByName(
    reopenedDeclarativeTextLineLayout,
    'packed_declarative_text_line',
  )?.line?.kind === 'none' &&
  JSON.stringify(reopenedTextLineByName(
    reopenedTextLineDeclarativeSlide,
    'packed_declarative_title_line',
  )?.line) === JSON.stringify(packedDeclarativePopulatedLine.line) &&
  reopenedTextShapeLineDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!textShapeLines) {
  throw new Error('Packed text shape lines failed');
}
await reopenedTextShapeLineDeck.writeFile('text-shape-line-smoke.pptx');
const textShapeArrowDeck = PptxDocument.create();
const textShapeArrowLayout = textShapeArrowDeck.layouts[0];
const textShapeArrowMaster = textShapeArrowDeck.masters[0];
const packedLayoutTextArrow = textShapeArrowLayout.addText('Packed layout text arrow', {
  name: 'packed_layout_text_arrow',
  arrows: { begin: 'none' },
});
const packedMasterTextArrow = textShapeArrowMaster.addRichText([{
  runs: [{ text: 'Packed master text arrow' }],
}], {
  name: 'packed_master_text_arrow',
  arrows: { end: 'triangle' },
});
const packedLayoutPlaceholderArrow = textShapeArrowLayout.addPlaceholder(
  'Packed text arrow prompt',
  {
    name: 'packed_title_arrow',
    type: 'title',
    index: 194,
    arrows: { begin: 'stealth', end: 'none' },
  },
);
const textShapeArrowSource = { begin: 'triangle', end: 'arrow' };
const textShapeArrowSlide = textShapeArrowDeck.addSlide({
  masterName: textShapeArrowLayout.name,
});
const packedPlainTextArrow = textShapeArrowSlide.addText('Packed plain text arrow', {
  name: 'packed_plain_text_arrow',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: textShapeArrowSource,
});
const packedRichTextArrow = textShapeArrowSlide.addRichText([{
  runs: [{ text: 'Packed rich text arrow' }],
}], {
  name: 'packed_rich_text_arrow',
  arrows: { begin: 'diamond' },
});
const packedPopulatedTextArrow = textShapeArrowSlide.addText('Packed populated text arrow', {
  placeholder: 'packed_title_arrow',
  arrows: { end: 'oval' },
});
const initialPackedPlainTextArrow = packedPlainTextArrow.arrows;
textShapeArrowSource.begin = 'oval';
textShapeArrowSource.end = 'diamond';
const detachedPackedPlainTextArrow = packedPlainTextArrow.arrows;
const packedDeclarativeTextArrowLayout = await textShapeArrowDeck.defineSlideMaster({
  title: 'PACKED-TEXT-ARROWS',
  objects: [
    {
      kind: 'text',
      text: 'Packed declarative text arrow',
      options: {
        name: 'packed_declarative_text_arrow',
        arrows: { begin: 'none', end: 'stealth' },
      },
    },
    {
      kind: 'placeholder',
      text: 'Packed declarative arrow prompt',
      options: {
        name: 'packed_declarative_title_arrow',
        type: 'title',
        index: 195,
        arrows: { begin: 'arrow', end: 'triangle' },
      },
    },
  ],
});
const packedDeclarativeTextArrow = packedDeclarativeTextArrowLayout.shapes.find(
  ({ name }) => name === 'packed_declarative_text_arrow',
);
const packedDeclarativePlaceholderArrow = packedDeclarativeTextArrowLayout.placeholders.find(
  ({ name }) => name === 'packed_declarative_title_arrow',
);
const packedDeclarativeTextArrowSlide = textShapeArrowDeck.addSlide({
  masterName: packedDeclarativeTextArrowLayout.name,
});
const packedDeclarativePopulatedArrow = packedDeclarativeTextArrowSlide.addText(
  'Packed declarative populated arrow',
  {
    placeholder: 'packed_declarative_title_arrow',
    arrows: { end: 'diamond' },
  },
);
const duplicateTextShapeArrowSlide = textShapeArrowDeck.duplicateSlide(0);
const duplicatePlainTextArrow = duplicateTextShapeArrowSlide.shapes.find(
  ({ name }) => name === 'packed_plain_text_arrow',
);
const duplicateRichTextArrow = duplicateTextShapeArrowSlide.shapes.find(
  ({ name }) => name === 'packed_rich_text_arrow',
);
if (!(duplicatePlainTextArrow instanceof ShapeModel) ||
    !(duplicateRichTextArrow instanceof ShapeModel)) {
  throw new Error('Packed duplicate text shape arrow failed');
}
duplicatePlainTextArrow.line = undefined;
const duplicatePlainArrowsAfterLineClear = duplicatePlainTextArrow.arrows;
duplicatePlainTextArrow.arrows = { begin: 'oval' };
duplicateRichTextArrow.arrows = undefined;
const reopenedTextShapeArrowDeck = await PptxDocument.open(await textShapeArrowDeck.write());
await reopenedTextShapeArrowDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedTextArrowSourceSlide = reopenedTextShapeArrowDeck.slides[0];
const reopenedTextArrowDeclarativeSlide = reopenedTextShapeArrowDeck.slides[1];
const reopenedTextArrowDuplicateSlide = reopenedTextShapeArrowDeck.slides[2];
const reopenedTextArrowLayout = reopenedTextShapeArrowDeck.layouts.find(
  ({ name }) => name === textShapeArrowLayout.name,
);
const reopenedDeclarativeTextArrowLayout = reopenedTextShapeArrowDeck.layouts.find(
  ({ name }) => name === 'PACKED-TEXT-ARROWS',
);
const reopenedTextArrowByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const textShapeArrows =
  packedLayoutTextArrow instanceof ShapeModel &&
  packedMasterTextArrow instanceof ShapeModel &&
  packedLayoutPlaceholderArrow instanceof ShapeModel &&
  packedPlainTextArrow instanceof ShapeModel &&
  packedRichTextArrow instanceof ShapeModel &&
  packedPopulatedTextArrow instanceof ShapeModel &&
  packedDeclarativeTextArrow instanceof ShapeModel &&
  packedDeclarativePlaceholderArrow instanceof ShapeModel &&
  packedDeclarativePopulatedArrow instanceof ShapeModel &&
  JSON.stringify(initialPackedPlainTextArrow) === JSON.stringify({
    begin: 'triangle',
    end: 'arrow',
  }) &&
  JSON.stringify(detachedPackedPlainTextArrow) ===
    JSON.stringify(initialPackedPlainTextArrow) &&
  JSON.stringify(duplicatePlainArrowsAfterLineClear) ===
    JSON.stringify(initialPackedPlainTextArrow) &&
  JSON.stringify(packedLayoutTextArrow.arrows) === JSON.stringify({ begin: 'none' }) &&
  JSON.stringify(packedMasterTextArrow.arrows) === JSON.stringify({ end: 'triangle' }) &&
  JSON.stringify(packedLayoutPlaceholderArrow.arrows) === JSON.stringify({
    begin: 'stealth',
    end: 'none',
  }) &&
  JSON.stringify(packedRichTextArrow.arrows) === JSON.stringify({ begin: 'diamond' }) &&
  JSON.stringify(packedPopulatedTextArrow.arrows) === JSON.stringify({ end: 'oval' }) &&
  JSON.stringify(packedDeclarativeTextArrow.arrows) === JSON.stringify({
    begin: 'none',
    end: 'stealth',
  }) &&
  JSON.stringify(packedDeclarativePlaceholderArrow.arrows) === JSON.stringify({
    begin: 'arrow',
    end: 'triangle',
  }) &&
  JSON.stringify(packedDeclarativePopulatedArrow.arrows) ===
    JSON.stringify({ end: 'diamond' }) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextArrowSourceSlide,
    'packed_plain_text_arrow',
  )?.arrows) === JSON.stringify(initialPackedPlainTextArrow) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextArrowSourceSlide,
    'packed_rich_text_arrow',
  )?.arrows) === JSON.stringify({ begin: 'diamond' }) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextArrowDuplicateSlide,
    'packed_plain_text_arrow',
  )?.arrows) === JSON.stringify({ begin: 'oval' }) &&
  reopenedTextArrowByName(
    reopenedTextArrowDuplicateSlide,
    'packed_plain_text_arrow',
  )?.line === undefined &&
  reopenedTextArrowByName(
    reopenedTextArrowDuplicateSlide,
    'packed_rich_text_arrow',
  )?.arrows === undefined &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextArrowLayout,
    'packed_layout_text_arrow',
  )?.arrows) === JSON.stringify({ begin: 'none' }) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextShapeArrowDeck.masters[0],
    'packed_master_text_arrow',
  )?.arrows) === JSON.stringify({ end: 'triangle' }) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedDeclarativeTextArrowLayout,
    'packed_declarative_text_arrow',
  )?.arrows) === JSON.stringify({ begin: 'none', end: 'stealth' }) &&
  JSON.stringify(reopenedTextArrowByName(
    reopenedTextArrowDeclarativeSlide,
    'packed_declarative_title_arrow',
  )?.arrows) === JSON.stringify({ end: 'diamond' }) &&
  reopenedTextShapeArrowDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!textShapeArrows) {
  throw new Error('Packed text shape arrows failed');
}
await reopenedTextShapeArrowDeck.writeFile('text-shape-arrows-smoke.pptx');
const textShapeShadowDeck = PptxDocument.create();
const textShapeShadowLayout = textShapeShadowDeck.layouts[0];
const textShapeShadowMaster = textShapeShadowDeck.masters[0];
const packedLayoutTextShadow = textShapeShadowLayout.addText('Packed layout text shadow', {
  name: 'packed_layout_text_shadow',
  shadow: {
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent2' },
    rotateWithShape: true,
  },
});
const packedMasterTextShadow = textShapeShadowMaster.addRichText([{
  runs: [{ text: 'Packed master text shadow' }],
}], {
  name: 'packed_master_text_shadow',
  shadow: {
    kind: 'inner',
    color: { kind: 'srgb', value: '445566' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
  },
});
const packedLayoutPlaceholderShadow = textShapeShadowLayout.addPlaceholder(
  'Packed text shadow prompt',
  {
    name: 'packed_title_shadow',
    type: 'title',
    index: 196,
    shadow: { kind: 'outer', color: { kind: 'srgb', value: '112233' } },
  },
);
const textShapeShadowColor = { kind: 'scheme', value: 'accent4' };
const textShapeShadowSource = {
  kind: 'outer',
  color: textShapeShadowColor,
  opacity: 0.4,
  blur: 2,
  angle: 45,
  distance: 3,
  rotateWithShape: true,
};
const textShapeShadowSlide = textShapeShadowDeck.addSlide({
  masterName: textShapeShadowLayout.name,
});
const packedPlainTextShadow = textShapeShadowSlide.addText('Packed plain text shadow', {
  name: 'packed_plain_text_shadow',
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  shadow: textShapeShadowSource,
});
const packedRichTextShadow = textShapeShadowSlide.addRichText([{
  runs: [{ text: 'Packed rich text shadow' }],
}], {
  name: 'packed_rich_text_shadow',
  shadow: {
    kind: 'inner',
    color: { kind: 'srgb', value: '667788' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
  },
});
const packedPopulatedTextShadow = textShapeShadowSlide.addText(
  'Packed populated text shadow',
  {
    placeholder: 'packed_title_shadow',
    shadow: {
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
      opacity: 0.5,
      blur: 1,
      angle: 90,
      distance: 2,
    },
  },
);
const initialPackedPlainTextShadow = packedPlainTextShadow.shadow;
const initialPackedPlainTextShadowAgain = packedPlainTextShadow.shadow;
textShapeShadowColor.value = 'accent6';
textShapeShadowSource.opacity = 0.9;
textShapeShadowSource.blur = 9;
textShapeShadowSource.angle = 180;
textShapeShadowSource.distance = 8;
textShapeShadowSource.rotateWithShape = false;
const detachedPackedPlainTextShadow = packedPlainTextShadow.shadow;
const textShapeShadowNoOpBytes = textShapeShadowDeck.opcPackage
  .requirePart(textShapeShadowSlide.partUri).bytes.slice();
const textShapeShadowNoOpJournal = textShapeShadowDeck.opcPackage.mutations.length;
packedPlainTextShadow.shadow = {
  kind: 'outer',
  color: { kind: 'scheme', value: 'accent4' },
  opacity: 0.4,
  blur: 2,
  angle: 45,
  distance: 3,
  rotateWithShape: true,
};
const textShapeShadowNoOpCurrent = textShapeShadowDeck.opcPackage
  .requirePart(textShapeShadowSlide.partUri).bytes;
const textShapeShadowNoOp =
  textShapeShadowNoOpJournal === textShapeShadowDeck.opcPackage.mutations.length &&
  textShapeShadowNoOpBytes.length === textShapeShadowNoOpCurrent.length &&
  textShapeShadowNoOpBytes.every(
    (value, index) => value === textShapeShadowNoOpCurrent[index],
  );
const packedDeclarativeTextShadowLayout = await textShapeShadowDeck.defineSlideMaster({
  title: 'PACKED-TEXT-SHADOWS',
  objects: [
    {
      kind: 'text',
      text: 'Packed declarative text shadow',
      options: {
        name: 'packed_declarative_text_shadow',
        shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent5' } },
      },
    },
    {
      kind: 'placeholder',
      text: 'Packed declarative shadow prompt',
      options: {
        name: 'packed_declarative_title_shadow',
        type: 'title',
        index: 197,
        shadow: { kind: 'inner', opacity: 0.25 },
      },
    },
  ],
});
const packedDeclarativeTextShadow = packedDeclarativeTextShadowLayout.shapes.find(
  ({ name }) => name === 'packed_declarative_text_shadow',
);
const packedDeclarativePlaceholderShadow = packedDeclarativeTextShadowLayout.placeholders.find(
  ({ name }) => name === 'packed_declarative_title_shadow',
);
const packedDeclarativeTextShadowSlide = textShapeShadowDeck.addSlide({
  masterName: packedDeclarativeTextShadowLayout.name,
});
const packedDeclarativePopulatedShadow = packedDeclarativeTextShadowSlide.addText(
  'Packed declarative populated shadow',
  {
    placeholder: 'packed_declarative_title_shadow',
    shadow: { kind: 'outer', rotateWithShape: true },
  },
);
const duplicateTextShapeShadowSlide = textShapeShadowDeck.duplicateSlide(0);
const duplicatePlainTextShadow = duplicateTextShapeShadowSlide.shapes.find(
  ({ name }) => name === 'packed_plain_text_shadow',
);
const duplicateRichTextShadow = duplicateTextShapeShadowSlide.shapes.find(
  ({ name }) => name === 'packed_rich_text_shadow',
);
if (!(duplicatePlainTextShadow instanceof ShapeModel) ||
    !(duplicateRichTextShadow instanceof ShapeModel)) {
  throw new Error('Packed duplicate text shape shadow failed');
}
duplicatePlainTextShadow.shadow = undefined;
const duplicatePlainLineAfterShadowClear = duplicatePlainTextShadow.line;
const duplicatePlainArrowsAfterShadowClear = duplicatePlainTextShadow.arrows;
duplicatePlainTextShadow.shadow = { kind: 'outer' };
duplicatePlainTextShadow.line = undefined;
duplicatePlainTextShadow.arrows = undefined;
const duplicatePlainShadowAfterLineArrowClear = duplicatePlainTextShadow.shadow;
duplicateRichTextShadow.shadow = undefined;
const reopenedTextShapeShadowDeck = await PptxDocument.open(await textShapeShadowDeck.write());
await reopenedTextShapeShadowDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedTextShadowSourceSlide = reopenedTextShapeShadowDeck.slides[0];
const reopenedTextShadowDeclarativeSlide = reopenedTextShapeShadowDeck.slides[1];
const reopenedTextShadowDuplicateSlide = reopenedTextShapeShadowDeck.slides[2];
const reopenedTextShadowLayout = reopenedTextShapeShadowDeck.layouts.find(
  ({ name }) => name === textShapeShadowLayout.name,
);
const reopenedDeclarativeTextShadowLayout = reopenedTextShapeShadowDeck.layouts.find(
  ({ name }) => name === 'PACKED-TEXT-SHADOWS',
);
const reopenedTextShadowByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const textShapeShadows =
  packedLayoutTextShadow instanceof ShapeModel &&
  packedMasterTextShadow instanceof ShapeModel &&
  packedLayoutPlaceholderShadow instanceof ShapeModel &&
  packedPlainTextShadow instanceof ShapeModel &&
  packedRichTextShadow instanceof ShapeModel &&
  packedPopulatedTextShadow instanceof ShapeModel &&
  packedDeclarativeTextShadow instanceof ShapeModel &&
  packedDeclarativePlaceholderShadow instanceof ShapeModel &&
  packedDeclarativePopulatedShadow instanceof ShapeModel &&
  initialPackedPlainTextShadow !== initialPackedPlainTextShadowAgain &&
  initialPackedPlainTextShadow?.color !== initialPackedPlainTextShadowAgain?.color &&
  Object.isFrozen(initialPackedPlainTextShadow) &&
  Object.isFrozen(initialPackedPlainTextShadow?.color) &&
  JSON.stringify(initialPackedPlainTextShadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent4' },
    opacity: 0.4,
    blur: 2,
    angle: 45,
    distance: 3,
    rotateWithShape: true,
  }) &&
  JSON.stringify(detachedPackedPlainTextShadow) ===
    JSON.stringify(initialPackedPlainTextShadow) &&
  textShapeShadowNoOp &&
  JSON.stringify(duplicatePlainLineAfterShadowClear) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  }) &&
  JSON.stringify(duplicatePlainArrowsAfterShadowClear) ===
    JSON.stringify({ begin: 'triangle', end: 'arrow' }) &&
  duplicatePlainShadowAfterLineArrowClear?.kind === 'outer' &&
  JSON.stringify(packedLayoutTextShadow.shadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent2' },
    opacity: 0.75,
    blur: 8,
    angle: 270,
    distance: 4,
    rotateWithShape: true,
  }) &&
  packedMasterTextShadow.shadow?.kind === 'inner' &&
  packedLayoutPlaceholderShadow.shadow?.kind === 'outer' &&
  packedRichTextShadow.shadow?.kind === 'inner' &&
  packedPopulatedTextShadow.shadow?.kind === 'inner' &&
  packedDeclarativeTextShadow.shadow?.kind === 'outer' &&
  packedDeclarativePlaceholderShadow.shadow?.kind === 'inner' &&
  packedDeclarativePopulatedShadow.shadow?.kind === 'outer' &&
  JSON.stringify(reopenedTextShadowByName(
    reopenedTextShadowSourceSlide,
    'packed_plain_text_shadow',
  )?.shadow) === JSON.stringify(initialPackedPlainTextShadow) &&
  reopenedTextShadowByName(
    reopenedTextShadowSourceSlide,
    'packed_rich_text_shadow',
  )?.shadow?.kind === 'inner' &&
  reopenedTextShadowByName(
    reopenedTextShadowDuplicateSlide,
    'packed_plain_text_shadow',
  )?.shadow?.kind === 'outer' &&
  reopenedTextShadowByName(
    reopenedTextShadowDuplicateSlide,
    'packed_plain_text_shadow',
  )?.line === undefined &&
  reopenedTextShadowByName(
    reopenedTextShadowDuplicateSlide,
    'packed_plain_text_shadow',
  )?.arrows === undefined &&
  reopenedTextShadowByName(
    reopenedTextShadowDuplicateSlide,
    'packed_rich_text_shadow',
  )?.shadow === undefined &&
  reopenedTextShadowByName(
    reopenedTextShadowLayout,
    'packed_layout_text_shadow',
  )?.shadow?.kind === 'outer' &&
  reopenedTextShadowByName(
    reopenedTextShapeShadowDeck.masters[0],
    'packed_master_text_shadow',
  )?.shadow?.kind === 'inner' &&
  reopenedTextShadowByName(
    reopenedDeclarativeTextShadowLayout,
    'packed_declarative_text_shadow',
  )?.shadow?.kind === 'outer' &&
  reopenedTextShadowByName(
    reopenedTextShadowDeclarativeSlide,
    'packed_declarative_title_shadow',
  )?.shadow?.kind === 'outer' &&
  reopenedTextShapeShadowDeck.diagnostics.every(({ severity }) => severity !== 'error');
if (!textShapeShadows) {
  throw new Error('Packed text shape shadows failed');
}
await reopenedTextShapeShadowDeck.writeFile('text-shape-shadows-smoke.pptx');
const textShapeHyperlinkDeck = PptxDocument.create();
const textShapeHyperlinkLayout = textShapeHyperlinkDeck.layouts[0];
const textShapeHyperlinkMaster = textShapeHyperlinkDeck.masters[0];
const textShapeHyperlinkTargetLayout = await textShapeHyperlinkDeck.defineSlideMaster({
  title: 'PACKED-TEXT-HYPERLINK-TARGET',
  objects: [],
});
const textShapeHyperlinkTargetSlide = textShapeHyperlinkDeck.addSlide({
  masterName: textShapeHyperlinkTargetLayout.name,
});
const packedLayoutTextHyperlink = textShapeHyperlinkLayout.addText(
  'Packed layout text hyperlink',
  {
    name: 'packed_layout_text_hyperlink',
    hyperlink: { url: 'https://layout.example', tooltip: 'Layout' },
  },
);
const packedMasterTextHyperlink = textShapeHyperlinkMaster.addRichText([{
  runs: [{ text: 'Packed master text hyperlink' }],
}], {
  name: 'packed_master_text_hyperlink',
  hyperlink: { slide: 1 },
});
const packedLayoutPlaceholderHyperlink = textShapeHyperlinkLayout.addPlaceholder(
  'Packed text hyperlink prompt',
  {
    name: 'packed_title_hyperlink',
    type: 'title',
    index: 198,
    hyperlink: { url: 'https://placeholder.example', tooltip: '' },
  },
);
const textShapeHyperlinkSourceSlide = textShapeHyperlinkDeck.addSlide({
  masterName: textShapeHyperlinkLayout.name,
});
const textShapeHyperlinkInput = {
  url: 'https://text.example/path?a=1&b=2',
  tooltip: 'Packed & linked',
};
const packedPlainTextHyperlink = textShapeHyperlinkSourceSlide.addText(
  'Packed plain text hyperlink\\nSecond line',
  {
    name: 'packed_plain_text_hyperlink',
    fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
    line: {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      width: 2,
      dash: 'dashDot',
    },
    arrows: { begin: 'triangle', end: 'arrow' },
    shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent4' } },
    hyperlink: textShapeHyperlinkInput,
  },
);
const packedRichTextHyperlink = textShapeHyperlinkSourceSlide.addRichText([{
  runs: [
    { text: 'Packed rich one' },
    { text: ' and two', style: { underline: false } },
  ],
}], {
  name: 'packed_rich_text_hyperlink',
  hyperlink: { slide: 1, tooltip: '' },
});
const packedPopulatedTextHyperlink = textShapeHyperlinkSourceSlide.addText(
  'Packed populated text hyperlink',
  {
    placeholder: 'packed_title_hyperlink',
    hyperlink: { slide: 2 },
  },
);
const packedEditableTextHyperlink = textShapeHyperlinkSourceSlide.addText(
  'Packed editable text hyperlink',
  {
    name: 'packed_editable_text_hyperlink',
    hyperlink: { url: 'https://runs.example', tooltip: 'Runs' },
  },
);
const initialPackedTextHyperlink = packedPlainTextHyperlink.hyperlink;
const initialPackedTextHyperlinkAgain = packedPlainTextHyperlink.hyperlink;
textShapeHyperlinkInput.url = 'https://changed.example';
textShapeHyperlinkInput.tooltip = 'Changed';
const detachedPackedTextHyperlink = packedPlainTextHyperlink.hyperlink;
const packedDeclarativeTextHyperlinkLayout = await textShapeHyperlinkDeck.defineSlideMaster({
  title: 'PACKED-TEXT-HYPERLINKS',
  objects: [
    {
      kind: 'text',
      text: 'Packed declarative text hyperlink',
      options: {
        name: 'packed_declarative_text_hyperlink',
        hyperlink: { url: 'https://declarative.example' },
      },
    },
    {
      kind: 'placeholder',
      text: 'Packed declarative hyperlink prompt',
      options: {
        name: 'packed_declarative_title_hyperlink',
        type: 'title',
        index: 199,
        hyperlink: { slide: 1, tooltip: 'Target' },
      },
    },
  ],
});
const packedDeclarativeTextHyperlink = packedDeclarativeTextHyperlinkLayout.shapes.find(
  ({ name }) => name === 'packed_declarative_text_hyperlink',
);
const packedDeclarativePlaceholderHyperlink =
  packedDeclarativeTextHyperlinkLayout.placeholders.find(
    ({ name }) => name === 'packed_declarative_title_hyperlink',
  );
const packedDeclarativeTextHyperlinkSlide = textShapeHyperlinkDeck.addSlide({
  masterName: packedDeclarativeTextHyperlinkLayout.name,
});
const packedDeclarativePopulatedHyperlink = packedDeclarativeTextHyperlinkSlide.addText(
  'Packed declarative populated hyperlink',
  {
    placeholder: 'packed_declarative_title_hyperlink',
    hyperlink: { slide: 3 },
  },
);
const packedTextHyperlinkShapeXml = (owner, name) => {
  const shape = owner.shapes.find(
    (candidate) => candidate instanceof ShapeModel && candidate.name === name,
  );
  if (!(shape instanceof ShapeModel)) throw new Error('Packed text hyperlink shape missing');
  const xml = new TextDecoder().decode(
    textShapeHyperlinkDeck.opcPackage.requirePart(owner.partUri).bytes,
  );
  const nameOffset = xml.indexOf('name="' + name + '"');
  const shapeStart = xml.lastIndexOf('<p:sp', nameOffset);
  const shapeEnd = xml.indexOf('</p:sp>', nameOffset);
  if (nameOffset < 0 || shapeStart < 0 || shapeEnd < 0) {
    throw new Error('Packed text hyperlink XML missing');
  }
  return xml.slice(shapeStart, shapeEnd + '</p:sp>'.length);
};
const packedTextHyperlinkIds = (xml) => xml.split('<a:hlinkClick').slice(1).map(
  (fragment) => fragment.split('r:id="')[1]?.split('"')[0],
);
const packedPlainTextHyperlinkXml = packedTextHyperlinkShapeXml(
  textShapeHyperlinkSourceSlide,
  'packed_plain_text_hyperlink',
);
const packedPlainTextHyperlinkIds = packedTextHyperlinkIds(packedPlainTextHyperlinkXml);
const packedRichTextHyperlinkXml = packedTextHyperlinkShapeXml(
  textShapeHyperlinkSourceSlide,
  'packed_rich_text_hyperlink',
);
const packedRichTextHyperlinkIds = packedTextHyperlinkIds(packedRichTextHyperlinkXml);
const packedPlainTextHyperlinkRelationship = textShapeHyperlinkSourceSlide.relationships.find(
  ({ id }) => id === packedPlainTextHyperlinkIds[0],
);
const packedRichTextHyperlinkRelationship = textShapeHyperlinkSourceSlide.relationships.find(
  ({ id }) => id === packedRichTextHyperlinkIds[0],
);
const packedTextHyperlinkDualOutput =
  packedPlainTextHyperlinkIds.length === 3 &&
  new Set(packedPlainTextHyperlinkIds).size === 1 &&
  packedPlainTextHyperlinkRelationship?.type.endsWith('/hyperlink') === true &&
  packedPlainTextHyperlinkRelationship.target === 'https://text.example/path?a=1&b=2' &&
  packedPlainTextHyperlinkRelationship.targetMode === 'External' &&
  packedPlainTextHyperlinkXml.includes('tooltip="Packed &amp; linked"') &&
  packedRichTextHyperlinkIds.length === 3 &&
  new Set(packedRichTextHyperlinkIds).size === 1 &&
  packedRichTextHyperlinkRelationship?.type.endsWith('/slide') === true &&
  packedRichTextHyperlinkRelationship.resolvedTarget === textShapeHyperlinkTargetSlide.partUri &&
  packedRichTextHyperlinkXml.split('<a:hlinkClick').slice(1).every(
    (fragment) => fragment.includes('action="ppaction://hlinksldjump"'),
  );
const packedEditableOriginalXml = packedTextHyperlinkShapeXml(
  textShapeHyperlinkSourceSlide,
  'packed_editable_text_hyperlink',
);
const packedEditableOriginalRelationshipId = packedTextHyperlinkIds(
  packedEditableOriginalXml,
)[0];
packedEditableTextHyperlink.hyperlink = { slide: 1, tooltip: '' };
const packedEditableReplacementXml = packedTextHyperlinkShapeXml(
  textShapeHyperlinkSourceSlide,
  'packed_editable_text_hyperlink',
);
const packedEditableReplacementRelationshipId = packedTextHyperlinkIds(
  packedEditableReplacementXml,
)[0];
const packedEditableReplacementRelationship = textShapeHyperlinkSourceSlide.relationships.find(
  ({ id }) => id === packedEditableReplacementRelationshipId,
);
packedEditableTextHyperlink.hyperlink = undefined;
const packedEditableClearedXml = packedTextHyperlinkShapeXml(
  textShapeHyperlinkSourceSlide,
  'packed_editable_text_hyperlink',
);
const packedTextHyperlinkRunOwnership =
  packedEditableTextHyperlink.hyperlink === undefined &&
  packedEditableReplacementRelationship !== undefined &&
  !textShapeHyperlinkSourceSlide.relationships.some(
    ({ id }) => id === packedEditableReplacementRelationship.id,
  ) &&
  textShapeHyperlinkSourceSlide.relationships.some(
    ({ id }) => id === packedEditableOriginalRelationshipId,
  ) &&
  packedEditableClearedXml.slice(
    0,
    packedEditableClearedXml.indexOf('</p:nvSpPr>'),
  ).split('<a:hlinkClick').length === 1 &&
  packedTextHyperlinkIds(packedEditableClearedXml).length === 1 &&
  packedTextHyperlinkIds(packedEditableClearedXml)[0] ===
    packedEditableOriginalRelationshipId;
const duplicateTextShapeHyperlinkSlide = textShapeHyperlinkDeck.duplicateSlide(1);
const duplicateTextShapeHyperlinkSelf = duplicateTextShapeHyperlinkSlide.shapes.find(
  ({ name }) => name === 'packed_title_hyperlink',
);
textShapeHyperlinkDeck.moveSlide(
  textShapeHyperlinkDeck.slides.indexOf(textShapeHyperlinkTargetSlide),
  1,
);
const textShapeHyperlinkMoveIdentity =
  JSON.stringify(packedRichTextHyperlink.hyperlink) === JSON.stringify({ slide: 2, tooltip: '' }) &&
  JSON.stringify(packedPopulatedTextHyperlink.hyperlink) === JSON.stringify({ slide: 1 }) &&
  JSON.stringify(duplicateTextShapeHyperlinkSelf?.hyperlink) === JSON.stringify({ slide: 4 });
const reopenedTextShapeHyperlinkDeck = await PptxDocument.open(
  await textShapeHyperlinkDeck.write(),
);
await reopenedTextShapeHyperlinkDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedTextShapeHyperlinkByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const reopenedTextShapeHyperlinkSource = reopenedTextShapeHyperlinkDeck.slides[0];
const reopenedTextShapeHyperlinkDuplicate = reopenedTextShapeHyperlinkDeck.slides[3];
const reopenedTextShapeHyperlinkLayout = reopenedTextShapeHyperlinkDeck.layouts.find(
  ({ name }) => name === textShapeHyperlinkLayout.name,
);
const reopenedDeclarativeTextHyperlinkLayout = reopenedTextShapeHyperlinkDeck.layouts.find(
  ({ name }) => name === 'PACKED-TEXT-HYPERLINKS',
);
const textShapeHyperlinkWarnings = reopenedTextShapeHyperlinkDeck.diagnostics.filter(
  ({ severity }) => severity === 'warning',
);
const textShapeHyperlinks =
  packedLayoutTextHyperlink instanceof ShapeModel &&
  packedMasterTextHyperlink instanceof ShapeModel &&
  packedLayoutPlaceholderHyperlink instanceof ShapeModel &&
  packedPlainTextHyperlink instanceof ShapeModel &&
  packedRichTextHyperlink instanceof ShapeModel &&
  packedPopulatedTextHyperlink instanceof ShapeModel &&
  packedDeclarativeTextHyperlink instanceof ShapeModel &&
  packedDeclarativePlaceholderHyperlink instanceof ShapeModel &&
  packedDeclarativePopulatedHyperlink instanceof ShapeModel &&
  initialPackedTextHyperlink !== initialPackedTextHyperlinkAgain &&
  Object.isFrozen(initialPackedTextHyperlink) &&
  JSON.stringify(initialPackedTextHyperlink) === JSON.stringify({
    url: 'https://text.example/path?a=1&b=2',
    tooltip: 'Packed & linked',
  }) &&
  JSON.stringify(detachedPackedTextHyperlink) === JSON.stringify(initialPackedTextHyperlink) &&
  packedTextHyperlinkDualOutput &&
  packedTextHyperlinkRunOwnership &&
  textShapeHyperlinkMoveIdentity &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedTextShapeHyperlinkSource,
    'packed_plain_text_hyperlink',
  )?.hyperlink) === JSON.stringify(initialPackedTextHyperlink) &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedTextShapeHyperlinkSource,
    'packed_rich_text_hyperlink',
  )?.hyperlink) === JSON.stringify({ slide: 2, tooltip: '' }) &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedTextShapeHyperlinkDuplicate,
    'packed_title_hyperlink',
  )?.hyperlink) === JSON.stringify({ slide: 4 }) &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedTextShapeHyperlinkLayout,
    'packed_layout_text_hyperlink',
  )?.hyperlink) === JSON.stringify({ url: 'https://layout.example', tooltip: 'Layout' }) &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedTextShapeHyperlinkDeck.masters[0],
    'packed_master_text_hyperlink',
  )?.hyperlink) === JSON.stringify({ slide: 2 }) &&
  JSON.stringify(reopenedTextShapeHyperlinkByName(
    reopenedDeclarativeTextHyperlinkLayout,
    'packed_declarative_text_hyperlink',
  )?.hyperlink) === JSON.stringify({ url: 'https://declarative.example' }) &&
  reopenedTextShapeHyperlinkDeck.diagnostics.every(
    ({ severity, code }) => severity !== 'error' &&
      (severity !== 'warning' || code === 'OPC_EXTERNAL_RELATIONSHIP'),
  ) &&
  textShapeHyperlinkWarnings.length > 0;
if (!textShapeHyperlinks) {
  throw new Error('Packed text shape hyperlinks failed: ' + JSON.stringify({
    initialPackedTextHyperlink,
    detachedPackedTextHyperlink,
    packedTextHyperlinkDualOutput,
    packedTextHyperlinkRunOwnership,
    textShapeHyperlinkMoveIdentity,
    reopenedPlain: reopenedTextShapeHyperlinkByName(
      reopenedTextShapeHyperlinkSource,
      'packed_plain_text_hyperlink',
    )?.hyperlink,
    reopenedRich: reopenedTextShapeHyperlinkByName(
      reopenedTextShapeHyperlinkSource,
      'packed_rich_text_hyperlink',
    )?.hyperlink,
    reopenedDuplicate: reopenedTextShapeHyperlinkByName(
      reopenedTextShapeHyperlinkDuplicate,
      'packed_title_hyperlink',
    )?.hyperlink,
    reopenedLayout: reopenedTextShapeHyperlinkByName(
      reopenedTextShapeHyperlinkLayout,
      'packed_layout_text_hyperlink',
    )?.hyperlink,
    reopenedMaster: reopenedTextShapeHyperlinkByName(
      reopenedTextShapeHyperlinkDeck.masters[0],
      'packed_master_text_hyperlink',
    )?.hyperlink,
    reopenedDeclarative: reopenedTextShapeHyperlinkByName(
      reopenedDeclarativeTextHyperlinkLayout,
      'packed_declarative_text_hyperlink',
    )?.hyperlink,
    diagnostics: reopenedTextShapeHyperlinkDeck.diagnostics,
  }));
}
await reopenedTextShapeHyperlinkDeck.writeFile('text-shape-hyperlinks-smoke.pptx');
const internalTextShapeHyperlinkDeck = PptxDocument.create();
const internalTextShapeHyperlinkSource = internalTextShapeHyperlinkDeck.addSlide();
internalTextShapeHyperlinkDeck.addSlide();
internalTextShapeHyperlinkSource.addText('Packed internal-only text hyperlink', {
  name: 'packed_internal_only_text_hyperlink',
  hyperlink: { slide: 2, tooltip: '' },
});
internalTextShapeHyperlinkSource.addRichText([{
  runs: [{ text: 'Packed internal-only rich hyperlink' }],
}], {
  name: 'packed_internal_only_rich_hyperlink',
  hyperlink: { slide: 1 },
});
const reopenedInternalTextShapeHyperlinkDeck = await PptxDocument.open(
  await internalTextShapeHyperlinkDeck.write(),
);
await reopenedInternalTextShapeHyperlinkDeck.write({ compatibility: 'powerpoint-2010' });
if (reopenedInternalTextShapeHyperlinkDeck.diagnostics.length !== 0) {
  throw new Error('Packed internal-only text shape hyperlink validation failed');
}
await reopenedInternalTextShapeHyperlinkDeck.writeFile(
  'text-shape-hyperlinks-internal-smoke.pptx',
);
const createdText = created.addSlide().addText('Smoke\\n\\nParagraph', { align: 'center', fit: 'shrink', valign: 'top', vert: 'vert270', wrap: false, bullet: true, level: 2, margin: 10, rtlMode: true, spacing: { before: 4, after: 6, line: { kind: 'exact', points: 20 } }, tabStops: [{ position: 1.25 }, { position: 2.5, alignment: 'right' }] });
const shapeLineDeck = PptxDocument.create();
const shapeLineSlide = shapeLineDeck.addSlide();
const shapeLineSourceColor = { kind: 'srgb', value: '#AA0000' };
const shapeLineSource = {
  kind: 'line',
  color: shapeLineSourceColor,
  transparency: 33.3334,
  width: 2.50001,
  dash: 'dashDot',
};
const packedSrgbLine = shapeLineSlide.addShape('rect', {
  name: 'Packed sRGB line',
  line: shapeLineSource,
});
const packedSchemeLine = shapeLineSlide.addShape('ellipse', {
  name: 'Packed scheme line',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
  },
});
const packedNoneLine = shapeLineSlide.addShape('star5', {
  name: 'Packed none line',
  line: { kind: 'none' },
});
const packedZeroWidthLine = shapeLineSlide.addShape('diamond', {
  name: 'Packed zero-width line',
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '112233' },
    width: 0,
    dash: 'sysDot',
  },
});
const packedLineDashes = [
  'solid', 'dash', 'dashDot', 'lgDash',
  'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot',
];
const packedDashLines = packedLineDashes.map((dash) => shapeLineSlide.addShape('roundRect', {
  name: 'Packed dash ' + dash,
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '445566' },
    dash,
  },
}));
const initialPackedSrgbLine = packedSrgbLine.line;
const initialPackedSrgbLineAgain = packedSrgbLine.line;
shapeLineSourceColor.value = '000000';
shapeLineSource.transparency = 1;
shapeLineSource.width = 9;
shapeLineSource.dash = 'solid';
const detachedPackedSrgbLine = packedSrgbLine.line;
const shapeLinePartCountBeforeEdit = shapeLineDeck.opcPackage.parts.length;
const shapeLineRelationshipCountBeforeEdit = shapeLineSlide.relationships.length;
const shapeLineNoOpBytes = shapeLineDeck.opcPackage.requirePart(shapeLineSlide.partUri).bytes.slice();
const shapeLineNoOpJournalLength = shapeLineDeck.opcPackage.mutations.length;
packedSrgbLine.line = {
  kind: 'line',
  color: { kind: 'srgb', value: 'AA0000' },
  transparency: 33.333,
  width: 31_750 / 12_700,
  dash: 'dashDot',
};
const currentShapeLineNoOpBytes = shapeLineDeck.opcPackage.requirePart(shapeLineSlide.partUri).bytes;
const shapeLineNoOp = shapeLineNoOpJournalLength === shapeLineDeck.opcPackage.mutations.length &&
  shapeLineNoOpBytes.length === currentShapeLineNoOpBytes.length &&
  shapeLineNoOpBytes.every((value, index) => value === currentShapeLineNoOpBytes[index]);
const linePart = shapeLineDeck.opcPackage.requirePart(shapeLineSlide.partUri);
const lineXml = new TextDecoder().decode(linePart.bytes);
const lineXmlWithArrows = lineXml.replace(
  '<a:prstDash val="dashDot"/></a:ln>',
  '<a:prstDash val="dashDot"/><a:round/><a:headEnd type="triangle"/>' +
  '<a:tailEnd type="arrow"/><a:extLst><a:ext uri="urn:packed-line">' +
  '<x:keep xmlns:x="urn:packed-line"/></a:ext></a:extLst></a:ln>',
);
if (lineXmlWithArrows === lineXml) throw new Error('Packed line arrow injection failed');
shapeLineDeck.opcPackage.setPart(shapeLineSlide.partUri, lineXmlWithArrows, linePart.contentType);
packedSrgbLine.line = { kind: 'none' };
const packedNoneLineReplacement = packedSrgbLine.line;
packedSrgbLine.line = undefined;
const packedClearedLine = packedSrgbLine.line;
packedSrgbLine.line = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent4' },
  transparency: 40,
  width: 2,
  dash: 'sysDash',
};
const shapeLineEditedXml = new TextDecoder().decode(
  shapeLineDeck.opcPackage.requirePart(shapeLineSlide.partUri).bytes,
);
const shapeLinePreservedSiblings = shapeLineEditedXml.includes('<a:round/>') &&
  shapeLineEditedXml.includes('<a:headEnd type="triangle"/>') &&
  shapeLineEditedXml.includes('<a:tailEnd type="arrow"/>') &&
  shapeLineEditedXml.includes('<x:keep xmlns:x="urn:packed-line"/>');
const shapeLineEditIsolation = shapeLineDeck.opcPackage.parts.length ===
  shapeLinePartCountBeforeEdit &&
  shapeLineSlide.relationships.length === shapeLineRelationshipCountBeforeEdit;
const duplicateShapeLineSlide = shapeLineDeck.duplicateSlide(0);
const duplicatePackedLine = duplicateShapeLineSlide.shapes[0];
if (!(duplicatePackedLine instanceof ShapeModel)) throw new Error('Packed duplicate line shape failed');
duplicatePackedLine.line = { kind: 'none' };
const reopenedShapeLineDeck = await PptxDocument.open(await shapeLineDeck.write());
const reopenedSourceLines = reopenedShapeLineDeck.slides[0].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.line : undefined);
const reopenedDuplicateLines = reopenedShapeLineDeck.slides[1].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.line : undefined);
const shapeLines = packedSrgbLine instanceof ShapeModel &&
  packedSchemeLine instanceof ShapeModel &&
  packedNoneLine instanceof ShapeModel &&
  packedZeroWidthLine instanceof ShapeModel &&
  initialPackedSrgbLine !== initialPackedSrgbLineAgain &&
  initialPackedSrgbLine?.kind === 'line' &&
  initialPackedSrgbLineAgain?.kind === 'line' &&
  initialPackedSrgbLine.color !== initialPackedSrgbLineAgain.color &&
  JSON.stringify(initialPackedSrgbLine) === JSON.stringify({
    kind: 'line',
    color: { kind: 'srgb', value: 'AA0000' },
    transparency: 33.333,
    width: 31_750 / 12_700,
    dash: 'dashDot',
  }) &&
  JSON.stringify(detachedPackedSrgbLine) === JSON.stringify(initialPackedSrgbLine) &&
  packedNoneLineReplacement?.kind === 'none' &&
  packedClearedLine === undefined &&
  shapeLineNoOp &&
  shapeLineEditIsolation &&
  shapeLinePreservedSiblings &&
  JSON.stringify(packedSrgbLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 40,
    width: 2,
    dash: 'sysDash',
  }) &&
  JSON.stringify(packedSchemeLine.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
    width: 1,
    dash: 'solid',
  }) &&
  packedNoneLine.line?.kind === 'none' &&
  packedZeroWidthLine.line?.kind === 'line' &&
  packedZeroWidthLine.line.width === 0 &&
  packedZeroWidthLine.line.dash === 'sysDot' &&
  packedDashLines.every((shape, index) => shape.line?.kind === 'line' &&
    shape.line.dash === packedLineDashes[index]) &&
  duplicatePackedLine.line?.kind === 'none' &&
  JSON.stringify(reopenedSourceLines[0]) === JSON.stringify(packedSrgbLine.line) &&
  JSON.stringify(reopenedDuplicateLines[0]) === JSON.stringify({ kind: 'none' }) &&
  reopenedSourceLines.slice(4).every((line, index) =>
    line?.kind === 'line' && line.dash === packedLineDashes[index]);
const shapeArrowDeck = PptxDocument.create();
const shapeArrowSlide = shapeArrowDeck.addSlide();
const shapeArrowSource = { begin: 'triangle', end: 'arrow' };
const packedBothArrows = shapeArrowSlide.addShape('line', {
  name: 'Packed both arrows',
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '112233' },
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: shapeArrowSource,
});
const packedArrowOnly = shapeArrowSlide.addShape('lineInv', {
  name: 'Packed arrow only',
  arrows: { begin: 'diamond' },
});
const packedArrowTypes = ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'];
const packedTypedArrows = packedArrowTypes.map((type, index) =>
  shapeArrowSlide.addShape(index % 2 === 0 ? 'line' : 'lineInv', {
    name: 'Packed arrow ' + type,
    arrows: index % 2 === 0 ? { begin: type } : { end: type },
  }));
const initialPackedArrows = packedBothArrows.arrows;
const initialPackedArrowsAgain = packedBothArrows.arrows;
shapeArrowSource.begin = 'none';
shapeArrowSource.end = 'none';
const detachedPackedArrows = packedBothArrows.arrows;
const shapeArrowPartCountBeforeEdit = shapeArrowDeck.opcPackage.parts.length;
const shapeArrowRelationshipCountBeforeEdit = shapeArrowSlide.relationships.length;
const shapeArrowNoOpBytes = shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri).bytes.slice();
const shapeArrowNoOpJournalLength = shapeArrowDeck.opcPackage.mutations.length;
packedBothArrows.arrows = { begin: 'triangle', end: 'arrow' };
const currentShapeArrowNoOpBytes = shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri).bytes;
const shapeArrowNoOp = shapeArrowNoOpJournalLength === shapeArrowDeck.opcPackage.mutations.length &&
  shapeArrowNoOpBytes.length === currentShapeArrowNoOpBytes.length &&
  shapeArrowNoOpBytes.every((value, index) => value === currentShapeArrowNoOpBytes[index]);
const arrowPart = shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri);
const arrowXml = new TextDecoder().decode(arrowPart.bytes);
const arrowXmlWithAdvancedLine = arrowXml.replace(
  '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
    '<a:prstDash val="dashDot"/><a:headEnd type="triangle"/>' +
    '<a:tailEnd type="arrow"/></a:ln>',
  '<a:gradFill><a:gsLst/></a:gradFill>' +
    '<a:custDash><a:ds d="1" sp="1"/></a:custDash><a:round/>' +
    '<a:headEnd type="triangle" w="lg" len="sm"/>' +
    '<a:tailEnd type="arrow" w="med" len="med"/>' +
    '<a:extLst><a:ext uri="urn:packed-arrows">' +
    '<x:keep xmlns:x="urn:packed-arrows"/></a:ext></a:extLst></a:ln>',
);
if (arrowXmlWithAdvancedLine === arrowXml) throw new Error('Packed arrow injection failed');
shapeArrowDeck.opcPackage.setPart(
  shapeArrowSlide.partUri,
  arrowXmlWithAdvancedLine,
  arrowPart.contentType,
);
const packedAdvancedLineBeforeArrowEdit = packedBothArrows.line;
packedBothArrows.arrows = { begin: 'diamond', end: 'oval' };
const shapeArrowSizedEditXml = new TextDecoder().decode(
  shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri).bytes,
);
const shapeArrowSizesPreserved =
  shapeArrowSizedEditXml.includes('<a:headEnd type="diamond" w="lg" len="sm"/>') &&
  shapeArrowSizedEditXml.includes('<a:tailEnd type="oval" w="med" len="med"/>');
packedBothArrows.arrows = { begin: 'stealth' };
const packedPartialArrows = packedBothArrows.arrows;
packedBothArrows.arrows = undefined;
const packedClearedArrows = packedBothArrows.arrows;
const shapeArrowClearedXml = new TextDecoder().decode(
  shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri).bytes,
);
const shapeArrowAdvancedLinePreserved =
  packedAdvancedLineBeforeArrowEdit === undefined &&
  shapeArrowClearedXml.includes('<a:gradFill><a:gsLst/></a:gradFill>') &&
  shapeArrowClearedXml.includes('<a:custDash><a:ds d="1" sp="1"/></a:custDash>') &&
  shapeArrowClearedXml.includes('<a:round/><a:extLst>') &&
  shapeArrowClearedXml.includes('<x:keep xmlns:x="urn:packed-arrows"/>') &&
  !shapeArrowClearedXml.includes('<a:headEnd type="stealth" w="lg" len="sm"') &&
  !shapeArrowClearedXml.includes('<a:tailEnd type="oval" w="med" len="med"');
packedBothArrows.arrows = { end: 'triangle' };
packedBothArrows.line = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent4' },
  transparency: 40,
  width: 2,
  dash: 'sysDash',
};
const shapeArrowLineEditPreservedArrows =
  JSON.stringify(packedBothArrows.arrows) === JSON.stringify({ end: 'triangle' });
packedBothArrows.arrows = undefined;
const packedLineAfterArrowClear = packedBothArrows.line;
packedBothArrows.arrows = { begin: 'none', end: 'arrow' };
const shapeArrowEditIsolation = shapeArrowDeck.opcPackage.parts.length ===
  shapeArrowPartCountBeforeEdit &&
  shapeArrowSlide.relationships.length === shapeArrowRelationshipCountBeforeEdit;
const duplicateShapeArrowSlide = shapeArrowDeck.duplicateSlide(0);
const duplicatePackedArrows = duplicateShapeArrowSlide.shapes[0];
if (!(duplicatePackedArrows instanceof ShapeModel)) {
  throw new Error('Packed duplicate arrow shape failed');
}
duplicatePackedArrows.arrows = { begin: 'diamond' };
duplicatePackedArrows.line = undefined;
const reopenedShapeArrowDeck = await PptxDocument.open(await shapeArrowDeck.write());
const reopenedSourceArrows = reopenedShapeArrowDeck.slides[0].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.arrows : undefined);
const reopenedDuplicateArrows = reopenedShapeArrowDeck.slides[1].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.arrows : undefined);
const reopenedSourceArrowLines = reopenedShapeArrowDeck.slides[0].shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.line : undefined);
const shapeArrowInitialXml = new TextDecoder().decode(
  shapeArrowDeck.opcPackage.requirePart(shapeArrowSlide.partUri).bytes,
);
const shapeArrowChecks = {
  models: packedBothArrows instanceof ShapeModel && packedArrowOnly instanceof ShapeModel,
  detached: initialPackedArrows !== initialPackedArrowsAgain &&
    Object.isFrozen(initialPackedArrows) &&
    JSON.stringify(initialPackedArrows) === JSON.stringify({ begin: 'triangle', end: 'arrow' }) &&
    JSON.stringify(detachedPackedArrows) === JSON.stringify(initialPackedArrows),
  noOp: shapeArrowNoOp,
  sizePreservation: shapeArrowSizesPreserved,
  partial: JSON.stringify(packedPartialArrows) === JSON.stringify({ begin: 'stealth' }),
  clear: packedClearedArrows === undefined,
  advancedLinePreservation: shapeArrowAdvancedLinePreserved,
  lineEditPreservation: shapeArrowLineEditPreservedArrows,
  arrowClearPreservesLine: JSON.stringify(packedLineAfterArrowClear) === JSON.stringify({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 40,
      width: 2,
      dash: 'sysDash',
    }),
  editIsolation: shapeArrowEditIsolation,
  finalSource: JSON.stringify(packedBothArrows.arrows) ===
    JSON.stringify({ begin: 'none', end: 'arrow' }),
  arrowOnly: JSON.stringify(packedArrowOnly.arrows) === JSON.stringify({ begin: 'diamond' }),
  allTypes: packedTypedArrows.every((shape, index) => {
    const expected = index % 2 === 0
      ? { begin: packedArrowTypes[index] }
      : { end: packedArrowTypes[index] };
    return JSON.stringify(shape.arrows) === JSON.stringify(expected);
  }),
  duplicate: duplicatePackedArrows.line === undefined &&
    JSON.stringify(duplicatePackedArrows.arrows) === JSON.stringify({ begin: 'diamond' }),
  reopenedSource: JSON.stringify(reopenedSourceArrows[0]) ===
    JSON.stringify({ begin: 'none', end: 'arrow' }),
  reopenedDuplicate: JSON.stringify(reopenedDuplicateArrows[0]) ===
    JSON.stringify({ begin: 'diamond' }),
  reopenedLine: JSON.stringify(reopenedSourceArrowLines[0]) ===
    JSON.stringify(packedBothArrows.line),
  noImplicitLineDefault: !shapeArrowInitialXml.includes('333333'),
};
const shapeArrows = Object.values(shapeArrowChecks).every((value) => value);
if (!shapeArrows) {
  throw new Error('Packed shape arrows failed: ' + JSON.stringify({
    checks: shapeArrowChecks,
    initial: initialPackedArrows,
    partial: packedPartialArrows,
    source: packedBothArrows.arrows,
    duplicate: duplicatePackedArrows.arrows,
    reopenedSource: reopenedSourceArrows[0],
    reopenedDuplicate: reopenedDuplicateArrows[0],
  }));
}
const initialTextWrap = createdText.textWrap;
const initialTextDirection = createdText.textDirection;
const initialTextFit = createdText.textFit;
createdText.text = 'Updated\\r\\nParagraph';
createdText.textMargins = { top: 4, left: 8 };
createdText.verticalAlignment = 'bottom';
createdText.textWrap = true;
const updatedTextWrap = createdText.textWrap;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
const updatedTextDirection = createdText.textDirection;
createdText.textDirection = undefined;
createdText.textFit = 'resize';
const updatedTextFit = createdText.textFit;
createdText.textFit = 'none';
createdText.textFit = undefined;
const richText = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: 1, spacing: { line: { kind: 'multiple', factor: 1.5 } }, tabStops: [{ position: 1.5, alignment: 'center' }], runs: [{ text: 'Bold', style: { bold: true, fontSize: 18, color: { kind: 'srgb', value: 'ff0000' }, glow: { color: { kind: 'srgb', value: '00ff00' }, opacity: 0.5, size: 8 }, highlight: { kind: 'srgb', value: 'ffff00' }, outline: { color: { kind: 'srgb', value: '0000ff' }, size: 1.5 }, underline: true, strike: true } }, { text: 'Blue', softBreakBefore: true, style: { lang: 'de-DE', color: { kind: 'scheme', value: 'accent1' }, glow: { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2.5 }, highlight: { kind: 'scheme', value: 'accent2' }, outline: { color: { kind: 'scheme', value: 'accent3' }, size: 2 }, underline: { style: 'dbl', color: { kind: 'srgb', value: '00ff00' } }, strike: 'dblStrike' } }] }, { rtl: false, runs: [{ text: 'LTR override' }] }], { lang: 'fr-CA', rtlMode: true });
const marginText = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, marginLeft: false, runs: [{ text: 'Bullet' }] }, { marginLeft: false, runs: [{ text: 'Absent' }] }], { paragraphMarginLeft: 24 });
const initialParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const bulletMarginIsolation = marginText.richText[1].marginLeft === undefined && marginText.richText[1].bullet.indent === 27;
marginText.richText = [{ marginLeft: 6, runs: [{ text: 'Six' }] }, { marginLeft: 0, runs: [{ text: 'Zero' }] }, { marginLeft: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }];
const updatedParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const rightMarginText = created.slides[0].addRichText([{ marginRight: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, runs: [{ text: 'Bullet' }] }, { marginRight: false, runs: [{ text: 'Absent' }] }], { paragraphMarginRight: 24 });
const initialParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const bulletRightMarginCoexistence = rightMarginText.richText[1].marginRight === 24 && rightMarginText.richText[1].bullet.indent === 27;
rightMarginText.richText = [{ marginRight: 6, runs: [{ text: 'Six' }] }, { marginRight: 0, runs: [{ text: 'Zero' }] }, { marginRight: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }, { bullet: true, marginRight: 9, runs: [{ text: 'Bullet' }] }];
const updatedParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const indentText = created.slides[0].addRichText([{ runs: [{ text: 'Default' }] }, { indent: -18, runs: [{ text: 'Hanging' }] }, { indent: false, runs: [{ text: 'Absent' }] }, { bullet: true, indent: false, runs: [{ text: 'Bullet' }] }], { paragraphIndent: 24 });
const initialParagraphIndents = indentText.richText.map(({ indent }) => indent);
const bulletIndentIsolation = indentText.richText[3].indent === undefined && indentText.richText[3].bullet.indent === 27;
indentText.richText = [{ indent: 6, runs: [{ text: 'Positive' }] }, { indent: -6, runs: [{ text: 'Negative' }] }, { indent: 0, runs: [{ text: 'Zero' }] }, { indent: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }];
const updatedParagraphIndents = indentText.richText.map(({ indent }) => indent);
const transparencyText = created.slides[0].addRichText([{ runs: [{ text: 'Quarter', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 25 } }, { text: 'Fractional', style: { transparency: 50.5555 } }, { text: 'Invisible', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 } }, { text: 'Default', style: { transparency: 60 } }] }]);
const initialTransparencies = transparencyText.richText[0].runs.map(({ style }) => style?.transparency);
transparencyText.richText = [{ runs: [{ text: 'Opaque', style: { transparency: 0 } }, { text: 'Mostly', style: { transparency: 75 } }, { text: 'Cleared' }] }];
const updatedTransparencies = transparencyText.richText[0].runs.map(({ style }) => style?.transparency);
const tableSlide = created.slides[0];
const creationColor = { kind: 'srgb', value: '#D9EAF7' };
const creationFill = { kind: 'solid', color: creationColor, transparency: 33.3334 };
const tableCreationFillColor = { kind: 'scheme', value: 'accent4' };
const tableCreationFill = { kind: 'solid', color: tableCreationFillColor, transparency: 40 };
const tableCreationBorderColor = { kind: 'scheme', value: 'accent4' };
const tableCreationBorder = { kind: 'line', color: tableCreationBorderColor, width: 1.5, style: 'dash' };
const creationBorderColor = { kind: 'srgb', value: '#C00000' };
const creationBorder = { kind: 'line', color: creationBorderColor, width: 2, style: 'solid' };
const creationMargin = { top: 4, left: 8 };
const createdTable = tableSlide.addTable([
  [
    { text: 'Region', options: { align: 'left', border: creationBorder, fill: creationFill, fit: 'shrink', margin: creationMargin, textDirection: 'vert', valign: 'top' } },
    { text: 'Revenue', options: {
      border: [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, undefined, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
      fit: 'resize',
      margin: [1, 2, 3, 4],
      textDirection: 'vert270',
      valign: 'middle',
    } },
  ],
  [
    { text: 'East', options: { align: 'right', border: { top: { kind: 'line', color: { kind: 'scheme', value: 'accent3' }, width: 1, style: 'dash' }, left: { kind: 'none' } }, fit: 'none', margin: 0, textDirection: 'wordArtVert', valign: 'bottom' } },
    { text: '', options: { align: 'justify', border: { kind: 'none' }, fill: { kind: 'none' }, margin: {}, textDirection: 'horz' } },
  ],
], { name: 'Created smoke table', align: 'center', columnWidths: [inches(1), inches(3)], rowHeights: [inches(0.5), inches(1.5)], fill: tableCreationFill, margin: { top: 9, left: 18 }, valign: 'middle' });
const tableBorderSlide = created.addSlide();
const createdTableBorderDefault = tableBorderSlide.addTable([[
  'Inherited border',
  { text: 'None override', options: { border: { kind: 'none' } } },
]], { name: 'Created table border default', columnWidths: inches(1), rowHeights: inches(0.5), border: tableCreationBorder });
const createdTableDirectionDefault = tableBorderSlide.addTable([[
  'Inherited direction',
  { text: 'Horizontal override', options: { textDirection: 'horz' } },
]], { name: 'Created table direction default', columnWidths: inches(1), rowHeights: inches(0.5), textDirection: 'vert270' });
const tableCellObjectCreation = JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ text }) => text))) === JSON.stringify([['Region', 'Revenue'], ['East', '']]);
const initialCreatedFill = createdTable.rows[0].cells[0].fill;
const initialTableDefaultFill = createdTable.rows[1].cells[0].fill;
const initialTableNoneOverride = createdTable.rows[1].cells[1].fill;
const initialCreatedBorders = createdTable.rows.map(({ cells }) => cells.map(({ borders }) => borders));
const initialTableDefaultBorders = createdTableBorderDefault.rows[0].cells[0].borders;
const initialTableBorderNoneOverride = createdTableBorderDefault.rows[0].cells[1].borders;
const initialCreatedMargins = createdTable.rows.map(({ cells }) => cells.map(({ margins }) => margins));
const initialCreatedAlignments = createdTable.rows.map(({ cells }) => cells.map(({ verticalAlignment }) => verticalAlignment));
const initialCreatedDirections = createdTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
const initialCreatedFits = createdTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
const initialTableDirectionDefault = createdTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection);
creationColor.value = '000000';
creationFill.transparency = 1;
tableCreationFillColor.value = 'accent6';
tableCreationFill.transparency = 1;
tableCreationBorderColor.value = 'accent6';
tableCreationBorder.width = 9;
creationBorderColor.value = '000000';
creationBorder.width = 9;
creationMargin.top = 99;
creationMargin.left = 99;
const detachedCreatedFill = createdTable.rows[0].cells[0].fill;
const detachedTableDefaultFill = createdTable.rows[1].cells[0].fill;
const detachedCreatedBorders = createdTable.rows[0].cells[0].borders;
const detachedTableDefaultBorders = createdTableBorderDefault.rows[0].cells[0].borders;
const detachedCreatedMargins = createdTable.rows[0].cells[0].margins;
const createdTableDefaults = createdTable instanceof TableModel && createdTable.transform.x === inches(0.5) && createdTable.transform.y === inches(0.5) && createdTable.rows[1].cells[1].margins?.top === 9 && createdTable.rows[1].cells[1].margins?.left === 18;
const createdTableXml = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
const createdTableCells = [...createdTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0]);
const createdTableHorizontalAlignments = createdTableCells.map((cellXml) => cellXml.match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const createdTableDirections = createdTableCells.map((cellXml) => cellXml.match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const createdTableFits = createdTableCells.map((cellXml) => cellXml.match(/<a:(normAutofit|spAutoFit)\\/>/)?.[1]);
const createdTableAnchors = createdTableCells.flatMap((cellXml) => [...cellXml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)].map((match) => match[1]));
const createdTableGrid = [...createdTableXml.matchAll(/<a:gridCol w="(\\d+)"\\/>/g)].map((match) => Number(match[1]));
const createdTableRows = [...createdTableXml.matchAll(/<a:tr h="(\\d+)">/g)].map((match) => Number(match[1]));
const initialTableColumnWidths = createdTable.columnWidths;
const initialTableRowHeights = createdTable.rowHeights;
createdTable.setColumnWidths([inches(1.5), inches(2.5)]);
createdTable.setRowHeights([inches(0.75), inches(1.25)]);
const explicitTableHeight = createdTable.transform.height;
createdTable.setRowHeights([0, inches(1.25)]);
const automaticTableHeightPreserved = createdTable.transform.height === explicitTableHeight;
createdTable.setRowHeights([inches(0.75), inches(1.25)]);
createdTable.setCellText(1, 0, 'Edited East');
createdTable.setCellFill(1, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 50 });
createdTable.setCellBorders(1, 0, { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, width: 1, style: 'solid' });
createdTable.setCellVerticalAlignment(0, 0, 'bottom');
createdTable.setCellVerticalAlignment(1, 1, undefined);
createdTable.setCellMargins(0, 0, { bottom: 9 });
createdTable.setCellMargins(1, 1, undefined);
createdTable.setCellFill(1, 1, undefined);
createdTableBorderDefault.setCellBorders(0, 0, undefined);
const reopenedCreated = await PptxDocument.open(await created.write());
const reopenedCreatedTable = reopenedCreated.slides[0].shapes.find((shape) => shape.name === 'Created smoke table');
const reopenedCreatedTableXml = new TextDecoder().decode(reopenedCreated.opcPackage.requirePart(reopenedCreated.slides[0].partUri).bytes);
const reopenedCreatedTableHorizontalAlignments = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const reopenedCreatedTableDirections = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const reopenedCreatedFits = reopenedCreatedTable instanceof TableModel
  ? reopenedCreatedTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))
  : undefined;
const reopenedTableBorderDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created table border default');
const reopenedTableDirectionDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created table direction default');
const tableCreation = createdTableDefaults && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rows[1].cells[0].text === 'Edited East' && reopenedCreatedTable.rows[1].cells[1].text === '' && reopenedCreatedTable.rows[1].cells[1].verticalAlignment === undefined;
const tableColumnWidths = createdTable.transform.width === inches(4) && createdTableGrid.length === 2 && createdTableGrid[0] === inches(1) && createdTableGrid[1] === inches(3) && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.transform.width === inches(4);
const tableColumnWidthEditing = initialTableColumnWidths?.join(',') === [inches(1), inches(3)].join(',') && createdTable.columnWidths?.join(',') === [inches(1.5), inches(2.5)].join(',') && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.columnWidths?.join(',') === [inches(1.5), inches(2.5)].join(',');
const tableRowHeights = createdTable.transform.height === inches(2) && createdTableRows.length === 2 && createdTableRows[0] === inches(0.5) && createdTableRows[1] === inches(1.5) && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.transform.height === inches(2);
const tableRowHeightEditing = initialTableRowHeights?.join(',') === [inches(0.5), inches(1.5)].join(',') && automaticTableHeightPreserved && createdTable.rowHeights?.join(',') === [inches(0.75), inches(1.25)].join(',') && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rowHeights?.join(',') === [inches(0.75), inches(1.25)].join(',');
const tableCellFillCreation = initialCreatedFill?.kind === 'solid' && initialCreatedFill.color.kind === 'srgb' && initialCreatedFill.color.value === 'D9EAF7' && initialCreatedFill.transparency === 33.333 && detachedCreatedFill?.kind === 'solid' && detachedCreatedFill.color.kind === 'srgb' && detachedCreatedFill.color.value === 'D9EAF7' && detachedCreatedFill.transparency === 33.333 && createdTable.rows[0].cells[1].fill?.kind === 'solid' && createdTable.rows[0].cells[1].fill.color.kind === 'scheme' && createdTable.rows[0].cells[1].fill.color.value === 'accent2' && createdTable.rows[0].cells[1].fill.transparency === 25 && initialTableNoneOverride?.kind === 'none' && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rows[0].cells[0].fill?.kind === 'solid' && reopenedCreatedTable.rows[0].cells[0].fill.color.value === 'D9EAF7' && reopenedCreatedTable.rows[0].cells[0].fill.transparency === 33.333 && reopenedCreatedTable.rows[0].cells[1].fill?.kind === 'solid' && reopenedCreatedTable.rows[0].cells[1].fill.color.kind === 'scheme' && reopenedCreatedTable.rows[0].cells[1].fill.color.value === 'accent2' && reopenedCreatedTable.rows[0].cells[1].fill.transparency === 25;
const tableFillCreation = tableCellFillCreation &&
  initialTableDefaultFill?.kind === 'solid' &&
  initialTableDefaultFill.color.kind === 'scheme' &&
  initialTableDefaultFill.color.value === 'accent4' &&
  initialTableDefaultFill.transparency === 40 &&
  detachedTableDefaultFill?.kind === 'solid' &&
  detachedTableDefaultFill.color.kind === 'scheme' &&
  detachedTableDefaultFill.color.value === 'accent4' &&
  detachedTableDefaultFill.transparency === 40 &&
  createdTableCells[2]?.includes('</a:lnB><a:solidFill><a:schemeClr val="accent4"><a:alpha val="60000"/></a:schemeClr></a:solidFill>') === true &&
  createdTable.rows[1].cells[0].fill?.kind === 'solid' &&
  createdTable.rows[1].cells[0].fill.color.kind === 'scheme' &&
  createdTable.rows[1].cells[0].fill.color.value === 'accent1' &&
  createdTable.rows[1].cells[0].fill.transparency === 50 &&
  createdTable.rows[1].cells[1].fill === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[1].cells[0].fill?.kind === 'solid' &&
  reopenedCreatedTable.rows[1].cells[0].fill.color.kind === 'scheme' &&
  reopenedCreatedTable.rows[1].cells[0].fill.color.value === 'accent1' &&
  reopenedCreatedTable.rows[1].cells[0].fill.transparency === 50 &&
  reopenedCreatedTable.rows[1].cells[1].fill === undefined;
const creationSides = ['top', 'right', 'bottom', 'left'];
const isCreationLine = (border, colorKind, colorValue, width, style) => border?.kind === 'line' && border.color.kind === colorKind && border.color.value === colorValue && border.width === width && border.style === style;
const allCreationLines = (borders, colorKind, colorValue, width, style) => borders !== undefined && creationSides.every((side) => isCreationLine(borders[side], colorKind, colorValue, width, style));
const allCreationNone = (borders) => borders !== undefined && creationSides.every((side) => borders[side]?.kind === 'none');
const initialScalarBorders = initialCreatedBorders[0][0];
const initialTupleBorders = initialCreatedBorders[0][1];
const initialNamedBorders = initialCreatedBorders[1][0];
const initialNoneBorders = initialCreatedBorders[1][1];
const tableCellBorderCreation = tableCellFillCreation &&
  allCreationLines(initialScalarBorders, 'srgb', 'C00000', 2, 'solid') &&
  allCreationLines(detachedCreatedBorders, 'srgb', 'C00000', 2, 'solid') &&
  isCreationLine(initialTupleBorders?.top, 'scheme', 'accent1', 1.5, 'dash') &&
  initialTupleBorders?.right?.kind === 'none' &&
  isCreationLine(initialTupleBorders?.bottom, 'srgb', '00FF00', 0, undefined) &&
  initialTupleBorders?.left?.kind === 'none' &&
  isCreationLine(initialNamedBorders?.top, 'scheme', 'accent3', 1, 'dash') &&
  initialNamedBorders?.right?.kind === 'none' && initialNamedBorders?.bottom?.kind === 'none' && initialNamedBorders?.left?.kind === 'none' &&
  allCreationNone(initialNoneBorders) &&
  allCreationLines(createdTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') &&
  reopenedCreatedTable instanceof TableModel &&
  allCreationLines(reopenedCreatedTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid') &&
  isCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.top, 'scheme', 'accent1', 1.5, 'dash') &&
  reopenedCreatedTable.rows[0].cells[1].borders?.right?.kind === 'none' &&
  isCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.bottom, 'srgb', '00FF00', 0, undefined) &&
  reopenedCreatedTable.rows[0].cells[1].borders?.left?.kind === 'none' &&
  allCreationLines(reopenedCreatedTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') &&
  allCreationNone(reopenedCreatedTable.rows[1].cells[1].borders);
const tableBorderCreation = tableCellBorderCreation &&
  allCreationLines(initialTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  allCreationNone(initialTableBorderNoneOverride) &&
  allCreationLines(detachedTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  createdTableBorderDefault.rows[0].cells[0].borders === undefined &&
  reopenedTableBorderDefault instanceof TableModel &&
  reopenedTableBorderDefault.rows[0].cells[0].borders === undefined &&
  allCreationNone(reopenedTableBorderDefault.rows[0].cells[1].borders);
const marginVector = (margins) => [margins?.top, margins?.right, margins?.bottom, margins?.left];
const tableCellMarginCreation = tableCellBorderCreation &&
  JSON.stringify(initialCreatedMargins.map((row) => row.map(marginVector))) === JSON.stringify([
    [[4, 7.2, 3.6, 8], [1, 2, 3, 4]],
    [[0, 0, 0, 0], [9, 7.2, 3.6, 18]],
  ]) &&
  JSON.stringify(marginVector(detachedCreatedMargins)) === JSON.stringify([4, 7.2, 3.6, 8]) &&
  createdTable.rows[0].cells[0].margins?.top === undefined &&
  createdTable.rows[0].cells[0].margins?.right === undefined &&
  createdTable.rows[0].cells[0].margins?.bottom === 9 &&
  createdTable.rows[0].cells[0].margins?.left === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[0].cells[0].margins?.top === undefined &&
  reopenedCreatedTable.rows[0].cells[0].margins?.right === undefined &&
  reopenedCreatedTable.rows[0].cells[0].margins?.bottom === 9 &&
  reopenedCreatedTable.rows[0].cells[0].margins?.left === undefined;
const tableMarginCreation = tableCellMarginCreation &&
  createdTableXml.includes('<a:tcPr marL="228600" marR="91440" marT="114300" marB="45720" anchor="ctr">') &&
  createdTable.rows[1].cells[1].margins === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[1].cells[1].margins === undefined;
const tableCellVerticalAlignmentCreation = tableCellMarginCreation &&
  JSON.stringify(initialCreatedAlignments) === JSON.stringify([
    ['top', 'middle'],
    ['bottom', 'middle'],
  ]) &&
  createdTable.rows[0].cells[0].verticalAlignment === 'bottom' &&
  createdTable.rows[0].cells[1].verticalAlignment === 'middle' &&
  createdTable.rows[1].cells[0].verticalAlignment === 'bottom' &&
  createdTable.rows[1].cells[1].verticalAlignment === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[0].cells[0].verticalAlignment === 'bottom' &&
  reopenedCreatedTable.rows[0].cells[1].verticalAlignment === 'middle' &&
  reopenedCreatedTable.rows[1].cells[0].verticalAlignment === 'bottom' &&
  reopenedCreatedTable.rows[1].cells[1].verticalAlignment === undefined;
const tableVerticalAlignmentCreation = tableCellVerticalAlignmentCreation &&
  JSON.stringify(createdTableAnchors) === JSON.stringify(['t', 'ctr', 'b', 'ctr']) &&
  createdTableCells.every((cellXml) => !/<a:bodyPr[^>]* anchor=/.test(cellXml));
const tableCellHorizontalAlignmentCreation = tableVerticalAlignmentCreation &&
  JSON.stringify(createdTableHorizontalAlignments) === JSON.stringify(['l', 'ctr', 'r', 'just']) &&
  JSON.stringify(reopenedCreatedTableHorizontalAlignments) === JSON.stringify(['l', 'ctr', 'r', 'just']);
const tableHorizontalAlignmentCreation = tableCellHorizontalAlignmentCreation &&
  createdTableHorizontalAlignments[1] === 'ctr' &&
  reopenedCreatedTableHorizontalAlignments[1] === 'ctr';
const tableCellTextDirectionCreation = tableHorizontalAlignmentCreation &&
  JSON.stringify(initialCreatedDirections) === JSON.stringify([
    ['vert', 'vert270'],
    ['wordArtVert', undefined],
  ]) &&
  JSON.stringify(createdTableDirections) === JSON.stringify([
    'vert',
    'vert270',
    'wordArtVert',
    undefined,
  ]) &&
  reopenedCreatedTable instanceof TableModel &&
  JSON.stringify(reopenedCreatedTable.rows.map(({ cells }) =>
    cells.map(({ textDirection }) => textDirection))) === JSON.stringify([
    ['vert', 'vert270'],
    ['wordArtVert', undefined],
  ]) &&
  JSON.stringify(reopenedCreatedTableDirections) === JSON.stringify([
    'vert',
    'vert270',
    'wordArtVert',
    undefined,
  ]);
const tableCellTextFitCreation = tableCellTextDirectionCreation &&
  JSON.stringify(initialCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]) &&
  JSON.stringify(createdTableFits) === JSON.stringify([
    'normAutofit',
    'spAutoFit',
    undefined,
    undefined,
  ]) &&
  JSON.stringify(reopenedCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]);
const tableTextDirectionCreation = tableCellTextDirectionCreation &&
  JSON.stringify(initialTableDirectionDefault) === JSON.stringify([
    'vert270',
    undefined,
  ]) &&
  reopenedTableDirectionDefault instanceof TableModel &&
  JSON.stringify(reopenedTableDirectionDefault.rows[0].cells.map(
    ({ textDirection }) => textDirection)) === JSON.stringify([
    'vert270',
    undefined,
  ]);
const tablePart = created.opcPackage.requirePart(tableSlide.partUri);
const tableXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Smoke table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:bodyPr custom="TARGET"><a:noAutofit/></a:bodyPr><a:p><a:pPr algn="ctr"/><a:r><a:t>Target</a:t></a:r></a:p></a:txBody><a:tcPr vert="horz" anchor="ctr" marL="12700" marR="25400" marT="38100" marB="50800"><a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="sysDash"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr custom="NEIGHBOR"><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>Neighbor</a:t></a:r></a:p></a:txBody><a:tcPr vert="vert" anchor="b" marL="63500" marR="76200" marT="88900" marB="101600" keep="ADJACENT"><a:lnL w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
created.opcPackage.setPart(tableSlide.partUri, new TextDecoder().decode(tablePart.bytes).replace('</p:spTree>', tableXml + '</p:spTree>'), tablePart.contentType);
const table = tableSlide.shapes.find((shape) => shape.name === 'Smoke table');
const initialCellDirection = table?.rows[0]?.cells[0]?.textDirection;
const initialCellFit = table?.rows[0]?.cells[0]?.textFit;
const initialCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
const initialCellMargins = table?.rows[0]?.cells[0]?.margins;
const initialCellFill = table?.rows[0]?.cells[0]?.fill;
const initialCellBorders = table?.rows[0]?.cells[0]?.borders;
const initialHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
const initialNeighborHorizontalAlignment = table?.rows[0]?.cells[1]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'left');
const leftHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'center');
const centerHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'right');
const rightHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'justify');
const justifyHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, undefined);
const clearedHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellBorders(0, 0, { kind: 'line', color: { kind: 'srgb', value: '#0000FF' }, width: 2, style: 'solid' });
const scalarCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }, undefined]);
const tupleCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, { left: { kind: 'none' } });
const partialCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, undefined);
const clearedCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '#FF0000' } });
const opaqueCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 0 });
const explicitOpaqueCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '112233' }, transparency: 33.333 });
const fractionalCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'none' });
const noneCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, undefined);
const clearedCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellMargins(0, 0, 4);
const uniformCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, [1, 2, 3, 4]);
const tupleCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, { top: 5, left: 7 });
const partialCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, undefined);
const clearedCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellTextFit(0, 0, 'shrink');
const shrinkCellFit = table?.rows[0]?.cells[0]?.textFit;
const beforeSameFit = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
table?.setCellTextFit(0, 0, 'shrink');
const sameFitPreserved = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes) === beforeSameFit;
table?.setCellTextFit(0, 0, 'resize');
const resizeCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'none');
const noneClearedCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'shrink');
table?.setCellTextFit(0, 0, undefined);
const undefinedClearedCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellVerticalAlignment(0, 0, 'top');
const topCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'middle');
const middleCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'bottom');
const bottomCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, undefined);
const clearedCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellTextDirection(0, 0, 'vert270');
const rotatedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'wordArtVert');
const stackedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'horz');
const horizontalCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, undefined);
const clearedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
const reopenedEdited = await PptxDocument.open(await created.write());
const reopenedEditedTable = reopenedEdited.slides[0].shapes.find(
  (shape) => shape.name === 'Smoke table',
);
const reopenedClearedHorizontalAlignment = reopenedEditedTable instanceof TableModel
  ? reopenedEditedTable.rows[0].cells[0].horizontalAlignment
  : null;
const reopenedNeighborHorizontalAlignment = reopenedEditedTable instanceof TableModel
  ? reopenedEditedTable.rows[0].cells[1].horizontalAlignment
  : null;
const inheritedLanguage = richText.richText[0].runs[0].style.lang;
const localLanguage = richText.richText[0].runs[1].style.lang;
const initialRtl = richText.richText.map(({ rtl }) => rtl);
const presentationRtlEnabled = created.rtlMode;
created.rtlMode = false;
const presentationRtlDisabled = created.rtlMode;
created.rtlMode = undefined;
const presentationRtlCleared = created.rtlMode;
const paragraphRtlAfterGlobalClear = richText.richText.map(({ rtl }) => rtl);
const metadata = PptxDocument.create({ title: 'Packed & <Title>' });
const createdPresentationTitle = metadata.title;
metadata.title = 'Edited title';
const editedPresentationTitle = metadata.title;
const reopenedMetadata = await PptxDocument.open(await metadata.write());
const reopenedPresentationTitle = reopenedMetadata.title;
metadata.title = '';
const emptyPresentationTitle = metadata.title;
metadata.title = undefined;
const clearedPresentationTitle = metadata.title;
const authorship = PptxDocument.create({ author: 'Packed & <Author>' });
const createdPresentationAuthor = authorship.author;
authorship.author = 'Edited author';
const editedPresentationAuthor = authorship.author;
const reopenedAuthorship = await PptxDocument.open(await authorship.write());
const reopenedPresentationAuthor = reopenedAuthorship.author;
authorship.author = '';
const emptyPresentationAuthor = authorship.author;
authorship.author = undefined;
const clearedPresentationAuthor = authorship.author;
const editorship = PptxDocument.create({ lastModifiedBy: 'Packed & <Editor>' });
const createdPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = 'Edited editor';
const editedPresentationLastModifiedBy = editorship.lastModifiedBy;
const reopenedEditorship = await PptxDocument.open(await editorship.write());
const reopenedPresentationLastModifiedBy = reopenedEditorship.lastModifiedBy;
editorship.lastModifiedBy = '';
const emptyPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = undefined;
const clearedPresentationLastModifiedBy = editorship.lastModifiedBy;
const chronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
});
const createdPresentationCreatedAt = chronology.createdAt;
const createdPresentationModifiedAt = chronology.modifiedAt;
chronology.createdAt = '2026-07-30T00:00:00Z';
const editedPresentationCreatedAt = chronology.createdAt;
chronology.modifiedAt = '2026-07-30T01:02:03Z';
const editedPresentationModifiedAt = chronology.modifiedAt;
const reopenedChronology = await PptxDocument.open(await chronology.write());
const reopenedPresentationCreatedAt = reopenedChronology.createdAt;
const reopenedPresentationModifiedAt = reopenedChronology.modifiedAt;
chronology.modifiedAt = undefined;
const clearedPresentationModifiedAt = chronology.modifiedAt;
const modifiedAtCreatedIsolation = chronology.createdAt;
chronology.createdAt = undefined;
const clearedPresentationCreatedAt = chronology.createdAt;
const subjectMatter = PptxDocument.create({ subject: 'Packed & <Subject>' });
const createdPresentationSubject = subjectMatter.subject;
subjectMatter.subject = 'Edited subject';
const editedPresentationSubject = subjectMatter.subject;
const reopenedSubjectMatter = await PptxDocument.open(await subjectMatter.write());
const reopenedPresentationSubject = reopenedSubjectMatter.subject;
subjectMatter.subject = '';
const emptyPresentationSubject = subjectMatter.subject;
subjectMatter.subject = undefined;
const clearedPresentationSubject = subjectMatter.subject;
const revisioned = PptxDocument.create({ revision: '007' });
const createdPresentationRevision = revisioned.revision;
revisioned.revision = '42';
const editedPresentationRevision = revisioned.revision;
const reopenedRevisioned = await PptxDocument.open(await revisioned.write());
const reopenedPresentationRevision = reopenedRevisioned.revision;
revisioned.revision = undefined;
const clearedPresentationRevision = revisioned.revision;
const organization = PptxDocument.create({ company: 'Packed & <Company>' });
const createdPresentationCompany = organization.company;
organization.company = 'Edited company';
const editedPresentationCompany = organization.company;
const reopenedOrganization = await PptxDocument.open(await organization.write());
const reopenedPresentationCompany = reopenedOrganization.company;
organization.company = '';
const emptyPresentationCompany = organization.company;
organization.company = undefined;
const clearedPresentationCompany = organization.company;
const themed = PptxDocument.create({
  theme: { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' },
});
const createdTheme = themed.theme;
themed.theme = { headFontFace: 'Noto Sans Display' };
const replacedTheme = themed.theme;
themed.masterLayoutTheme.presentationTheme.setFonts({ minorLatin: 'Noto Sans' });
const reopenedTheme = (await PptxDocument.open(await themed.write())).theme;
const sectioned = PptxDocument.create();
const firstSection = sectioned.addSection({ title: 'Packed & <Intro>' });
const packedSectionEscaped = new TextDecoder().decode(
  sectioned.opcPackage.requirePart(sectioned.presentationPartUri).bytes,
).includes('name="Packed &amp; &lt;Intro&gt;"');
const assignedSlide = sectioned.addSlide({ sectionTitle: 'Packed & <Intro>' });
const automaticSlide = sectioned.addSlide();
const secondAssignedSlide = sectioned.addSlide({ sectionTitle: 'Packed & <Intro>' });
const dataSection = sectioned.addSection({ title: 'Data', order: 1 });
sectioned.assignSlideToSection(1, dataSection.id);
const defaultSection = sectioned.sections.find(({ title }) => title === 'Default-1');
if (!defaultSection) throw new Error('Packed default section was not created');
sectioned.deleteSection(defaultSection.id);
sectioned.renameSection(firstSection.id, 'Edited intro');
sectioned.moveSection(dataSection.id, 0);
const detachedSections = sectioned.sections;
detachedSections[0].title = 'Detached caller title';
detachedSections[0].slideIds.push(999);
const currentSections = sectioned.sections;
const reopenedSections = (await PptxDocument.open(await sectioned.write())).sections;
const hiddenDeck = PptxDocument.create();
const packedVisibleSlide = hiddenDeck.addSlide();
const packedHiddenSlide = hiddenDeck.addSlide();
packedHiddenSlide.hidden = true;
const packedHiddenDuplicate = hiddenDeck.duplicateSlide(1);
packedVisibleSlide.hidden = true;
packedHiddenSlide.hidden = false;
const reopenedHiddenDeck = await PptxDocument.open(await hiddenDeck.write());
const reopenedHiddenStates = reopenedHiddenDeck.slides.map(({ hidden }) => hidden);
const reopenedHiddenRootStates = reopenedHiddenDeck.slides.map(({ partUri }) =>
  /<p:sld\\b[^>]*\\sshow="0"/.test(new TextDecoder().decode(
    reopenedHiddenDeck.opcPackage.requirePart(partUri).bytes,
  )));
const speakerNotesDeck = PptxDocument.create();
const packedLazyNotesSlide = speakerNotesDeck.addSlide();
const packedLazyNotesInitial = packedLazyNotesSlide.notes;
packedLazyNotesSlide.addNotes('Temporary notes');
packedLazyNotesSlide.notes = undefined;
const packedEmptyNotesSlide = speakerNotesDeck.addSlide();
packedEmptyNotesSlide.notes = '';
const packedOriginalNotesSlide = speakerNotesDeck.addSlide().addNotes('Original');
const packedNotesDuplicate = speakerNotesDeck.duplicateSlide(2);
packedOriginalNotesSlide.notes = 'Edited';
const reopenedNotesDeck = await PptxDocument.open(await speakerNotesDeck.write());
const reopenedNotesSnapshots = reopenedNotesDeck.slides.map(({ notes }) => notes);
const reopenedNotesUris = reopenedNotesDeck.slides.flatMap((slide) => {
  const uri = slide.relationships.find(({ type }) => type.endsWith('/notesSlide'))?.resolvedTarget;
  return uri === undefined ? [] : [uri];
});
const reopenedNotesParts = reopenedNotesDeck.opcPackage.parts.filter(
  ({ contentType }) => contentType === 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
);
const reopenedNotesRetargeted = reopenedNotesDeck.slides.every((slide) => {
  const uri = slide.relationships.find(({ type }) => type.endsWith('/notesSlide'))?.resolvedTarget;
  if (uri === undefined) return slide.notes === undefined;
  const backlinks = reopenedNotesDeck.opcPackage.relationships(uri).filter(
    ({ type }) => type.endsWith('/slide'),
  );
  return backlinks.length === 1 && backlinks[0].resolvedTarget === slide.partUri;
});
const reopenedNotesMasterUris = reopenedNotesUris.map((uri) =>
  reopenedNotesDeck.opcPackage.relationships(uri).find(
    ({ type }) => type.endsWith('/notesMaster'),
  )?.resolvedTarget);
richText.richText = [{ align: 'justify', bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 }, level: 3, spacing: { before: 5, after: 7, line: { kind: 'exact', points: 22 } }, tabStops: [{ position: 2.75, alignment: 'decimal' }], runs: [{ text: 'Updated rich', style: { lang: 'ja-JP', baseline: 'superscript', characterSpacing: 2.5, italic: true, glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 }, highlight: { kind: 'srgb', value: '00ff00' }, outline: { color: { kind: 'scheme', value: 'accent1' }, size: 0.75 }, underline: { style: 'wavyHeavy', color: { kind: 'scheme', value: 'accent2' } }, strike: false } }] }];
const custom = PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
custom.slideSize = { width: inches(10), height: inches(7.5) };
const customXml = new TextDecoder().decode(custom.opcPackage.requirePart('/ppt/presentation.xml').bytes);
const slideBackgroundDeck = PptxDocument.create();
const noFillBackgroundSlide = slideBackgroundDeck.addSlide();
noFillBackgroundSlide.background = { kind: 'none' };
const clearedSolidBackgroundSlide = slideBackgroundDeck.addSlide();
clearedSolidBackgroundSlide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
};
const gradientBackgroundSlide = slideBackgroundDeck.addSlide();
gradientBackgroundSlide.background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: 'FF0000' },
    { offset: 1, color: '0000FF', alpha: 0.5 },
  ],
};
const packedBackgroundPngDataUri =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
const imageBackgroundSlide = slideBackgroundDeck.addSlide();
await slideBackgroundDeck.setSlideBackgroundImage(3, packedBackgroundPngDataUri);
const duplicateImageBackgroundSlide = slideBackgroundDeck.duplicateSlide(3);
const backgroundImageTarget = (slide) => slide.relationships.find(
  ({ type, targetMode }) => type.endsWith('/image') && targetMode === 'Internal',
)?.resolvedTarget;
const sharedBackgroundTarget = backgroundImageTarget(imageBackgroundSlide);
const packedBackgroundSharedBeforeReplace = sharedBackgroundTarget !== undefined &&
  backgroundImageTarget(duplicateImageBackgroundSlide) === sharedBackgroundTarget;
duplicateImageBackgroundSlide.background = {
  kind: 'image',
  contentType: 'image/jpeg',
  bytes: Uint8Array.of(255, 216, 255, 224, 0, 16, 74, 70, 73, 70, 255, 217),
};
const duplicateBackgroundTarget = backgroundImageTarget(duplicateImageBackgroundSlide);
const packedBackgroundCloneOnWrite = duplicateBackgroundTarget !== undefined &&
  duplicateBackgroundTarget !== sharedBackgroundTarget &&
  backgroundImageTarget(imageBackgroundSlide) === sharedBackgroundTarget &&
  slideBackgroundDeck.opcPackage.hasPart(sharedBackgroundTarget) &&
  slideBackgroundDeck.opcPackage.hasPart(duplicateBackgroundTarget);
const clearedImageBackgroundSlide = slideBackgroundDeck.addSlide();
await slideBackgroundDeck.setSlideBackgroundImage(5, packedBackgroundPngDataUri);
const clearedBackgroundTarget = backgroundImageTarget(clearedImageBackgroundSlide);
clearedImageBackgroundSlide.background = undefined;
clearedSolidBackgroundSlide.background = undefined;
const packedBackgroundCleanup = clearedBackgroundTarget !== undefined &&
  !slideBackgroundDeck.opcPackage.hasPart(clearedBackgroundTarget) &&
  backgroundImageTarget(clearedImageBackgroundSlide) === undefined;
const reopenedSlideBackgroundDeck = await PptxDocument.open(await slideBackgroundDeck.write());
await reopenedSlideBackgroundDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedBackgroundKinds = reopenedSlideBackgroundDeck.slides.map(
  ({ background }) => background?.kind,
);
const reopenedBackgroundTargets = reopenedSlideBackgroundDeck.slides
  .map(backgroundImageTarget)
  .filter((target) => target !== undefined);
const reopenedBackgroundMediaParts = reopenedSlideBackgroundDeck.opcPackage.parts.filter(
  ({ uri }) => uri.startsWith('/ppt/media/background'),
);
const reopenedBackgroundOrphans = reopenedBackgroundMediaParts.filter(({ uri }) =>
  (reopenedSlideBackgroundDeck.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) === 0);
const slideBackgrounds = packedBackgroundSharedBeforeReplace && packedBackgroundCloneOnWrite &&
  packedBackgroundCleanup &&
  JSON.stringify(reopenedBackgroundKinds) ===
    JSON.stringify(['none', undefined, 'linear-gradient', 'image', 'image', undefined]) &&
  reopenedBackgroundTargets.length === 2 &&
  new Set(reopenedBackgroundTargets).size === 2 &&
  reopenedBackgroundMediaParts.length === 2 &&
  reopenedBackgroundMediaParts.some(({ contentType }) => contentType === 'image/png') &&
  reopenedBackgroundMediaParts.some(({ contentType }) => contentType === 'image/jpeg') &&
  reopenedBackgroundOrphans.length === 0 &&
  reopenedSlideBackgroundDeck.diagnostics.filter(({ severity }) => severity === 'error').length === 0;
await reopenedSlideBackgroundDeck.writeFile('slide-background-smoke.pptx');
const nativeChartDeck = PptxDocument.create();
const nativeChartModels = [];
for (const type of CHART_TYPES) {
  const slide = nativeChartDeck.addSlide();
  const series = type === 'scatter'
    ? [{ name: 'Forecast', xValues: [1, 2, 3], values: [120, 150, 135] }]
    : type === 'bubble'
      ? [{ name: 'Portfolio', xValues: [1, 2, 3], values: [120, 150, 135], sizes: [8, 12, 10] }]
      : [{ name: 'Revenue', categories: ['North', 'South', 'West'], values: [120, 150, 135] }];
  const chart = await slide.addChart(type, series, {
    name: 'Packed ' + type + ' chart',
    altText: type + ' chart with regional business data',
    x: inches(1),
    y: inches(1),
    width: inches(8),
    height: inches(4.5),
  });
  await chart.replaceDefinition({
    groups: chart.definition.groups,
    options: { ...chart.definition.options, title: { text: 'Native ' + type + ' chart' } },
  });
  nativeChartModels.push(chart);
}
const nativeComboSlide = nativeChartDeck.addSlide();
const nativeCombo = await nativeComboSlide.addChart([
  {
    type: 'bar',
    series: [{ name: 'Revenue', categories: ['Q1', 'Q2', 'Q3'], values: [100, 130, 160] }],
  },
  {
    type: 'line',
    axis: 'secondary',
    series: [{ name: 'Margin', categories: ['Q1', 'Q2', 'Q3'], values: [24, 28, 31] }],
  },
], {
  name: 'Packed primary-secondary combination chart',
  x: inches(1),
  y: inches(1),
  width: inches(8),
  height: inches(4.5),
});
await nativeCombo.replaceDefinition({
  groups: nativeCombo.definition.groups,
  options: { ...nativeCombo.definition.options, title: { text: 'Revenue and margin' } },
});
await nativeChartModels[0].replaceSeries([{
  name: 'Revenue edited',
  categories: ['North', 'South', 'West'],
  values: [125, 155, 140],
}]);
await nativeChartModels[1].replaceDefinition({
  groups: [{
    type: 'line',
    series: [{ name: 'Converted', categories: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }],
  }],
  options: { title: { text: 'Bar converted to line' } },
});
const nativeDuplicateSlide = nativeChartDeck.duplicateSlide(nativeChartDeck.slides.length - 1);
const nativeDuplicate = nativeDuplicateSlide.shapes.find((shape) => shape instanceof ChartModel);
const nativeDuplicatePartUri = nativeDuplicate.chartPartUri;
const nativeComboPartUri = nativeCombo.chartPartUri;
nativeDuplicate.remove();
nativeDuplicateSlide.addText('Duplicate chart removed; source remains intact', {
  x: inches(1), y: inches(3), width: inches(8), height: inches(1), align: 'center',
});
const nativeChartBytes = await nativeChartDeck.write();
const reopenedNativeCharts = await PptxDocument.open(nativeChartBytes);
await reopenedNativeCharts.write();
const reopenedNativeChartModels = reopenedNativeCharts.slides.flatMap(({ shapes }) => shapes)
  .filter((shape) => shape instanceof ChartModel);
const nativeChartWorkbookResults = await Promise.all(reopenedNativeChartModels.map((chart) =>
  chartWorkbookMatches(
    reopenedNativeCharts.opcPackage.requirePart(chart.workbookPartUri).bytes,
    chart.definition,
    chart.xml,
  )));
const nativeChartWorkbooksMatch = nativeChartWorkbookResults.every(Boolean);
const nativeChartTypes = new Set(reopenedNativeChartModels.flatMap(({ definition }) =>
  definition.groups.map(({ type }) => type)));
const nativeChartIdsUnique = reopenedNativeCharts.slides.every((slide) => {
  const ids = slide.shapes.map(({ id }) => id);
  return new Set(ids).size === ids.length;
});
const nativeChartOrphans = reopenedNativeCharts.opcPackage.parts
  .filter(({ contentType }) =>
    contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    || contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  .filter(({ uri }) =>
    (reopenedNativeCharts.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) === 0);
const nativeCharts = reopenedNativeChartModels.length === 10
  && CHART_TYPES.every((type) => nativeChartTypes.has(type))
  && nativeChartWorkbooksMatch
  && nativeChartIdsUnique
  && nativeChartOrphans.length === 0
  && !nativeChartDeck.opcPackage.hasPart(nativeDuplicatePartUri)
  && nativeChartDeck.opcPackage.hasPart(nativeComboPartUri)
  && reopenedNativeCharts.diagnostics.filter(({ code }) => code.startsWith('CHART_')).length === 0;
await reopenedNativeCharts.writeFile('native-charts-smoke.pptx');
const masterLayoutDeck = PptxDocument.create({ slideSize: 'wide', firstSlideNumber: 3 });
const defaultMasterLayout = masterLayoutDeck.layouts[0];
masterLayoutDeck.masters[0].background = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'F3F6FA' },
};
const masterLayout = await masterLayoutDeck.defineSlideMaster({
  title: 'PACKED-MASTER-LAYOUT',
  background: { kind: 'image-source', source: packedFallbackPng },
  margin: [inches(0.1), inches(0.2), inches(0.3), inches(0.4)],
  slideNumber: { x: inches(12), y: inches(7), width: inches(0.8), height: inches(0.3) },
  objects: [
    {
      kind: 'rect',
      options: { x: inches(0.25), y: inches(0.2), width: inches(12.8), height: inches(0.08) },
    },
    {
      kind: 'line',
      options: { x: inches(0.5), y: inches(6.8), width: inches(11.8), height: 1 },
    },
    {
      kind: 'text',
      text: 'Packed master layout gallery',
      options: { x: inches(9.4), y: inches(6.85), width: inches(2.4), height: inches(0.3) },
    },
    {
      kind: 'placeholder',
      text: 'Title prompt',
      options: {
        name: 'packed_title', type: 'title', index: 101,
        x: inches(0.5), y: inches(0.4), width: inches(12.3), height: inches(0.6),
      },
    },
    {
      kind: 'placeholder',
      text: 'Body prompt',
      options: {
        name: 'packed_body', type: 'body', index: 102,
        x: inches(0.5), y: inches(1.2), width: inches(4), height: inches(0.8),
      },
    },
    {
      kind: 'placeholder',
      text: 'Picture prompt',
      options: {
        name: 'packed_picture', type: 'pic', index: 103,
        x: inches(0.5), y: inches(2.2), width: inches(3), height: inches(2),
      },
    },
    {
      kind: 'placeholder',
      text: 'Chart prompt',
      options: {
        name: 'packed_chart', type: 'chart', index: 104,
        x: inches(3.75), y: inches(2.2), width: inches(4), height: inches(2),
      },
    },
    {
      kind: 'placeholder',
      text: 'Table prompt',
      options: {
        name: 'packed_table', type: 'tbl', index: 105,
        x: inches(8), y: inches(2.2), width: inches(4.8), height: inches(2),
      },
    },
    {
      kind: 'placeholder',
      text: 'Media prompt',
      options: {
        name: 'packed_media', type: 'media', index: 106,
        x: inches(0.5), y: inches(4.6), width: inches(2.5), height: inches(1.4),
      },
    },
    {
      kind: 'image',
      source: packedFallbackPng,
      options: {
        name: 'Packed layout image object',
        x: inches(3.4), y: inches(4.8), width: inches(1), height: inches(1),
      },
    },
    {
      kind: 'chart',
      groups: [{
        type: 'line',
        series: [{ name: 'Layout trend', categories: ['Q1', 'Q2'], values: [8, 12] }],
      }],
      options: {
        name: 'Packed layout chart object',
        x: inches(4.8), y: inches(4.6), width: inches(3), height: inches(1.5),
      },
    },
  ],
});
const masterLayoutMargin = masterLayout.margin;
const masterLayoutSlide = masterLayoutDeck.addSlide({ masterName: masterLayout.name });
masterLayoutSlide.addText('Packed master/layout support', { placeholder: 'packed_title' });
masterLayoutSlide.addRichText([{
  runs: [
    { text: 'All six ', style: { bold: true } },
    { text: 'placeholder domains are populated.' },
  ],
}], { placeholder: { type: 'body', index: 102 } });
await masterLayoutDeck.addImage(0, packedFallbackPng, {
  placeholder: { type: 'pic', index: 103 },
});
const masterLayoutChart = await masterLayoutDeck.addChart(0, 'bar', [{
  name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
}], { placeholder: 'packed_chart' });
masterLayoutSlide.addTable([
  ['Quarter', 'Revenue'],
  ['Q1', '10'],
  ['Q2', '20'],
], { placeholder: { type: 'tbl', index: 105 } });
await masterLayoutDeck.addAudio(0, Uint8Array.of(1, 2, 3, 4), {
  placeholder: 'packed_media',
  contentType: 'audio/mpeg',
  poster: packedFallbackPng,
  posterContentType: 'image/png',
});
const masterLayoutLiveIdentity = masterLayout instanceof SlideLayoutModel &&
  masterLayoutDeck.masters[0] instanceof SlideMasterModel &&
  masterLayoutDeck.layouts.find(({ partUri }) => partUri === masterLayout.partUri) === masterLayout &&
  masterLayoutDeck.masters[0].layouts.some((layout) => layout === masterLayout);
const masterLayoutSelectedTarget = masterLayoutSlide.relationships.find(
  ({ type }) => type.endsWith('/slideLayout'),
)?.resolvedTarget;
const lifecycleLayout = await masterLayoutDeck.defineSlideMaster({
  title: 'PACKED-LIFECYCLE',
  objects: [{ kind: 'text', text: 'Lifecycle original' }],
});
const lifecycleSlide = masterLayoutDeck.addSlide({ masterName: lifecycleLayout.name });
await masterLayoutDeck.replaceSlideMaster(lifecycleLayout, {
  title: 'PACKED-LIFECYCLE-REPLACED',
  background: { kind: 'solid', color: { kind: 'srgb', value: 'DDEBF7' } },
  objects: [{ kind: 'text', text: 'Lifecycle replacement' }],
});
const masterLayoutReplaced = lifecycleLayout.name === 'PACKED-LIFECYCLE-REPLACED' &&
  lifecycleLayout.background?.kind === 'solid' && lifecycleLayout.shapes.length === 1;
masterLayoutDeck.deleteSlideMaster(lifecycleLayout, defaultMasterLayout);
const masterLayoutRetargeted = lifecycleSlide.relationships.find(
  ({ type }) => type.endsWith('/slideLayout'),
)?.resolvedTarget === defaultMasterLayout.partUri &&
  !masterLayoutDeck.layouts.includes(lifecycleLayout);
const reopenedMasterLayoutDeck = await PptxDocument.open(await masterLayoutDeck.write());
await reopenedMasterLayoutDeck.write({ compatibility: 'powerpoint-2010' });
const reopenedMasterLayout = reopenedMasterLayoutDeck.layouts.find(
  ({ name }) => name === 'PACKED-MASTER-LAYOUT',
);
const reopenedMasterLayoutSlide = reopenedMasterLayoutDeck.slides[0];
const reopenedMasterLayoutChart = reopenedMasterLayoutSlide.shapes.find(
  (shape) => shape instanceof ChartModel,
);
const masterLayoutWorkbookMatches = reopenedMasterLayoutChart instanceof ChartModel &&
  reopenedMasterLayoutChart.workbookPartUri !== undefined &&
  await chartWorkbookMatches(
    reopenedMasterLayoutDeck.opcPackage.requirePart(
      reopenedMasterLayoutChart.workbookPartUri,
    ).bytes,
    reopenedMasterLayoutChart.definition,
    reopenedMasterLayoutChart.xml,
  );
const masterLayoutDependencyOrphans = reopenedMasterLayoutDeck.opcPackage.parts
  .filter(({ contentType }) =>
    contentType.startsWith('image/') || contentType.startsWith('audio/') ||
    contentType.startsWith('video/') ||
    contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' ||
    contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  .filter(({ uri }) =>
    (reopenedMasterLayoutDeck.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) === 0);
const masterLayoutOwnerIdsUnique = [
  ...reopenedMasterLayoutDeck.masters,
  ...reopenedMasterLayoutDeck.layouts,
  ...reopenedMasterLayoutDeck.slides,
].every((owner) => {
  const ids = owner.shapes.map(({ id }) => id);
  return new Set(ids).size === ids.length;
});
const masterLayoutChecks = {
  liveIdentity: masterLayoutLiveIdentity,
  replaced: masterLayoutReplaced,
  retargeted: masterLayoutRetargeted,
  selected: masterLayoutSelectedTarget === masterLayout.partUri,
  types: PLACEHOLDER_TYPES.join(',') === 'title,body,pic,chart,tbl,media',
  margin: masterLayoutMargin?.top === inches(0.1) &&
    masterLayoutMargin.right === inches(0.2) &&
    masterLayoutMargin.bottom === inches(0.3) && masterLayoutMargin.left === inches(0.4),
  liveChart: masterLayoutChart.placeholder?.type === 'chart',
  layoutClass: reopenedMasterLayout instanceof SlideLayoutModel,
  masterClass: reopenedMasterLayoutDeck.masters[0] instanceof SlideMasterModel,
  transientMargin: reopenedMasterLayout.margin === undefined,
  layoutBackground: reopenedMasterLayout.background?.kind === 'image',
  masterBackground: reopenedMasterLayoutDeck.masters[0].background?.kind === 'solid',
  layoutPlaceholders: reopenedMasterLayout.placeholders
    .map(({ placeholder }) => placeholder?.type).join(',') === PLACEHOLDER_TYPES.join(','),
  slidePlaceholders: reopenedMasterLayoutSlide.placeholders
    .map(({ placeholder }) => placeholder?.type).join(',') === PLACEHOLDER_TYPES.join(','),
  slideKinds: reopenedMasterLayoutSlide.shapes.slice(0, 6)
    .map(({ kind }) => kind).join(',') === 'text,text,image,chart,table,audio',
  slideTargets: reopenedMasterLayoutDeck.slides[0].relationships.some(
    ({ type, resolvedTarget }) =>
      type.endsWith('/slideLayout') && resolvedTarget === reopenedMasterLayout.partUri,
  ) && reopenedMasterLayoutDeck.slides[1].relationships.some(
    ({ type, resolvedTarget }) =>
      type.endsWith('/slideLayout') && resolvedTarget === defaultMasterLayout.partUri,
  ),
  chartClass: reopenedMasterLayoutChart instanceof ChartModel,
  chartDefinition: reopenedMasterLayoutChart?.definition?.groups[0]?.type === 'bar',
  chartWorkbook: masterLayoutWorkbookMatches,
  noOrphans: masterLayoutDependencyOrphans.length === 0,
  uniqueIds: masterLayoutOwnerIdsUnique,
  validation: reopenedMasterLayoutDeck.diagnostics
    .filter(({ severity }) => severity === 'error').length === 0,
};
const masterLayouts = Object.values(masterLayoutChecks).every(Boolean);
if (!masterLayouts) throw new Error(JSON.stringify({
  masterLayoutChecks,
  layoutPlaceholderTypes: reopenedMasterLayout.placeholders
    .map(({ placeholder }) => placeholder?.type),
  slidePlaceholderTypes: reopenedMasterLayoutSlide.placeholders
    .map(({ placeholder }) => placeholder?.type),
  slideKinds: reopenedMasterLayoutSlide.shapes.map(({ kind }) => kind),
  dependencyOrphans: masterLayoutDependencyOrphans.map(({ uri }) => uri),
  diagnostics: reopenedMasterLayoutDeck.diagnostics,
}));
await reopenedMasterLayoutDeck.writeFile('master-layout-smoke.pptx');
const checks = {
  slideNumbers,
  slideDefaultColor,
  masterLayouts,
  PptxDocument: typeof PptxDocument === 'function',
  presetShapes,
  customGeometryPaths,
  customGeometryGuideFormulas,
  customGeometryAdjustmentHandles,
  customGeometryConnectionSites,
  customGeometryTextRectangles,
  customGeometryEvaluator,
  shapeAdjustments,
  shapeShadows,
  shapeFills,
  textShapeFills,
  textShapeLines,
  textShapeArrows,
  textShapeShadows,
  textShapeHyperlinks,
  shapeLines,
  shapeArrows,
  shapeHyperlinks,
  embeddedRasterImages,
  svgImages,
  embeddedMedia,
  stableMediaLifecycle,
  nativeMediaTiming,
  nativeCharts,
  slideBackgrounds,
  presentationRtl: presentationRtlEnabled === true && presentationRtlDisabled === false && presentationRtlCleared === undefined && paragraphRtlAfterGlobalClear[0] === true && paragraphRtlAfterGlobalClear[1] === false,
  presentationTitle: createdPresentationTitle === 'Packed & <Title>' && editedPresentationTitle === 'Edited title' && reopenedPresentationTitle === 'Edited title' && emptyPresentationTitle === '' && clearedPresentationTitle === undefined,
  presentationAuthor: createdPresentationAuthor === 'Packed & <Author>' && editedPresentationAuthor === 'Edited author' && reopenedPresentationAuthor === 'Edited author' && emptyPresentationAuthor === '' && clearedPresentationAuthor === undefined,
  presentationLastModifiedBy: createdPresentationLastModifiedBy === 'Packed & <Editor>' && editedPresentationLastModifiedBy === 'Edited editor' && reopenedPresentationLastModifiedBy === 'Edited editor' && emptyPresentationLastModifiedBy === '' && clearedPresentationLastModifiedBy === undefined,
  presentationCreatedAt: createdPresentationCreatedAt === '2024-02-29T12:34:56.123+05:30' && editedPresentationCreatedAt === '2026-07-30T00:00:00Z' && reopenedPresentationCreatedAt === '2026-07-30T00:00:00Z' && clearedPresentationCreatedAt === undefined,
  presentationModifiedAt: createdPresentationModifiedAt === '2024-03-01T01:02:03.456+08:00' && editedPresentationModifiedAt === '2026-07-30T01:02:03Z' && reopenedPresentationModifiedAt === '2026-07-30T01:02:03Z' && clearedPresentationModifiedAt === undefined && modifiedAtCreatedIsolation === '2026-07-30T00:00:00Z',
  presentationSubject: createdPresentationSubject === 'Packed & <Subject>' && editedPresentationSubject === 'Edited subject' && reopenedPresentationSubject === 'Edited subject' && emptyPresentationSubject === '' && clearedPresentationSubject === undefined,
  presentationRevision: createdPresentationRevision === '007' && editedPresentationRevision === '42' && reopenedPresentationRevision === '42' && clearedPresentationRevision === undefined,
  presentationCompany: createdPresentationCompany === 'Packed & <Company>' && editedPresentationCompany === 'Edited company' && reopenedPresentationCompany === 'Edited company' && emptyPresentationCompany === '' && clearedPresentationCompany === undefined,
  presentationThemeFonts: createdTheme?.headFontFace === 'Aptos Display' && createdTheme.bodyFontFace === 'Aptos' && replacedTheme?.headFontFace === 'Noto Sans Display' && replacedTheme.bodyFontFace === 'Calibri' && reopenedTheme?.headFontFace === 'Noto Sans Display' && reopenedTheme.bodyFontFace === 'Noto Sans',
  presentationSections: packedSectionEscaped && assignedSlide.slideId === 256 && automaticSlide.slideId === 257 && secondAssignedSlide.slideId === 258 && /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/.test(firstSection.id) && /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/.test(dataSection.id) && currentSections.length === 2 && currentSections[0].id === dataSection.id && currentSections[0].title === 'Data' && currentSections[0].slideIds.join(',') === String(automaticSlide.slideId) && currentSections[1].id === firstSection.id && currentSections[1].title === 'Edited intro' && currentSections[1].slideIds.join(',') === [assignedSlide.slideId, secondAssignedSlide.slideId].join(',') && reopenedSections?.length === 2 && reopenedSections[0].id === dataSection.id && reopenedSections[0].slideIds.join(',') === String(automaticSlide.slideId) && reopenedSections[1].id === firstSection.id && reopenedSections[1].title === 'Edited intro' && reopenedSections[1].slideIds.join(',') === [assignedSlide.slideId, secondAssignedSlide.slideId].join(','),
  hiddenSlides: packedHiddenDuplicate.hidden === true && reopenedHiddenStates.join(',') === 'true,false,true' && reopenedHiddenRootStates.join(',') === 'true,false,true',
  speakerNotes: packedLazyNotesInitial === undefined && packedNotesDuplicate.notes === 'Original' && JSON.stringify(reopenedNotesSnapshots) === JSON.stringify([undefined, '', 'Edited', 'Original']) && reopenedNotesParts.length === 3 && reopenedNotesUris.length === 3 && new Set(reopenedNotesUris).size === 3 && reopenedNotesRetargeted && reopenedNotesMasterUris.every((uri) => uri !== undefined) && new Set(reopenedNotesMasterUris).size === 1,
  paragraphMarginLeft: initialParagraphMargins[0] === 12 && initialParagraphMargins[1] === undefined && initialParagraphMargins[2] === undefined && bulletMarginIsolation && updatedParagraphMargins[0] === 6 && updatedParagraphMargins[1] === 0 && updatedParagraphMargins[2] === undefined && updatedParagraphMargins[3] === undefined,
  paragraphMarginRight: initialParagraphRightMargins[0] === 12 && initialParagraphRightMargins[1] === 24 && initialParagraphRightMargins[2] === undefined && bulletRightMarginCoexistence && updatedParagraphRightMargins[0] === 6 && updatedParagraphRightMargins[1] === 0 && updatedParagraphRightMargins[2] === undefined && updatedParagraphRightMargins[3] === undefined && updatedParagraphRightMargins[4] === 9,
  paragraphIndent: initialParagraphIndents[0] === 24 && initialParagraphIndents[1] === -18 && initialParagraphIndents[2] === undefined && initialParagraphIndents[3] === undefined && bulletIndentIsolation && updatedParagraphIndents[0] === 6 && updatedParagraphIndents[1] === -6 && updatedParagraphIndents[2] === 0 && updatedParagraphIndents[3] === undefined && updatedParagraphIndents[4] === undefined,
  richTextTransparency: initialTransparencies[0] === 25 && initialTransparencies[1] === 50.555 && initialTransparencies[2] === 100 && initialTransparencies[3] === 60 && updatedTransparencies[0] === 0 && updatedTransparencies[1] === 75 && updatedTransparencies[2] === undefined,
  tableCreation,
  tableCellObjectCreation,
  tableCellFillCreation,
  tableFillCreation,
  tableCellBorderCreation,
  tableBorderCreation,
  tableCellMarginCreation,
  tableMarginCreation,
  tableCellVerticalAlignmentCreation,
  tableVerticalAlignmentCreation,
  tableCellHorizontalAlignmentCreation,
  tableHorizontalAlignmentCreation,
  tableCellTextDirectionCreation,
  tableTextDirectionCreation,
  tableColumnWidths,
  tableColumnWidthEditing,
  tableRowHeights,
  tableRowHeightEditing,
  tableCellTextDirection: table instanceof TableModel && initialCellDirection === 'horz' && rotatedCellDirection === 'vert270' && stackedCellDirection === 'wordArtVert' && horizontalCellDirection === 'horz' && clearedCellDirection === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].textDirection === 'vert' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellTextFitCreation,
  tableCellTextFit: table instanceof TableModel && initialCellFit === 'none' && shrinkCellFit === 'shrink' && sameFitPreserved && resizeCellFit === 'resize' && noneClearedCellFit === undefined && undefinedClearedCellFit === undefined && table.rows[0].cells[0].textFit === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].textFit === 'resize' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellVerticalAlignment: table instanceof TableModel && initialCellAlignment === 'middle' && topCellAlignment === 'top' && middleCellAlignment === 'middle' && bottomCellAlignment === 'bottom' && clearedCellAlignment === undefined && table.rows[0].cells[0].verticalAlignment === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].verticalAlignment === 'bottom' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellHorizontalAlignmentEditing: table instanceof TableModel && initialHorizontalAlignment === 'center' && initialNeighborHorizontalAlignment === undefined && leftHorizontalAlignment === 'left' && centerHorizontalAlignment === 'center' && rightHorizontalAlignment === 'right' && justifyHorizontalAlignment === 'justify' && clearedHorizontalAlignment === undefined && table.rows[0].cells[0].horizontalAlignment === undefined && table.rows[0].cells[1].horizontalAlignment === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].text === 'Neighbor' && reopenedClearedHorizontalAlignment === undefined && reopenedNeighborHorizontalAlignment === undefined,
  tableCellMargins: table instanceof TableModel && initialCellMargins?.top === 3 && initialCellMargins?.right === 2 && initialCellMargins?.bottom === 4 && initialCellMargins?.left === 1 && uniformCellMargins?.top === 4 && uniformCellMargins?.right === 4 && uniformCellMargins?.bottom === 4 && uniformCellMargins?.left === 4 && tupleCellMargins?.top === 1 && tupleCellMargins?.right === 2 && tupleCellMargins?.bottom === 3 && tupleCellMargins?.left === 4 && partialCellMargins?.top === 5 && partialCellMargins?.right === undefined && partialCellMargins?.bottom === undefined && partialCellMargins?.left === 7 && clearedCellMargins === undefined && table.rows[0].cells[0].margins === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].margins?.top === 7 && table.rows[0].cells[1].margins?.right === 6 && table.rows[0].cells[1].margins?.bottom === 8 && table.rows[0].cells[1].margins?.left === 5 && table.rows[0].cells[1].text === 'Neighbor',
  tableCellBorders: table instanceof TableModel && initialCellBorders?.top?.kind === 'line' && initialCellBorders.top.color.kind === 'scheme' && initialCellBorders.top.color.value === 'accent2' && initialCellBorders.top.width === 1.5 && initialCellBorders.top.style === 'dash' && initialCellBorders.right?.kind === 'none' && initialCellBorders.bottom?.kind === 'none' && initialCellBorders.left?.kind === 'line' && initialCellBorders.left.color.kind === 'srgb' && initialCellBorders.left.color.value === 'FF0000' && initialCellBorders.left.width === 1 && initialCellBorders.left.style === 'solid' && scalarCellBorders?.top?.kind === 'line' && scalarCellBorders.top.color.kind === 'srgb' && scalarCellBorders.top.color.value === '0000FF' && scalarCellBorders.top.width === 2 && scalarCellBorders.top.style === 'solid' && scalarCellBorders.right?.kind === 'line' && scalarCellBorders.bottom?.kind === 'line' && scalarCellBorders.left?.kind === 'line' && tupleCellBorders?.top?.kind === 'line' && tupleCellBorders.top.color.kind === 'scheme' && tupleCellBorders.top.color.value === 'accent1' && tupleCellBorders.top.width === 1.5 && tupleCellBorders.top.style === 'dash' && tupleCellBorders.right?.kind === 'line' && tupleCellBorders.right.color.kind === 'srgb' && tupleCellBorders.right.color.value === '00FF00' && tupleCellBorders.right.width === 0 && tupleCellBorders.right.style === undefined && tupleCellBorders.bottom?.kind === 'none' && tupleCellBorders.left === undefined && partialCellBorders?.left?.kind === 'none' && partialCellBorders.top === undefined && partialCellBorders.right === undefined && partialCellBorders.bottom === undefined && clearedCellBorders === undefined && table.rows[0].cells[0].borders === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].borders?.left?.kind === 'line' && table.rows[0].cells[1].borders.left.color.kind === 'srgb' && table.rows[0].cells[1].borders.left.color.value === '333333' && table.rows[0].cells[1].borders.left.width === 2 && table.rows[0].cells[1].borders.left.style === 'solid' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellFill: table instanceof TableModel && initialCellFill?.kind === 'solid' && initialCellFill.color.kind === 'scheme' && initialCellFill.color.value === 'accent1' && initialCellFill.transparency === 25 && opaqueCellFill?.kind === 'solid' && opaqueCellFill.color.kind === 'srgb' && opaqueCellFill.color.value === 'FF0000' && opaqueCellFill.transparency === undefined && explicitOpaqueCellFill?.kind === 'solid' && explicitOpaqueCellFill.color.kind === 'scheme' && explicitOpaqueCellFill.color.value === 'accent2' && explicitOpaqueCellFill.transparency === 0 && fractionalCellFill?.kind === 'solid' && fractionalCellFill.color.kind === 'srgb' && fractionalCellFill.color.value === '112233' && fractionalCellFill.transparency === 33.333 && noneCellFill?.kind === 'none' && clearedCellFill === undefined && table.rows[0].cells[0].fill === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].fill?.kind === 'solid' && table.rows[0].cells[1].fill.color.kind === 'srgb' && table.rows[0].cells[1].fill.color.value === '70AD47' && table.rows[0].cells[1].fill.transparency === 50 && table.rows[0].cells[1].text === 'Neighbor',
  createText: createdText.text === 'Updated\\nParagraph' && initialTextWrap === false && updatedTextWrap === true && createdText.textWrap === undefined && initialTextDirection === 'vert270' && updatedTextDirection === 'wordArtVert' && createdText.textDirection === undefined && initialTextFit === 'shrink' && updatedTextFit === 'resize' && createdText.textFit === undefined && createdText.verticalAlignment === 'bottom' && createdText.textMargins.top === 4 && createdText.textMargins.left === 8 && createdText.textMargins.right === undefined && createdText.richText.every(({ align, bullet, level, rtl, spacing, tabStops }) => align === 'center' && bullet?.kind === 'bullet' && bullet.indent === 27 && level === 2 && rtl === true && spacing?.line?.kind === 'exact' && Array.isArray(tabStops) && tabStops[0]?.position === 1.25 && tabStops[1]?.alignment === 'right') && created.slides[0].shapes[0] === createdText,
  richText: inheritedLanguage === 'fr-CA' && localLanguage === 'de-DE' && initialRtl[0] === true && initialRtl[1] === false && richText.text === 'Updated rich' && richText.richText[0].rtl === undefined && richText.richText[0].align === 'justify' && richText.richText[0].bullet.style === 'romanUcPeriod' && richText.richText[0].level === 3 && richText.richText[0].spacing.line.kind === 'exact' && Array.isArray(richText.richText[0].tabStops) && richText.richText[0].tabStops[0].alignment === 'decimal' && richText.richText[0].runs[0].style.lang === 'ja-JP' && richText.richText[0].runs[0].style.baseline === 'superscript' && richText.richText[0].runs[0].style.characterSpacing === 2.5 && richText.richText[0].runs[0].style.italic === true && richText.richText[0].runs[0].style.glow.color.value === 'accent3' && richText.richText[0].runs[0].style.glow.opacity === 0.25 && richText.richText[0].runs[0].style.glow.size === 6 && richText.richText[0].runs[0].style.highlight.value === '00FF00' && richText.richText[0].runs[0].style.outline.color.value === 'accent1' && richText.richText[0].runs[0].style.outline.size === 0.75 && richText.richText[0].runs[0].style.underline.style === 'wavyHeavy' && richText.richText[0].runs[0].style.underline.color.value === 'accent2' && richText.richText[0].runs[0].style.strike === false,
  customSlideSize: custom.slideSize.width === inches(10) && customXml.includes('<p:sldSz cx="9144000" cy="6858000"/>'),
  GradientCodec: typeof GradientCodec === 'function',
  importPptxGenJS: typeof importPptxGenJS === 'function',
  transitions: typeof transitions.TransitionCodec === 'function',
  animations: typeof animations.AnimationTimingCodec === 'function',
  advancedCharts: typeof advancedCharts.AdvancedChartCodec === 'function',
  smartArt: typeof smartArt.SmartArtDiagramCodec === 'function',
};
if (Object.values(checks).some((value) => !value)) throw new Error(JSON.stringify(checks));
process.stdout.write(JSON.stringify(checks));
`,
  );
  const apiResult = run(process.execPath, ['smoke.mjs'], directory);
  const apiChecks = JSON.parse(apiResult.stdout);

  await writeFile(
    join(directory, 'browser-smoke.mjs'),
    `import { CustomGeometryEvaluationError, evaluateCustomGeometry, ImageModel, inches, PRESET_SHAPE_TYPES, PptxDocument, ShapeModel, TableModel, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const resolved = import.meta.resolve('@jiayunxie/pptx');
if (!resolved.endsWith('/dist/browser.js')) throw new Error('Browser condition resolved to ' + resolved);
const checks = [PptxDocument, transitions.TransitionCodec, animations.AnimationTimingCodec, advancedCharts.AdvancedChartCodec, smartArt.SmartArtDiagramCodec];
if (checks.some((value) => typeof value !== 'function')) throw new Error('Browser API surface is incomplete');
if (PRESET_SHAPE_TYPES.length !== 178 || !Object.isFrozen(PRESET_SHAPE_TYPES)) {
  throw new Error('Browser preset catalog failed');
}
const browserRasterDeck = PptxDocument.create();
const browserSlideNumberDeck = PptxDocument.create({ firstSlideNumber: -2 });
const browserNumberSource = browserSlideNumberDeck.addSlide();
browserNumberSource.slideNumber = {
  align: 'center',
  rtl: true,
  style: { italic: true, color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
};
browserSlideNumberDeck.layouts[0].slideNumber = { x: 200 };
browserSlideNumberDeck.masters[0].slideNumber = { x: 300 };
const browserNumberDuplicate = browserSlideNumberDeck.duplicateSlide(0);
browserSlideNumberDeck.moveSlide(browserSlideNumberDeck.slides.indexOf(browserNumberDuplicate), 0);
const reopenedBrowserSlideNumbers = await PptxDocument.open(await browserSlideNumberDeck.writeBlob());
await reopenedBrowserSlideNumbers.write({ compatibility: 'powerpoint-current' });
const browserNumberCache = (partUri) => {
  const xml = new TextDecoder().decode(reopenedBrowserSlideNumbers.opcPackage.requirePart(partUri).bytes);
  const fieldStart = xml.indexOf('type="slidenum"');
  const textStart = xml.indexOf('<a:t>', fieldStart);
  const textEnd = xml.indexOf('</a:t>', textStart);
  return fieldStart < 0 || textStart < 0 || textEnd < 0
    ? undefined
    : xml.slice(textStart + 5, textEnd);
};
if (reopenedBrowserSlideNumbers.firstSlideNumber !== -2 ||
    reopenedBrowserSlideNumbers.slides.length !== 2 ||
    reopenedBrowserSlideNumbers.slides.some(({ slideNumber }) =>
      slideNumber?.align !== 'center' || slideNumber.rtl !== true ||
      slideNumber.style.italic !== true || slideNumber.style.transparency !== 25) ||
    reopenedBrowserSlideNumbers.slides.map(({ partUri }) => browserNumberCache(partUri)).join(',') !== '-2,-1' ||
    reopenedBrowserSlideNumbers.layouts[0].slideNumber?.x !== 200 ||
    reopenedBrowserSlideNumbers.masters[0].slideNumber?.x !== 300 ||
    reopenedBrowserSlideNumbers.diagnostics.some(({ code }) => code.startsWith('SLIDE_NUMBER_'))) {
  throw new Error('Browser slide-number round trip failed');
}
const browserRasterBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const browserRasterSlide = browserRasterDeck.addSlide();
const browserRasterImage = browserRasterSlide.addImage(browserRasterBytes, {
  contentType: 'image/png',
  name: 'Browser PNG',
  width: inches(2),
  height: inches(1),
});
if (!(browserRasterImage instanceof ImageModel) ||
    browserRasterSlide.shapes[0] !== browserRasterImage ||
    browserRasterDeck.opcPackage.requirePart(browserRasterImage.sourcePartUri).contentType !==
      'image/png') {
  throw new Error('Browser embedded raster image creation failed');
}
browserRasterBytes.fill(0);
const browserRasterBlob = await browserRasterDeck.writeBlob();
if (!(browserRasterBlob instanceof Blob)) throw new Error('Browser raster writeBlob failed');
const reopenedBrowserRasterDeck = await PptxDocument.open(browserRasterBlob);
const reopenedBrowserRasterImage = reopenedBrowserRasterDeck.slides[0].shapes[0];
if (!(reopenedBrowserRasterImage instanceof ImageModel) ||
    reopenedBrowserRasterImage.name !== 'Browser PNG' ||
    reopenedBrowserRasterImage.transform.width !== inches(2) ||
    reopenedBrowserRasterDeck.opcPackage
      .requirePart(reopenedBrowserRasterImage.sourcePartUri).bytes[0] !== 137) {
  throw new Error('Browser embedded raster image round trip failed');
}
const browserSvgDeck = PptxDocument.create();
browserSvgDeck.addSlide();
const browserSvgBytes = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"/>',
);
const browserFallbackPng = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
  0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
  39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);
const browserSvgImage = await browserSvgDeck.addImage(0, browserSvgBytes, {
  fallback: browserFallbackPng,
  sizing: { type: 'cover', width: inches(4), height: inches(3) },
  name: 'Browser bundle SVG',
});
const reopenedBrowserSvgDeck = await PptxDocument.open(await browserSvgDeck.writeBlob());
const reopenedBrowserSvgImage = reopenedBrowserSvgDeck.slides[0].shapes[0];
if (!(reopenedBrowserSvgImage instanceof ImageModel) || !reopenedBrowserSvgImage.isSvg ||
    reopenedBrowserSvgImage.name !== 'Browser bundle SVG' ||
    reopenedBrowserSvgImage.sourceRectangle?.left !== 12.5 ||
    reopenedBrowserSvgDeck.opcPackage
      .requirePart(reopenedBrowserSvgImage.fallbackPartUri).contentType !== 'image/png' ||
    reopenedBrowserSvgDeck.opcPackage
      .requirePart(reopenedBrowserSvgImage.svgPartUri).contentType !== 'image/svg+xml' ||
    browserSvgImage !== browserSvgDeck.slides[0].shapes[0]) {
  throw new Error('Browser SVG image round trip failed');
}
const browserShapeDeck = PptxDocument.create();
const browserShape = browserShapeDeck.addSlide().addShape('foldedCorner');
browserShape.presetType = 'star5';
const browserCustomGeometryDeck = PptxDocument.create();
const browserFormulaGeometry = {
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
    { name: 'a1', formula: { operator: 'at2', operands: ['h', 'x1'] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 0 } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 'a1',
        sweepAngle: 'cd2',
      },
      { kind: 'close' },
    ],
  }],
};
const browserFormulaReplacement = {
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [50_000] } }],
  guides: [{ name: 'x1', formula: { operator: 'pin', operands: [0, 'adj1', 100_000] } }],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 't' } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 0,
        sweepAngle: 'cd4',
      },
    ],
  }],
};
const browserCustomGeometryShape = browserCustomGeometryDeck.addSlide()
  .addCustomShape(browserFormulaGeometry);
const initialBrowserFormulaGeometry = browserCustomGeometryShape.customGeometry;
if (!Object.isFrozen(initialBrowserFormulaGeometry) ||
    !Object.isFrozen(initialBrowserFormulaGeometry?.adjustments) ||
    !Object.isFrozen(initialBrowserFormulaGeometry?.adjustments?.[0]?.formula.operands) ||
    !Object.isFrozen(initialBrowserFormulaGeometry?.guides) ||
    !Object.isFrozen(initialBrowserFormulaGeometry?.guides?.[0]?.formula.operands) ||
    !Object.isFrozen(initialBrowserFormulaGeometry?.paths[0]?.commands) ||
    JSON.stringify(initialBrowserFormulaGeometry) !== JSON.stringify(browserFormulaGeometry)) {
  throw new Error('Browser custom geometry guide formula snapshot failed');
}
browserCustomGeometryShape.customGeometry = browserFormulaReplacement;
browserCustomGeometryShape.presetType = 'ellipse';
browserCustomGeometryShape.customGeometry = browserFormulaReplacement;
const browserEvaluatorContext = { width: inches(1), height: inches(1) };
const browserPureEvaluatorResult = evaluateCustomGeometry(
  browserFormulaReplacement,
  browserEvaluatorContext,
);
const browserLiveEvaluatorResult = browserCustomGeometryShape.evaluateCustomGeometry();
let browserTypedEvaluatorFailure = false;
try {
  evaluateCustomGeometry({
    paths: [{
      width: 1,
      height: 1,
      commands: [{ kind: 'moveTo', point: { x: 'missing', y: 0 } }],
    }],
  }, { width: 1, height: 1 });
} catch (error) {
  browserTypedEvaluatorFailure = error instanceof CustomGeometryEvaluationError &&
    error.code === 'unknown-token' && error.token === 'missing';
}
const reopenedBrowserCustomGeometryShape = (await PptxDocument.open(
  await browserCustomGeometryDeck.writeBlob(),
)).slides[0].shapes[0];
const reopenedBrowserEvaluatorResult = reopenedBrowserCustomGeometryShape instanceof ShapeModel
  ? reopenedBrowserCustomGeometryShape.evaluateCustomGeometry()
  : undefined;
if (!(reopenedBrowserCustomGeometryShape instanceof ShapeModel) ||
    reopenedBrowserCustomGeometryShape.presetType !== undefined ||
    !Object.isFrozen(reopenedBrowserCustomGeometryShape.customGeometry?.guides?.[0]?.formula.operands) ||
    JSON.stringify(reopenedBrowserCustomGeometryShape.customGeometry) !==
      JSON.stringify(browserFormulaReplacement) ||
    reopenedBrowserCustomGeometryShape.customGeometry?.paths[0]?.commands[1]?.kind !== 'arcTo' ||
    reopenedBrowserCustomGeometryShape.customGeometry.paths[0].commands[1].sweepAngle !== 'cd4' ||
    JSON.stringify(browserLiveEvaluatorResult) !== JSON.stringify(browserPureEvaluatorResult) ||
    JSON.stringify(reopenedBrowserEvaluatorResult) !== JSON.stringify(browserPureEvaluatorResult) ||
    browserPureEvaluatorResult.guides?.[0]?.value !== 50_000 ||
    browserPureEvaluatorResult.textRectangle.right !== inches(1) ||
    browserPureEvaluatorResult.textRectangle.bottom !== inches(1) ||
    !Object.isFrozen(browserPureEvaluatorResult) ||
    !Object.isFrozen(browserPureEvaluatorResult.context) ||
    !Object.isFrozen(browserPureEvaluatorResult.paths[0]?.commands[1]) ||
    !browserTypedEvaluatorFailure) {
  throw new Error('Browser custom geometry lifecycle failed');
}
const browserHandleSource = {
  adjustments: [
    { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adjY', formula: { operator: 'val', operands: [50_000] } },
    { name: 'adjR', formula: { operator: 'val', operands: [30_000] } },
    { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
  ],
  handles: [
    {
      kind: 'xy',
      position: { x: 'adjX', y: 'adjY' },
      xGuide: 'adjX',
      maxX: 100_000,
    },
    {
      kind: 'polar',
      position: { x: 'hc', y: 'vc' },
      radiusGuide: 'adjR',
      minRadius: 0,
      angleGuide: 'adjAng',
      maxAngle: 'cd',
    },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'adjX', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'adjY' } },
    ],
  }],
};
const browserHandleExpected = structuredClone(browserHandleSource);
const browserHandleReplacement = {
  ...browserHandleExpected,
  handles: [
    {
      kind: 'polar',
      position: { x: 'x1', y: 'y1' },
      maxRadius: 'ss',
      angleGuide: 'adjAng',
      minAngle: 0,
    },
    {
      kind: 'xy',
      position: { x: 'adjX', y: 'adjY' },
      xGuide: 'adjX',
      minX: 0,
      yGuide: 'adjY',
      maxY: 'b',
    },
  ],
};
const browserHandleDeck = PptxDocument.create();
const browserHandleShape = browserHandleDeck.addSlide().addCustomShape(browserHandleSource);
browserHandleSource.handles[0].position.x = 'changed';
browserHandleSource.handles.reverse();
const initialBrowserHandleGeometry = browserHandleShape.customGeometry;
if (!Object.isFrozen(initialBrowserHandleGeometry) ||
    !Object.isFrozen(initialBrowserHandleGeometry?.handles) ||
    !initialBrowserHandleGeometry?.handles?.every((handle) =>
      Object.isFrozen(handle) && Object.isFrozen(handle.position)) ||
    JSON.stringify(initialBrowserHandleGeometry) !== JSON.stringify(browserHandleExpected) ||
    Object.hasOwn(initialBrowserHandleGeometry.handles[0], 'minX') ||
    !Object.hasOwn(initialBrowserHandleGeometry.handles[0], 'maxX') ||
    Object.hasOwn(initialBrowserHandleGeometry.handles[1], 'maxRadius') ||
    !Object.hasOwn(initialBrowserHandleGeometry.handles[1], 'maxAngle')) {
  throw new Error('Browser custom geometry adjustment handle snapshot failed');
}
browserHandleShape.customGeometry = browserHandleReplacement;
const reopenedBrowserHandleShape = (await PptxDocument.open(
  await browserHandleDeck.writeBlob(),
)).slides[0].shapes[0];
const reopenedBrowserHandles = reopenedBrowserHandleShape instanceof ShapeModel
  ? reopenedBrowserHandleShape.customGeometry?.handles
  : undefined;
if (!(reopenedBrowserHandleShape instanceof ShapeModel) ||
    !Object.isFrozen(reopenedBrowserHandles) ||
    !reopenedBrowserHandles?.every((handle) =>
      Object.isFrozen(handle) && Object.isFrozen(handle.position)) ||
    JSON.stringify(reopenedBrowserHandleShape.customGeometry) !==
      JSON.stringify(browserHandleReplacement) ||
    reopenedBrowserHandles[0]?.kind !== 'polar' ||
    reopenedBrowserHandles[1]?.kind !== 'xy' ||
    Object.hasOwn(reopenedBrowserHandles[0], 'radiusGuide') ||
    !Object.hasOwn(reopenedBrowserHandles[0], 'maxRadius') ||
    Object.hasOwn(reopenedBrowserHandles[1], 'maxX') ||
    !Object.hasOwn(reopenedBrowserHandles[1], 'maxY')) {
  throw new Error('Browser custom geometry adjustment handle lifecycle failed');
}
const browserConnectionSource = {
  adjustments: [{ name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } }],
  connectionSites: [
    { position: { x: 'hc', y: 't' }, angle: 0 },
    { position: { x: 'r', y: 'vc' }, angle: 'adjAng' },
    { position: { x: 25_000, y: 100_000 }, angle: -5_400_000 },
    { position: { x: 'hc', y: 't' }, angle: 0 },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'b' } },
    ],
  }],
};
const browserConnectionExpected = structuredClone(browserConnectionSource);
const browserConnectionReplacement = {
  ...browserConnectionExpected,
  connectionSites: [
    { position: { x: 25_000, y: 100_000 }, angle: -5_400_000 },
    { position: { x: 'hc', y: 't' }, angle: 0 },
    { position: { x: 'l', y: 'vc' }, angle: 'adjAng' },
    { position: { x: 'hc', y: 't' }, angle: 0 },
  ],
};
const browserConnectionDeck = PptxDocument.create();
const browserConnectionShape = browserConnectionDeck.addSlide()
  .addCustomShape(browserConnectionSource, { name: 'Browser connection sites' });
browserConnectionSource.connectionSites[0].angle = 1;
browserConnectionSource.connectionSites[0].position.x = 'changed';
browserConnectionSource.connectionSites.reverse();
const initialBrowserConnectionGeometry = browserConnectionShape.customGeometry;
if (!Object.isFrozen(initialBrowserConnectionGeometry) ||
    !Object.isFrozen(initialBrowserConnectionGeometry?.connectionSites) ||
    !initialBrowserConnectionGeometry?.connectionSites?.every((site) =>
      Object.isFrozen(site) && Object.isFrozen(site.position)) ||
    JSON.stringify(initialBrowserConnectionGeometry) !==
      JSON.stringify(browserConnectionExpected)) {
  throw new Error('Browser custom geometry connection site snapshot failed');
}
browserConnectionShape.customGeometry = browserConnectionReplacement;
browserConnectionShape.presetType = 'diamond';
browserConnectionShape.customGeometry = browserConnectionReplacement;
const reopenedBrowserConnectionShape = (await PptxDocument.open(
  await browserConnectionDeck.writeBlob(),
)).slides[0].shapes[0];
const reopenedBrowserConnections = reopenedBrowserConnectionShape instanceof ShapeModel
  ? reopenedBrowserConnectionShape.customGeometry?.connectionSites
  : undefined;
if (!(reopenedBrowserConnectionShape instanceof ShapeModel) ||
    reopenedBrowserConnectionShape.name !== 'Browser connection sites' ||
    reopenedBrowserConnectionShape.presetType !== undefined ||
    !Object.isFrozen(reopenedBrowserConnections) ||
    !reopenedBrowserConnections?.every((site) =>
      Object.isFrozen(site) && Object.isFrozen(site.position)) ||
    JSON.stringify(reopenedBrowserConnectionShape.customGeometry) !==
      JSON.stringify(browserConnectionReplacement) ||
    reopenedBrowserConnections[0]?.angle !== -5_400_000 ||
    reopenedBrowserConnections[2]?.angle !== 'adjAng' ||
    reopenedBrowserConnections[2]?.position.x !== 'l') {
  throw new Error('Browser custom geometry connection site lifecycle failed');
}
const browserTextRectangleSource = {
  guides: [
    { name: 'textLeft', formula: { operator: 'val', operands: [20_000] } },
    { name: 'textRight', formula: { operator: 'val', operands: [80_000] } },
  ],
  connectionSites: [{ position: { x: 'hc', y: 't' }, angle: 0 }],
  textRectangle: {
    left: 'textLeft',
    top: 12_500,
    right: 'textRight',
    bottom: 87_500,
  },
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'b' } },
    ],
  }],
};
const browserTextRectangleExpected = structuredClone(browserTextRectangleSource);
const browserTextRectangleReplacement = {
  ...browserTextRectangleExpected,
  textRectangle: {
    left: 0,
    top: 't',
    right: 90_000,
    bottom: 'b',
  },
};
const browserTextRectangleDeck = PptxDocument.create();
const browserTextRectangleShape = browserTextRectangleDeck.addSlide()
  .addCustomShape(browserTextRectangleSource, { name: 'Browser text rectangle' });
browserTextRectangleSource.textRectangle.left = 'changed';
browserTextRectangleSource.textRectangle.top = 1;
browserTextRectangleSource.textRectangle.right = 2;
browserTextRectangleSource.textRectangle.bottom = 3;
const initialBrowserTextRectangleGeometry = browserTextRectangleShape.customGeometry;
if (!Object.isFrozen(initialBrowserTextRectangleGeometry) ||
    !Object.isFrozen(initialBrowserTextRectangleGeometry?.textRectangle) ||
    JSON.stringify(initialBrowserTextRectangleGeometry) !==
      JSON.stringify(browserTextRectangleExpected)) {
  throw new Error('Browser custom geometry text rectangle snapshot failed');
}
browserTextRectangleShape.customGeometry = browserTextRectangleReplacement;
const { textRectangle: ignoredBrowserTextRectangle, ...browserDefaultTextRectangleGeometry } =
  browserTextRectangleReplacement;
browserTextRectangleShape.customGeometry = browserDefaultTextRectangleGeometry;
if (Object.hasOwn(browserTextRectangleShape.customGeometry, 'textRectangle') ||
    ignoredBrowserTextRectangle === undefined) {
  throw new Error('Browser custom geometry text rectangle reset failed');
}
browserTextRectangleShape.customGeometry = browserTextRectangleReplacement;
browserTextRectangleShape.presetType = 'diamond';
browserTextRectangleShape.customGeometry = browserTextRectangleReplacement;
const reopenedBrowserTextRectangleShape = (await PptxDocument.open(
  await browserTextRectangleDeck.writeBlob(),
)).slides[0].shapes[0];
if (!(reopenedBrowserTextRectangleShape instanceof ShapeModel) ||
    reopenedBrowserTextRectangleShape.name !== 'Browser text rectangle' ||
    reopenedBrowserTextRectangleShape.presetType !== undefined ||
    !Object.isFrozen(reopenedBrowserTextRectangleShape.customGeometry) ||
    !Object.isFrozen(reopenedBrowserTextRectangleShape.customGeometry?.textRectangle) ||
    JSON.stringify(reopenedBrowserTextRectangleShape.customGeometry) !==
      JSON.stringify(browserTextRectangleReplacement)) {
  throw new Error('Browser custom geometry text rectangle lifecycle failed');
}
const browserAdjustmentDeck = PptxDocument.create();
const browserAdjustmentSlide = browserAdjustmentDeck.addSlide();
const browserAdjustedShape = browserAdjustmentSlide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
const browserInitialAdjustments = browserAdjustedShape.adjustments;
const browserInitialAdjustmentsAgain = browserAdjustedShape.adjustments;
if (!Array.isArray(browserInitialAdjustments) ||
    !Object.isFrozen(browserInitialAdjustments) ||
    !browserInitialAdjustments.every((adjustment) => Object.isFrozen(adjustment)) ||
    browserInitialAdjustments === browserInitialAdjustmentsAgain) {
  throw new Error('Browser shape adjustment snapshot immutability failed');
}
const browserAdjustmentNoOpBytes = browserAdjustmentDeck.opcPackage
  .requirePart(browserAdjustmentSlide.partUri).bytes.slice();
const browserAdjustmentNoOpJournal = browserAdjustmentDeck.opcPackage.mutations.length;
browserAdjustedShape.adjustments = [
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
];
const browserAdjustmentNoOpCurrent = browserAdjustmentDeck.opcPackage
  .requirePart(browserAdjustmentSlide.partUri).bytes;
if (browserAdjustmentNoOpJournal !== browserAdjustmentDeck.opcPackage.mutations.length ||
    browserAdjustmentNoOpBytes.length !== browserAdjustmentNoOpCurrent.length ||
    !browserAdjustmentNoOpBytes.every(
      (value, index) => value === browserAdjustmentNoOpCurrent[index],
    )) {
  throw new Error('Browser shape adjustment no-op failed');
}
browserAdjustedShape.adjustments = [
  { name: 'adj1', value: 10_800_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 20_000 },
];
if (JSON.stringify(browserAdjustedShape.adjustments) !== JSON.stringify([
  { name: 'adj1', value: 10_800_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 20_000 },
])) {
  throw new Error('Browser shape adjustment edit failed');
}
browserAdjustedShape.adjustments = [];
if (!Array.isArray(browserAdjustedShape.adjustments) ||
    !Object.isFrozen(browserAdjustedShape.adjustments) ||
    browserAdjustedShape.adjustments.length !== 0) {
  throw new Error('Browser shape adjustment clear failed');
}
browserAdjustedShape.adjustments = browserInitialAdjustments;
const reopenedBrowserAdjustmentDeck = await PptxDocument.open(
  await browserAdjustmentDeck.writeBlob(),
);
const reopenedBrowserAdjustedShape = reopenedBrowserAdjustmentDeck.slides[0].shapes[0];
const reopenedBrowserAdjustments = reopenedBrowserAdjustedShape.adjustments;
if (!(reopenedBrowserAdjustedShape instanceof ShapeModel) ||
    !Array.isArray(reopenedBrowserAdjustments) ||
    !Object.isFrozen(reopenedBrowserAdjustments) ||
    !reopenedBrowserAdjustments.every((adjustment) => Object.isFrozen(adjustment)) ||
    JSON.stringify(reopenedBrowserAdjustments) !==
    JSON.stringify(browserInitialAdjustments)) {
  throw new Error('Browser shape adjustment reopen failed');
}
const browserHyperlinkDeck = PptxDocument.create();
const browserHyperlinkSlide = browserHyperlinkDeck.addSlide();
browserHyperlinkDeck.addSlide();
const browserHyperlinkInput = {
  url: 'https://browser.example',
  tooltip: 'Browser link',
};
const browserUrlHyperlink = browserHyperlinkSlide.addShape('rect', {
  hyperlink: browserHyperlinkInput,
});
const browserInternalHyperlink = browserHyperlinkSlide.addShape('actionButtonForwardNext', {
  hyperlink: { slide: 2 },
});
const browserClearedHyperlink = browserHyperlinkSlide.addShape('ellipse', {
  hyperlink: { url: 'https://clear.browser.example' },
});
const browserInitialHyperlink = browserUrlHyperlink.hyperlink;
browserHyperlinkInput.url = 'https://changed.browser.example';
browserHyperlinkInput.tooltip = 'Changed';
if (!(browserUrlHyperlink instanceof ShapeModel) ||
    !Object.isFrozen(browserInitialHyperlink) ||
    JSON.stringify(browserInitialHyperlink) !== JSON.stringify({
      url: 'https://browser.example',
      tooltip: 'Browser link',
    }) ||
    JSON.stringify(browserUrlHyperlink.hyperlink) !== JSON.stringify(browserInitialHyperlink) ||
    JSON.stringify(browserInternalHyperlink.hyperlink) !== JSON.stringify({ slide: 2 })) {
  throw new Error('Browser shape hyperlink create/read failed');
}
browserUrlHyperlink.hyperlink = { url: 'mailto:browser@example.com', tooltip: '' };
browserInternalHyperlink.hyperlink = { url: 'https://temporary.browser.example' };
browserInternalHyperlink.hyperlink = { slide: 2, tooltip: '' };
browserClearedHyperlink.hyperlink = undefined;
if (browserClearedHyperlink.hyperlink !== undefined) {
  throw new Error('Browser shape hyperlink clear failed');
}
const reopenedBrowserHyperlinkDeck = await PptxDocument.open(
  await browserHyperlinkDeck.writeBlob(),
);
const reopenedBrowserHyperlinks = reopenedBrowserHyperlinkDeck.slides[0]?.shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.hyperlink : undefined);
if (JSON.stringify(reopenedBrowserHyperlinks) !== JSON.stringify([
  { url: 'mailto:browser@example.com', tooltip: '' },
  { slide: 2, tooltip: '' },
  undefined,
])) {
  throw new Error('Browser shape hyperlink edit/reopen failed: ' +
    JSON.stringify(reopenedBrowserHyperlinks));
}
const browserTextHyperlinkDeck = PptxDocument.create();
const browserTextHyperlinkSource = browserTextHyperlinkDeck.addSlide();
browserTextHyperlinkDeck.addSlide();
const browserTextHyperlinkInput = {
  url: 'https://browser-text.example',
  tooltip: 'Browser text',
};
const browserPlainTextHyperlink = browserTextHyperlinkSource.addText(
  'Browser text hyperlink\\nSecond line',
  {
    name: 'browser_plain_text_hyperlink',
    hyperlink: browserTextHyperlinkInput,
  },
);
const browserRichTextHyperlink = browserTextHyperlinkSource.addRichText([{
  runs: [{ text: 'Browser rich one' }, { text: ' and two' }],
}], {
  name: 'browser_rich_text_hyperlink',
  hyperlink: { slide: 2, tooltip: '' },
});
const browserInitialTextHyperlink = browserPlainTextHyperlink.hyperlink;
browserTextHyperlinkInput.url = 'https://changed.browser-text.example';
browserTextHyperlinkInput.tooltip = 'Changed';
const reopenedBrowserTextHyperlinkDeck = await PptxDocument.open(
  await browserTextHyperlinkDeck.writeBlob(),
);
const reopenedBrowserTextHyperlinks = reopenedBrowserTextHyperlinkDeck.slides[0].shapes.map(
  (shape) => shape instanceof ShapeModel ? shape.hyperlink : undefined,
);
if (!(browserPlainTextHyperlink instanceof ShapeModel) ||
    !(browserRichTextHyperlink instanceof ShapeModel) ||
    !Object.isFrozen(browserInitialTextHyperlink) ||
    JSON.stringify(browserInitialTextHyperlink) !== JSON.stringify({
      url: 'https://browser-text.example',
      tooltip: 'Browser text',
    }) ||
    JSON.stringify(browserPlainTextHyperlink.hyperlink) !==
      JSON.stringify(browserInitialTextHyperlink) ||
    JSON.stringify(reopenedBrowserTextHyperlinks) !== JSON.stringify([
      { url: 'https://browser-text.example', tooltip: 'Browser text' },
      { slide: 2, tooltip: '' },
    ]) ||
    reopenedBrowserTextHyperlinkDeck.diagnostics.some(({ severity }) => severity === 'error')) {
  throw new Error('Browser text shape hyperlink lifecycle failed');
}
const browserShadowDeck = PptxDocument.create();
const browserShadowSlide = browserShadowDeck.addSlide();
const browserShadowColor = { kind: 'srgb', value: '#123abc' };
const browserShadowInput = {
  kind: 'outer',
  color: browserShadowColor,
  opacity: 0.42,
  blur: 7.25,
  angle: 123.4,
  distance: 5.5,
  rotateWithShape: true,
};
const browserOuterShadow = browserShadowSlide.addShape('roundRect', {
  shadow: browserShadowInput,
});
const browserInnerShadow = browserShadowSlide.addShape('ellipse', {
  shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent2' } },
});
const browserClearedShadow = browserShadowSlide.addShape('rect', {
  shadow: { kind: 'outer' },
});
const browserInitialShadow = browserOuterShadow.shadow;
const browserInitialShadowAgain = browserOuterShadow.shadow;
browserShadowColor.value = 'FFFFFF';
browserShadowInput.opacity = 1;
browserShadowInput.blur = 20;
browserOuterShadow.shadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent5' },
  opacity: 0,
  blur: 0,
  angle: 0,
  distance: 0,
};
browserInnerShadow.shadow = {
  kind: 'outer',
  color: { kind: 'srgb', value: '445566' },
  opacity: 0.5,
  blur: 2,
  angle: 45,
  distance: 1,
  rotateWithShape: false,
};
browserClearedShadow.shadow = undefined;
const reopenedBrowserShadowDeck = await PptxDocument.open(await browserShadowDeck.writeBlob());
const reopenedBrowserShadows = reopenedBrowserShadowDeck.slides[0]?.shapes.map((shape) =>
  shape instanceof ShapeModel ? shape.shadow : undefined);
const browserShadowChecks = {
  initial: JSON.stringify(browserInitialShadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'srgb', value: '123ABC' },
    opacity: 0.42,
    blur: 7.25,
    angle: 123.4,
    distance: 5.5,
    rotateWithShape: true,
  }),
  detached: browserInitialShadow !== browserInitialShadowAgain &&
    browserInitialShadow?.color !== browserInitialShadowAgain?.color &&
    Object.isFrozen(browserInitialShadow) &&
    Object.isFrozen(browserInitialShadow?.color),
  edited: JSON.stringify(browserOuterShadow.shadow) === JSON.stringify({
    kind: 'inner',
    color: { kind: 'scheme', value: 'accent5' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
  }) && JSON.stringify(browserInnerShadow.shadow) === JSON.stringify({
    kind: 'outer',
    color: { kind: 'srgb', value: '445566' },
    opacity: 0.5,
    blur: 2,
    angle: 45,
    distance: 1,
    rotateWithShape: false,
  }),
  cleared: browserClearedShadow.shadow === undefined,
  reopened: JSON.stringify(reopenedBrowserShadows) === JSON.stringify([
    {
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent5' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    },
    {
      kind: 'outer',
      color: { kind: 'srgb', value: '445566' },
      opacity: 0.5,
      blur: 2,
      angle: 45,
      distance: 1,
      rotateWithShape: false,
    },
    undefined,
  ]),
};
if (Object.values(browserShadowChecks).some((value) => !value)) {
  throw new Error('Browser shape shadow failed: ' + JSON.stringify({
    checks: browserShadowChecks,
    initial: browserInitialShadow,
    reopened: reopenedBrowserShadows,
  }));
}
const browserShapeFillColor = { kind: 'srgb', value: '#224466' };
const browserShapeFillSource = {
  kind: 'solid',
  color: browserShapeFillColor,
  transparency: 12.3456,
};
const browserFilledShape = browserShapeDeck.slides[0].addShape('rect', {
  fill: browserShapeFillSource,
});
const browserInitialShapeFill = browserFilledShape.fill;
browserShapeFillColor.value = 'FFFFFF';
browserShapeFillSource.transparency = 90;
const browserDetachedShapeFill = browserFilledShape.fill;
const browserShapeFillPartCount = browserShapeDeck.opcPackage.parts.length;
const browserShapeFillRelationshipCount = browserShapeDeck.slides[0].relationships.length;
browserFilledShape.fill = { kind: 'none' };
browserFilledShape.fill = undefined;
browserFilledShape.fill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent5' },
  transparency: 50,
};
const reopenedBrowserShape = await PptxDocument.open(await browserShapeDeck.writeBlob());
if (reopenedBrowserShape.slides[0]?.shapes[0]?.presetType !== 'star5') {
  throw new Error('Browser preset shape failed');
}
const reopenedBrowserFill = reopenedBrowserShape.slides[0]?.shapes[1]?.fill;
const browserShapeFillChecks = {
  initial: browserInitialShapeFill?.kind === 'solid' &&
    browserInitialShapeFill.color.kind === 'srgb' &&
    browserInitialShapeFill.color.value === '224466' &&
    Math.abs((browserInitialShapeFill.transparency ?? 0) - 12.346) < 1e-9,
  detached: JSON.stringify(browserDetachedShapeFill) === JSON.stringify(browserInitialShapeFill),
  partIsolation: browserShapeDeck.opcPackage.parts.length === browserShapeFillPartCount,
  relationshipIsolation: browserShapeDeck.slides[0].relationships.length ===
    browserShapeFillRelationshipCount,
  edited: JSON.stringify(browserFilledShape.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent5' },
    transparency: 50,
  }),
  reopened: JSON.stringify(reopenedBrowserFill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent5' },
    transparency: 50,
  }),
};
if (Object.values(browserShapeFillChecks).some((value) => !value)) {
  throw new Error('Browser shape fill failed: ' + JSON.stringify({
    checks: browserShapeFillChecks,
    initial: browserInitialShapeFill,
  }));
}
const browserTextShapeFillDeck = PptxDocument.create();
const browserTextShapeFillLayout = browserTextShapeFillDeck.layouts[0];
const browserTextShapeFillPlaceholder = browserTextShapeFillLayout.addPlaceholder(
  'Browser text fill prompt',
  {
    name: 'browser_text_fill_placeholder',
    type: 'title',
    index: 190,
    fill: {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 100,
    },
  },
);
const browserTextShapeFillSlide = browserTextShapeFillDeck.addSlide({
  masterName: browserTextShapeFillLayout.name,
});
const browserTextShapeFillSource = {
  kind: 'solid',
  color: { kind: 'srgb', value: '#AB12CD' },
  transparency: 25,
};
const browserPlainTextShapeFill = browserTextShapeFillSlide.addText(
  'Browser plain text fill',
  { name: 'browser_plain_text_fill', fill: browserTextShapeFillSource },
);
const browserRichTextShapeFill = browserTextShapeFillSlide.addRichText([{
  runs: [{ text: 'Browser rich text fill' }],
}], {
  name: 'browser_rich_text_fill',
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const browserPopulatedTextShapeFill = browserTextShapeFillSlide.addText(
  'Browser populated text fill',
  { placeholder: 'browser_text_fill_placeholder', fill: { kind: 'none' } },
);
const browserPlainTextShapeFillSnapshot = browserPlainTextShapeFill.fill;
browserTextShapeFillSource.color.value = 'FFFFFF';
browserTextShapeFillSource.transparency = 90;
const reopenedBrowserTextShapeFillDeck = await PptxDocument.open(
  await browserTextShapeFillDeck.writeBlob(),
);
await reopenedBrowserTextShapeFillDeck.write({ compatibility: 'powerpoint-current' });
const reopenedBrowserTextShapeFillSlide = reopenedBrowserTextShapeFillDeck.slides[0];
const reopenedBrowserTextShapeFillLayout = reopenedBrowserTextShapeFillDeck.layouts[0];
const browserTextShapeFillByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const browserTextShapeFillChecks = {
  detached: JSON.stringify(browserPlainTextShapeFill.fill) ===
    JSON.stringify(browserPlainTextShapeFillSnapshot),
  plain: JSON.stringify(browserTextShapeFillByName(
    reopenedBrowserTextShapeFillSlide,
    'browser_plain_text_fill',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'srgb', value: 'AB12CD' },
    transparency: 25,
  }),
  rich: JSON.stringify(browserTextShapeFillByName(
    reopenedBrowserTextShapeFillSlide,
    'browser_rich_text_fill',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  }),
  populated: browserTextShapeFillByName(
    reopenedBrowserTextShapeFillSlide,
    'browser_text_fill_placeholder',
  )?.fill?.kind === 'none',
  placeholder: JSON.stringify(browserTextShapeFillByName(
    reopenedBrowserTextShapeFillLayout,
    'browser_text_fill_placeholder',
  )?.fill) === JSON.stringify({
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 100,
  }),
  live: browserTextShapeFillPlaceholder instanceof ShapeModel &&
    browserRichTextShapeFill instanceof ShapeModel &&
    browserPopulatedTextShapeFill instanceof ShapeModel,
  validation: reopenedBrowserTextShapeFillDeck.diagnostics.every(
    ({ severity }) => severity !== 'error',
  ),
};
if (Object.values(browserTextShapeFillChecks).some((value) => !value)) {
  throw new Error('Browser text shape fill failed: ' + JSON.stringify(browserTextShapeFillChecks));
}
const browserTextShapeLineDeck = PptxDocument.create();
const browserTextShapeLineLayout = browserTextShapeLineDeck.layouts[0];
const browserTextShapeLinePlaceholder = browserTextShapeLineLayout.addPlaceholder(
  'Browser text line prompt',
  {
    name: 'browser_text_line_placeholder',
    type: 'title',
    index: 191,
    line: {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 100,
      width: 0,
      dash: 'sysDot',
    },
  },
);
const browserTextShapeLineSlide = browserTextShapeLineDeck.addSlide({
  masterName: browserTextShapeLineLayout.name,
});
const browserTextShapeLineSource = {
  kind: 'line',
  color: { kind: 'srgb', value: '#AB12CD' },
  transparency: 25,
  width: 2.5,
  dash: 'dashDot',
};
const browserPlainTextShapeLine = browserTextShapeLineSlide.addText(
  'Browser plain text line',
  { name: 'browser_plain_text_line', line: browserTextShapeLineSource },
);
const browserRichTextShapeLine = browserTextShapeLineSlide.addRichText([{
  runs: [{ text: 'Browser rich text line' }],
}], {
  name: 'browser_rich_text_line',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const browserPopulatedTextShapeLine = browserTextShapeLineSlide.addText(
  'Browser populated text line',
  { placeholder: 'browser_text_line_placeholder', line: { kind: 'none' } },
);
const browserPlainTextShapeLineSnapshot = browserPlainTextShapeLine.line;
browserTextShapeLineSource.color.value = 'FFFFFF';
browserTextShapeLineSource.transparency = 90;
browserTextShapeLineSource.width = 9;
browserTextShapeLineSource.dash = 'solid';
const reopenedBrowserTextShapeLineDeck = await PptxDocument.open(
  await browserTextShapeLineDeck.writeBlob(),
);
await reopenedBrowserTextShapeLineDeck.write({ compatibility: 'powerpoint-current' });
const reopenedBrowserTextShapeLineSlide = reopenedBrowserTextShapeLineDeck.slides[0];
const reopenedBrowserTextShapeLineLayout = reopenedBrowserTextShapeLineDeck.layouts[0];
const browserTextShapeLineByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const browserTextShapeLineChecks = {
  detached: JSON.stringify(browserPlainTextShapeLine.line) ===
    JSON.stringify(browserPlainTextShapeLineSnapshot),
  plain: JSON.stringify(browserTextShapeLineByName(
    reopenedBrowserTextShapeLineSlide,
    'browser_plain_text_line',
  )?.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'srgb', value: 'AB12CD' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  }),
  rich: JSON.stringify(browserTextShapeLineByName(
    reopenedBrowserTextShapeLineSlide,
    'browser_rich_text_line',
  )?.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
    width: 1,
    dash: 'solid',
  }),
  populated: browserTextShapeLineByName(
    reopenedBrowserTextShapeLineSlide,
    'browser_text_line_placeholder',
  )?.line?.kind === 'none',
  placeholder: JSON.stringify(browserTextShapeLineByName(
    reopenedBrowserTextShapeLineLayout,
    'browser_text_line_placeholder',
  )?.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 100,
    width: 0,
    dash: 'sysDot',
  }),
  live: browserTextShapeLinePlaceholder instanceof ShapeModel &&
    browserRichTextShapeLine instanceof ShapeModel &&
    browserPopulatedTextShapeLine instanceof ShapeModel,
  validation: reopenedBrowserTextShapeLineDeck.diagnostics.every(
    ({ severity }) => severity !== 'error',
  ),
};
if (Object.values(browserTextShapeLineChecks).some((value) => !value)) {
  throw new Error('Browser text shape line failed: ' + JSON.stringify(browserTextShapeLineChecks));
}
const browserTextShapeArrowDeck = PptxDocument.create();
const browserTextShapeArrowLayout = browserTextShapeArrowDeck.layouts[0];
const browserTextShapeArrowPlaceholder = browserTextShapeArrowLayout.addPlaceholder(
  'Browser text arrow prompt',
  {
    name: 'browser_text_arrow_placeholder',
    type: 'title',
    index: 196,
    arrows: { begin: 'none', end: 'stealth' },
  },
);
const browserTextShapeArrowSlide = browserTextShapeArrowDeck.addSlide({
  masterName: browserTextShapeArrowLayout.name,
});
const browserTextShapeArrowSource = { begin: 'triangle', end: 'arrow' };
const browserPlainTextShapeArrow = browserTextShapeArrowSlide.addText(
  'Browser plain text arrow',
  {
    name: 'browser_plain_text_arrow',
    line: {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      width: 2,
      dash: 'dashDot',
    },
    arrows: browserTextShapeArrowSource,
  },
);
const browserRichTextShapeArrow = browserTextShapeArrowSlide.addRichText([{
  runs: [{ text: 'Browser rich text arrow' }],
}], {
  name: 'browser_rich_text_arrow',
  arrows: { end: 'diamond' },
});
const browserPopulatedTextShapeArrow = browserTextShapeArrowSlide.addText(
  'Browser populated text arrow',
  { placeholder: 'browser_text_arrow_placeholder', arrows: { begin: 'arrow' } },
);
const browserPlainTextShapeArrowSnapshot = browserPlainTextShapeArrow.arrows;
browserTextShapeArrowSource.begin = 'oval';
browserTextShapeArrowSource.end = 'diamond';
const reopenedBrowserTextShapeArrowDeck = await PptxDocument.open(
  await browserTextShapeArrowDeck.writeBlob(),
);
await reopenedBrowserTextShapeArrowDeck.write({ compatibility: 'powerpoint-current' });
const reopenedBrowserTextShapeArrowSlide = reopenedBrowserTextShapeArrowDeck.slides[0];
const reopenedBrowserTextShapeArrowLayout = reopenedBrowserTextShapeArrowDeck.layouts[0];
const browserTextShapeArrowByName = (owner, name) => owner.shapes.find(
  (shape) => shape instanceof ShapeModel && shape.name === name,
);
const browserTextShapeArrowChecks = {
  detached: JSON.stringify(browserPlainTextShapeArrow.arrows) ===
    JSON.stringify(browserPlainTextShapeArrowSnapshot),
  plain: JSON.stringify(browserTextShapeArrowByName(
    reopenedBrowserTextShapeArrowSlide,
    'browser_plain_text_arrow',
  )?.arrows) === JSON.stringify({ begin: 'triangle', end: 'arrow' }),
  rich: JSON.stringify(browserTextShapeArrowByName(
    reopenedBrowserTextShapeArrowSlide,
    'browser_rich_text_arrow',
  )?.arrows) === JSON.stringify({ end: 'diamond' }),
  populated: JSON.stringify(browserTextShapeArrowByName(
    reopenedBrowserTextShapeArrowSlide,
    'browser_text_arrow_placeholder',
  )?.arrows) === JSON.stringify({ begin: 'arrow' }),
  placeholder: JSON.stringify(browserTextShapeArrowByName(
    reopenedBrowserTextShapeArrowLayout,
    'browser_text_arrow_placeholder',
  )?.arrows) === JSON.stringify({ begin: 'none', end: 'stealth' }),
  line: browserTextShapeArrowByName(
    reopenedBrowserTextShapeArrowSlide,
    'browser_plain_text_arrow',
  )?.line?.kind === 'line',
  live: browserTextShapeArrowPlaceholder instanceof ShapeModel &&
    browserRichTextShapeArrow instanceof ShapeModel &&
    browserPopulatedTextShapeArrow instanceof ShapeModel,
  validation: reopenedBrowserTextShapeArrowDeck.diagnostics.every(
    ({ severity }) => severity !== 'error',
  ),
};
if (Object.values(browserTextShapeArrowChecks).some((value) => !value)) {
  throw new Error(
    'Browser text shape arrow failed: ' + JSON.stringify(browserTextShapeArrowChecks),
  );
}
const browserLineDeck = PptxDocument.create();
const browserLineSlide = browserLineDeck.addSlide();
const browserLineColor = { kind: 'srgb', value: '#335577' };
const browserLineSource = {
  kind: 'line',
  color: browserLineColor,
  transparency: 12.3456,
  width: 2.50001,
  dash: 'dashDot',
};
const browserLinedShape = browserLineSlide.addShape('ellipse', {
  line: browserLineSource,
});
const browserInitialLine = browserLinedShape.line;
const browserInitialLineAgain = browserLinedShape.line;
browserLineColor.value = 'FFFFFF';
browserLineSource.transparency = 90;
browserLineSource.width = 9;
browserLineSource.dash = 'solid';
const browserDetachedLine = browserLinedShape.line;
const browserLinePartCount = browserLineDeck.opcPackage.parts.length;
const browserLineRelationshipCount = browserLineSlide.relationships.length;
browserLinedShape.line = { kind: 'none' };
browserLinedShape.line = undefined;
browserLinedShape.line = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent5' },
  transparency: 50,
  width: 0,
  dash: 'sysDot',
};
const reopenedBrowserLineDeck = await PptxDocument.open(await browserLineDeck.writeBlob());
const reopenedBrowserLine = reopenedBrowserLineDeck.slides[0]?.shapes[0]?.line;
const browserLineChecks = {
  initial: browserInitialLine?.kind === 'line' &&
    browserInitialLine.color.kind === 'srgb' &&
    browserInitialLine.color.value === '335577' &&
    Math.abs((browserInitialLine.transparency ?? 0) - 12.346) < 1e-9 &&
    browserInitialLine.width === 31_750 / 12_700 &&
    browserInitialLine.dash === 'dashDot',
  detached: browserInitialLine !== browserInitialLineAgain &&
    browserInitialLine?.kind === 'line' &&
    browserInitialLineAgain?.kind === 'line' &&
    browserInitialLine.color !== browserInitialLineAgain.color &&
    JSON.stringify(browserDetachedLine) === JSON.stringify(browserInitialLine),
  partIsolation: browserLineDeck.opcPackage.parts.length === browserLinePartCount,
  relationshipIsolation: browserLineSlide.relationships.length ===
    browserLineRelationshipCount,
  edited: JSON.stringify(browserLinedShape.line) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent5' },
    transparency: 50,
    width: 0,
    dash: 'sysDot',
  }),
  reopened: JSON.stringify(reopenedBrowserLine) === JSON.stringify({
    kind: 'line',
    color: { kind: 'scheme', value: 'accent5' },
    transparency: 50,
    width: 0,
    dash: 'sysDot',
  }),
};
if (Object.values(browserLineChecks).some((value) => !value)) {
  throw new Error('Browser shape line failed: ' + JSON.stringify({
    checks: browserLineChecks,
    initial: browserInitialLine,
  }));
}
const browserArrowDeck = PptxDocument.create();
const browserArrowSlide = browserArrowDeck.addSlide();
const browserArrowSource = { begin: 'triangle', end: 'arrow' };
const browserArrowShape = browserArrowSlide.addShape('line', {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '224466' },
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: browserArrowSource,
});
const browserInitialArrows = browserArrowShape.arrows;
const browserInitialArrowsAgain = browserArrowShape.arrows;
browserArrowSource.begin = 'none';
browserArrowSource.end = 'none';
const browserDetachedArrows = browserArrowShape.arrows;
const browserArrowPartCount = browserArrowDeck.opcPackage.parts.length;
const browserArrowRelationshipCount = browserArrowSlide.relationships.length;
const browserArrowNoOpBytes = browserArrowDeck.opcPackage
  .requirePart(browserArrowSlide.partUri).bytes.slice();
const browserArrowNoOpJournal = browserArrowDeck.opcPackage.mutations.length;
browserArrowShape.arrows = { begin: 'triangle', end: 'arrow' };
const browserArrowCurrentBytes = browserArrowDeck.opcPackage
  .requirePart(browserArrowSlide.partUri).bytes;
const browserArrowNoOp = browserArrowNoOpJournal === browserArrowDeck.opcPackage.mutations.length &&
  browserArrowNoOpBytes.length === browserArrowCurrentBytes.length &&
  browserArrowNoOpBytes.every((value, index) => value === browserArrowCurrentBytes[index]);
browserArrowShape.arrows = { begin: 'diamond' };
const browserPartialArrows = browserArrowShape.arrows;
browserArrowShape.arrows = undefined;
const browserClearedArrows = browserArrowShape.arrows;
const browserLineAfterArrowClear = browserArrowShape.line;
browserArrowShape.arrows = { begin: 'none', end: 'stealth' };
browserArrowShape.line = undefined;
const browserArrowsAfterLineClear = browserArrowShape.arrows;
const reopenedBrowserArrowDeck = await PptxDocument.open(await browserArrowDeck.writeBlob());
const reopenedBrowserArrows = reopenedBrowserArrowDeck.slides[0]?.shapes[0]?.arrows;
const browserArrowChecks = {
  initial: JSON.stringify(browserInitialArrows) ===
    JSON.stringify({ begin: 'triangle', end: 'arrow' }),
  detached: browserInitialArrows !== browserInitialArrowsAgain &&
    Object.isFrozen(browserInitialArrows) &&
    JSON.stringify(browserDetachedArrows) === JSON.stringify(browserInitialArrows),
  noOp: browserArrowNoOp,
  partial: JSON.stringify(browserPartialArrows) === JSON.stringify({ begin: 'diamond' }),
  clearPreservesLine: browserClearedArrows === undefined &&
    JSON.stringify(browserLineAfterArrowClear) === JSON.stringify({
      kind: 'line',
      color: { kind: 'srgb', value: '224466' },
      width: 2.5,
      dash: 'dashDot',
    }),
  lineClearPreservesArrows: JSON.stringify(browserArrowsAfterLineClear) ===
    JSON.stringify({ begin: 'none', end: 'stealth' }),
  partIsolation: browserArrowDeck.opcPackage.parts.length === browserArrowPartCount,
  relationshipIsolation: browserArrowSlide.relationships.length ===
    browserArrowRelationshipCount,
  reopened: JSON.stringify(reopenedBrowserArrows) ===
    JSON.stringify({ begin: 'none', end: 'stealth' }),
};
if (Object.values(browserArrowChecks).some((value) => !value)) {
  throw new Error('Browser shape arrows failed: ' + JSON.stringify({
    checks: browserArrowChecks,
    initial: browserInitialArrows,
  }));
}
const created = PptxDocument.create({ rtlMode: true, slideSize: '16:9' });
const browserText = created.addSlide().addText('Browser\\nText', { align: 'center', fit: 'resize', valign: 'bottom', vert: 'vert', wrap: false, bullet: true, level: 2, margin: [0, 0, 0, 0], rtlMode: true, spacing: { line: { kind: 'multiple', factor: 1.25 } }, tabStops: [{ position: 1.25 }] });
if (browserText.textWrap !== false || browserText.verticalAlignment !== 'bottom' || browserText.textDirection !== 'vert' || browserText.textFit !== 'resize' || browserText.richText.some(({ rtl }) => rtl !== true) || browserText.richText[0].tabStops[0].position !== 1.25 || browserText.textMargins.top !== 0 || browserText.textMargins.right !== 0 || browserText.textMargins.bottom !== 0 || browserText.textMargins.left !== 0) throw new Error('Browser create-text API failed');
const browserRich = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'number', style: 'alphaUcPeriod' }, level: 3, spacing: { before: 4, after: 6 }, tabStops: [{ position: 2.5, alignment: 'decimal' }], runs: [{ text: 'Rich', style: { lang: 'ja-JP', baseline: 'subscript', characterSpacing: 0, bold: true, glow: { opacity: 0.75, size: 4 }, highlight: { kind: 'scheme', value: 'accent1' }, outline: { color: { kind: 'srgb', value: 'ff0000' }, size: 1.25 }, underline: { style: 'wavyDbl' }, strike: 'dblStrike' } }] }], { rtlMode: true }).richText[0];
if (browserRich.rtl !== true || browserRich.tabStops[0].alignment !== 'decimal' || browserRich.runs[0].style.lang !== 'ja-JP' || browserRich.runs[0].style.baseline !== 'subscript' || browserRich.runs[0].style.characterSpacing !== 0 || browserRich.runs[0].style.glow.color.value !== 'FFFFFF' || browserRich.runs[0].style.glow.opacity !== 0.75 || browserRich.runs[0].style.glow.size !== 4 || browserRich.runs[0].style.highlight.value !== 'accent1' || browserRich.runs[0].style.outline.color.value !== 'FF0000' || browserRich.runs[0].style.outline.size !== 1.25 || browserRich.runs[0].style.underline.style !== 'wavyDbl' || browserRich.runs[0].style.strike !== 'dblStrike') throw new Error('Browser rich-text API failed');
const browserTransparency = created.slides[0].addRichText([{ runs: [{ text: 'Half', style: { transparency: 50 } }] }]);
if (browserTransparency.richText[0].runs[0].style.transparency !== 50) throw new Error('Browser transparency create failed');
browserTransparency.richText = [{ runs: [{ text: 'Cleared' }] }];
if (browserTransparency.richText[0].runs[0].style?.transparency !== undefined) throw new Error('Browser transparency clear failed');
const browserMargin = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginLeft: 24 });
if (browserMargin.richText[0].marginLeft !== 12) throw new Error('Browser paragraph margin create failed');
browserMargin.richText = [{ marginLeft: false, runs: [{ text: 'Cleared' }] }];
if (browserMargin.richText[0].marginLeft !== undefined) throw new Error('Browser paragraph margin clear failed');
const browserRightMargin = created.slides[0].addRichText([{ bullet: true, marginRight: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginRight: 24 });
if (browserRightMargin.richText[0].marginRight !== 12 || browserRightMargin.richText[0].bullet.indent !== 27) throw new Error('Browser paragraph right margin create failed');
browserRightMargin.richText = [{ marginRight: false, runs: [{ text: 'Cleared' }] }];
if (browserRightMargin.richText[0].marginRight !== undefined) throw new Error('Browser paragraph right margin clear failed');
const browserIndent = created.slides[0].addRichText([{ indent: -12, runs: [{ text: 'Indent' }] }, { bullet: true, indent: false, runs: [{ text: 'Bullet' }] }], { paragraphIndent: 24 });
if (browserIndent.richText[0].indent !== -12 || browserIndent.richText[1].indent !== undefined || browserIndent.richText[1].bullet.indent !== 27) throw new Error('Browser paragraph indent create failed');
browserIndent.richText = [{ indent: false, runs: [{ text: 'Cleared' }] }];
if (browserIndent.richText[0].indent !== undefined) throw new Error('Browser paragraph indent clear failed');
const tableSlide = created.slides[0];
const browserCreationColor = { kind: 'srgb', value: '#D9EAF7' };
const browserCreationFill = { kind: 'solid', color: browserCreationColor, transparency: 33.3334 };
const browserTableFillColor = { kind: 'scheme', value: 'accent4' };
const browserTableFill = { kind: 'solid', color: browserTableFillColor, transparency: 40 };
const browserTableBorderColor = { kind: 'scheme', value: 'accent4' };
const browserTableBorder = { kind: 'line', color: browserTableBorderColor, width: 1.5, style: 'dash' };
const browserCreationBorderColor = { kind: 'srgb', value: '#C00000' };
const browserCreationBorder = { kind: 'line', color: browserCreationBorderColor, width: 2, style: 'solid' };
const browserCreationMargin = { top: 4, left: 8 };
const createdTable = tableSlide.addTable([
  [
    { text: 'Region', options: { align: 'left', border: browserCreationBorder, fill: browserCreationFill, fit: 'shrink', margin: browserCreationMargin, textDirection: 'vert', valign: 'top' } },
    { text: 'Revenue', options: {
      border: [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, undefined, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
      fit: 'resize',
      margin: [1, 2, 3, 4],
      textDirection: 'vert270',
      valign: 'middle',
    } },
  ],
  [
    { text: 'West', options: { align: 'right', border: { top: { kind: 'line', color: { kind: 'scheme', value: 'accent3' }, width: 1, style: 'dash' }, left: { kind: 'none' } }, fit: 'none', margin: 0, textDirection: 'wordArtVert', valign: 'bottom' } },
    { text: '', options: { align: 'justify', border: { kind: 'none' }, fill: { kind: 'none' }, margin: {}, textDirection: 'horz' } },
  ],
], { name: 'Created browser table', align: 'center', columnWidths: inches(1.25), rowHeights: inches(0.75), fill: browserTableFill, margin: { top: 9, left: 18 }, valign: 'middle' });
const browserTableDefaultsSlide = created.addSlide();
const browserTableBorderDefault = browserTableDefaultsSlide.addTable([[
  'Inherited border',
  { text: 'None override', options: { border: { kind: 'none' } } },
]], { name: 'Created browser table border default', columnWidths: inches(1), rowHeights: inches(0.5), border: browserTableBorder });
const browserTableDirectionDefault = browserTableDefaultsSlide.addTable([[
  'Inherited direction',
  { text: 'Horizontal override', options: { textDirection: 'horz' } },
]], { name: 'Created browser table direction default', columnWidths: inches(1), rowHeights: inches(0.5), textDirection: 'vert270' });
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ text }) => text))) !== JSON.stringify([['Region', 'Revenue'], ['West', '']])) throw new Error('Browser table cell object creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ verticalAlignment }) => verticalAlignment))) !== JSON.stringify([['top', 'middle'], ['bottom', 'middle']])) throw new Error('Browser table vertical alignment creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))) !== JSON.stringify([['vert', 'vert270'], ['wordArtVert', undefined]])) throw new Error('Browser table cell text direction creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))) !== JSON.stringify([['shrink', 'resize'], [undefined, undefined]])) throw new Error('Browser table cell text fit creation failed');
if (JSON.stringify(browserTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection)) !== JSON.stringify(['vert270', undefined])) throw new Error('Browser table text direction creation failed');
const browserMarginVector = (margins) => [margins?.top, margins?.right, margins?.bottom, margins?.left];
const browserInitialMargins = createdTable.rows.map(({ cells }) => cells.map(({ margins }) => browserMarginVector(margins)));
if (JSON.stringify(browserInitialMargins) !== JSON.stringify([[[4, 7.2, 3.6, 8], [1, 2, 3, 4]], [[0, 0, 0, 0], [9, 7.2, 3.6, 18]]])) throw new Error('Browser table cell margin creation failed');
if (createdTable.rows[0].cells[0].fill?.kind !== 'solid' || createdTable.rows[0].cells[0].fill.color.kind !== 'srgb' || createdTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || createdTable.rows[0].cells[0].fill.transparency !== 33.333 || createdTable.rows[0].cells[1].fill?.kind !== 'solid' || createdTable.rows[0].cells[1].fill.color.kind !== 'scheme' || createdTable.rows[0].cells[1].fill.color.value !== 'accent2' || createdTable.rows[0].cells[1].fill.transparency !== 25 || createdTable.rows[1].cells[1].fill?.kind !== 'none') throw new Error('Browser table cell fill creation failed');
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent4' || createdTable.rows[1].cells[0].fill.transparency !== 40) throw new Error('Browser table fill inheritance failed');
const browserCreationSides = ['top', 'right', 'bottom', 'left'];
const browserIsCreationLine = (border, colorKind, colorValue, width, style) => border?.kind === 'line' && border.color.kind === colorKind && border.color.value === colorValue && border.width === width && border.style === style;
const browserAllCreationLines = (borders, colorKind, colorValue, width, style) => borders !== undefined && browserCreationSides.every((side) => browserIsCreationLine(borders[side], colorKind, colorValue, width, style));
const browserAllCreationNone = (borders) => borders !== undefined && browserCreationSides.every((side) => borders[side]?.kind === 'none');
const browserScalarBorders = createdTable.rows[0].cells[0].borders;
const browserTupleBorders = createdTable.rows[0].cells[1].borders;
const browserNamedBorders = createdTable.rows[1].cells[0].borders;
if (!browserAllCreationLines(browserScalarBorders, 'srgb', 'C00000', 2, 'solid') || !browserIsCreationLine(browserTupleBorders?.top, 'scheme', 'accent1', 1.5, 'dash') || browserTupleBorders?.right?.kind !== 'none' || !browserIsCreationLine(browserTupleBorders?.bottom, 'srgb', '00FF00', 0, undefined) || browserTupleBorders?.left?.kind !== 'none' || !browserIsCreationLine(browserNamedBorders?.top, 'scheme', 'accent3', 1, 'dash') || browserNamedBorders?.right?.kind !== 'none' || browserNamedBorders?.bottom?.kind !== 'none' || browserNamedBorders?.left?.kind !== 'none' || !browserAllCreationNone(createdTable.rows[1].cells[1].borders)) throw new Error('Browser table cell border creation failed');
const browserInitialTableDefaultBorders = browserTableBorderDefault.rows[0].cells[0].borders;
if (!browserAllCreationLines(browserInitialTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') || !browserAllCreationNone(browserTableBorderDefault.rows[0].cells[1].borders)) throw new Error('Browser table border inheritance failed');
browserCreationColor.value = '000000';
browserCreationFill.transparency = 1;
browserTableFillColor.value = 'accent6';
browserTableFill.transparency = 1;
browserTableBorderColor.value = 'accent6';
browserTableBorder.width = 9;
browserCreationBorderColor.value = '000000';
browserCreationBorder.width = 9;
browserCreationMargin.top = 99;
browserCreationMargin.left = 99;
if (createdTable.rows[0].cells[0].fill?.kind !== 'solid' || createdTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || createdTable.rows[0].cells[0].fill.transparency !== 33.333) throw new Error('Browser table cell fill creation retained source state');
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent4' || createdTable.rows[1].cells[0].fill.transparency !== 40) throw new Error('Browser table fill creation retained source state');
if (!browserAllCreationLines(createdTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid')) throw new Error('Browser table cell border creation retained source state');
if (!browserAllCreationLines(browserTableBorderDefault.rows[0].cells[0].borders, 'scheme', 'accent4', 1.5, 'dash')) throw new Error('Browser table border creation retained source state');
if (JSON.stringify(browserMarginVector(createdTable.rows[0].cells[0].margins)) !== JSON.stringify([4, 7.2, 3.6, 8])) throw new Error('Browser table cell margin creation retained source state');
const createdTablePartXml = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
const browserCreatedTableCells = [...createdTablePartXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0]);
const browserCreatedTableHorizontalAlignments = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const browserCreatedTableDirections = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const browserCreatedTableFits = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:(normAutofit|spAutoFit)\\/>/)?.[1]);
const createdTableGrid = [...createdTablePartXml.matchAll(/<a:gridCol w="(\\d+)"\\/>/g)].map((match) => Number(match[1]));
const createdTableRows = [...createdTablePartXml.matchAll(/<a:tr h="(\\d+)">/g)].map((match) => Number(match[1]));
if (!createdTablePartXml.includes('<a:tcPr marL="228600" marR="91440" marT="114300" marB="45720" anchor="ctr">')) throw new Error('Browser table margin XML creation failed');
if (JSON.stringify(browserCreatedTableHorizontalAlignments) !== JSON.stringify(['l', 'ctr', 'r', 'just'])) throw new Error('Browser table horizontal alignment creation failed');
if (JSON.stringify(browserCreatedTableDirections) !== JSON.stringify(['vert', 'vert270', 'wordArtVert', undefined])) throw new Error('Browser table cell text direction XML creation failed');
if (JSON.stringify(browserCreatedTableFits) !== JSON.stringify(['normAutofit', 'spAutoFit', undefined, undefined])) throw new Error('Browser table cell text fit XML creation failed');
if (!(createdTable instanceof TableModel) || createdTable.transform.x !== inches(0.5) || createdTable.transform.y !== inches(0.5) || createdTable.transform.width !== inches(2.5) || createdTable.transform.height !== inches(1.5) || createdTableGrid.length !== 2 || createdTableGrid.some((width) => width !== inches(1.25)) || createdTableRows.length !== 2 || createdTableRows.some((height) => height !== inches(0.75)) || createdTable.rows[1].cells[1].margins?.top !== 9 || createdTable.rows[1].cells[1].margins?.left !== 18) throw new Error('Browser table sizing creation failed');
createdTable.setColumnWidths([inches(1), inches(1.5)]);
if (createdTable.columnWidths?.join(',') !== [inches(1), inches(1.5)].join(',') || createdTable.transform.width !== inches(2.5)) throw new Error('Browser table column-width editing failed');
createdTable.setRowHeights([inches(0.5), inches(1)]);
if (createdTable.rowHeights?.join(',') !== [inches(0.5), inches(1)].join(',') || createdTable.transform.height !== inches(1.5)) throw new Error('Browser table row-height editing failed');
createdTable.setCellText(1, 0, 'Edited West');
createdTable.setCellFill(1, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 50 });
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent1' || createdTable.rows[1].cells[0].fill.transparency !== 50) throw new Error('Browser table fill override failed');
createdTable.setCellBorders(1, 0, { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, width: 1, style: 'solid' });
createdTable.setCellVerticalAlignment(0, 0, 'bottom');
createdTable.setCellVerticalAlignment(1, 1, undefined);
createdTable.setCellMargins(0, 0, { bottom: 9 });
createdTable.setCellMargins(1, 1, undefined);
createdTable.setCellFill(1, 1, undefined);
if (createdTable.rows[1].cells[1].fill !== undefined) throw new Error('Browser table fill clear re-inherited');
browserTableBorderDefault.setCellBorders(0, 0, undefined);
if (browserTableBorderDefault.rows[0].cells[0].borders !== undefined) throw new Error('Browser table border clear re-inherited');
const reopenedCreated = await PptxDocument.open(await created.write());
const reopenedCreatedTable = reopenedCreated.slides[0].shapes.find((shape) => shape.name === 'Created browser table');
const reopenedCreatedTableXml = new TextDecoder().decode(reopenedCreated.opcPackage.requirePart(reopenedCreated.slides[0].partUri).bytes);
const reopenedCreatedTableHorizontalAlignments = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const reopenedCreatedTableDirections = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const reopenedBrowserTableBorderDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created browser table border default');
const reopenedBrowserTableDirectionDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created browser table direction default');
if (!(reopenedCreatedTable instanceof TableModel) || reopenedCreatedTable.columnWidths?.join(',') !== [inches(1), inches(1.5)].join(',') || reopenedCreatedTable.rowHeights?.join(',') !== [inches(0.5), inches(1)].join(',') || reopenedCreatedTable.transform.width !== inches(2.5) || reopenedCreatedTable.transform.height !== inches(1.5) || reopenedCreatedTable.rows[1].cells[0].text !== 'Edited West' || reopenedCreatedTable.rows[1].cells[1].text !== '' || reopenedCreatedTable.rows[0].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || reopenedCreatedTable.rows[0].cells[1].fill?.kind !== 'solid' || reopenedCreatedTable.rows[0].cells[1].fill.color.kind !== 'scheme' || reopenedCreatedTable.rows[0].cells[1].fill.transparency !== 25 || reopenedCreatedTable.rows[1].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[1].cells[0].fill.color.value !== 'accent1' || reopenedCreatedTable.rows[1].cells[0].fill.transparency !== 50 || reopenedCreatedTable.rows[1].cells[1].fill !== undefined || !browserAllCreationLines(reopenedCreatedTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid') || !browserIsCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.top, 'scheme', 'accent1', 1.5, 'dash') || reopenedCreatedTable.rows[0].cells[1].borders?.right?.kind !== 'none' || !browserIsCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.bottom, 'srgb', '00FF00', 0, undefined) || reopenedCreatedTable.rows[0].cells[1].borders?.left?.kind !== 'none' || !browserAllCreationLines(reopenedCreatedTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') || !browserAllCreationNone(reopenedCreatedTable.rows[1].cells[1].borders)) throw new Error('Browser table creation round trip failed');
if (reopenedCreatedTable.rows[1].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[1].cells[0].fill.color.kind !== 'scheme' || reopenedCreatedTable.rows[1].cells[0].fill.color.value !== 'accent1' || reopenedCreatedTable.rows[1].cells[0].fill.transparency !== 50 || reopenedCreatedTable.rows[1].cells[1].fill !== undefined) throw new Error('Browser table fill round trip failed');
if (!(reopenedBrowserTableBorderDefault instanceof TableModel) || reopenedBrowserTableBorderDefault.rows[0].cells[0].borders !== undefined || !browserAllCreationNone(reopenedBrowserTableBorderDefault.rows[0].cells[1].borders)) throw new Error('Browser table border round trip failed');
if (reopenedCreatedTable.rows[0].cells[0].margins?.top !== undefined || reopenedCreatedTable.rows[0].cells[0].margins?.right !== undefined || reopenedCreatedTable.rows[0].cells[0].margins?.bottom !== 9 || reopenedCreatedTable.rows[0].cells[0].margins?.left !== undefined) throw new Error('Browser table cell margin creation round trip failed');
if (reopenedCreatedTable.rows[1].cells[1].margins !== undefined) throw new Error('Browser table margin clear re-inherited');
if (reopenedCreatedTable.rows[0].cells[0].verticalAlignment !== 'bottom' || reopenedCreatedTable.rows[0].cells[1].verticalAlignment !== 'middle' || reopenedCreatedTable.rows[1].cells[0].verticalAlignment !== 'bottom' || reopenedCreatedTable.rows[1].cells[1].verticalAlignment !== undefined) throw new Error('Browser table cell vertical alignment creation round trip failed');
if (JSON.stringify(reopenedCreatedTableHorizontalAlignments) !== JSON.stringify(['l', 'ctr', 'r', 'just'])) throw new Error('Browser table horizontal alignment creation round trip failed');
if (JSON.stringify(reopenedCreatedTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))) !== JSON.stringify([['vert', 'vert270'], ['wordArtVert', undefined]]) || JSON.stringify(reopenedCreatedTableDirections) !== JSON.stringify(['vert', 'vert270', 'wordArtVert', undefined])) throw new Error('Browser table cell text direction creation round trip failed');
if (!(reopenedBrowserTableDirectionDefault instanceof TableModel) || JSON.stringify(reopenedBrowserTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection)) !== JSON.stringify(['vert270', undefined])) throw new Error('Browser table text direction creation round trip failed');
const tablePart = created.opcPackage.requirePart(tableSlide.partUri);
const tableXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Browser table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:bodyPr custom="TARGET"><a:noAutofit/></a:bodyPr><a:p><a:r><a:t>Browser target</a:t></a:r></a:p></a:txBody><a:tcPr vert="horz" anchor="ctr" marL="12700" marR="25400" marT="38100" marB="50800"><a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="sysDash"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr custom="NEIGHBOR"><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>Browser neighbor</a:t></a:r></a:p></a:txBody><a:tcPr vert="vert" anchor="b" marL="63500" marR="76200" marT="88900" marB="101600"><a:lnL w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
created.opcPackage.setPart(tableSlide.partUri, new TextDecoder().decode(tablePart.bytes).replace('</p:spTree>', tableXml + '</p:spTree>'), tablePart.contentType);
const table = tableSlide.shapes.find((shape) => shape.name === 'Browser table');
if (!(table instanceof TableModel) || table.rows[0].cells[0].textFit !== 'none' || table.rows[0].cells[1].textFit !== 'resize') throw new Error('Browser table-cell fit read failed');
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined) {
  throw new Error('Browser table-cell horizontal alignment initial read failed');
}
table.setCellHorizontalAlignment(0, 0, 'left');
if (table.rows[0].cells[0].horizontalAlignment !== 'left') {
  throw new Error('Browser table-cell left alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'center');
if (table.rows[0].cells[0].horizontalAlignment !== 'center') {
  throw new Error('Browser table-cell center alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'right');
if (table.rows[0].cells[0].horizontalAlignment !== 'right') {
  throw new Error('Browser table-cell right alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'justify');
if (table.rows[0].cells[0].horizontalAlignment !== 'justify') {
  throw new Error('Browser table-cell justify alignment failed');
}
table.setCellHorizontalAlignment(0, 0, undefined);
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined ||
    table.rows[0].cells[0].text !== 'Browser target' ||
    table.rows[0].cells[1].text !== 'Browser neighbor') {
  throw new Error('Browser table-cell horizontal alignment clear failed');
}
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'scheme' || table.rows[0].cells[0].borders.top.color.value !== 'accent2' || table.rows[0].cells[0].borders.top.width !== 1.5 || table.rows[0].cells[0].borders.top.style !== 'dash' || table.rows[0].cells[0].borders.right?.kind !== 'none' || table.rows[0].cells[0].borders.bottom?.kind !== 'none' || table.rows[0].cells[0].borders.left?.kind !== 'line') throw new Error('Browser table-cell border read failed');
table.setCellBorders(0, 0, { kind: 'line', color: { kind: 'srgb', value: '#0000FF' }, width: 2, style: 'solid' });
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'srgb' || table.rows[0].cells[0].borders.top.color.value !== '0000FF' || table.rows[0].cells[0].borders.top.width !== 2 || table.rows[0].cells[0].borders.top.style !== 'solid' || table.rows[0].cells[0].borders.right?.kind !== 'line' || table.rows[0].cells[0].borders.bottom?.kind !== 'line' || table.rows[0].cells[0].borders.left?.kind !== 'line') throw new Error('Browser table-cell scalar border failed');
table.setCellBorders(0, 0, [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }, undefined]);
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'scheme' || table.rows[0].cells[0].borders.top.color.value !== 'accent1' || table.rows[0].cells[0].borders.top.style !== 'dash' || table.rows[0].cells[0].borders.right?.kind !== 'line' || table.rows[0].cells[0].borders.right.width !== 0 || table.rows[0].cells[0].borders.right.style !== undefined || table.rows[0].cells[0].borders.bottom?.kind !== 'none' || table.rows[0].cells[0].borders.left !== undefined) throw new Error('Browser table-cell tuple border failed');
table.setCellBorders(0, 0, { left: { kind: 'none' } });
if (table.rows[0].cells[0].borders?.left?.kind !== 'none' || table.rows[0].cells[0].borders.top !== undefined || table.rows[0].cells[0].borders.right !== undefined || table.rows[0].cells[0].borders.bottom !== undefined) throw new Error('Browser table-cell partial border failed');
table.setCellBorders(0, 0, undefined);
if (table.rows[0].cells[0].borders !== undefined || table.rows[0].cells[1].borders?.left?.kind !== 'line' || table.rows[0].cells[1].borders.left.color.kind !== 'srgb' || table.rows[0].cells[1].borders.left.color.value !== '333333' || table.rows[0].cells[1].borders.left.width !== 2 || table.rows[0].cells[1].borders.left.style !== 'solid') throw new Error('Browser table-cell border clear failed');
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'scheme' || table.rows[0].cells[0].fill.color.value !== 'accent1' || table.rows[0].cells[0].fill.transparency !== 25) throw new Error('Browser table-cell fill read failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '#FF0000' } });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'srgb' || table.rows[0].cells[0].fill.color.value !== 'FF0000' || table.rows[0].cells[0].fill.transparency !== undefined) throw new Error('Browser table-cell opaque fill failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 0 });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'scheme' || table.rows[0].cells[0].fill.color.value !== 'accent2' || table.rows[0].cells[0].fill.transparency !== 0) throw new Error('Browser table-cell explicit opaque fill failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '112233' }, transparency: 33.333 });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'srgb' || table.rows[0].cells[0].fill.transparency !== 33.333) throw new Error('Browser table-cell fractional fill failed');
table.setCellFill(0, 0, { kind: 'none' });
if (table.rows[0].cells[0].fill?.kind !== 'none') throw new Error('Browser table-cell no fill failed');
table.setCellFill(0, 0, undefined);
if (table.rows[0].cells[0].fill !== undefined || table.rows[0].cells[1].fill?.kind !== 'solid' || table.rows[0].cells[1].fill.color.kind !== 'srgb' || table.rows[0].cells[1].fill.color.value !== '70AD47' || table.rows[0].cells[1].fill.transparency !== 50) throw new Error('Browser table-cell fill clear failed');
if (table.rows[0].cells[0].margins?.top !== 3 || table.rows[0].cells[0].margins?.right !== 2 || table.rows[0].cells[0].margins?.bottom !== 4 || table.rows[0].cells[0].margins?.left !== 1) throw new Error('Browser table-cell margin read failed');
table.setCellMargins(0, 0, 4);
if (table.rows[0].cells[0].margins?.top !== 4 || table.rows[0].cells[0].margins?.right !== 4 || table.rows[0].cells[0].margins?.bottom !== 4 || table.rows[0].cells[0].margins?.left !== 4) throw new Error('Browser table-cell scalar margin failed');
table.setCellMargins(0, 0, [1, 2, 3, 4]);
if (table.rows[0].cells[0].margins?.top !== 1 || table.rows[0].cells[0].margins?.right !== 2 || table.rows[0].cells[0].margins?.bottom !== 3 || table.rows[0].cells[0].margins?.left !== 4) throw new Error('Browser table-cell tuple margin failed');
table.setCellMargins(0, 0, { top: 5, left: 7 });
if (table.rows[0].cells[0].margins?.top !== 5 || table.rows[0].cells[0].margins?.right !== undefined || table.rows[0].cells[0].margins?.bottom !== undefined || table.rows[0].cells[0].margins?.left !== 7) throw new Error('Browser table-cell partial margin failed');
table.setCellMargins(0, 0, undefined);
if (table.rows[0].cells[0].margins !== undefined || table.rows[0].cells[1].margins?.top !== 7 || table.rows[0].cells[1].margins?.right !== 6 || table.rows[0].cells[1].margins?.bottom !== 8 || table.rows[0].cells[1].margins?.left !== 5) throw new Error('Browser table-cell margin clear failed');
table.setCellTextFit(0, 0, 'shrink');
const beforeSameFit = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
table.setCellTextFit(0, 0, 'shrink');
if (table.rows[0].cells[0].textFit !== 'shrink' || new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes) !== beforeSameFit) throw new Error('Browser table-cell shrink fit failed');
table.setCellTextFit(0, 0, 'resize');
if (table.rows[0].cells[0].textFit !== 'resize') throw new Error('Browser table-cell resize fit failed');
table.setCellTextFit(0, 0, 'none');
if (table.rows[0].cells[0].textFit !== undefined) throw new Error('Browser table-cell none fit clear failed');
table.setCellTextFit(0, 0, 'shrink');
table.setCellTextFit(0, 0, undefined);
if (table.rows[0].cells[0].textFit !== undefined || table.rows[0].cells[1].textFit !== 'resize') throw new Error('Browser table-cell undefined fit clear failed');
if (table.rows[0].cells[0].verticalAlignment !== 'middle' || table.rows[0].cells[1].verticalAlignment !== 'bottom') throw new Error('Browser table-cell vertical alignment read failed');
table.setCellVerticalAlignment(0, 0, 'top');
if (table.rows[0].cells[0].verticalAlignment !== 'top') throw new Error('Browser table-cell top alignment failed');
table.setCellVerticalAlignment(0, 0, 'middle');
if (table.rows[0].cells[0].verticalAlignment !== 'middle') throw new Error('Browser table-cell middle alignment failed');
table.setCellVerticalAlignment(0, 0, 'bottom');
if (table.rows[0].cells[0].verticalAlignment !== 'bottom') throw new Error('Browser table-cell bottom alignment failed');
table.setCellVerticalAlignment(0, 0, undefined);
if (table.rows[0].cells[0].verticalAlignment !== undefined || table.rows[0].cells[1].verticalAlignment !== 'bottom') throw new Error('Browser table-cell vertical alignment clear failed');
table?.setCellTextDirection(0, 0, 'vert270');
if (!(table instanceof TableModel) || table.rows[0].cells[0].textDirection !== 'vert270' || table.rows[0].cells[1].textDirection !== 'vert') throw new Error('Browser table-cell direction edit failed');
table.setCellTextDirection(0, 0, 'wordArtVert');
if (table.rows[0].cells[0].textDirection !== 'wordArtVert') throw new Error('Browser table-cell stacked direction edit failed');
table.setCellTextDirection(0, 0, undefined);
if (table.rows[0].cells[0].textDirection !== undefined || table.rows[0].cells[0].text !== 'Browser target' || table.rows[0].cells[1].text !== 'Browser neighbor') throw new Error('Browser table-cell direction clear failed');
const reopenedBrowserEdited = await PptxDocument.open(await created.write());
const reopenedBrowserTable = reopenedBrowserEdited.slides[0].shapes.find(
  (shape) => shape.name === 'Browser table',
);
if (!(reopenedBrowserTable instanceof TableModel) ||
    reopenedBrowserTable.rows[0].cells[0].horizontalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].horizontalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[0].text !== 'Browser target' ||
    reopenedBrowserTable.rows[0].cells[1].text !== 'Browser neighbor' ||
    reopenedBrowserTable.rows[0].cells[0].textDirection !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].textDirection !== 'vert' ||
    reopenedBrowserTable.rows[0].cells[0].textFit !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].textFit !== 'resize' ||
    reopenedBrowserTable.rows[0].cells[0].verticalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].verticalAlignment !== 'bottom' ||
    reopenedBrowserTable.rows[0].cells[0].margins !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].margins?.left !== 5 ||
    reopenedBrowserTable.rows[0].cells[0].borders !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].borders?.left?.kind !== 'line' ||
    reopenedBrowserTable.rows[0].cells[0].fill !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].fill?.kind !== 'solid') {
  throw new Error('Browser table-cell horizontal alignment reopen failed');
}
if (created.rtlMode !== true) throw new Error('Browser presentation RTL create failed');
created.rtlMode = false;
if (created.rtlMode !== false || browserRich.rtl !== true) throw new Error('Browser presentation RTL edit failed');
created.rtlMode = undefined;
if (created.rtlMode !== undefined || browserRich.rtl !== true) throw new Error('Browser presentation RTL clear failed');
const browserMetadata = PptxDocument.create({ title: 'Browser title' });
if (browserMetadata.title !== 'Browser title') throw new Error('Browser presentation title create failed');
browserMetadata.title = 'Edited browser title';
if (browserMetadata.title !== 'Edited browser title') throw new Error('Browser presentation title edit failed');
const reopenedBrowserMetadata = await PptxDocument.open(await browserMetadata.write());
if (reopenedBrowserMetadata.title !== 'Edited browser title') throw new Error('Browser presentation title reopen failed');
browserMetadata.title = '';
if (browserMetadata.title !== '') throw new Error('Browser presentation title empty failed');
browserMetadata.title = undefined;
if (browserMetadata.title !== undefined) throw new Error('Browser presentation title clear failed');
const browserAuthorship = PptxDocument.create({ author: 'Browser author' });
if (browserAuthorship.author !== 'Browser author') throw new Error('Browser presentation author create failed');
browserAuthorship.author = 'Edited browser author';
if (browserAuthorship.author !== 'Edited browser author') throw new Error('Browser presentation author edit failed');
const reopenedBrowserAuthorship = await PptxDocument.open(await browserAuthorship.write());
if (reopenedBrowserAuthorship.author !== 'Edited browser author') throw new Error('Browser presentation author reopen failed');
browserAuthorship.author = '';
if (browserAuthorship.author !== '') throw new Error('Browser presentation author empty failed');
browserAuthorship.author = undefined;
if (browserAuthorship.author !== undefined) throw new Error('Browser presentation author clear failed');
const browserEditorship = PptxDocument.create({ lastModifiedBy: 'Browser editor' });
if (browserEditorship.lastModifiedBy !== 'Browser editor') throw new Error('Browser presentation lastModifiedBy create failed');
browserEditorship.lastModifiedBy = 'Edited browser editor';
if (browserEditorship.lastModifiedBy !== 'Edited browser editor') throw new Error('Browser presentation lastModifiedBy edit failed');
const reopenedBrowserEditorship = await PptxDocument.open(await browserEditorship.write());
if (reopenedBrowserEditorship.lastModifiedBy !== 'Edited browser editor') throw new Error('Browser presentation lastModifiedBy reopen failed');
browserEditorship.lastModifiedBy = '';
if (browserEditorship.lastModifiedBy !== '') throw new Error('Browser presentation lastModifiedBy empty failed');
browserEditorship.lastModifiedBy = undefined;
if (browserEditorship.lastModifiedBy !== undefined) throw new Error('Browser presentation lastModifiedBy clear failed');
const browserChronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
});
if (browserChronology.createdAt !== '2024-02-29T12:34:56.123+05:30') throw new Error('Browser presentation createdAt create failed');
if (browserChronology.modifiedAt !== '2024-03-01T01:02:03.456+08:00') throw new Error('Browser presentation modifiedAt create failed');
browserChronology.createdAt = '2026-07-30T00:00:00Z';
if (browserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation createdAt edit failed');
browserChronology.modifiedAt = '2026-07-30T01:02:03Z';
if (browserChronology.modifiedAt !== '2026-07-30T01:02:03Z') throw new Error('Browser presentation modifiedAt edit failed');
const reopenedBrowserChronology = await PptxDocument.open(await browserChronology.write());
if (reopenedBrowserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation createdAt reopen failed');
if (reopenedBrowserChronology.modifiedAt !== '2026-07-30T01:02:03Z') throw new Error('Browser presentation modifiedAt reopen failed');
browserChronology.modifiedAt = undefined;
if (browserChronology.modifiedAt !== undefined) throw new Error('Browser presentation modifiedAt clear failed');
if (browserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation modifiedAt changed createdAt');
browserChronology.createdAt = undefined;
if (browserChronology.createdAt !== undefined) throw new Error('Browser presentation createdAt clear failed');
const browserSubjectMatter = PptxDocument.create({ subject: 'Browser subject' });
if (browserSubjectMatter.subject !== 'Browser subject') throw new Error('Browser presentation subject create failed');
browserSubjectMatter.subject = 'Edited browser subject';
if (browserSubjectMatter.subject !== 'Edited browser subject') throw new Error('Browser presentation subject edit failed');
const reopenedBrowserSubjectMatter = await PptxDocument.open(await browserSubjectMatter.write());
if (reopenedBrowserSubjectMatter.subject !== 'Edited browser subject') throw new Error('Browser presentation subject reopen failed');
browserSubjectMatter.subject = '';
if (browserSubjectMatter.subject !== '') throw new Error('Browser presentation subject empty failed');
browserSubjectMatter.subject = undefined;
if (browserSubjectMatter.subject !== undefined) throw new Error('Browser presentation subject clear failed');
const browserRevisioned = PptxDocument.create({ revision: '007' });
if (browserRevisioned.revision !== '007') throw new Error('Browser presentation revision create failed');
browserRevisioned.revision = '42';
if (browserRevisioned.revision !== '42') throw new Error('Browser presentation revision edit failed');
const reopenedBrowserRevisioned = await PptxDocument.open(await browserRevisioned.write());
if (reopenedBrowserRevisioned.revision !== '42') throw new Error('Browser presentation revision reopen failed');
browserRevisioned.revision = undefined;
if (browserRevisioned.revision !== undefined) throw new Error('Browser presentation revision clear failed');
const browserOrganization = PptxDocument.create({ company: 'Browser company' });
if (browserOrganization.company !== 'Browser company') throw new Error('Browser presentation company create failed');
browserOrganization.company = 'Edited browser company';
if (browserOrganization.company !== 'Edited browser company') throw new Error('Browser presentation company edit failed');
const reopenedBrowserOrganization = await PptxDocument.open(await browserOrganization.write());
if (reopenedBrowserOrganization.company !== 'Edited browser company') throw new Error('Browser presentation company reopen failed');
browserOrganization.company = '';
if (browserOrganization.company !== '') throw new Error('Browser presentation company empty failed');
browserOrganization.company = undefined;
if (browserOrganization.company !== undefined) throw new Error('Browser presentation company clear failed');
const browserThemed = PptxDocument.create({
  theme: { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' },
});
const browserCreatedTheme = browserThemed.theme;
if (browserCreatedTheme?.headFontFace !== 'Aptos Display' || browserCreatedTheme.bodyFontFace !== 'Aptos') throw new Error('Browser presentation theme create failed');
browserCreatedTheme.headFontFace = 'Detached caller value';
if (browserThemed.theme?.headFontFace !== 'Aptos Display') throw new Error('Browser presentation theme snapshot was not detached');
browserThemed.theme = { headFontFace: 'Noto Sans Display' };
if (browserThemed.theme?.headFontFace !== 'Noto Sans Display' || browserThemed.theme.bodyFontFace !== 'Calibri') throw new Error('Browser presentation theme replacement failed');
browserThemed.masterLayoutTheme.presentationTheme.setFonts({ minorLatin: 'Noto Sans' });
const reopenedBrowserThemed = await PptxDocument.open(await browserThemed.writeBlob());
if (reopenedBrowserThemed.theme?.headFontFace !== 'Noto Sans Display' || reopenedBrowserThemed.theme.bodyFontFace !== 'Noto Sans') throw new Error('Browser presentation theme reopen failed');
const browserSectioned = PptxDocument.create();
const browserIntro = browserSectioned.addSection({ title: 'Browser intro' });
const browserSectionSlide = browserSectioned.addSlide({ sectionTitle: 'Browser intro' });
const browserData = browserSectioned.addSection({ title: 'Browser data' });
browserSectioned.assignSlideToSection(0, browserData.id);
browserSectioned.renameSection(browserIntro.id, 'Browser edited');
browserSectioned.moveSection(browserData.id, 0);
const reopenedBrowserSections = (await PptxDocument.open(await browserSectioned.write())).sections;
if (reopenedBrowserSections?.length !== 2 || reopenedBrowserSections[0].id !== browserData.id || reopenedBrowserSections[0].title !== 'Browser data' || reopenedBrowserSections[0].slideIds[0] !== browserSectionSlide.slideId || reopenedBrowserSections[1].id !== browserIntro.id || reopenedBrowserSections[1].title !== 'Browser edited' || reopenedBrowserSections[1].slideIds.length !== 0) throw new Error('Browser presentation sections failed');
const browserHiddenDeck = PptxDocument.create();
const browserHiddenSlide = browserHiddenDeck.addSlide();
browserHiddenSlide.hidden = true;
const reopenedBrowserHidden = await PptxDocument.open(await browserHiddenDeck.writeBlob());
if (reopenedBrowserHidden.slides[0]?.hidden !== true) throw new Error('Browser hidden slide failed');
const browserNotesDeck = PptxDocument.create();
const browserNotesSlide = browserNotesDeck.addSlide().addNotes('Browser notes');
if (browserNotesSlide.notes !== 'Browser notes') throw new Error('Browser notes immediate state failed');
const reopenedBrowserNotes = await PptxDocument.open(await browserNotesDeck.writeBlob());
if (reopenedBrowserNotes.slides[0]?.notes !== 'Browser notes') {
  throw new Error('Browser speaker notes failed');
}
PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
created.slideSize = { width: inches(10), height: inches(7.5) };
process.stdout.write(resolved);
`,
  );
  run(process.execPath, ['--conditions=browser', 'browser-smoke.mjs'], directory);

  await writeFile(
    join(directory, 'smoke.ts'),
    `import {
  CustomGeometryEvaluationError,
  ChartModel,
  chartWorkbookMatches,
  degrees,
  evaluateCustomGeometry,
  ImageModel,
  PLACEHOLDER_TYPES,
  PRESET_SHAPE_TYPES,
  PptxDocument,
  ShapeModel,
  SlideLayoutModel,
  SlideMasterModel,
  TableModel,
  inches,
  type AddCustomShapeOptions,
  type ChartDefinitionInput,
  type ChartDiagnostic,
  type ChartGroupInput,
  type ChartSeriesInput,
  type ChartType,
  type AddImageSourceOptions,
  type AddImageOptions,
  type AddSvgImageOptions,
  type AddShapeOptions,
  type AddTextOptions,
  type CustomGeometry,
  type CustomGeometryCommand,
  type CustomGeometryConnectionSite,
  type CustomGeometryEvaluationContext,
  type CustomGeometryEvaluationErrorCode,
  type CustomGeometryTextRectangle,
  type CustomGeometryFormula,
  type CustomGeometryGuide,
  type CustomGeometryHandle,
  type CustomGeometryPolarHandle,
  type CustomGeometryPath,
  type CustomGeometryPathFill,
  type CustomGeometryPoint,
  type CustomGeometryValue,
  type CustomGeometryXyHandle,
  type EvaluatedCustomGeometry,
  type EvaluatedCustomGeometryCommand,
  type EvaluatedCustomGeometryConnectionSite,
  type EvaluatedCustomGeometryGuide,
  type EvaluatedCustomGeometryHandle,
  type EvaluatedCustomGeometryPath,
  type EvaluatedCustomGeometryPoint,
  type EvaluatedCustomGeometryPolarHandle,
  type EvaluatedCustomGeometryTextRectangle,
  type EvaluatedCustomGeometryXyHandle,
  type Hyperlink,
  type ImageByteChunk,
  type ImageByteStream,
  type ImageContentType,
  type ImageCropRegion,
  type ImageInfo,
  type ImageSizing,
  type ImageSizingResult,
  type ImageSource,
  type AddMediaOptions,
  type MediaByteChunk,
  type MediaByteStream,
  type MediaKind,
  type MediaModel,
  type MediaPlaybackSettings,
  type MediaSource,
  type DefineSlideMasterOptions,
  type PlaceholderSelector,
  type PlaceholderType,
  type ReplaceMediaPosterOptions,
  type ReplaceMediaSourceOptions,
  type ResolvedImageSource,
  type ShapeArrows,
  type ShapeArrowType,
  type ShapeFill,
  type ShapeLine,
  type ShapeLineDash,
  type ShapeAdjustment,
  type ShapeShadow,
  type PresetShapeType,
  type RasterImageContentType,
  type SetSlideBackgroundImageOptions,
  type SimpleFill,
  type SlideBackground,
  type SlideBackgroundImage,
  type SvgImageContentType,
  type SvgImageInfo,
  type SlideModel,
  type SlideMasterMargin,
  type SlideMasterObject,
  type SlideNumber,
  type SlideNumberColor,
  type SlideNumberMarginInput,
  type SlideNumberMargins,
  type SlideNumberOptions,
  type SlideNumberTextStyle,
  type SlideNumberTextStyleOptions,
  type CustomSlideSize,
  type AddSectionOptions,
  type AddSlideOptions,
  type PresentationSection,
  type PresentationTheme,
  type PresentationThemeOptions,
  type ThemeFontSnapshot,
  type ThemeFontUpdate,
  type RichTextColor,
  type RichTextParagraph,
  type TextAlignment,
  type NumberingStyle,
  type ParagraphBullet,
  type ParagraphLineSpacing,
  type ParagraphSpacing,
  type ParagraphTabStop,
  type ParagraphTabStopAlignment,
  type RichTextUnderline,
  type RichTextUnderlineStyle,
  type RichTextStrikeStyle,
  type RichTextOutline,
  type RichTextGlow,
  type RichTextBaseline,
  type RichTextRunStyle,
  type TextBoxMarginInput,
  type TextBoxMargins,
  type TextBoxFit,
  type TextBoxTextDirection,
  type AddTableCell,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
  type TableCellTextDirection,
  type TableCellBorder,
  type TableCellBorderInput,
  type TableCellBorders,
  type TableCellBorderStyle,
  type TableCellFill,
  type TextBoxVerticalAlignment,
  GradientCodec,
  importPptxGenJS,
  transitions,
  animations,
  advancedCharts,
  smartArt,
  calculateImageSizing,
  inspectImage,
  inspectSvgImage,
  resolveImageSource,
} from '@jiayunxie/pptx';

const documentPromise: Promise<PptxDocument> = PptxDocument.open(new Uint8Array());
const createdDocument: PptxDocument = PptxDocument.create({ format: 'pptx', slideSize: 'wide' });
const typedPlaceholderType: PlaceholderType = PLACEHOLDER_TYPES[3];
const typedPlaceholderName: PlaceholderSelector = 'typed_title';
const typedPlaceholderIdentity: PlaceholderSelector = { type: 'pic', index: 102 };
const typedMasterObjects: readonly SlideMasterObject[] = [
  {
    kind: 'placeholder',
    text: 'Typed title',
    options: { name: 'typed_title', type: 'title', index: 101 },
  },
  {
    kind: 'placeholder',
    text: 'Typed picture',
    options: { name: 'typed_picture', type: 'pic', index: 102 },
  },
  {
    kind: 'placeholder',
    text: 'Typed chart',
    options: { name: 'typed_chart', type: typedPlaceholderType, index: 103 },
  },
];
const typedMasterDefinition: DefineSlideMasterOptions = {
  title: 'TYPED-MASTER',
  margin: inches(0.25),
  objects: typedMasterObjects,
};
const typedMasterDocument = PptxDocument.create();
const typedMasterWrite: Promise<Uint8Array> = typedMasterDocument
  .defineSlideMaster(typedMasterDefinition)
  .then(async (layout: SlideLayoutModel) => {
    const master: SlideMasterModel = typedMasterDocument.masters[0];
    const margin: Readonly<SlideMasterMargin> | undefined = layout.margin;
    const masterLayouts: readonly SlideLayoutModel[] = master.layouts;
    const layoutShapes = layout.shapes;
    const masterShapes = master.shapes;
    const slide = typedMasterDocument.addSlide({ masterName: layout.name });
    slide.addText('Typed title value', { placeholder: typedPlaceholderName });
    await typedMasterDocument.addImage(0, Uint8Array.of(137, 80, 78, 71), {
      contentType: 'image/png',
      placeholder: typedPlaceholderIdentity,
    });
    await typedMasterDocument.addChart(0, 'bar', [{
      name: 'Typed revenue',
      categories: ['Q1'],
      values: [1],
    }], { placeholder: 'typed_chart' });
    void [margin, masterLayouts, layoutShapes, masterShapes];
    return typedMasterDocument.write();
  });
const typedDefaultColor: RichTextColor = { kind: 'scheme', value: 'accent1' };
const typedDefaultColorSlide = createdDocument.addSlide();
typedDefaultColorSlide.color = typedDefaultColor;
const currentDefaultColor: Readonly<RichTextColor> | undefined = typedDefaultColorSlide.color;
// @ts-expect-error slide default colors require structured RichTextColor values
typedDefaultColorSlide.color = 'accent1';
void [currentDefaultColor, typedDefaultColorSlide];
const typedSlideNumberColor: SlideNumberColor = { kind: 'scheme', value: 'accent1' };
const typedSlideNumberMargin: SlideNumberMarginInput = [1, 2, 3, 4];
const typedSlideNumberMargins: SlideNumberMargins = { top: 1, left: 4 };
const typedSlideNumberStyleOptions: SlideNumberTextStyleOptions = {
  italic: true,
  color: typedSlideNumberColor,
  transparency: 20,
};
const typedSlideNumberOptions: SlideNumberOptions = {
  align: 'justify',
  rtl: true,
  margin: typedSlideNumberMargin,
  style: typedSlideNumberStyleOptions,
};
const typedSlideNumberStyle: SlideNumberTextStyle = {
  lang: 'en-US',
  bold: false,
  italic: true,
  color: typedSlideNumberColor,
};
const typedSlideNumber: SlideNumber = {
  x: 0,
  y: 0,
  width: 800_000,
  height: 300_000,
  align: 'justify',
  rtl: true,
  margin: typedSlideNumberMargins,
  style: typedSlideNumberStyle,
};
const typedSlideNumberDocument = PptxDocument.create({ firstSlideNumber: 0 });
typedSlideNumberDocument.addSlide().slideNumber = typedSlideNumberOptions;
void [typedSlideNumber, typedSlideNumberDocument];
const typedChartType: ChartType = 'bar';
const typedChartSeries: readonly ChartSeriesInput[] = [{
  name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
}];
const typedChartGroups: readonly ChartGroupInput[] = [{
  type: typedChartType,
  series: typedChartSeries,
}, {
  type: 'line',
  axis: 'secondary',
  series: [{ name: 'Margin', categories: ['Q1', 'Q2'], values: [25, 30] }],
}];
const typedChartDefinition: ChartDefinitionInput = { groups: typedChartGroups };
const typedChartPromise: Promise<ChartModel> = createdDocument.addChart(0, typedChartGroups);
const typedChartDiagnostics: Promise<readonly ChartDiagnostic[]> = typedChartPromise.then(
  (chart) => chart.diagnostics(),
);
const typedChartWorkbookCheck: Promise<boolean> = typedChartPromise.then((chart) =>
  chartWorkbookMatches(new Uint8Array(), chart.definition!, chart.xml));
// @ts-expect-error native chart types exclude PptxGenJS aliases
const invalidChartType: ChartType = 'column';
const invalidChartAxis: ChartGroupInput = {
  type: 'line',
  // @ts-expect-error chart groups support only primary or secondary axes
  axis: 'tertiary',
  series: typedChartSeries,
};
const invalidChartValues: ChartSeriesInput = {
  name: 'Invalid',
  // @ts-expect-error chart values must be numeric
  values: ['10'],
};
const typedSimpleBackground: SimpleFill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};
const typedImageBackground: SlideBackgroundImage = {
  kind: 'image',
  contentType: 'image/png',
  bytes: Uint8Array.of(137, 80, 78, 71),
};
const typedSlideBackground: SlideBackground = typedImageBackground;
const typedSlideBackgroundOptions: SetSlideBackgroundImageOptions = {
  contentType: 'image/png',
  signal: new AbortController().signal,
};
const typedBackgroundSlide: SlideModel = createdDocument.addSlide();
typedBackgroundSlide.background = typedSimpleBackground;
const typedBackgroundPromise: Promise<void> = createdDocument.setSlideBackgroundImage(
  createdDocument.slides.length - 1,
  typedImageBackground.bytes,
  typedSlideBackgroundOptions,
);
const typedRasterContentType: RasterImageContentType = 'image/png';
const typedRasterOptions: AddImageOptions = {
  contentType: typedRasterContentType,
  width: inches(2),
  height: inches(1),
};
const typedRasterImage: ImageModel = createdDocument.addSlide()
  .addImage(new Uint8Array([1]), typedRasterOptions);
typedRasterImage.setTransform({ x: inches(2) });
typedRasterImage.replaceData(new Uint8Array([2]), 'image/png');
// @ts-expect-error raster image content types exclude SVG
const invalidRasterSvg: RasterImageContentType = 'image/svg+xml';
// @ts-expect-error embedded raster image options require contentType
const invalidRasterMissingType: AddImageOptions = {};
// @ts-expect-error embedded raster creation excludes path loading
const invalidRasterPath: AddImageOptions = { contentType: 'image/png', path: './image.png' };
// @ts-expect-error embedded raster creation excludes data-URI loading
const invalidRasterData: AddImageOptions = { contentType: 'image/png', data: 'data:image/png;base64,AQ==' };
const typedSvgContentType: SvgImageContentType = 'image/svg+xml';
const typedImageContentType: ImageContentType = typedSvgContentType;
const typedSvgBytes = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"/>',
);
const typedSvgInfo: SvgImageInfo = inspectSvgImage(typedSvgBytes);
const typedImageInfo: ImageInfo = inspectImage(typedSvgBytes);
const typedCropRegion: ImageCropRegion = { x: 0, y: 0, width: 640, height: 360 };
const typedImageSizing: ImageSizing = {
  type: 'crop',
  width: inches(4),
  height: inches(3),
  source: typedCropRegion,
};
const typedImageSizingResult: Readonly<ImageSizingResult> = calculateImageSizing(
  typedImageInfo,
  typedImageSizing,
);
const typedImageSource: ImageSource = typedSvgBytes;
const typedImageChunk: ImageByteChunk = typedSvgBytes;
const typedImageStream: ImageByteStream = new ReadableStream<ImageByteChunk>({
  start(controller) {
    controller.enqueue(typedImageChunk);
    controller.close();
  },
});
const typedImageSourceOptions: AddImageSourceOptions = {
  contentType: typedImageContentType,
  fallback: typedImageSource,
  sizing: typedImageSizing,
};
const typedResolvedImage: Promise<ResolvedImageSource> = resolveImageSource(typedImageStream);
const typedSvgOptions: AddSvgImageOptions = { name: 'Typed SVG' };
const typedSvgSlide = createdDocument.addSlide();
const typedSvgImage: ImageModel = typedSvgSlide.addSvgImage(
  typedSvgBytes,
  new Uint8Array([1]),
  typedSvgOptions,
);
const typedHighLevelSvgImage: Promise<ImageModel> = createdDocument.addImage(
  createdDocument.slides.length - 1,
  typedImageSource,
  typedImageSourceOptions,
);
const typedMediaKind: MediaKind = 'audio';
const typedMediaChunk: MediaByteChunk = Uint8Array.of(1, 2, 3);
const typedMediaStream: MediaByteStream = new ReadableStream<MediaByteChunk>({
  start(controller) {
    controller.enqueue(typedMediaChunk);
    controller.close();
  },
});
const typedMediaSources: readonly MediaSource[] = [
  'data:audio/mpeg;base64,AQID',
  './typed-audio.mp3',
  Uint8Array.of(1),
  Uint8Array.of(2).buffer,
  new Blob([Uint8Array.of(3)], { type: 'audio/mpeg' }),
  typedMediaStream,
];
const typedPlayback: MediaPlaybackSettings = {
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.5,
};
const typedMediaOptions: AddMediaOptions = {
  name: 'Typed media',
  altText: 'Typed media description',
  contentType: 'audio/mpeg',
  fileName: 'typed-media.mp3',
  poster: 'data:image/png;base64,AQID',
  posterContentType: 'image/png',
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(4),
  ...typedPlayback,
  transcode: async (bytes, contentType, kind) => ({
    bytes,
    contentType,
    extension: kind === 'audio' ? '.mp3' : '.mp4',
  }),
};
const typedMediaPromise: Promise<MediaModel> = createdDocument.addAudio(
  0,
  typedMediaSources[0],
  typedMediaOptions,
);
const typedVideoPromise: Promise<MediaModel> = createdDocument.addVideo(
  0,
  new Blob([Uint8Array.of(4)], { type: 'video/mp4' }),
  { contentType: 'video/mp4', poster: typedMediaStream },
);
const typedReplaceMediaSourceOptions: ReplaceMediaSourceOptions = {
  contentType: 'audio/wav',
  fileName: 'typed-replacement.wav',
};
const typedReplaceMediaPosterOptions: ReplaceMediaPosterOptions = {
  contentType: 'image/gif',
  fileName: 'typed-poster.gif',
};
const typedMediaLifecycle: Promise<void> = typedMediaPromise.then(async (media) => {
  media.name = 'Typed media edited';
  media.altText = undefined;
  media.settings = typedPlayback;
  media.setTransform({ x: inches(2) });
  const sourceResult: MediaModel = await media.replaceSource(
    Uint8Array.of(5),
    typedReplaceMediaSourceOptions,
  );
  const posterResult: MediaModel = await media.replacePoster(
    Uint8Array.of(6),
    typedReplaceMediaPosterOptions,
  );
  await media.replacePoster();
  media.remove();
  createdDocument.slides[0].deleteMedia(sourceResult.shapeId);
  void posterResult;
});
// @ts-expect-error media kind accepts only audio or video
const invalidMediaKind: MediaKind = 'online';
// @ts-expect-error media name must be a string
const invalidMediaName: AddMediaOptions = { name: 1 };
// @ts-expect-error media poster must be a supported media source
const invalidMediaPoster: AddMediaOptions = { poster: {} };
// @ts-expect-error media playback accepts only click or auto
const invalidMediaPlayback: AddMediaOptions = { play: 'hover' };
const invalidMediaTranscode: AddMediaOptions = {
  // @ts-expect-error media transcoder bytes must be Uint8Array
  transcode: async () => ({ bytes: 'bad', contentType: 'audio/mpeg' }),
};
// @ts-expect-error media source must be a supported source
const invalidMediaSource: MediaSource = {};
// @ts-expect-error media source replacement excludes placement
const invalidMediaSourceReplacement: ReplaceMediaSourceOptions = { x: inches(1) };
// @ts-expect-error media poster replacement excludes transcoders
const invalidMediaPosterReplacement: ReplaceMediaPosterOptions = { transcode: async () => undefined };
// @ts-expect-error low-level raster image options exclude SVG
const invalidLowLevelSvgOptions: AddImageOptions = { contentType: 'image/svg+xml' };
const invalidSvgFallback: AddImageSourceOptions = {
  // @ts-expect-error SVG fallback must be a supported image source
  fallback: {},
};
const typedPreset: PresetShapeType = 'foldedCorner';
const typedCustomPoint: CustomGeometryPoint = { x: 1, y: 2 };
const typedCustomCommand: CustomGeometryCommand = {
  kind: 'quadraticBezierTo',
  control: typedCustomPoint,
  end: { x: 3, y: 4 },
};
const typedCustomFill: CustomGeometryPathFill = 'darken';
const typedCustomPath: CustomGeometryPath = {
  width: 100,
  height: 200,
  fill: typedCustomFill,
  commands: [typedCustomCommand],
};
const typedNumericConnectionSite: CustomGeometryConnectionSite = {
  angle: -5_400_000,
  position: { x: 25_000, y: 100_000 },
};
const typedTokenConnectionSite: CustomGeometryConnectionSite = {
  angle: 'adjAng',
  position: { x: 'hc', y: 't' },
};
const typedNumericTextRectangle: CustomGeometryTextRectangle = {
  left: 0,
  top: 12_500,
  right: 100_000,
  bottom: 87_500,
};
const typedTokenTextRectangle: CustomGeometryTextRectangle = {
  left: 'textLeft',
  top: 't',
  right: 'textRight',
  bottom: 'b',
};
const typedCustomGeometry: CustomGeometry = {
  connectionSites: [typedNumericConnectionSite, typedTokenConnectionSite],
  textRectangle: typedTokenTextRectangle,
  paths: [typedCustomPath],
};
const typedCustomValue: CustomGeometryValue = 'x1';
const typedUnaryFormula: CustomGeometryFormula = { operator: 'val', operands: [25_000] };
const typedBinaryFormula: CustomGeometryFormula = { operator: 'at2', operands: ['h', 'x1'] };
const typedTernaryFormula: CustomGeometryFormula = {
  operator: '*/',
  operands: ['w', 'adj1', 100_000],
};
const typedCustomGuide: CustomGeometryGuide = { name: 'x1', formula: typedTernaryFormula };
const typedFormulaGeometry: CustomGeometry = {
  adjustments: [{ name: 'adj1', formula: typedUnaryFormula }],
  guides: [typedCustomGuide, { name: 'a1', formula: typedBinaryFormula }],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: typedCustomValue, y: 't' } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 'a1',
        sweepAngle: 'cd2',
      },
    ],
  }],
};
const typedXyHandle: CustomGeometryXyHandle = {
  kind: 'xy',
  position: { x: 'adjX', y: 'adjY' },
  xGuide: 'adjX',
  minX: 0,
  maxX: 100_000,
};
const typedPolarHandle: CustomGeometryPolarHandle = {
  kind: 'polar',
  position: { x: 'hc', y: 'vc' },
  radiusGuide: 'adjR',
  minRadius: 0,
  maxRadius: 'ss',
  angleGuide: 'adjAng',
  minAngle: 0,
  maxAngle: 'cd',
};
const typedCustomHandles: readonly CustomGeometryHandle[] = [
  typedXyHandle,
  typedPolarHandle,
];
const typedHandleGeometry: CustomGeometry = {
  handles: typedCustomHandles,
  paths: [typedCustomPath],
};
const typedEvaluationContext: CustomGeometryEvaluationContext = { width: 200, height: 100 };
const typedEvaluatedGuide: EvaluatedCustomGeometryGuide = { name: 'x1', value: 50 };
const typedEvaluatedPoint: EvaluatedCustomGeometryPoint = { x: 50, y: 25 };
const typedEvaluatedXyHandle: EvaluatedCustomGeometryXyHandle = {
  kind: 'xy',
  position: typedEvaluatedPoint,
  xGuide: 'adjX',
  minX: 0,
  maxX: 200,
};
const typedEvaluatedPolarHandle: EvaluatedCustomGeometryPolarHandle = {
  kind: 'polar',
  position: typedEvaluatedPoint,
  radiusGuide: 'adjR',
  minRadius: 1,
  maxRadius: 100,
  angleGuide: 'adjAng',
  minAngle: 0,
  maxAngle: 5_400_000,
};
const typedEvaluatedHandles: readonly EvaluatedCustomGeometryHandle[] = [
  typedEvaluatedXyHandle,
  typedEvaluatedPolarHandle,
];
const typedEvaluatedConnectionSite: EvaluatedCustomGeometryConnectionSite = {
  position: typedEvaluatedPoint,
  angle: 5_400_000,
};
const typedEvaluatedTextRectangle: EvaluatedCustomGeometryTextRectangle = {
  left: 0,
  top: 0,
  right: 200,
  bottom: 100,
};
const typedEvaluatedCommand: EvaluatedCustomGeometryCommand = {
  kind: 'lineTo',
  point: typedEvaluatedPoint,
};
const typedEvaluatedPath: EvaluatedCustomGeometryPath = {
  width: 100,
  height: 100,
  commands: [typedEvaluatedCommand],
};
const typedEvaluatedGeometry: EvaluatedCustomGeometry = {
  context: typedEvaluationContext,
  guides: [typedEvaluatedGuide],
  handles: typedEvaluatedHandles,
  connectionSites: [typedEvaluatedConnectionSite],
  textRectangle: typedEvaluatedTextRectangle,
  paths: [typedEvaluatedPath],
};
const typedPureEvaluation: EvaluatedCustomGeometry = evaluateCustomGeometry(
  typedCustomGeometry,
  typedEvaluationContext,
);
const typedEvaluationShape: ShapeModel = createdDocument.addSlide()
  .addCustomShape(typedCustomGeometry);
const typedLiveEvaluation: EvaluatedCustomGeometry | undefined =
  typedEvaluationShape.evaluateCustomGeometry();
const typedEvaluationErrorCode: CustomGeometryEvaluationErrorCode = 'unknown-token';
const typedEvaluationError = new CustomGeometryEvaluationError(
  typedEvaluationErrorCode,
  'Typed evaluator error',
  'x1',
  'missing',
);
// @ts-expect-error Evaluation context requires height.
const invalidMissingEvaluationHeight: CustomGeometryEvaluationContext = { width: 1 };
const invalidStringEvaluationWidth: CustomGeometryEvaluationContext = {
  // @ts-expect-error Evaluation context width is numeric.
  width: '1',
  height: 1,
};
const invalidExtraEvaluationContext: CustomGeometryEvaluationContext = {
  width: 1,
  height: 1,
  // @ts-expect-error Evaluation context rejects extra fields.
  extra: true,
};
const invalidStringEvaluatedPoint: EvaluatedCustomGeometryPoint = {
  // @ts-expect-error Evaluated coordinates are numeric.
  x: 'x1',
  y: 0,
};
// @ts-expect-error Evaluation error codes use a closed union.
const invalidEvaluationErrorCode: CustomGeometryEvaluationErrorCode = 'invalid';
const invalidTokenEvaluatedPath: EvaluatedCustomGeometryPath = {
  width: 1,
  height: 1,
  commands: [{
    kind: 'moveTo',
    point: {
      // @ts-expect-error Evaluated paths cannot retain guide tokens.
      x: 'x1',
      y: 0,
    },
  }],
};
// @ts-expect-error Live shape evaluation accepts no explicit context argument.
const invalidShapeEvaluationArguments = typedEvaluationShape.evaluateCustomGeometry(typedEvaluationContext);
// @ts-expect-error Adjustment handles require a position.
const invalidMissingHandlePosition: CustomGeometryHandle = { kind: 'polar' };
const invalidXyPolarField: CustomGeometryXyHandle = {
  kind: 'xy',
  position: { x: 0, y: 0 },
  // @ts-expect-error XY handles do not expose polar guide fields.
  radiusGuide: 'adjR',
};
// @ts-expect-error Connection sites require an angle.
const invalidMissingConnectionAngle: CustomGeometryConnectionSite = { position: { x: 0, y: 0 } };
// @ts-expect-error Connection sites require a position.
const invalidMissingConnectionPosition: CustomGeometryConnectionSite = { angle: 0 };
// @ts-expect-error Connection sites reject extra fields.
const invalidExtraConnectionField: CustomGeometryConnectionSite = { angle: 0, position: { x: 0, y: 0 }, extra: true };
// @ts-expect-error Connection-site angles are numeric or token values.
const invalidConnectionAngleType: CustomGeometryConnectionSite = { angle: false, position: { x: 0, y: 0 } };
// @ts-expect-error Connection-site entries cannot be undefined.
const invalidUndefinedConnectionSite: CustomGeometryConnectionSite = undefined;
// @ts-expect-error Text rectangles require left.
const invalidMissingTextRectangleLeft: CustomGeometryTextRectangle = { top: 0, right: 1, bottom: 1 };
// @ts-expect-error Text rectangles require top.
const invalidMissingTextRectangleTop: CustomGeometryTextRectangle = { left: 0, right: 1, bottom: 1 };
// @ts-expect-error Text rectangles require right.
const invalidMissingTextRectangleRight: CustomGeometryTextRectangle = { left: 0, top: 0, bottom: 1 };
// @ts-expect-error Text rectangles require bottom.
const invalidMissingTextRectangleBottom: CustomGeometryTextRectangle = { left: 0, top: 0, right: 1 };
// @ts-expect-error Text rectangles reject extra fields.
const invalidExtraTextRectangleField: CustomGeometryTextRectangle = { left: 0, top: 0, right: 1, bottom: 1, extra: true };
// @ts-expect-error Text rectangle values are numeric or token values.
const invalidTextRectangleValue: CustomGeometryTextRectangle = { left: false, top: 0, right: 1, bottom: 1 };
// @ts-expect-error Text rectangles cannot be undefined.
const invalidUndefinedTextRectangle: CustomGeometryTextRectangle = undefined;
// @ts-expect-error Custom geometry formulas reject unknown operators.
const invalidCustomFormulaOperator: CustomGeometryFormula = { operator: 'unknown', operands: [1] };
// @ts-expect-error The val operator requires exactly one operand.
const invalidCustomFormulaArity: CustomGeometryFormula = { operator: 'val', operands: [1, 2] };
const typedCustomOptions: AddCustomShapeOptions = { name: 'Typed custom geometry' };
const typedCustomShape: ShapeModel = createdDocument.addSlide().addCustomShape(
  typedCustomGeometry,
  typedCustomOptions,
);
const typedCustomGeometryRead: CustomGeometry | undefined = typedCustomShape.customGeometry;
typedCustomShape.customGeometry = typedCustomGeometry;
const typedNoneShapeFill: ShapeFill = { kind: 'none' };
const typedSolidShapeFill: ShapeFill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 25,
};
const typedTextShapeNoneFill: ShapeFill = { kind: 'none' };
const typedTextShapeSrgbFill: ShapeFill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'A1B2C3' },
  transparency: 25,
};
const typedTextShapeSchemeFill: ShapeFill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent3' },
  transparency: 0,
};
const typedTextShapeLine: ShapeLine = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent4' },
  transparency: 25,
  width: 2.5,
  dash: 'dashDot',
};
const typedTextShapeArrows: ShapeArrows = {
  begin: 'none',
  end: 'triangle',
};
const typedTextShapeShadow: ShapeShadow = {
  kind: 'outer',
  color: { kind: 'scheme', value: 'accent5' },
  opacity: 0.4,
  blur: 2,
  angle: 45,
  distance: 3,
  rotateWithShape: true,
};
const typedTextShapeInnerShadow: ShapeShadow = {
  kind: 'inner',
  opacity: 0,
  blur: 0,
  angle: 0,
  distance: 0,
};
const typedTextShapeOptions: AddTextOptions = {
  name: 'Typed text shape fill, line, arrows, and shadow',
  fill: typedTextShapeSrgbFill,
  line: typedTextShapeLine,
  arrows: typedTextShapeArrows,
  shadow: typedTextShapeShadow,
};
const typedTextShapeSlide = createdDocument.addSlide();
const typedPlainTextShape: ShapeModel = typedTextShapeSlide.addText(
  'Typed plain text shape fill',
  typedTextShapeOptions,
);
const typedRichTextShape: ShapeModel = typedTextShapeSlide.addRichText([{
  runs: [{ text: 'Typed rich text shape fill' }],
}], {
  fill: typedTextShapeSchemeFill,
  line: typedTextShapeLine,
  arrows: { begin: 'diamond' },
  shadow: typedTextShapeInnerShadow,
});
const typedLayoutTextShape: ShapeModel = createdDocument.layouts[0].addText(
  'Typed layout text shape fill',
  {
    fill: typedTextShapeNoneFill,
    line: { kind: 'none' },
    arrows: { end: 'arrow' },
    shadow: typedTextShapeShadow,
  },
);
const typedMasterTextShape: ShapeModel = createdDocument.masters[0].addText(
  'Typed master text shape fill',
  {
    fill: typedTextShapeSchemeFill,
    line: typedTextShapeLine,
    arrows: { begin: 'stealth', end: 'oval' },
    shadow: typedTextShapeInnerShadow,
  },
);
const typedPlaceholderTextShape: ShapeModel = createdDocument.layouts[0].addPlaceholder(
  'Typed placeholder text shape fill',
  {
    name: 'typed_text_fill_placeholder',
    type: 'title',
    index: 190,
    fill: typedTextShapeSrgbFill,
    line: typedTextShapeLine,
    arrows: { begin: 'arrow', end: 'none' },
    shadow: typedTextShapeShadow,
  },
);
const typedDeclarativeTextFillObject: SlideMasterObject = {
  kind: 'text',
  text: 'Typed declarative text shape fill',
  options: {
    fill: typedTextShapeSchemeFill,
    line: typedTextShapeLine,
    arrows: typedTextShapeArrows,
    shadow: typedTextShapeInnerShadow,
  },
};
const typedTextShapeLineRead: ShapeLine | undefined = typedPlainTextShape.line;
const typedTextShapeArrowsRead: ShapeArrows | undefined = typedPlainTextShape.arrows;
const typedTextShapeShadowRead: ShapeShadow | undefined = typedPlainTextShape.shadow;
const typedShapeLineDash: ShapeLineDash = 'lgDashDotDot';
const typedNoneShapeLine: ShapeLine = { kind: 'none' };
const typedSolidShapeLine: ShapeLine = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent3' },
  transparency: 25,
  width: 2.5,
  dash: typedShapeLineDash,
};
const typedShapeArrowType: ShapeArrowType = 'triangle';
const typedShapeArrows: ShapeArrows = {
  begin: typedShapeArrowType,
  end: 'arrow',
};
const typedShapeAdjustments: readonly ShapeAdjustment[] = [
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
];
const typedOuterShapeShadow: ShapeShadow = {
  kind: 'outer',
  color: { kind: 'srgb', value: '123ABC' },
  opacity: 0,
  blur: 0,
  angle: 0,
  distance: 0,
  rotateWithShape: true,
};
const typedInnerShapeShadow: ShapeShadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.5,
  blur: 3,
  angle: 270,
  distance: 2,
};
const typedUrlHyperlink: Hyperlink = {
  url: 'https://example.com',
  tooltip: 'Typed URL',
};
const typedSlideHyperlink: Hyperlink = { slide: 2, tooltip: '' };
const typedTextHyperlinkOptions: AddTextOptions = {
  hyperlink: typedUrlHyperlink,
};
const typedTextHyperlinkShape: ShapeModel = typedTextShapeSlide.addText(
  'Typed text hyperlink',
  typedTextHyperlinkOptions,
);
const typedTextHyperlinkRead: Hyperlink | undefined = typedTextHyperlinkShape.hyperlink;
const invalidBothTextHyperlink: AddTextOptions = {
  // @ts-expect-error text shape hyperlinks require exactly one target
  hyperlink: { url: 'https://example.com', slide: 2 },
};
const invalidAliasTextHyperlink: AddTextOptions = {
  // @ts-expect-error text shape hyperlinks expose no target alias
  hyperlink: { target: 'https://example.com' },
};
const typedShapeOptions: AddShapeOptions = {
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(4),
  rotation: degrees(45),
  flipHorizontal: true,
  name: 'Typed shape',
  fill: typedSolidShapeFill,
  line: typedSolidShapeLine,
  arrows: typedShapeArrows,
  adjustments: typedShapeAdjustments,
  hyperlink: typedUrlHyperlink,
  shadow: typedOuterShapeShadow,
};
const typedShape: ShapeModel = createdDocument.addSlide().addShape(
  typedPreset,
  typedShapeOptions,
);
const typedPresetRead: PresetShapeType | undefined = typedShape.presetType;
typedShape.presetType = 'rect';
const typedShapeAdjustmentsRead: readonly ShapeAdjustment[] | undefined =
  typedShape.adjustments;
typedShape.adjustments = typedShapeAdjustments;
typedShape.adjustments = [];
const typedShapeFillRead: ShapeFill | undefined = typedShape.fill;
typedShape.fill = typedNoneShapeFill;
typedShape.fill = typedSolidShapeFill;
typedShape.fill = undefined;
const typedShapeLineRead: ShapeLine | undefined = typedShape.line;
typedShape.line = typedNoneShapeLine;
typedShape.line = typedSolidShapeLine;
typedShape.line = undefined;
const typedShapeArrowsRead: ShapeArrows | undefined = typedShape.arrows;
typedShape.arrows = { begin: 'diamond' };
typedShape.arrows = typedShapeArrows;
typedShape.arrows = undefined;
const typedShapeHyperlinkRead: Hyperlink | undefined = typedShape.hyperlink;
typedShape.hyperlink = typedSlideHyperlink;
typedShape.hyperlink = typedUrlHyperlink;
typedShape.hyperlink = undefined;
const typedShapeShadowRead: ShapeShadow | undefined = typedShape.shadow;
typedShape.shadow = typedInnerShapeShadow;
typedShape.shadow = typedOuterShapeShadow;
typedShape.shadow = undefined;
const typedPresetCatalog: readonly PresetShapeType[] = PRESET_SHAPE_TYPES;
// @ts-expect-error folderCorner is not a canonical OOXML preset
createdDocument.addSlide().addShape('folderCorner');
// @ts-expect-error custGeom belongs to the custom-geometry API
createdDocument.addSlide().addShape('custGeom');
// @ts-expect-error unknown preset-shape options are rejected
createdDocument.addSlide().addShape('rect', { color: 'FF0000' });
// @ts-expect-error transforms use numeric native units
createdDocument.addSlide().addShape('rect', { width: '3', rotation: '45' });
// @ts-expect-error shape adjustments require both name and value
const invalidMissingShapeAdjustmentValue: ShapeAdjustment = { name: 'adj' };
// @ts-expect-error shape adjustment values are numeric direct OOXML integers
const invalidShapeAdjustmentValue: ShapeAdjustment = { name: 'adj', value: '25000' };
const invalidShapeAdjustmentOptions: AddShapeOptions = {
  // @ts-expect-error shape adjustments are supplied as a list
  adjustments: { name: 'adj', value: 25_000 },
};
// @ts-expect-error gradient is not a simple shape fill kind
const invalidShapeFillKind: ShapeFill = { kind: 'gradient' };
// @ts-expect-error shape fill colors use srgb or scheme
const invalidShapeFillColor: ShapeFill = { kind: 'solid', color: { kind: 'rgb', value: 'FF0000' } };
// @ts-expect-error transparency is numeric
const invalidShapeFillTransparency: ShapeFill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: '50' };
const invalidPptxGenJSTextFill: AddTextOptions = {
  // @ts-expect-error PptxGenJS-style text fill objects are intentionally unsupported
  fill: { color: 'FF0000' },
};
const invalidTextFillKind: AddTextOptions = {
  // @ts-expect-error text shape fill kind must be none or solid
  fill: { kind: 'gradient' },
};
const invalidTextFillMissingColor: AddTextOptions = {
  // @ts-expect-error solid text shape fills require a color
  fill: { kind: 'solid' },
};
const invalidTextFillTransparency: AddTextOptions = {
  fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FF0000' },
    // @ts-expect-error text shape fill transparency is numeric
    transparency: '25',
  },
};
const invalidTextFillUnknownKey: AddTextOptions = {
  fill: {
    kind: 'none',
    // @ts-expect-error text shape none fills reject unknown fields
    extra: true,
  },
};
const invalidPptxGenJSTextLine: AddTextOptions = {
  // @ts-expect-error PptxGenJS-style text line objects are intentionally unsupported
  line: { color: 'FF0000', dashType: 'dash' },
};
const invalidTextLineKind: AddTextOptions = {
  // @ts-expect-error text shape line kind must be none or line
  line: { kind: 'solid' },
};
const invalidTextLineMissingColor: AddTextOptions = {
  // @ts-expect-error solid text shape lines require a color
  line: { kind: 'line' },
};
const invalidTextLineWidth: AddTextOptions = {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: 'FF0000' },
    // @ts-expect-error text shape line width is numeric points
    width: '2',
  },
};
const invalidTextLineDash: AddTextOptions = {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: 'FF0000' },
    // @ts-expect-error text shape line dash union is closed
    dash: 'dot',
  },
};
const invalidPptxGenJSTextArrows: AddTextOptions = {
  // @ts-expect-error PptxGenJS-style text arrow aliases are intentionally unsupported
  arrows: { beginArrowType: 'triangle' },
};
const invalidDeprecatedTextArrows: AddTextOptions = {
  // @ts-expect-error deprecated text arrow aliases are intentionally unsupported
  arrows: { lineHead: 'triangle' },
};
const invalidTextArrowToken: AddTextOptions = {
  // @ts-expect-error text arrow endpoint tokens use a closed union
  arrows: { end: 'bogus' },
};
const invalidEmptyTextArrow: AddTextOptions = {
  // @ts-expect-error empty text arrow endpoint tokens are rejected
  arrows: { begin: '' },
};
const invalidTextArrowUnknownKey: AddTextOptions = {
  arrows: {
    begin: 'triangle',
    // @ts-expect-error text arrow values reject unknown fields
    extra: true,
  },
};
const invalidPptxGenJSTextShadow: AddTextOptions = {
  // @ts-expect-error PptxGenJS-style text shadow aliases are intentionally unsupported
  shadow: { type: 'outer' },
};
const invalidNoneTextShadow: AddTextOptions = {
  // @ts-expect-error text shadow kind must be outer or inner
  shadow: { kind: 'none' },
};
const invalidTextShadowOffset: AddTextOptions = {
  shadow: {
    kind: 'outer',
    // @ts-expect-error PptxGenJS offset is intentionally unsupported
    offset: 4,
  },
};
const invalidTextInnerShadowRotate: AddTextOptions = {
  shadow: {
    kind: 'inner',
    // @ts-expect-error inner text shadow cannot rotate with the shape
    rotateWithShape: true,
  },
};
const invalidTextShadowOpacity: AddTextOptions = {
  shadow: {
    kind: 'outer',
    // @ts-expect-error text shadow opacity is numeric
    opacity: '0.5',
  },
};
// @ts-expect-error solid is not the native shape-line discriminator
const invalidShapeLineKind: ShapeLine = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
// @ts-expect-error shape line colors use srgb or scheme
const invalidShapeLineColor: ShapeLine = { kind: 'line', color: { kind: 'rgb', value: 'FF0000' } };
// @ts-expect-error shape line transparency is numeric
const invalidShapeLineTransparency: ShapeLine = { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, transparency: '50' };
// @ts-expect-error shape line width is numeric points
const invalidShapeLineWidth: ShapeLine = { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: '2' };
// @ts-expect-error shape line dash union is closed
const invalidShapeLineDash: ShapeLineDash = 'dot';
// @ts-expect-error PptxGenJS dashType is not a native alias
const invalidShapeLineAlias: ShapeLine = { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, dashType: 'dash' };
// @ts-expect-error shape arrow type union is closed
const invalidShapeArrowType: ShapeArrowType = 'open';
// @ts-expect-error shape arrow values use the closed canonical token union
const invalidShapeArrowValue: ShapeArrows = { begin: '' };
// @ts-expect-error PptxGenJS beginArrowType is not a native alias
const invalidShapeArrowAlias: ShapeArrows = { beginArrowType: 'arrow' };
// @ts-expect-error shape hyperlink requires exactly one target
const invalidMissingHyperlink: Hyperlink = {};
// @ts-expect-error shape hyperlink targets are mutually exclusive
const invalidBothHyperlink: Hyperlink = { url: 'https://example.com', slide: 2 };
// @ts-expect-error shape hyperlink URL must be a string
const invalidUrlHyperlink: Hyperlink = { url: 42 };
// @ts-expect-error shape hyperlink slide must be numeric
const invalidSlideHyperlink: Hyperlink = { slide: '2' };
// @ts-expect-error shape hyperlink tooltip must be a string
const invalidTooltipHyperlink: Hyperlink = { slide: 2, tooltip: 7 };
// @ts-expect-error shape hyperlink public value has no relationship ID escape hatch
const invalidUnknownHyperlink: Hyperlink = { url: 'https://example.com', _rId: 'rId9' };
// @ts-expect-error shape shadow requires an outer or inner kind
const invalidMissingShapeShadowKind: ShapeShadow = {};
// @ts-expect-error none is represented by undefined, not a shadow kind
const invalidNoneShapeShadowKind: ShapeShadow = { kind: 'none' };
// @ts-expect-error rotateWithShape is outer-only
const invalidInnerShapeShadowRotate: ShapeShadow = { kind: 'inner', rotateWithShape: false };
// @ts-expect-error PptxGenJS offset is not a native alias
const invalidShapeShadowOffset: ShapeShadow = { kind: 'outer', offset: 4 };
// @ts-expect-error PptxGenJS type is not a native alias
const invalidShapeShadowType: ShapeShadow = { kind: 'outer', type: 'outer' };
// @ts-expect-error unknown shape shadow fields are rejected
const invalidUnknownShapeShadow: ShapeShadow = { kind: 'outer', size: 2 };
// @ts-expect-error shape shadow opacity is numeric
const invalidShapeShadowFieldType: ShapeShadow = { kind: 'outer', opacity: '0.5' };
const addSectionOptions: AddSectionOptions = { title: 'Typed', order: 0 };
const typedSection: PresentationSection = createdDocument.addSection(addSectionOptions);
const addSlideOptions: AddSlideOptions = { sectionTitle: typedSection.title };
createdDocument.addSlide(addSlideOptions);
const sectionSnapshot: readonly PresentationSection[] | undefined = createdDocument.sections;
createdDocument.renameSection(typedSection.id, 'Renamed');
createdDocument.moveSection(typedSection.id, 0);
createdDocument.assignSlideToSection(0, typedSection.id);
createdDocument.deleteSection(typedSection.id);
const typedVisibilitySlide = createdDocument.addSlide();
const hiddenSnapshot: boolean | undefined = typedVisibilitySlide.hidden;
typedVisibilitySlide.hidden = true;
typedVisibilitySlide.hidden = false;
const typedNotesSlide = createdDocument.addSlide();
const notesSnapshot: string | undefined = typedNotesSlide.notes;
typedNotesSlide.notes = 'Typed notes';
typedNotesSlide.notes = '';
typedNotesSlide.notes = undefined;
const returnedNotesSlide: SlideModel = typedNotesSlide.addNotes('Returned');
const globalRtl: PptxDocument = PptxDocument.create({ rtlMode: true });
const globalRtlSnapshot: boolean | undefined = globalRtl.rtlMode;
globalRtl.rtlMode = false;
globalRtl.rtlMode = undefined;
const titledDocument: PptxDocument = PptxDocument.create({ title: 'Typed title' });
const titleSnapshot: string | undefined = titledDocument.title;
titledDocument.title = 'Edited typed title';
titledDocument.title = '';
titledDocument.title = undefined;
const authoredDocument: PptxDocument = PptxDocument.create({ author: 'Typed author' });
const authorSnapshot: string | undefined = authoredDocument.author;
authoredDocument.author = 'Edited typed author';
authoredDocument.author = '';
authoredDocument.author = undefined;
const lastModifiedDocument: PptxDocument = PptxDocument.create({
  lastModifiedBy: 'Typed editor',
});
const lastModifiedSnapshot: string | undefined = lastModifiedDocument.lastModifiedBy;
lastModifiedDocument.lastModifiedBy = 'Edited typed editor';
lastModifiedDocument.lastModifiedBy = '';
lastModifiedDocument.lastModifiedBy = undefined;
const createdAtDocument: PptxDocument = PptxDocument.create({
  createdAt: '2026-07-30T00:00:00Z',
});
const createdAtSnapshot: string | undefined = createdAtDocument.createdAt;
createdAtDocument.createdAt = '2024-02-29T12:34:56.123+05:30';
createdAtDocument.createdAt = undefined;
const modifiedAtDocument: PptxDocument = PptxDocument.create({
  modifiedAt: '2026-07-30T01:02:03Z',
});
const modifiedAtSnapshot: string | undefined = modifiedAtDocument.modifiedAt;
modifiedAtDocument.modifiedAt = '2024-03-01T01:02:03.456+08:00';
modifiedAtDocument.modifiedAt = undefined;
const subjectDocument: PptxDocument = PptxDocument.create({ subject: 'Typed subject' });
const subjectSnapshot: string | undefined = subjectDocument.subject;
subjectDocument.subject = 'Edited typed subject';
subjectDocument.subject = '';
subjectDocument.subject = undefined;
const revisionDocument: PptxDocument = PptxDocument.create({ revision: '007' });
const revisionSnapshot: string | undefined = revisionDocument.revision;
revisionDocument.revision = '42';
revisionDocument.revision = undefined;
const companyDocument: PptxDocument = PptxDocument.create({ company: 'Typed company' });
const companySnapshot: string | undefined = companyDocument.company;
companyDocument.company = 'Edited typed company';
companyDocument.company = '';
companyDocument.company = undefined;
const themeOptions: PresentationThemeOptions = {
  headFontFace: 'Aptos Display',
};
const themedDocument = PptxDocument.create({ theme: themeOptions });
const themeSnapshot: PresentationTheme | undefined = themedDocument.theme;
const fontSnapshot: ThemeFontSnapshot | undefined =
  themedDocument.masterLayoutTheme.presentationTheme?.fonts;
const fontUpdate: ThemeFontUpdate = { minorLatin: 'Aptos' };
themedDocument.masterLayoutTheme.presentationTheme?.setFonts(fontUpdate);
const customSlideSize: CustomSlideSize = { width: inches(11.7), height: inches(8.3) };
const customDocument: PptxDocument = PptxDocument.create({ slideSize: customSlideSize });
customDocument.slideSize = { width: inches(10), height: inches(7.5) };
const alignment: TextAlignment = 'center';
const bullet: ParagraphBullet = { kind: 'bullet', character: '◆', indent: 18 };
const numbering: NumberingStyle = 'romanUcPeriod';
const lineSpacing: ParagraphLineSpacing = { kind: 'multiple', factor: 1.5 };
const spacing: ParagraphSpacing = { before: 4, after: 6, line: lineSpacing };
const tabAlignment: ParagraphTabStopAlignment = 'right';
const tabStops: readonly ParagraphTabStop[] = [{ position: 1.25 }, { position: 2.5, alignment: tabAlignment }];
const underlineStyle: RichTextUnderlineStyle = 'dotDashHeavy';
const underline: RichTextUnderline = { style: underlineStyle, color: { kind: 'srgb', value: 'FF0000' } };
const strike: RichTextStrikeStyle = 'dblStrike';
const outline: RichTextOutline = { color: { kind: 'scheme', value: 'accent1' }, size: 1.5 };
const glow: RichTextGlow = { color: { kind: 'scheme', value: 'accent2' }, opacity: 0.5, size: 8 };
const baseline: RichTextBaseline = 'superscript';
const characterStyle: RichTextRunStyle = { baseline, characterSpacing: 2.5, lang: 'de-DE' };
const margin: TextBoxMarginInput = [4, 8, 4, 8];
const verticalAlignment: TextBoxVerticalAlignment = 'top';
const direction: TextBoxTextDirection = 'vert270';
const cellDirection: TableCellTextDirection = 'vert270';
const fit: TextBoxFit = 'shrink';
const cellFit: TextBoxFit = 'shrink';
const cellAlignment: TextBoxVerticalAlignment = 'middle';
const cellHorizontalAlignment: TextAlignment = 'center';
const tableHorizontalAlignment: TextAlignment = 'center';
const cellMargins: TextBoxMarginInput = { top: 4, left: 8 };
const cellBorderStyle: TableCellBorderStyle = 'dash';
const cellBorder: TableCellBorder = { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: cellBorderStyle };
const cellBorderInput: TableCellBorderInput = [cellBorder, { kind: 'none' }, undefined, cellBorder];
const cellFill: TableCellFill = { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 };
const createdText = createdDocument.addSlide().addText('Typed\\ntext', { align: alignment, fit, valign: verticalAlignment, vert: direction, wrap: true, bullet, level: 2, margin, spacing, tabStops });
const creationBorder: TableCellBorderInput = [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'none' },
  undefined,
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 2 },
];
const creationMargin: TextBoxMarginInput = { top: 4, left: 8 };
const creationOptions: AddTableCellOptions = {
  align: cellHorizontalAlignment,
  border: creationBorder,
  fill: cellFill,
  fit: cellFit,
  margin: creationMargin,
  textDirection: cellDirection,
  valign: cellAlignment,
};
const objectCell: AddTableCell = { text: 'Revenue', options: creationOptions };
const tableRows: readonly (readonly AddTableCellInput[])[] = [['Region', objectCell], [{ text: 'East' }, { text: '' }]];
const tableOptions: AddTableOptions = { align: tableHorizontalAlignment, name: 'Typed table', x: inches(1), columnWidths: [inches(1), inches(3)], rowHeights: [inches(0.5), inches(1.5)], border: cellBorderInput, fill: cellFill, margin: cellMargins, textDirection: cellDirection, valign: cellAlignment };
const typedTable: TableModel = createdDocument.slides[0].addTable(tableRows, tableOptions);
const widthSnapshot: readonly number[] | undefined = typedTable.columnWidths;
const heightSnapshot: readonly number[] | undefined = typedTable.rowHeights;
typedTable.setColumnWidths(inches(2));
typedTable.setColumnWidths([inches(1.5), inches(2.5)]);
typedTable.setRowHeights(inches(1));
typedTable.setRowHeights([0, inches(1.5)]);
const table = createdDocument.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const snapshotDirection: TableCellTextDirection | undefined =
  table?.rows[0]?.cells[0]?.textDirection;
const snapshotFit: TextBoxFit | undefined = table?.rows[0]?.cells[0]?.textFit;
const snapshotAlignment: TextBoxVerticalAlignment | undefined =
  table?.rows[0]?.cells[0]?.verticalAlignment;
const snapshotHorizontalAlignment: TextAlignment | undefined =
  table?.rows[0]?.cells[0]?.horizontalAlignment;
const snapshotCellMargins: TextBoxMargins | undefined = table?.rows[0]?.cells[0]?.margins;
const snapshotCellBorders: TableCellBorders | undefined = table?.rows[0]?.cells[0]?.borders;
const snapshotCellFill: TableCellFill | undefined = table?.rows[0]?.cells[0]?.fill;
table?.setCellTextDirection(0, 0, cellDirection);
table?.setCellTextDirection(0, 0, undefined);
table?.setCellTextFit(0, 0, cellFit);
table?.setCellTextFit(0, 0, 'none');
table?.setCellTextFit(0, 0, undefined);
table?.setCellVerticalAlignment(0, 0, cellAlignment);
table?.setCellVerticalAlignment(0, 0, undefined);
table?.setCellHorizontalAlignment(0, 0, cellHorizontalAlignment);
table?.setCellHorizontalAlignment(0, 0, undefined);
table?.setCellMargins(0, 0, cellMargins);
table?.setCellMargins(0, 0, [3.6, 7.2, 10.8, 14.4]);
table?.setCellMargins(0, 0, undefined);
table?.setCellBorders(0, 0, cellBorder);
table?.setCellBorders(0, 0, cellBorderInput);
table?.setCellBorders(0, 0, { top: cellBorder, left: { kind: 'none' } });
table?.setCellBorders(0, 0, undefined);
table?.setCellFill(0, 0, cellFill);
table?.setCellFill(0, 0, { kind: 'none' });
table?.setCellFill(0, 0, undefined);
const marginSnapshot: TextBoxMargins | undefined = createdText.textMargins;
const wrapSnapshot: boolean | undefined = createdText.textWrap;
const directionSnapshot: TextBoxTextDirection | undefined = createdText.textDirection;
const fitSnapshot: TextBoxFit | undefined = createdText.textFit;
createdText.textMargins = { top: 3, left: 6 };
createdText.verticalAlignment = 'bottom';
createdText.verticalAlignment = undefined;
createdText.textWrap = false;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
createdText.textDirection = undefined;
createdText.textFit = 'resize';
createdText.textFit = 'none';
createdText.textFit = undefined;
createdText.text = 'Updated\\n\\ntyped text';
const paragraphs: readonly RichTextParagraph[] = [{ align: 'justify', bullet: { kind: 'number', style: numbering, startAt: 3 }, level: 3, spacing: { line: { kind: 'exact', points: 20 } }, tabStops, runs: [{ text: 'Typed rich', style: { ...characterStyle, fontSize: 12.5, bold: true, color: { kind: 'scheme', value: 'tx1' }, glow, highlight: { kind: 'srgb', value: 'FFFF00' }, outline, underline, strike } }] }];
const richText = createdDocument.slides[0].addRichText(paragraphs, { lang: 'fr-CA' });
richText.richText = paragraphs;
const transparentParagraphs: readonly RichTextParagraph[] = [{
  runs: [
    { text: 'Opaque', style: { transparency: 0 } },
    { text: 'Quarter', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 25 } },
    { text: 'Theme', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 } },
  ],
}];
createdDocument.addSlide().addRichText(transparentParagraphs);
const rtlParagraphs: readonly RichTextParagraph[] = [
  { rtl: true, runs: [{ text: 'RTL' }] },
  { rtl: false, runs: [{ text: 'LTR' }] },
];
createdDocument.addSlide().addRichText(rtlParagraphs, { rtlMode: true });
const paragraphMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginLeft: 12, runs: [{ text: 'Override' }] },
  { marginLeft: false, runs: [{ text: 'Suppressed' }] },
];
createdDocument.addSlide().addRichText(paragraphMargins, { paragraphMarginLeft: 24 });
const paragraphRightMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginRight: 12, runs: [{ text: 'Override' }] },
  { marginRight: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, marginRight: 18, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphRightMargins, { paragraphMarginRight: 24 });
const paragraphIndents: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { indent: 18, runs: [{ text: 'First-line' }] },
  { indent: -18, runs: [{ text: 'Hanging' }] },
  { indent: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, indent: false, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphIndents, { paragraphIndent: 24 });
const gradientConstructor: typeof GradientCodec = GradientCodec;
const adapter: typeof importPptxGenJS = importPptxGenJS;
const transition: transitions.SlideTransition = { effect: 'fade' };
const animationConstructor: typeof animations.AnimationTimingCodec = animations.AnimationTimingCodec;
const chartConstructor: typeof advancedCharts.AdvancedChartCodec = advancedCharts.AdvancedChartCodec;
const smartArtConstructor: typeof smartArt.SmartArtDiagramCodec = smartArt.SmartArtDiagramCodec;
documentPromise.then((document) => {
  transitions.installTransitionPlugin(document);
  animations.installAnimationPlugin(document);
  advancedCharts.installAdvancedChartPlugin(document);
  smartArt.installSmartArtPlugin(document);
});
void [typedNotesSlide, notesSnapshot, returnedNotesSlide];
void [typedPreset, typedNoneShapeFill, typedSolidShapeFill,
  typedTextShapeNoneFill, typedTextShapeSrgbFill, typedTextShapeSchemeFill,
  typedTextShapeOptions, typedTextShapeSlide, typedPlainTextShape, typedRichTextShape,
  typedLayoutTextShape, typedMasterTextShape, typedPlaceholderTextShape,
  typedDeclarativeTextFillObject, typedTextShapeLine, typedTextShapeLineRead,
  typedTextShapeShadow, typedTextShapeInnerShadow, typedTextShapeShadowRead,
  invalidPptxGenJSTextFill, invalidTextFillKind, invalidTextFillMissingColor,
  invalidTextFillTransparency, invalidTextFillUnknownKey, invalidPptxGenJSTextLine,
  invalidTextLineKind, invalidTextLineMissingColor, invalidTextLineWidth,
  invalidTextLineDash, invalidPptxGenJSTextShadow, invalidNoneTextShadow,
  invalidTextShadowOffset, invalidTextInnerShadowRotate, invalidTextShadowOpacity,
  typedShapeOptions, typedShape,
  typedCustomPoint, typedCustomCommand, typedCustomFill, typedCustomPath, typedCustomGeometry,
  typedCustomValue, typedUnaryFormula, typedBinaryFormula, typedTernaryFormula,
  typedCustomGuide, typedFormulaGeometry, typedXyHandle, typedPolarHandle,
  typedCustomHandles, typedHandleGeometry, invalidMissingHandlePosition,
  invalidXyPolarField, typedNumericConnectionSite, typedTokenConnectionSite,
  typedNumericTextRectangle, typedTokenTextRectangle,
  invalidMissingConnectionAngle, invalidMissingConnectionPosition,
  invalidExtraConnectionField, invalidConnectionAngleType, invalidUndefinedConnectionSite,
  invalidMissingTextRectangleLeft, invalidMissingTextRectangleTop,
  invalidMissingTextRectangleRight, invalidMissingTextRectangleBottom,
  invalidExtraTextRectangleField, invalidTextRectangleValue, invalidUndefinedTextRectangle,
  invalidCustomFormulaOperator, invalidCustomFormulaArity,
  typedCustomOptions, typedCustomShape, typedCustomGeometryRead,
  typedEvaluationContext, typedEvaluatedGuide, typedEvaluatedPoint,
  typedEvaluatedXyHandle, typedEvaluatedPolarHandle, typedEvaluatedHandles,
  typedEvaluatedConnectionSite, typedEvaluatedTextRectangle, typedEvaluatedCommand,
  typedEvaluatedPath, typedEvaluatedGeometry, typedPureEvaluation, typedEvaluationShape,
  typedLiveEvaluation,
  typedEvaluationErrorCode, typedEvaluationError, invalidMissingEvaluationHeight,
  invalidStringEvaluationWidth, invalidExtraEvaluationContext, invalidStringEvaluatedPoint,
  invalidEvaluationErrorCode, invalidTokenEvaluatedPath, invalidShapeEvaluationArguments,
  typedPresetRead, typedShapeAdjustments, typedShapeAdjustmentsRead,
  invalidMissingShapeAdjustmentValue, invalidShapeAdjustmentValue,
  invalidShapeAdjustmentOptions, typedShapeFillRead, typedPresetCatalog, invalidShapeFillKind,
  invalidShapeFillColor, invalidShapeFillTransparency, typedShapeLineDash,
  typedNoneShapeLine, typedSolidShapeLine, typedShapeLineRead, invalidShapeLineKind,
  invalidShapeLineColor, invalidShapeLineTransparency, invalidShapeLineWidth,
  invalidShapeLineDash, invalidShapeLineAlias, typedUrlHyperlink, typedSlideHyperlink,
  typedTextHyperlinkOptions, typedTextHyperlinkShape, typedTextHyperlinkRead,
  invalidBothTextHyperlink, invalidAliasTextHyperlink,
  typedShapeHyperlinkRead, invalidMissingHyperlink, invalidBothHyperlink,
  invalidUrlHyperlink, invalidSlideHyperlink, invalidTooltipHyperlink,
  invalidUnknownHyperlink, typedOuterShapeShadow, typedInnerShapeShadow,
  typedShapeShadowRead, invalidMissingShapeShadowKind, invalidNoneShapeShadowKind,
  invalidInnerShapeShadowRotate, invalidShapeShadowOffset, invalidShapeShadowType,
  invalidUnknownShapeShadow, invalidShapeShadowFieldType];
void [documentPromise, createdDocument, typedMasterWrite, typedChartDefinition, typedChartPromise,
  typedChartDiagnostics, typedChartWorkbookCheck, invalidChartType, invalidChartAxis,
  invalidChartValues, typedSimpleBackground, typedImageBackground, typedSlideBackground,
  typedSlideBackgroundOptions, typedBackgroundSlide, typedBackgroundPromise,
  typedRasterContentType, typedRasterOptions, typedRasterImage,
  invalidRasterSvg, invalidRasterMissingType, invalidRasterPath, invalidRasterData, typedSvgContentType, typedImageContentType, typedSvgInfo, typedImageInfo, typedCropRegion, typedImageSizing, typedImageSizingResult, typedImageSource, typedImageChunk, typedImageStream, typedImageSourceOptions, typedResolvedImage, typedSvgOptions, typedSvgImage, typedHighLevelSvgImage, typedMediaKind, typedMediaChunk, typedMediaStream, typedMediaSources, typedPlayback, typedMediaOptions, typedMediaPromise, typedVideoPromise, typedReplaceMediaSourceOptions, typedReplaceMediaPosterOptions, typedMediaLifecycle, invalidMediaKind, invalidMediaName, invalidMediaPoster, invalidMediaPlayback, invalidMediaTranscode, invalidMediaSource, invalidMediaSourceReplacement, invalidMediaPosterReplacement, invalidLowLevelSvgOptions, invalidSvgFallback, addSectionOptions, typedSection, addSlideOptions, sectionSnapshot, typedVisibilitySlide, hiddenSnapshot, globalRtl, globalRtlSnapshot, titledDocument, titleSnapshot, authoredDocument, authorSnapshot, lastModifiedDocument, lastModifiedSnapshot, createdAtDocument, createdAtSnapshot, modifiedAtDocument, modifiedAtSnapshot, subjectDocument, subjectSnapshot, revisionDocument, revisionSnapshot, companyDocument, companySnapshot, themedDocument, themeSnapshot, fontSnapshot, fontUpdate, customDocument, createdText, creationBorder, creationMargin, creationOptions, objectCell, tableRows, tableOptions, typedTable, widthSnapshot, heightSnapshot, table, snapshotDirection, snapshotFit, snapshotAlignment, snapshotHorizontalAlignment, tableHorizontalAlignment, snapshotCellMargins, snapshotCellBorders, snapshotCellFill, cellDirection, cellFit, cellAlignment, cellHorizontalAlignment, cellMargins, cellBorderStyle, cellBorder, cellBorderInput, cellFill, marginSnapshot, wrapSnapshot, directionSnapshot, fitSnapshot, fit, direction, verticalAlignment, richText, transparentParagraphs, rtlParagraphs, paragraphMargins, paragraphRightMargins, paragraphIndents, gradientConstructor, adapter, transition, animationConstructor, chartConstructor, smartArtConstructor];
`,
  );
  run(
    process.execPath,
    [
      join(repositoryRoot, 'node_modules/typescript/bin/tsc'),
      'smoke.ts',
      '--noEmit',
      '--strict',
      '--target',
      'ES2022',
      '--module',
      'NodeNext',
      '--moduleResolution',
      'NodeNext',
      '--typeRoots',
      join(repositoryRoot, 'node_modules/@types'),
      '--types',
      'node',
    ],
    directory,
  );

  const bin = join(directory, 'node_modules', '.bin', process.platform === 'win32' ? 'pptx-inspect.cmd' : 'pptx-inspect');
  const cliResult = run(bin, ['--json', 'doctor'], directory);
  const doctor = JSON.parse(cliResult.stdout);
  if (!doctor.ok || doctor.data?.version !== '0.1.0') throw new Error(`CLI smoke failed: ${cliResult.stdout}`);
  const masterLayoutDeckPath = join(directory, 'master-layout-smoke.pptx');
  const masterLayoutInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', masterLayoutDeckPath],
    directory,
  );
  const masterLayoutInspected = JSON.parse(masterLayoutInspectResult.stdout);
  const masterLayoutContentTypes = masterLayoutInspected.data?.contentTypes ?? {};
  if (!masterLayoutInspected.ok ||
      masterLayoutContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1 ||
      masterLayoutContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 2 ||
      masterLayoutContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 2 ||
      masterLayoutContentTypes[
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      ] !== 2 ||
      masterLayoutContentTypes[
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ] !== 2) {
    throw new Error(`CLI master-layout inspect failed: ${masterLayoutInspectResult.stdout}`);
  }
  const masterLayoutValidateResult = run(
    bin,
    ['--json', 'package', 'validate', masterLayoutDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const masterLayoutValidated = JSON.parse(masterLayoutValidateResult.stdout);
  if (!masterLayoutValidated.ok || !masterLayoutValidated.data?.valid ||
      masterLayoutValidated.data.errorCount !== 0 ||
      masterLayoutValidated.data.warningCount !== 0) {
    throw new Error(`CLI master-layout validation failed: ${masterLayoutValidateResult.stdout}`);
  }
  const masterLayoutSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', masterLayoutDeckPath],
    directory,
  );
  const masterLayoutSlides = JSON.parse(masterLayoutSlidesResult.stdout);
  if (!masterLayoutSlides.ok || masterLayoutSlides.data?.length !== 2 ||
      masterLayoutSlides.data[0]?.shapeCount !== 7 ||
      masterLayoutSlides.data[1]?.shapeCount !== 0) {
    throw new Error(`CLI master-layout slide listing failed: ${masterLayoutSlidesResult.stdout}`);
  }
  const masterLayoutPart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', masterLayoutDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const masterLayoutLayoutXml = masterLayoutPart('/ppt/slideLayouts/slideLayout2.xml');
  const masterLayoutMasterXml = masterLayoutPart('/ppt/slideMasters/slideMaster1.xml');
  const masterLayoutSlideXml = masterLayoutPart(masterLayoutSlides.data[0].partUri);
  if (!masterLayoutLayoutXml.includes('name="PACKED-MASTER-LAYOUT"') ||
      !placeholderTypes.every((type) => masterLayoutLayoutXml.includes('type="' + type + '"')) ||
      !masterLayoutMasterXml.includes('sldNum="1"') ||
      !masterLayoutSlideXml.includes('Packed master/layout support') ||
      !placeholderTypes.every((type) => masterLayoutSlideXml.includes('type="' + type + '"'))) {
    throw new Error('CLI master-layout part inspection failed');
  }
  const masterLayoutDiffResult = run(
    bin,
    [
      '--json', 'package', 'diff', masterLayoutDeckPath,
      join(directory, 'slide-number-smoke.pptx'),
    ],
    directory,
  );
  const masterLayoutDiff = JSON.parse(masterLayoutDiffResult.stdout);
  if (!masterLayoutDiff.ok ||
      (masterLayoutDiff.data?.added?.length ?? 0) +
        (masterLayoutDiff.data?.removed?.length ?? 0) +
        (masterLayoutDiff.data?.changed?.length ?? 0) === 0) {
    throw new Error(`CLI master-layout diff failed: ${masterLayoutDiffResult.stdout}`);
  }
  const svgDeckPath = join(directory, 'svg-smoke.pptx');
  const inspectResult = run(bin, ['--json', 'package', 'inspect', svgDeckPath], directory);
  const inspected = JSON.parse(inspectResult.stdout);
  if (!inspected.ok || inspected.data?.contentTypes?.['image/svg+xml'] < 1 ||
      inspected.data?.contentTypes?.['image/png'] < 1) {
    throw new Error(`CLI SVG inspect failed: ${inspectResult.stdout}`);
  }
  const validateResult = run(
    bin,
    ['--json', 'package', 'validate', svgDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const validated = JSON.parse(validateResult.stdout);
  if (!validated.ok || !validated.data?.valid || validated.data.errorCount !== 0 ||
      validated.data.warningCount !== 0) {
    throw new Error(`CLI SVG validation failed: ${validateResult.stdout}`);
  }
  const mediaDeckPath = join(directory, 'media-smoke.pptx');
  const mediaInspectResult = run(bin, ['--json', 'package', 'inspect', mediaDeckPath], directory);
  const mediaInspected = JSON.parse(mediaInspectResult.stdout);
  const mediaContentTypes = mediaInspected.data?.contentTypes ?? {};
  if (!mediaInspected.ok || mediaContentTypes['audio/mpeg'] !== 2 ||
      mediaContentTypes['audio/wav'] !== 1 || mediaContentTypes['audio/mp4'] !== 1 ||
      mediaContentTypes['video/mp4'] !== 1 || mediaContentTypes['video/quicktime'] !== 1 ||
      mediaContentTypes['image/png'] < 1 || mediaContentTypes['image/jpeg'] < 1) {
    throw new Error(`CLI media inspect failed: ${mediaInspectResult.stdout}`);
  }
  const mediaValidateResult = run(
    bin,
    ['--json', 'package', 'validate', mediaDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const mediaValidated = JSON.parse(mediaValidateResult.stdout);
  if (!mediaValidated.ok || !mediaValidated.data?.valid ||
      mediaValidated.data.errorCount !== 0 || mediaValidated.data.warningCount !== 0) {
    throw new Error(`CLI media validation failed: ${mediaValidateResult.stdout}`);
  }
  const stableMediaDeckPath = join(directory, 'stable-media-smoke.pptx');
  const stableMediaInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', stableMediaDeckPath],
    directory,
  );
  const stableMediaInspected = JSON.parse(stableMediaInspectResult.stdout);
  const stableMediaContentTypes = stableMediaInspected.data?.contentTypes ?? {};
  if (!stableMediaInspected.ok || stableMediaContentTypes['audio/wav'] !== 1 ||
      stableMediaContentTypes['image/png'] !== 1) {
    throw new Error(`CLI stable media inspect failed: ${stableMediaInspectResult.stdout}`);
  }
  const stableMediaValidateResult = run(
    bin,
    ['--json', 'package', 'validate', stableMediaDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const stableMediaValidated = JSON.parse(stableMediaValidateResult.stdout);
  if (!stableMediaValidated.ok || !stableMediaValidated.data?.valid ||
      stableMediaValidated.data.errorCount !== 0 || stableMediaValidated.data.warningCount !== 0) {
    throw new Error(`CLI stable media validation failed: ${stableMediaValidateResult.stdout}`);
  }
  const nativeChartDeckPath = join(directory, 'native-charts-smoke.pptx');
  const nativeChartInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', nativeChartDeckPath],
    directory,
  );
  const nativeChartInspected = JSON.parse(nativeChartInspectResult.stdout);
  const nativeChartContentTypes = nativeChartInspected.data?.contentTypes ?? {};
  if (!nativeChartInspected.ok ||
      nativeChartContentTypes['application/vnd.openxmlformats-officedocument.drawingml.chart+xml'] !== 10 ||
      nativeChartContentTypes['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'] !== 10) {
    throw new Error(`CLI native chart inspect failed: ${nativeChartInspectResult.stdout}`);
  }
  const nativeChartValidateResult = run(
    bin,
    ['--json', 'package', 'validate', nativeChartDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const nativeChartValidated = JSON.parse(nativeChartValidateResult.stdout);
  if (!nativeChartValidated.ok || !nativeChartValidated.data?.valid ||
      nativeChartValidated.data.errorCount !== 0 || nativeChartValidated.data.warningCount !== 0) {
    throw new Error(`CLI native chart validation failed: ${nativeChartValidateResult.stdout}`);
  }
  const nativeChartSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', nativeChartDeckPath],
    directory,
  );
  const nativeChartSlides = JSON.parse(nativeChartSlidesResult.stdout);
  if (!nativeChartSlides.ok || nativeChartSlides.data?.length !== 11 ||
      nativeChartSlides.data[10]?.shapeCount !== 1) {
    throw new Error(`CLI native chart slide listing failed: ${nativeChartSlidesResult.stdout}`);
  }
  const nativeChartPartResult = run(
    bin,
    ['--json', 'part', 'read', nativeChartDeckPath, '/ppt/charts/chart1.xml'],
    directory,
  );
  const nativeChartPart = JSON.parse(nativeChartPartResult.stdout);
  if (!nativeChartPart.ok || !nativeChartPart.data?.content?.includes('<c:chartSpace')) {
    throw new Error(`CLI native chart part read failed: ${nativeChartPartResult.stdout}`);
  }
  const slideBackgroundDeckPath = join(directory, 'slide-background-smoke.pptx');
  const slideBackgroundInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', slideBackgroundDeckPath],
    directory,
  );
  const slideBackgroundInspected = JSON.parse(slideBackgroundInspectResult.stdout);
  if (!slideBackgroundInspected.ok ||
      slideBackgroundInspected.data?.contentTypes?.['image/png'] !== 1 ||
      slideBackgroundInspected.data?.contentTypes?.['image/jpeg'] !== 1) {
    throw new Error(`CLI slide background inspect failed: ${slideBackgroundInspectResult.stdout}`);
  }
  const slideBackgroundValidateResult = run(
    bin,
    ['--json', 'package', 'validate', slideBackgroundDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const slideBackgroundValidated = JSON.parse(slideBackgroundValidateResult.stdout);
  if (!slideBackgroundValidated.ok || !slideBackgroundValidated.data?.valid ||
      slideBackgroundValidated.data.errorCount !== 0 ||
      slideBackgroundValidated.data.warningCount !== 0) {
    throw new Error(`CLI slide background validation failed: ${slideBackgroundValidateResult.stdout}`);
  }
  const slideNumberDeckPath = join(directory, 'slide-number-smoke.pptx');
  const slideNumberInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', slideNumberDeckPath],
    directory,
  );
  const slideNumberInspected = JSON.parse(slideNumberInspectResult.stdout);
  if (!slideNumberInspected.ok ||
      slideNumberInspected.data?.contentTypes?.[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 2) {
    throw new Error(`CLI slide-number inspect failed: ${slideNumberInspectResult.stdout}`);
  }
  const slideNumberValidateResult = run(
    bin,
    ['--json', 'package', 'validate', slideNumberDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const slideNumberValidated = JSON.parse(slideNumberValidateResult.stdout);
  if (!slideNumberValidated.ok || !slideNumberValidated.data?.valid ||
      slideNumberValidated.data.errorCount !== 0 ||
      slideNumberValidated.data.warningCount !== 0) {
    throw new Error(`CLI slide-number validation failed: ${slideNumberValidateResult.stdout}`);
  }
  const slideNumberSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', slideNumberDeckPath],
    directory,
  );
  const slideNumberSlides = JSON.parse(slideNumberSlidesResult.stdout);
  if (!slideNumberSlides.ok || slideNumberSlides.data?.length !== 2 ||
      slideNumberSlides.data.some(({ title, shapeCount }) => title !== '' || shapeCount !== 1)) {
    throw new Error(`CLI slide-number slide listing failed: ${slideNumberSlidesResult.stdout}`);
  }
  const slideNumberPart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', slideNumberDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const slideNumberSlideXml = slideNumberPart(slideNumberSlides.data[0].partUri);
  const slideNumberLayoutXml = slideNumberPart('/ppt/slideLayouts/slideLayout1.xml');
  const slideNumberMasterXml = slideNumberPart('/ppt/slideMasters/slideMaster1.xml');
  if (!slideNumberSlideXml.includes('type="slidenum"') ||
      !slideNumberSlideXml.includes('<a:t>5</a:t>') ||
      !slideNumberLayoutXml.includes('<a:t>‹#›</a:t>') ||
      !slideNumberMasterXml.includes('<a:t>‹#›</a:t>') ||
      !slideNumberMasterXml.includes('sldNum="1"')) {
    throw new Error('CLI slide-number part inspection failed');
  }
  const slideDefaultColorDeckPath = join(directory, 'slide-default-color-smoke.pptx');
  const slideDefaultColorInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', slideDefaultColorDeckPath],
    directory,
  );
  const slideDefaultColorInspected = JSON.parse(slideDefaultColorInspectResult.stdout);
  if (!slideDefaultColorInspected.ok ||
      slideDefaultColorInspected.data?.contentTypes?.[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 2) {
    throw new Error(`CLI slide-default-color inspect failed: ${slideDefaultColorInspectResult.stdout}`);
  }
  const slideDefaultColorValidateResult = run(
    bin,
    ['--json', 'package', 'validate', slideDefaultColorDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const slideDefaultColorValidated = JSON.parse(slideDefaultColorValidateResult.stdout);
  if (!slideDefaultColorValidated.ok || !slideDefaultColorValidated.data?.valid ||
      slideDefaultColorValidated.data.errorCount !== 0 ||
      slideDefaultColorValidated.data.warningCount !== 0) {
    throw new Error(`CLI slide-default-color validation failed: ${slideDefaultColorValidateResult.stdout}`);
  }
  const slideDefaultColorSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', slideDefaultColorDeckPath],
    directory,
  );
  const slideDefaultColorSlides = JSON.parse(slideDefaultColorSlidesResult.stdout);
  if (!slideDefaultColorSlides.ok || slideDefaultColorSlides.data?.length !== 2 ||
      slideDefaultColorSlides.data[0]?.shapeCount !== 2 ||
      slideDefaultColorSlides.data[1]?.shapeCount !== 3) {
    throw new Error(`CLI slide-default-color slide listing failed: ${slideDefaultColorSlidesResult.stdout}`);
  }
  const slideDefaultColorPartResult = run(
    bin,
    ['--json', 'part', 'read', slideDefaultColorDeckPath, slideDefaultColorSlides.data[0].partUri],
    directory,
  );
  const slideDefaultColorPart = JSON.parse(slideDefaultColorPartResult.stdout);
  const slideDefaultColorXml = slideDefaultColorPart.data?.content ?? '';
  if (!slideDefaultColorPart.ok ||
      !slideDefaultColorXml.includes('<a:srgbClr val="FF3399"/>') ||
      !slideDefaultColorXml.includes('<a:schemeClr val="accent1"/>') ||
      !slideDefaultColorXml.includes(
        '<a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr>',
      )) {
    throw new Error(`CLI slide-default-color part inspection failed: ${slideDefaultColorPartResult.stdout}`);
  }
  const textShapeFillDeckPath = join(directory, 'text-shape-fill-smoke.pptx');
  const textShapeFillInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', textShapeFillDeckPath],
    directory,
  );
  const textShapeFillInspected = JSON.parse(textShapeFillInspectResult.stdout);
  const textShapeFillContentTypes = textShapeFillInspected.data?.contentTypes ?? {};
  if (!textShapeFillInspected.ok ||
      textShapeFillContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 3 ||
      textShapeFillContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 2 ||
      textShapeFillContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1) {
    throw new Error(`CLI text-shape-fill inspect failed: ${textShapeFillInspectResult.stdout}`);
  }
  const textShapeFillValidateResult = run(
    bin,
    ['--json', 'package', 'validate', textShapeFillDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const textShapeFillValidated = JSON.parse(textShapeFillValidateResult.stdout);
  if (!textShapeFillValidated.ok || !textShapeFillValidated.data?.valid ||
      textShapeFillValidated.data.errorCount !== 0 ||
      textShapeFillValidated.data.warningCount !== 0) {
    throw new Error(`CLI text-shape-fill validation failed: ${textShapeFillValidateResult.stdout}`);
  }
  const textShapeFillSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', textShapeFillDeckPath],
    directory,
  );
  const textShapeFillSlides = JSON.parse(textShapeFillSlidesResult.stdout);
  if (!textShapeFillSlides.ok || textShapeFillSlides.data?.length !== 3 ||
      textShapeFillSlides.data[0]?.shapeCount !== 3 ||
      textShapeFillSlides.data[1]?.shapeCount !== 1 ||
      textShapeFillSlides.data[2]?.shapeCount !== 3) {
    throw new Error(`CLI text-shape-fill slide listing failed: ${textShapeFillSlidesResult.stdout}`);
  }
  const textShapeFillPart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', textShapeFillDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const textShapeFillSourceXml = textShapeFillPart(textShapeFillSlides.data[0].partUri);
  const textShapeFillDuplicateXml = textShapeFillPart(textShapeFillSlides.data[2].partUri);
  const textShapeFillLayoutXml = textShapeFillPart('/ppt/slideLayouts/slideLayout1.xml');
  const textShapeFillMasterXml = textShapeFillPart('/ppt/slideMasters/slideMaster1.xml');
  if (!textShapeFillSourceXml.includes(
        '<a:srgbClr val="AB12CD"><a:alpha val="75000"/></a:srgbClr>',
      ) ||
      !textShapeFillSourceXml.includes(
        '<a:schemeClr val="accent2"><a:alpha val="100000"/></a:schemeClr>',
      ) ||
      !textShapeFillDuplicateXml.includes('<a:noFill/>') ||
      !textShapeFillLayoutXml.includes('name="packed_layout_text_fill"') ||
      !textShapeFillLayoutXml.includes('<a:noFill/>') ||
      !textShapeFillMasterXml.includes(
        '<a:schemeClr val="accent6"><a:alpha val="0"/></a:schemeClr>',
      )) {
    throw new Error('CLI text-shape-fill part inspection failed');
  }
  const textShapeLineDeckPath = join(directory, 'text-shape-line-smoke.pptx');
  const textShapeLineInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', textShapeLineDeckPath],
    directory,
  );
  const textShapeLineInspected = JSON.parse(textShapeLineInspectResult.stdout);
  const textShapeLineContentTypes = textShapeLineInspected.data?.contentTypes ?? {};
  if (!textShapeLineInspected.ok ||
      textShapeLineContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 3 ||
      textShapeLineContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 2 ||
      textShapeLineContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1) {
    throw new Error(`CLI text-shape-line inspect failed: ${textShapeLineInspectResult.stdout}`);
  }
  const textShapeLineValidateResult = run(
    bin,
    ['--json', 'package', 'validate', textShapeLineDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const textShapeLineValidated = JSON.parse(textShapeLineValidateResult.stdout);
  if (!textShapeLineValidated.ok || !textShapeLineValidated.data?.valid ||
      textShapeLineValidated.data.errorCount !== 0 ||
      textShapeLineValidated.data.warningCount !== 0) {
    throw new Error(`CLI text-shape-line validation failed: ${textShapeLineValidateResult.stdout}`);
  }
  const textShapeLineSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', textShapeLineDeckPath],
    directory,
  );
  const textShapeLineSlides = JSON.parse(textShapeLineSlidesResult.stdout);
  if (!textShapeLineSlides.ok || textShapeLineSlides.data?.length !== 3 ||
      textShapeLineSlides.data[0]?.shapeCount !== 3 ||
      textShapeLineSlides.data[1]?.shapeCount !== 1 ||
      textShapeLineSlides.data[2]?.shapeCount !== 3) {
    throw new Error(`CLI text-shape-line slide listing failed: ${textShapeLineSlidesResult.stdout}`);
  }
  const textShapeLinePart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', textShapeLineDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const textShapeLineSourceXml = textShapeLinePart(textShapeLineSlides.data[0].partUri);
  const textShapeLineDuplicateXml = textShapeLinePart(textShapeLineSlides.data[2].partUri);
  const textShapeLineLayoutXml = textShapeLinePart('/ppt/slideLayouts/slideLayout1.xml');
  const textShapeLineMasterXml = textShapeLinePart('/ppt/slideMasters/slideMaster1.xml');
  if (!textShapeLineSourceXml.includes(
        '<a:ln w="31750"><a:solidFill><a:srgbClr val="AB12CD">' +
        '<a:alpha val="75000"/></a:srgbClr></a:solidFill>' +
        '<a:prstDash val="dashDot"/></a:ln>',
      ) ||
      !textShapeLineSourceXml.includes(
        '<a:ln w="12700"><a:solidFill><a:schemeClr val="accent2">' +
        '<a:alpha val="100000"/></a:schemeClr></a:solidFill>' +
        '<a:prstDash val="solid"/></a:ln>',
      ) ||
      !textShapeLineDuplicateXml.includes('<a:ln><a:noFill/></a:ln>') ||
      !textShapeLineDuplicateXml.includes('<a:ln></a:ln>') ||
      !textShapeLineLayoutXml.includes('name="packed_layout_text_line"') ||
      !textShapeLineLayoutXml.includes('<a:ln><a:noFill/></a:ln>') ||
      !textShapeLineMasterXml.includes(
        '<a:ln w="0"><a:solidFill><a:schemeClr val="accent6">' +
        '<a:alpha val="0"/></a:schemeClr></a:solidFill>' +
        '<a:prstDash val="sysDot"/></a:ln>',
      )) {
    throw new Error('CLI text-shape-line part inspection failed');
  }
  const textShapeArrowDeckPath = join(directory, 'text-shape-arrows-smoke.pptx');
  const textShapeArrowInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', textShapeArrowDeckPath],
    directory,
  );
  const textShapeArrowInspected = JSON.parse(textShapeArrowInspectResult.stdout);
  const textShapeArrowContentTypes = textShapeArrowInspected.data?.contentTypes ?? {};
  if (!textShapeArrowInspected.ok ||
      textShapeArrowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 3 ||
      textShapeArrowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 2 ||
      textShapeArrowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1) {
    throw new Error(`CLI text-shape-arrow inspect failed: ${textShapeArrowInspectResult.stdout}`);
  }
  const textShapeArrowValidateResult = run(
    bin,
    ['--json', 'package', 'validate', textShapeArrowDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const textShapeArrowValidated = JSON.parse(textShapeArrowValidateResult.stdout);
  if (!textShapeArrowValidated.ok || !textShapeArrowValidated.data?.valid ||
      textShapeArrowValidated.data.errorCount !== 0 ||
      textShapeArrowValidated.data.warningCount !== 0) {
    throw new Error(`CLI text-shape-arrow validation failed: ${textShapeArrowValidateResult.stdout}`);
  }
  const textShapeArrowSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', textShapeArrowDeckPath],
    directory,
  );
  const textShapeArrowSlides = JSON.parse(textShapeArrowSlidesResult.stdout);
  if (!textShapeArrowSlides.ok || textShapeArrowSlides.data?.length !== 3 ||
      textShapeArrowSlides.data[0]?.shapeCount !== 3 ||
      textShapeArrowSlides.data[1]?.shapeCount !== 1 ||
      textShapeArrowSlides.data[2]?.shapeCount !== 3) {
    throw new Error(`CLI text-shape-arrow slide listing failed: ${textShapeArrowSlidesResult.stdout}`);
  }
  const textShapeArrowPart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', textShapeArrowDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const textShapeArrowSourceXml = textShapeArrowPart(textShapeArrowSlides.data[0].partUri);
  const textShapeArrowDuplicateXml = textShapeArrowPart(textShapeArrowSlides.data[2].partUri);
  const textShapeArrowLayoutXml = textShapeArrowPart('/ppt/slideLayouts/slideLayout1.xml');
  const textShapeArrowMasterXml = textShapeArrowPart('/ppt/slideMasters/slideMaster1.xml');
  if (!textShapeArrowSourceXml.includes(
        '<a:ln w="25400"><a:solidFill><a:schemeClr val="accent2"/>' +
        '</a:solidFill><a:prstDash val="dashDot"/>' +
        '<a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
      ) ||
      !textShapeArrowSourceXml.includes(
        '<a:ln><a:noFill/><a:headEnd type="diamond"/></a:ln>',
      ) ||
      !textShapeArrowSourceXml.includes(
        '<a:ln><a:noFill/><a:tailEnd type="oval"/></a:ln>',
      ) ||
      !textShapeArrowDuplicateXml.includes(
        '<a:ln><a:headEnd type="oval"/></a:ln>',
      ) ||
      !textShapeArrowDuplicateXml.includes('<a:ln><a:noFill/></a:ln>') ||
      !textShapeArrowLayoutXml.includes('name="packed_layout_text_arrow"') ||
      !textShapeArrowLayoutXml.includes('<a:headEnd type="none"/>') ||
      !textShapeArrowLayoutXml.includes(
        '<a:headEnd type="stealth"/><a:tailEnd type="none"/>',
      ) ||
      !textShapeArrowMasterXml.includes('<a:tailEnd type="triangle"/>')) {
    throw new Error('CLI text-shape-arrow part inspection failed');
  }
  const textShapeShadowDeckPath = join(directory, 'text-shape-shadows-smoke.pptx');
  const textShapeShadowInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', textShapeShadowDeckPath],
    directory,
  );
  const textShapeShadowInspected = JSON.parse(textShapeShadowInspectResult.stdout);
  const textShapeShadowContentTypes = textShapeShadowInspected.data?.contentTypes ?? {};
  if (!textShapeShadowInspected.ok ||
      textShapeShadowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 3 ||
      textShapeShadowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 2 ||
      textShapeShadowContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1) {
    throw new Error(`CLI text-shape-shadow inspect failed: ${textShapeShadowInspectResult.stdout}`);
  }
  const textShapeShadowValidateResult = run(
    bin,
    ['--json', 'package', 'validate', textShapeShadowDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const textShapeShadowValidated = JSON.parse(textShapeShadowValidateResult.stdout);
  if (!textShapeShadowValidated.ok || !textShapeShadowValidated.data?.valid ||
      textShapeShadowValidated.data.errorCount !== 0 ||
      textShapeShadowValidated.data.warningCount !== 0) {
    throw new Error(
      `CLI text-shape-shadow validation failed: ${textShapeShadowValidateResult.stdout}`,
    );
  }
  const textShapeShadowSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', textShapeShadowDeckPath],
    directory,
  );
  const textShapeShadowSlides = JSON.parse(textShapeShadowSlidesResult.stdout);
  if (!textShapeShadowSlides.ok || textShapeShadowSlides.data?.length !== 3 ||
      textShapeShadowSlides.data[0]?.shapeCount !== 3 ||
      textShapeShadowSlides.data[1]?.shapeCount !== 1 ||
      textShapeShadowSlides.data[2]?.shapeCount !== 3) {
    throw new Error(
      `CLI text-shape-shadow slide listing failed: ${textShapeShadowSlidesResult.stdout}`,
    );
  }
  const textShapeShadowPart = (uri) => JSON.parse(run(
    bin,
    ['--json', 'part', 'read', textShapeShadowDeckPath, uri],
    directory,
  ).stdout).data?.content ?? '';
  const textShapeShadowSourceXml = textShapeShadowPart(textShapeShadowSlides.data[0].partUri);
  const textShapeShadowDuplicateXml = textShapeShadowPart(
    textShapeShadowSlides.data[2].partUri,
  );
  const textShapeShadowLayoutXml = textShapeShadowPart('/ppt/slideLayouts/slideLayout1.xml');
  const textShapeShadowMasterXml = textShapeShadowPart('/ppt/slideMasters/slideMaster1.xml');
  if (!textShapeShadowSourceXml.includes(
        '<a:ln w="25400"><a:solidFill><a:schemeClr val="accent2"/>' +
        '</a:solidFill><a:prstDash val="dashDot"/>' +
        '<a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>' +
        '<a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" ' +
        'algn="bl" rotWithShape="1" blurRad="25400" dist="38100" dir="2700000">' +
        '<a:schemeClr val="accent4"><a:alpha val="40000"/></a:schemeClr>' +
        '</a:outerShdw></a:effectLst>',
      ) ||
      !textShapeShadowSourceXml.includes(
        '<a:innerShdw blurRad="0" dist="0" dir="0">' +
        '<a:srgbClr val="667788"><a:alpha val="0"/></a:srgbClr></a:innerShdw>',
      ) ||
      !textShapeShadowSourceXml.includes(
        '<a:innerShdw blurRad="12700" dist="25400" dir="5400000">' +
        '<a:schemeClr val="accent3"><a:alpha val="50000"/></a:schemeClr>' +
        '</a:innerShdw>',
      ) ||
      !textShapeShadowDuplicateXml.includes(
        '<a:ln></a:ln><a:effectLst><a:outerShdw sx="100000" sy="100000"',
      ) ||
      !textShapeShadowDuplicateXml.includes('<a:effectLst></a:effectLst>') ||
      !textShapeShadowLayoutXml.includes('name="packed_layout_text_shadow"') ||
      !textShapeShadowLayoutXml.includes(
        '<a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr>',
      ) ||
      !textShapeShadowLayoutXml.includes('name="packed_title_shadow"') ||
      !textShapeShadowMasterXml.includes(
        '<a:innerShdw blurRad="0" dist="0" dir="0">' +
        '<a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:innerShdw>',
      )) {
    throw new Error('CLI text-shape-shadow part inspection failed');
  }
  const textShapeHyperlinkDeckPath = join(directory, 'text-shape-hyperlinks-smoke.pptx');
  const textShapeHyperlinkInspectResult = run(
    bin,
    ['--json', 'package', 'inspect', textShapeHyperlinkDeckPath],
    directory,
  );
  const textShapeHyperlinkInspected = JSON.parse(textShapeHyperlinkInspectResult.stdout);
  const textShapeHyperlinkContentTypes = textShapeHyperlinkInspected.data?.contentTypes ?? {};
  if (!textShapeHyperlinkInspected.ok ||
      textShapeHyperlinkContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
      ] !== 4 ||
      textShapeHyperlinkContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
      ] !== 3 ||
      textShapeHyperlinkContentTypes[
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
      ] !== 1) {
    throw new Error(
      `CLI text-shape-hyperlink inspect failed: ${textShapeHyperlinkInspectResult.stdout}`,
    );
  }
  const textShapeHyperlinkValidateResult = run(
    bin,
    ['--json', 'package', 'validate', textShapeHyperlinkDeckPath, '--profile', 'powerpoint-2010'],
    directory,
  );
  const textShapeHyperlinkValidated = JSON.parse(textShapeHyperlinkValidateResult.stdout);
  if (!textShapeHyperlinkValidated.ok || !textShapeHyperlinkValidated.data?.valid ||
      textShapeHyperlinkValidated.data.errorCount !== 0 ||
      textShapeHyperlinkValidated.data.warningCount < 1 ||
      !textShapeHyperlinkValidated.data.diagnostics.every(
        ({ severity, code }) =>
          severity !== 'error' &&
          (severity !== 'warning' || code === 'OPC_EXTERNAL_RELATIONSHIP'),
      )) {
    throw new Error(
      `CLI text-shape-hyperlink validation failed: ${textShapeHyperlinkValidateResult.stdout}`,
    );
  }
  const textShapeHyperlinkSlidesResult = run(
    bin,
    ['--json', 'slides', 'list', textShapeHyperlinkDeckPath],
    directory,
  );
  const textShapeHyperlinkSlides = JSON.parse(textShapeHyperlinkSlidesResult.stdout);
  if (!textShapeHyperlinkSlides.ok || textShapeHyperlinkSlides.data?.length !== 4 ||
      textShapeHyperlinkSlides.data.map(({ shapeCount }) => shapeCount).join(',') !== '4,0,1,4') {
    throw new Error(
      `CLI text-shape-hyperlink slide listing failed: ${textShapeHyperlinkSlidesResult.stdout}`,
    );
  }
  const textShapeHyperlinkPartResult = run(
    bin,
    [
      '--json', 'part', 'read', textShapeHyperlinkDeckPath,
      textShapeHyperlinkSlides.data[0].partUri,
    ],
    directory,
  );
  const textShapeHyperlinkPart = JSON.parse(textShapeHyperlinkPartResult.stdout);
  const textShapeHyperlinkSourceXml = textShapeHyperlinkPart.data?.content ?? '';
  const textShapeHyperlinkShapeXml = (name) => {
    const nameOffset = textShapeHyperlinkSourceXml.indexOf('name="' + name + '"');
    const shapeStart = textShapeHyperlinkSourceXml.lastIndexOf('<p:sp', nameOffset);
    const shapeEnd = textShapeHyperlinkSourceXml.indexOf('</p:sp>', nameOffset);
    return nameOffset < 0 || shapeStart < 0 || shapeEnd < 0
      ? ''
      : textShapeHyperlinkSourceXml.slice(shapeStart, shapeEnd + '</p:sp>'.length);
  };
  const textShapeHyperlinkClickIds = (xml) => xml.split('<a:hlinkClick').slice(1).map(
    (fragment) => fragment.split('r:id="')[1]?.split('"')[0],
  );
  const textShapeHyperlinkPlainXml = textShapeHyperlinkShapeXml(
    'packed_plain_text_hyperlink',
  );
  const textShapeHyperlinkRichXml = textShapeHyperlinkShapeXml(
    'packed_rich_text_hyperlink',
  );
  const textShapeHyperlinkPlainIds = textShapeHyperlinkClickIds(
    textShapeHyperlinkPlainXml,
  );
  const textShapeHyperlinkRichIds = textShapeHyperlinkClickIds(textShapeHyperlinkRichXml);
  if (!textShapeHyperlinkPart.ok || textShapeHyperlinkPlainIds.length !== 3 ||
      new Set(textShapeHyperlinkPlainIds).size !== 1 ||
      !textShapeHyperlinkPlainXml.includes('tooltip="Packed &amp; linked"') ||
      textShapeHyperlinkRichIds.length !== 3 ||
      new Set(textShapeHyperlinkRichIds).size !== 1 ||
      !textShapeHyperlinkRichXml.split('<a:hlinkClick').slice(1).every(
        (fragment) => fragment.includes('action="ppaction://hlinksldjump"'),
      )) {
    throw new Error(`CLI text-shape-hyperlink part read failed: ${textShapeHyperlinkPartResult.stdout}`);
  }
  const internalTextShapeHyperlinkDeckPath = join(
    directory,
    'text-shape-hyperlinks-internal-smoke.pptx',
  );
  const internalTextShapeHyperlinkValidateResult = run(
    bin,
    [
      '--json', 'package', 'validate', internalTextShapeHyperlinkDeckPath,
      '--profile', 'powerpoint-2010',
    ],
    directory,
  );
  const internalTextShapeHyperlinkValidated = JSON.parse(
    internalTextShapeHyperlinkValidateResult.stdout,
  );
  if (!internalTextShapeHyperlinkValidated.ok ||
      !internalTextShapeHyperlinkValidated.data?.valid ||
      internalTextShapeHyperlinkValidated.data.errorCount !== 0 ||
      internalTextShapeHyperlinkValidated.data.warningCount !== 0 ||
      internalTextShapeHyperlinkValidated.data.diagnostics.length !== 0) {
    throw new Error(
      'CLI internal-only text-shape-hyperlink validation failed: ' +
      internalTextShapeHyperlinkValidateResult.stdout,
    );
  }
  if (process.env.PPTX_SLIDE_BACKGROUND_GALLERY_OUT) {
    const galleryOutput = resolve(process.env.PPTX_SLIDE_BACKGROUND_GALLERY_OUT);
    await mkdir(dirname(galleryOutput), { recursive: true });
    await writeFile(galleryOutput, await readFile(slideBackgroundDeckPath));
  }
  if (process.env.PPTX_CHART_GALLERY_OUT) {
    const galleryOutput = resolve(process.env.PPTX_CHART_GALLERY_OUT);
    await mkdir(dirname(galleryOutput), { recursive: true });
    await writeFile(galleryOutput, await readFile(nativeChartDeckPath));
  }
  if (process.env.PPTX_SLIDE_NUMBER_GALLERY_OUT) {
    const galleryOutput = resolve(process.env.PPTX_SLIDE_NUMBER_GALLERY_OUT);
    await mkdir(dirname(galleryOutput), { recursive: true });
    await writeFile(galleryOutput, await readFile(slideNumberDeckPath));
  }
  if (process.env.PPTX_SLIDE_DEFAULT_COLOR_GALLERY_OUT) {
    const galleryOutput = resolve(process.env.PPTX_SLIDE_DEFAULT_COLOR_GALLERY_OUT);
    await mkdir(dirname(galleryOutput), { recursive: true });
    await writeFile(galleryOutput, await readFile(slideDefaultColorDeckPath));
  }
  if (process.env.PPTX_MASTER_LAYOUT_GALLERY_OUT) {
    const galleryOutput = resolve(process.env.PPTX_MASTER_LAYOUT_GALLERY_OUT);
    await mkdir(dirname(galleryOutput), { recursive: true });
    await writeFile(galleryOutput, await readFile(masterLayoutDeckPath));
  }

  process.stdout.write(
    `${JSON.stringify({ ok: true, tarball: basename(tarball), api: apiChecks, masterLayouts: apiChecks.masterLayouts, slideNumbers: apiChecks.slideNumbers, slideDefaultColor: apiChecks.slideDefaultColor, presetShapes: apiChecks.presetShapes, customGeometryPaths: apiChecks.customGeometryPaths, customGeometryGuideFormulas: apiChecks.customGeometryGuideFormulas, customGeometryAdjustmentHandles: apiChecks.customGeometryAdjustmentHandles, customGeometryConnectionSites: apiChecks.customGeometryConnectionSites, customGeometryTextRectangles: apiChecks.customGeometryTextRectangles, customGeometryEvaluator: apiChecks.customGeometryEvaluator, shapeAdjustments: apiChecks.shapeAdjustments, shapeShadows: apiChecks.shapeShadows, shapeFills: apiChecks.shapeFills, textShapeFills: apiChecks.textShapeFills, textShapeLines: apiChecks.textShapeLines, textShapeArrows: apiChecks.textShapeArrows, textShapeShadows: apiChecks.textShapeShadows, textShapeHyperlinks: apiChecks.textShapeHyperlinks, shapeLines: apiChecks.shapeLines, shapeArrows: apiChecks.shapeArrows, shapeHyperlinks: apiChecks.shapeHyperlinks, embeddedRasterImages: apiChecks.embeddedRasterImages, svgImages: apiChecks.svgImages, embeddedMedia: apiChecks.embeddedMedia, stableMediaLifecycle: apiChecks.stableMediaLifecycle, nativeMediaTiming: apiChecks.nativeMediaTiming, nativeCharts: apiChecks.nativeCharts, slideBackgrounds: apiChecks.slideBackgrounds, types: true, cli: doctor.data.version, masterLayoutInspect: true, masterLayoutValidate: true, masterLayoutSlides: true, masterLayoutPartRead: true, masterLayoutDiff: true, svgInspect: true, svgValidate: true, mediaInspect: true, mediaValidate: true, stableMediaInspect: true, stableMediaValidate: true, nativeChartInspect: true, nativeChartValidate: true, nativeChartSlides: true, nativeChartPartRead: true, slideBackgroundInspect: true, slideBackgroundValidate: true, slideNumberInspect: true, slideNumberValidate: true, slideNumberSlides: true, slideNumberPartRead: true, slideDefaultColorInspect: true, slideDefaultColorValidate: true, slideDefaultColorSlides: true, slideDefaultColorPartRead: true, textShapeFillInspect: true, textShapeFillValidate: true, textShapeFillSlides: true, textShapeFillPartRead: true, textShapeLineInspect: true, textShapeLineValidate: true, textShapeLineSlides: true, textShapeLinePartRead: true, textShapeArrowInspect: true, textShapeArrowValidate: true, textShapeArrowSlides: true, textShapeArrowPartRead: true, textShapeShadowInspect: true, textShapeShadowValidate: true, textShapeShadowSlides: true, textShapeShadowPartRead: true, textShapeHyperlinkInspect: true, textShapeHyperlinkValidate: true, textShapeHyperlinkSlides: true, textShapeHyperlinkPartRead: true, textShapeHyperlinkInternalValidate: true })}\n`,
  );
} finally {
  await rm(directory, { recursive: true, force: true });
}

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8' });
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} failed (${result.status})\n${result.stdout}\n${result.stderr}`);
  }
  return result;
}

async function listFiles(directory) {
  const files = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await listFiles(path)));
    else files.push(path);
  }
  return files;
}
