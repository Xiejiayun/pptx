import {
  CodecRegistry,
  GradientCodec,
  MasterLayoutThemeCodec,
  MediaCodec,
  slideNumberDiagnostics,
  type AddMediaOptions,
  type CodecDiagnostic,
  type LayoutModel as RawLayoutModel,
  type MasterModel as RawMasterModel,
  type MediaSource,
  type SlideNumber,
} from '@pptx/codecs';
import {
  ChartModel,
  MediaModel,
  PresentationModel,
  chartDiagnostics,
  type AddChartOptions,
  type AddSlideOptions,
  type AddSvgImageOptions,
  type ChartGroupInput,
  type ChartSeriesInput,
  type ChartType,
  type ImageModel,
  type PlaceholderSelector,
  type SlideBackground,
  type SlideModel,
} from '@pptx/model';
import {
  prepareChartCreation,
  type PreparedChartCreation,
} from '@pptx/model/internal/chart-create';
import { escapeXmlAttribute } from '@pptx/lossless-xml';
import { OpcPackage, type PackageOpenOptions } from '@pptx/opc';
import {
  ValidationError,
  validateMasterLayoutPlaceholders,
  validatePackage,
  type CompatibilityProfile,
  type Diagnostic,
} from '@pptx/validator';
import { createPresentationPackage, type CreatePresentationOptions } from './create.js';
import { type OutputType, type WriteOutput } from './output-type.js';
import { convertWriteOutput, resolveWriteOutputType } from './write-output.js';
import {
  normalizePresentationTheme,
  type PresentationTheme,
  type PresentationThemeOptions,
} from './presentation-theme.js';
import { PPTX_VERSION, type PptxVersion } from './version.js';
import {
  presentationLayoutFromSlideSize,
  type PresentationLayout,
} from './presentation-layout.js';
import {
  assertImageContentType,
  normalizeAddImageSourceOptions,
  resolveImageSource,
  type AddImageSourceOptions,
  type ImageSource,
  type RasterImageSource,
  type ResolvedImageSource,
} from './raster-image-source.js';
import { calculateImageSizing, type ImageSizing } from './raster-image-sizing.js';
import {
  resolveSlideBackgroundImage,
  type SetSlideBackgroundImageOptions,
} from './slide-background-source.js';
import { resolveSvgFallback } from './svg-image-fallback.js';
import {
  normalizeDefineSlideMasterOptions,
  SlideLayoutModel,
  SlideMasterModel,
  type DefineSlideMasterOptions,
  type NormalizedDefineSlideMasterOptions,
  type NormalizedSlideMasterObject,
  type SlideMasterMargin,
} from './master-layout.js';
import {
  normalizeTableToSlidesRequest,
  prepareTableToSlidesContent,
  snapshotHtmlTableById,
  type TableToSlidesMargins,
  type TableToSlidesOptions,
} from './table-to-slides.js';

const SLIDE_LAYOUT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';

export * from '@pptx/codecs';
export * from '@pptx/model';
export { MediaModel } from '@pptx/model';
export type { BuiltInSlideSize, CreatePresentationOptions, CustomSlideSize } from './create.js';
export type { PresentationTheme, PresentationThemeOptions } from './presentation-theme.js';
export { OUTPUT_TYPES } from './output-type.js';
export type { OutputType, WriteOutput } from './output-type.js';
export { PPTX_VERSION } from './version.js';
export type { PptxVersion } from './version.js';
export type { PresentationLayout, PresentationLayoutName } from './presentation-layout.js';
export {
  assertImageContentType,
  inspectImage,
  inspectRasterImage,
  resolveImageSource,
} from './raster-image-source.js';
export type {
  AddImageSourceOptions,
  ImageByteChunk,
  ImageByteStream,
  ImageContentType,
  ImageInfo,
  ImageSource,
  RasterImageByteChunk,
  RasterImageByteStream,
  RasterImageInfo,
  RasterImageSource,
  ResolvedImageSource,
} from './raster-image-source.js';
export { inspectSvgImage } from './svg-image-source.js';
export type { SvgImageInfo } from './svg-image-source.js';
export { calculateImageSizing, calculateRasterImageSizing } from './raster-image-sizing.js';
export type {
  ImageCropRegion,
  ImageSizing,
  ImageSizingResult,
  RasterImageCropRegion,
  RasterImageSizing,
  RasterImageSizingResult,
} from './raster-image-sizing.js';
export type { SetSlideBackgroundImageOptions } from './slide-background-source.js';
export { SlideLayoutModel, SlideMasterModel } from './master-layout.js';
export type {
  DefineSlideMasterOptions,
  SlideMasterBackground,
  SlideMasterMargin,
  SlideMasterMarginInput,
  SlideMasterObject,
} from './master-layout.js';
export { PackageError } from '@pptx/opc';
export type { PackageOpenOptions } from '@pptx/opc';
export { ValidationError } from '@pptx/validator';
export type { CompatibilityProfile, Diagnostic } from '@pptx/validator';
export { ModelParseError as ParseError, SlideModel as Slide, SlideTitleModel as SlideTitle } from '@pptx/model';

export type PptxByteChunk = number | Uint8Array | ArrayBuffer | ArrayBufferView;
export type PptxByteStream = ReadableStream<PptxByteChunk> | AsyncIterable<PptxByteChunk>;
export type PptxInput = string | Uint8Array | ArrayBuffer | Blob | PptxByteStream;

export interface PptxNodeReadableStream extends AsyncIterable<Uint8Array> {
  readonly destroyed: boolean;
  readonly readable: boolean;
  readonly readableEnded: boolean;
  readonly readableObjectMode: false;
  pipe<TDestination>(
    destination: TDestination,
    options?: { readonly end?: boolean },
  ): TDestination;
  pause(): this;
  resume(): this;
  isPaused(): boolean;
  read(size?: number): Uint8Array | null;
  destroy(error?: Error): this;
  on(event: 'data', listener: (chunk: Uint8Array) => void): this;
  on(event: 'end' | 'close', listener: () => void): this;
  on(event: 'error', listener: (error: Error) => void): this;
  once(event: 'data', listener: (chunk: Uint8Array) => void): this;
  once(event: 'end' | 'close', listener: () => void): this;
  once(event: 'error', listener: (error: Error) => void): this;
}

export interface WriteBaseOptions {
  readonly compatibility?: CompatibilityProfile;
  readonly compression?: boolean;
  readonly mode?: 'strict' | 'permissive';
}

export interface WriteOptions<
  TOutputType extends OutputType = 'uint8array',
> extends WriteBaseOptions {
  readonly outputType?: TOutputType;
}

type PreparedSlideMasterObject =
  | Exclude<NormalizedSlideMasterObject, { readonly kind: 'image' | 'chart' }>
  | {
      readonly kind: 'image';
      readonly source: Readonly<ResolvedImageSource>;
      readonly fallbackPngBytes?: Uint8Array;
      readonly options: Readonly<AddSvgImageOptions>;
    }
  | {
      readonly kind: 'chart';
      readonly chart: Readonly<PreparedChartCreation>;
    };

interface PreparedSlideMasterDefinition {
  readonly background?: SlideBackground;
  readonly objects: readonly PreparedSlideMasterObject[];
}

interface CanonicalSlideMasterDefinition {
  readonly title: string;
  readonly masterPartUri: string;
  readonly margin?: Readonly<SlideMasterMargin>;
  readonly slideNumber?: Readonly<SlideNumber>;
  readonly prepared: Readonly<PreparedSlideMasterDefinition>;
}

interface SlideMasterDefinitionRelationshipState {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
  readonly resolvedTarget?: string;
}

interface SlideMasterDefinitionPartState {
  readonly uri: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly relationships: readonly Readonly<SlideMasterDefinitionRelationshipState>[];
}

type SlideMasterDefinitionState = readonly Readonly<SlideMasterDefinitionPartState>[];

interface StoredSlideMasterDefinition {
  readonly definition: Readonly<CanonicalSlideMasterDefinition>;
  readonly state: SlideMasterDefinitionState;
}

export class OpaqueMutationError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'OpaqueMutationError';
  }
}

export class PptxDocument extends PresentationModel {
  readonly diagnostics: Diagnostic[] = [];
  readonly codecRegistry = new CodecRegistry();
  readonly #masterLayoutTheme: MasterLayoutThemeCodec;
  readonly #layoutModels = new Map<string, SlideLayoutModel>();
  readonly #masterModels = new Map<string, SlideMasterModel>();
  readonly #layoutMargins = new Map<string, Readonly<SlideMasterMargin>>();
  readonly #layoutDefinitions = new Map<string, Readonly<StoredSlideMasterDefinition>>();
  #deletedLastLayout = false;

  get version(): PptxVersion {
    return PPTX_VERSION;
  }

  get presLayout(): PresentationLayout {
    return presentationLayoutFromSlideSize(this.slideSize);
  }

  /** @internal */
  tableAutoPageMarginsForSlide(
    slide: SlideModel,
  ): Readonly<SlideMasterMargin> | undefined {
    if (slide.presentation !== this) return undefined;
    const relationships = slide.relationships.filter((relationship) =>
      relationship.type === SLIDE_LAYOUT_RELATIONSHIP
      && relationship.targetMode === 'Internal'
      && relationship.resolvedTarget !== undefined);
    if (relationships.length !== 1) return undefined;
    return this.#layoutMargins.get(relationships[0]!.resolvedTarget!);
  }

  private constructor(opcPackage: OpcPackage) {
    super(opcPackage);
    this.#masterLayoutTheme = new MasterLayoutThemeCodec(opcPackage, this.presentationPartUri);
    this.codecRegistry.register(this.#masterLayoutTheme);
    this.codecRegistry.register(new GradientCodec());
    this.codecRegistry.register(new MediaCodec(opcPackage));
  }

  static async open(input: PptxInput, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    const bytes = await readInput(input, options.signal);
    const pkg = await OpcPackage.open(bytes, options);
    return new PptxDocument(pkg);
  }

  static create(options: CreatePresentationOptions = {}): PptxDocument {
    const themeInput = options.theme;
    const theme = themeInput === undefined
      ? undefined
      : normalizePresentationTheme(themeInput);
    const document = new PptxDocument(createPresentationPackage(options));
    if (theme !== undefined) document.theme = theme;
    if (options.author !== undefined) document.author = options.author;
    if (options.company !== undefined) document.company = options.company;
    if (options.createdAt !== undefined) document.createdAt = options.createdAt;
    if (options.lastModifiedBy !== undefined) document.lastModifiedBy = options.lastModifiedBy;
    if (options.modifiedAt !== undefined) document.modifiedAt = options.modifiedAt;
    if (options.revision !== undefined) document.revision = options.revision;
    if (options.title !== undefined) document.title = options.title;
    if (options.subject !== undefined) document.subject = options.subject;
    return document;
  }

  static async fromBuffer(input: Uint8Array | ArrayBuffer, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    return this.open(input, options);
  }

  transaction<T>(operation: (document: this) => T): T {
    return this.opcPackage.transaction(() => {
      const result = operation(this);
      const diagnostics = validatePackage(this.opcPackage);
      if (diagnostics.some(({ severity }) => severity === 'error')) throw new ValidationError(diagnostics);
      return result;
    });
  }

  async write<TOutputType extends OutputType = 'uint8array'>(
    options: WriteOptions<TOutputType> = {},
  ): Promise<WriteOutput<TOutputType>> {
    const outputType = resolveWriteOutputType(options.outputType);
    const bytes = await this.#writeBytes(options);
    return convertWriteOutput(bytes, outputType as TOutputType);
  }

  async stream(options: WriteBaseOptions = {}): Promise<PptxNodeReadableStream> {
    if (!isNodeRuntime()) {
      throw new Error('PptxDocument.stream() is only supported in Node.js');
    }
    const bytes = await this.#writeBytes(options);
    const { Readable } = await loadNodeModule<NodeStreamModule>(['node:stream'].join('/'));
    return Readable.from(chunkPptxBytes(bytes), { objectMode: false }) as PptxNodeReadableStream;
  }

  async #writeBytes(options: WriteBaseOptions): Promise<Uint8Array> {
    const compression = resolveOutputCompression(options.compression);
    const compatibility = options.compatibility ?? 'powerpoint-current';
    const diagnostics: Diagnostic[] = [
      ...validatePackage(this.opcPackage),
      ...validateMasterLayoutPlaceholders(this.opcPackage, compatibility),
    ];
    const gradients = new GradientCodec();
    const media = new MediaCodec(this.opcPackage);
    const firstSlideNumber = this.firstSlideNumber ?? 1;
    for (const partUri of this.#layoutMargins.keys()) {
      diagnostics.push({
        severity: 'info',
        code: 'LAYOUT_MARGIN_TRANSIENT',
        message: 'Layout margin is runtime-only and is not serialized into the presentation',
        partUri,
        compatibility,
        suggestion: 'Reapply the margin after reopen before using runtime layout helpers.',
      });
    }
    for (const [index, slide] of this.slides.entries()) {
      const background = slide.background;
      if (background?.kind === 'linear-gradient' || background?.kind === 'path-gradient') {
        diagnostics.push(...gradients.diagnostics(background, compatibility, slide.partUri));
      }
      diagnostics.push(...media.diagnosticsForSlide(slide.partUri, compatibility));
      diagnostics.push(...await chartDiagnostics(this.opcPackage, slide.partUri));
      appendCodecDiagnostics(diagnostics, slideNumberDiagnostics(
        this.opcPackage,
        slide.partUri,
        'slide',
        String(firstSlideNumber + index),
        compatibility,
      ), compatibility);
    }
    for (const layout of this.layouts) {
      appendCodecDiagnostics(diagnostics, slideNumberDiagnostics(
        this.opcPackage,
        layout.partUri,
        'layout',
        '‹#›',
        compatibility,
      ), compatibility);
    }
    for (const master of this.masters) {
      appendCodecDiagnostics(diagnostics, slideNumberDiagnostics(
        this.opcPackage,
        master.partUri,
        'master',
        '‹#›',
        compatibility,
      ), compatibility);
    }
    this.diagnostics.splice(0, this.diagnostics.length, ...diagnostics);
    if ((options.mode ?? 'strict') === 'strict' && diagnostics.some(({ severity }) => severity === 'error')) {
      throw new ValidationError(diagnostics);
    }
    return this.opcPackage.write(
      compression === undefined ? {} : { compression },
    );
  }

  async writeFile(path: string, options: WriteBaseOptions = {}): Promise<void> {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    await fs.writeFile(path, await this.#writeBytes(options));
  }

  async writeBlob(options: WriteBaseOptions = {}): Promise<Blob> {
    const bytes = await this.#writeBytes(options);
    return new Blob([new Uint8Array(bytes).buffer], { type: this.formatProfile.fileContentType });
  }

  async download(
    fileName = `presentation${this.formatProfile.extension}`,
    options: WriteBaseOptions = {},
  ): Promise<void> {
    if (typeof document === 'undefined' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
      throw new Error('PptxDocument.download() requires a browser DOM; use writeFile() in Node.js');
    }
    const url = URL.createObjectURL(await this.writeBlob(options));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.style.display = 'none';
    document.body?.appendChild(anchor);
    try {
      anchor.click();
    } finally {
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 0);
    }
  }

  get masters(): readonly SlideMasterModel[] {
    return this.#masterLayoutTheme.masters.map((master) => this.modelForMaster(master));
  }

  get layouts(): readonly SlideLayoutModel[] {
    return this.#masterLayoutTheme.layouts.map((layout) => this.modelForLayout(layout));
  }

  get themes() {
    return this.#masterLayoutTheme.themes;
  }

  get theme(): PresentationTheme | undefined {
    const fonts = this.#masterLayoutTheme.presentationTheme?.fonts;
    return fonts === undefined
      ? undefined
      : { headFontFace: fonts.majorLatin, bodyFontFace: fonts.minorLatin };
  }

  set theme(value: PresentationThemeOptions) {
    const normalized = normalizePresentationTheme(value);
    const theme = this.#masterLayoutTheme.presentationTheme;
    if (!theme) throw new Error('Presentation does not have one editable direct theme');
    theme.setFonts({
      majorLatin: normalized.headFontFace,
      minorLatin: normalized.bodyFontFace,
    });
  }

  get masterLayoutTheme(): MasterLayoutThemeCodec {
    return this.#masterLayoutTheme;
  }

  override addSlide(options: AddSlideOptions = {}): SlideModel {
    if (this.#deletedLastLayout && this.#masterLayoutTheme.layouts.length === 0) {
      throw new Error('Presentation does not have a usable slide layout');
    }
    return super.addSlide(options);
  }

  async defineSlideMaster(options: DefineSlideMasterOptions): Promise<SlideLayoutModel> {
    const normalized = normalizeDefineSlideMasterOptions(options, this.slideSize);
    if (
      normalized.master !== undefined
      && !this.masters.includes(normalized.master)
    ) {
      throw new TypeError('Slide master definition master belongs to another document or is detached');
    }
    const rawMaster = this.#masterLayoutTheme.requireAttachedMaster(
      normalized.master?.partUri,
    );
    if (this.#masterLayoutTheme.hasAttachedLayoutName(normalized.title)) {
      throw new RangeError(`Slide master title ${normalized.title} is already in use`);
    }
    const prepared = await prepareSlideMasterDefinition(normalized);

    const rawLayout = this.opcPackage.transaction(() => {
      if (
        normalized.master !== undefined
        && !this.masters.includes(normalized.master)
      ) {
        throw new TypeError('Slide master definition master belongs to another document or is detached');
      }
      this.#masterLayoutTheme.requireAttachedMaster(rawMaster.partUri);
      if (this.#masterLayoutTheme.hasAttachedLayoutName(normalized.title)) {
        throw new RangeError(`Slide master title ${normalized.title} is already in use`);
      }
      const raw = this.#masterLayoutTheme.createLayout(
        rawMaster.partUri,
        blankNamedLayoutXml(normalized.title),
      );
      const layout = new SlideLayoutModel(
        this,
        raw,
        (partUri) => this.#layoutMargins.get(partUri),
      );
      if (prepared.background !== undefined) {
        layout.background = prepared.background;
      }
      for (const object of prepared.objects) addSlideMasterObject(layout, object);
      if (normalized.slideNumber !== undefined) {
        layout.slideNumber = normalized.slideNumber;
        this.#masterLayoutTheme.enableMasterSlideNumbers(rawMaster.partUri);
      }
      return raw;
    });
    const layout = this.modelForLayout(rawLayout);
    if (normalized.margin !== undefined) {
      this.#layoutMargins.set(rawLayout.partUri, normalized.margin);
    }
    const definition = Object.freeze({
      title: normalized.title,
      masterPartUri: rawMaster.partUri,
      ...(normalized.margin === undefined ? {} : { margin: normalized.margin }),
      ...(normalized.slideNumber === undefined ? {} : { slideNumber: normalized.slideNumber }),
      prepared,
    });
    this.#layoutDefinitions.set(rawLayout.partUri, Object.freeze({
      definition,
      state: captureLayoutDefinitionState(this.opcPackage, rawLayout.partUri),
    }));
    this.#deletedLastLayout = false;
    return layout;
  }

  async replaceSlideMaster(
    layout: SlideLayoutModel,
    options: DefineSlideMasterOptions,
  ): Promise<void> {
    const current = this.layouts.find(({ partUri }) => partUri === layout?.partUri);
    if (!(layout instanceof SlideLayoutModel) || current !== layout) {
      throw new TypeError('Slide master replacement layout belongs to another document or is detached');
    }
    const normalized = normalizeDefineSlideMasterOptions(options, this.slideSize);
    if (
      normalized.master !== undefined
      && !this.masters.includes(normalized.master)
    ) {
      throw new TypeError('Slide master replacement master belongs to another document or is detached');
    }
    const rawMaster = this.#masterLayoutTheme.requireAttachedMaster(
      normalized.master?.partUri,
    );
    this.#masterLayoutTheme.requireAttachedLayout(layout.partUri);
    if (this.#masterLayoutTheme.hasAttachedLayoutNameExcept(
      normalized.title,
      layout.partUri,
    )) {
      throw new RangeError(`Slide master title ${normalized.title} is already in use`);
    }
    const prepared = await prepareSlideMasterDefinition(normalized);
    const definition = Object.freeze({
      title: normalized.title,
      masterPartUri: rawMaster.partUri,
      ...(normalized.margin === undefined ? {} : { margin: normalized.margin }),
      ...(normalized.slideNumber === undefined ? {} : { slideNumber: normalized.slideNumber }),
      prepared,
    });
    const existing = this.#layoutDefinitions.get(layout.partUri);
    if (
      existing
      && dataValuesEqual(existing.definition, definition)
      && dataValuesEqual(
        existing.state,
        captureLayoutDefinitionState(this.opcPackage, layout.partUri),
      )
    ) return;

    this.opcPackage.transaction(() => {
      if (layout.isStale()) {
        throw new TypeError('Slide master replacement layout belongs to another document or is detached');
      }
      if (
        normalized.master !== undefined
        && !this.masters.includes(normalized.master)
      ) {
        throw new TypeError('Slide master replacement master belongs to another document or is detached');
      }
      this.#masterLayoutTheme.requireAttachedMaster(rawMaster.partUri);
      this.#masterLayoutTheme.requireAttachedLayout(layout.partUri);
      if (this.#masterLayoutTheme.hasAttachedLayoutNameExcept(
        normalized.title,
        layout.partUri,
      )) {
        throw new RangeError(`Slide master title ${normalized.title} is already in use`);
      }
      const replacementRaw = this.#masterLayoutTheme.resetLayout(
        layout.partUri,
        rawMaster.partUri,
        blankNamedLayoutXml(normalized.title),
      );
      const replacementLayout = new SlideLayoutModel(
        this,
        replacementRaw,
        (partUri) => this.#layoutMargins.get(partUri),
      );
      if (prepared.background !== undefined) replacementLayout.background = prepared.background;
      for (const object of prepared.objects) addSlideMasterObject(replacementLayout, object);
      if (normalized.slideNumber !== undefined) {
        replacementLayout.slideNumber = normalized.slideNumber;
        this.#masterLayoutTheme.enableMasterSlideNumbers(rawMaster.partUri);
      }
    });
    layout.rotateContentGeneration();
    if (normalized.margin === undefined) this.#layoutMargins.delete(layout.partUri);
    else this.#layoutMargins.set(layout.partUri, normalized.margin);
    this.#layoutDefinitions.set(layout.partUri, Object.freeze({
      definition,
      state: captureLayoutDefinitionState(this.opcPackage, layout.partUri),
    }));
  }

  deleteSlideMaster(
    layout: SlideLayoutModel,
    replacement?: SlideLayoutModel,
  ): void {
    const current = this.layouts.find(({ partUri }) => partUri === layout?.partUri);
    if (!(layout instanceof SlideLayoutModel) || current !== layout) {
      throw new TypeError('Slide master deletion layout belongs to another document or is detached');
    }
    if (replacement !== undefined) {
      const currentReplacement = this.layouts.find(
        ({ partUri }) => partUri === replacement.partUri,
      );
      if (
        !(replacement instanceof SlideLayoutModel)
        || currentReplacement !== replacement
        || replacement === layout
      ) {
        throw new TypeError('Slide master deletion replacement is invalid or detached');
      }
    }
    this.#masterLayoutTheme.deleteLayout(layout.partUri, replacement?.partUri);
    layout.rotateContentGeneration();
    this.#layoutMargins.delete(layout.partUri);
    this.#layoutDefinitions.delete(layout.partUri);
    this.#layoutModels.delete(layout.partUri);
    this.#deletedLastLayout = this.#masterLayoutTheme.layouts.length === 0;
  }

  private modelForLayout(raw: RawLayoutModel): SlideLayoutModel {
    const existing = this.#layoutModels.get(raw.partUri);
    if (existing && !existing.isStale()) return existing;
    if (existing) {
      this.#layoutModels.delete(raw.partUri);
      this.#layoutMargins.delete(raw.partUri);
      this.#layoutDefinitions.delete(raw.partUri);
    }
    const created = new SlideLayoutModel(
      this,
      raw,
      (partUri) => this.#layoutMargins.get(partUri),
    );
    this.#layoutModels.set(raw.partUri, created);
    return created;
  }

  private modelForMaster(raw: RawMasterModel): SlideMasterModel {
    const existing = this.#masterModels.get(raw.partUri);
    if (existing && !existing.isStale()) return existing;
    if (existing) this.#masterModels.delete(raw.partUri);
    const created = new SlideMasterModel(
      this,
      raw,
      (layout) => this.modelForLayout(layout),
    );
    this.#masterModels.set(raw.partUri, created);
    return created;
  }

  async tableToSlides(
    elementId: string,
    options: TableToSlidesOptions = {},
  ): Promise<readonly SlideModel[]> {
    const request = normalizeTableToSlidesRequest(elementId, options);
    const snapshot = snapshotHtmlTableById(request.id);
    const layoutMargins = resolveTableToSlidesLayoutMargins(this, request.masterSlideName);
    const prepared = prepareTableToSlidesContent(
      request,
      snapshot,
      this.slideSize,
      layoutMargins,
    );
    const existingPartUris = new Set(this.slides.map(({ partUri }) => partUri));
    const createdPartUris = new Set<string>();
    try {
      return this.opcPackage.transaction(() => {
        try {
          const first = this.addSlide(
            request.masterSlideName === undefined
              ? {}
              : { masterName: request.masterSlideName },
          );
          first.addTable(prepared.rows, prepared.tableOptions);
          const htmlPages = Object.freeze([first, ...first.newAutoPagedSlides]);
          for (const page of htmlPages) {
            if (request.addShape !== undefined) {
              page.addShape(request.addShape.type, request.addShape.options);
            }
            if (request.addTable !== undefined) {
              page.addTable(request.addTable.rows, request.addTable.options);
            }
            if (request.addText !== undefined) {
              if (typeof request.addText.text === 'string') {
                page.addText(request.addText.text, request.addText.options);
              } else {
                page.addRichText(request.addText.text, request.addText.options);
              }
            }
          }
          return htmlPages;
        } finally {
          for (const slide of this.slides) {
            if (!existingPartUris.has(slide.partUri)) createdPartUris.add(slide.partUri);
          }
        }
      });
    } catch (error) {
      for (const partUri of createdPartUris) discardTableToSlidesModel(this, partUri);
      throw error;
    }
  }

  async addImage(
    slideIndex: number,
    source: ImageSource,
    options: AddImageSourceOptions = {},
  ): Promise<ImageModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    const normalized = normalizeAddImageSourceOptions(options);
    const resolved = await resolveImageSource(source, normalized.signal);
    assertImageContentType(normalized.contentType, resolved);
    const placement = normalized.sizing
      ? calculateImageSizing(
          resolved.info,
          imageSizingForPlaceholder(
            slide,
            normalized.imageOptions.placeholder,
            normalized.sizing,
          ),
        ) as Pick<
          AddSvgImageOptions,
          'width' | 'height' | 'sourceRectangle'
        >
      : undefined;
    if (resolved.info.contentType !== 'image/svg+xml') {
      if (normalized.fallback !== undefined) {
        throw new TypeError('fallback is only valid for SVG images');
      }
      return slide.addImage(resolved.bytes, {
        ...normalized.imageOptions,
        ...(placement ?? {}),
        contentType: resolved.info.contentType,
      });
    }
    const fallback = await resolveSvgFallback(
      resolved,
      normalized.fallback,
      normalized.signal,
    );
    return slide.addSvgImage(resolved.bytes, fallback, {
      ...normalized.imageOptions,
      ...(placement ?? {}),
    });
  }

  async setSlideBackgroundImage(
    slideIndex: number,
    source: RasterImageSource,
    options: SetSlideBackgroundImageOptions = {},
  ): Promise<void> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    slide.background = await resolveSlideBackgroundImage(source, options);
  }

  addChart(
    slideIndex: number,
    type: ChartType,
    series: readonly ChartSeriesInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  addChart(
    slideIndex: number,
    groups: readonly ChartGroupInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  async addChart(
    slideIndex: number,
    typeOrGroups: ChartType | readonly ChartGroupInput[],
    seriesOrOptions?: readonly ChartSeriesInput[] | AddChartOptions,
    options: AddChartOptions = {},
  ): Promise<ChartModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return Array.isArray(typeOrGroups)
      ? slide.addChart(typeOrGroups, seriesOrOptions as AddChartOptions | undefined)
      : slide.addChart(
          typeOrGroups as ChartType,
          seriesOrOptions as readonly ChartSeriesInput[],
          options,
        );
  }

  async addAudio(slideIndex: number, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return slide.addAudio(source, options);
  }

  async addVideo(slideIndex: number, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return slide.addVideo(source, options);
  }

  media(slideIndex: number): readonly MediaModel[] {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return slide.media;
  }
}

function resolveTableToSlidesLayoutMargins(
  document: PptxDocument,
  masterSlideName: string | undefined,
): Readonly<TableToSlidesMargins> | undefined {
  const layouts = document.layouts;
  if (masterSlideName !== undefined) {
    const matches = layouts.filter(({ name }) => name === masterSlideName);
    if (matches.length === 0) {
      throw new RangeError(`Slide master ${masterSlideName} was not found`);
    }
    if (matches.length !== 1) {
      throw new TypeError(`Slide master ${masterSlideName} is ambiguous`);
    }
    return matches[0]!.margin;
  }
  const byPartUri = new Map(layouts.map((layout) => [layout.partUri, layout]));
  const inherited = document.slides[0]?.relationships.filter((relationship) =>
    relationship.type === SLIDE_LAYOUT_RELATIONSHIP
    && relationship.targetMode === 'Internal'
    && relationship.resolvedTarget !== undefined
    && byPartUri.has(relationship.resolvedTarget)) ?? [];
  if (inherited.length > 1) {
    throw new TypeError('Source slide layout is ambiguous');
  }
  return inherited[0]?.resolvedTarget === undefined
    ? layouts[0]?.margin
    : byPartUri.get(inherited[0].resolvedTarget)?.margin;
}

function discardTableToSlidesModel(document: PptxDocument, partUri: string): void {
  const presentation = document as PptxDocument & {
    discardDetachedSlideModel(uri: string): void;
  };
  presentation.discardDetachedSlideModel(partUri);
}

async function prepareSlideMasterDefinition(
  definition: Readonly<NormalizedDefineSlideMasterOptions>,
): Promise<Readonly<PreparedSlideMasterDefinition>> {
  const background = definition.background?.kind === 'image-source'
    ? resolveSlideBackgroundImage(definition.background.source, {
        ...(definition.background.contentType === undefined
          ? {}
          : { contentType: definition.background.contentType }),
        ...(definition.background.signal === undefined
          ? {}
          : { signal: definition.background.signal }),
      })
    : Promise.resolve(definition.background);
  const objects = Promise.all(definition.objects.map(async (
    object,
  ): Promise<PreparedSlideMasterObject> => {
    if (object.kind === 'chart') {
      return Object.freeze({
        kind: 'chart',
        chart: await prepareChartCreation(object.groups, object.options),
      });
    }
    if (object.kind !== 'image') return object;
    const resolved = await resolveImageSource(object.source, object.options.signal);
    assertImageContentType(object.options.contentType, resolved);
    const placement = object.options.sizing === undefined
      ? undefined
      : calculateImageSizing(resolved.info, object.options.sizing);
    const imageOptions = Object.freeze({
      ...object.options.imageOptions,
      ...(placement ?? {}),
    }) as unknown as Readonly<AddSvgImageOptions>;
    if (resolved.info.contentType !== 'image/svg+xml') {
      if (object.options.fallback !== undefined) {
        throw new TypeError('fallback is only valid for SVG images');
      }
      return Object.freeze({
        kind: 'image',
        source: resolved,
        options: imageOptions,
      });
    }
    return Object.freeze({
      kind: 'image',
      source: resolved,
      fallbackPngBytes: await resolveSvgFallback(
        resolved,
        object.options.fallback,
        object.options.signal,
      ),
      options: imageOptions,
    });
  }));
  const [preparedBackground, preparedObjects] = await Promise.all([background, objects]);
  return Object.freeze({
    ...(preparedBackground === undefined ? {} : { background: preparedBackground }),
    objects: Object.freeze(preparedObjects),
  });
}

function addSlideMasterObject(
  layout: SlideLayoutModel,
  object: PreparedSlideMasterObject,
): void {
  switch (object.kind) {
    case 'rect':
    case 'line':
      layout.addShape(object.kind, object.options);
      return;
    case 'text':
      if (typeof object.text === 'string') layout.addText(object.text, object.options);
      else layout.addRichText(object.text, object.options);
      return;
    case 'placeholder':
      layout.addPlaceholder(object.text, object.options);
      return;
    case 'image':
      if (object.source.info.contentType === 'image/svg+xml') {
        layout.addSvgImage(
          object.source.bytes,
          object.fallbackPngBytes!,
          object.options,
        );
      } else {
        layout.addImage(object.source.bytes, {
          ...object.options,
          contentType: object.source.info.contentType,
        });
      }
      return;
    case 'chart':
      layout.commitPreparedChart(object.chart);
  }
}

function dataValuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (left instanceof Uint8Array || right instanceof Uint8Array) {
    return left instanceof Uint8Array
      && right instanceof Uint8Array
      && byteViewsEqual(left, right);
  }
  if (left instanceof ArrayBuffer || right instanceof ArrayBuffer) {
    if (!(left instanceof ArrayBuffer) || !(right instanceof ArrayBuffer)) return false;
    return byteViewsEqual(new Uint8Array(left), new Uint8Array(right));
  }
  if (ArrayBuffer.isView(left) || ArrayBuffer.isView(right)) {
    if (!ArrayBuffer.isView(left) || !ArrayBuffer.isView(right)) return false;
    return byteViewsEqual(
      new Uint8Array(left.buffer, left.byteOffset, left.byteLength),
      new Uint8Array(right.buffer, right.byteOffset, right.byteLength),
    );
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left)
      && Array.isArray(right)
      && left.length === right.length
      && left.every((value, index) => dataValuesEqual(value, right[index]));
  }
  if (
    !left
    || !right
    || typeof left !== 'object'
    || typeof right !== 'object'
  ) return false;
  const leftKeys = Reflect.ownKeys(left);
  const rightKeys = Reflect.ownKeys(right);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) =>
    rightKeys.includes(key)
    && dataValuesEqual(
      (left as Record<PropertyKey, unknown>)[key],
      (right as Record<PropertyKey, unknown>)[key],
    ));
}

function byteViewsEqual(left: Uint8Array, right: Uint8Array): boolean {
  return left.byteLength === right.byteLength
    && left.every((value, index) => value === right[index]);
}

function captureLayoutDefinitionState(
  pkg: OpcPackage,
  layoutPartUri: string,
): SlideMasterDefinitionState {
  const slideMasterRelationshipType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
  const visited = new Set<string>();
  const parts: Readonly<SlideMasterDefinitionPartState>[] = [];
  const visit = (partUri: string): void => {
    if (visited.has(partUri)) return;
    visited.add(partUri);
    const part = pkg.requirePart(partUri);
    const sourceRelationships = [...pkg.relationships(partUri)].sort(compareRelationshipState);
    const relationships = Object.freeze(sourceRelationships.map((relationship) => Object.freeze({
      id: relationship.id,
      type: relationship.type,
      target: relationship.target,
      targetMode: relationship.targetMode,
      ...(relationship.resolvedTarget === undefined
        ? {}
        : { resolvedTarget: relationship.resolvedTarget }),
    })));
    parts.push(Object.freeze({
      uri: part.uri,
      contentType: part.contentType,
      bytes: new Uint8Array(part.bytes),
      relationships,
    }));
    for (const relationship of sourceRelationships) {
      if (
        relationship.type === slideMasterRelationshipType
        || relationship.targetMode !== 'Internal'
        || relationship.resolvedTarget === undefined
        || !pkg.hasPart(relationship.resolvedTarget)
      ) continue;
      visit(relationship.resolvedTarget);
    }
  };
  visit(layoutPartUri);
  parts.sort(({ uri: left }, { uri: right }) => left < right ? -1 : left > right ? 1 : 0);
  return Object.freeze(parts);
}

function compareRelationshipState(
  left: Readonly<SlideMasterDefinitionRelationshipState>,
  right: Readonly<SlideMasterDefinitionRelationshipState>,
): number {
  const leftKey = [
    left.id,
    left.type,
    left.target,
    left.targetMode,
    left.resolvedTarget ?? '',
  ].join('\u0000');
  const rightKey = [
    right.id,
    right.type,
    right.target,
    right.targetMode,
    right.resolvedTarget ?? '',
  ].join('\u0000');
  return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
}

function blankNamedLayoutXml(title: string): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    + 'type="blank" preserve="1">'
    + `<p:cSld name="${escapeXmlAttribute(title)}"><p:spTree>`
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
    + '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
    + '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
    + '</a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
}

function imageSizingForPlaceholder(
  slide: SlideModel,
  selector: PlaceholderSelector | undefined,
  sizing: Readonly<ImageSizing>,
): Readonly<ImageSizing> {
  if (selector === undefined) return sizing;
  const owner = slide.placeholders.find((shape) => {
    const identity = shape.placeholder;
    return typeof selector === 'string'
      ? shape.name === selector
      : identity?.type === selector.type && identity.index === selector.index;
  });
  if (owner?.placeholder?.type !== 'pic') return sizing;
  return Object.freeze({
    ...sizing,
    width: owner.transform.width,
    height: owner.transform.height,
  });
}

function appendCodecDiagnostics(
  target: Diagnostic[],
  source: readonly CodecDiagnostic[],
  compatibility: CompatibilityProfile,
): void {
  for (const { severity, code, message, partUri } of source) {
    target.push({
      severity,
      code,
      message,
      ...(partUri === undefined ? {} : { partUri }),
      compatibility,
    });
  }
}

export async function* openPptxStream(path: string): AsyncIterable<Uint8Array> {
  const fs = await loadNodeModule<NodeFs>(['node:fs'].join('/'));
  for await (const chunk of fs.createReadStream(path)) yield normalizeByteChunk(chunk);
}

async function readInput(input: PptxInput, signal?: AbortSignal): Promise<Uint8Array> {
  if (typeof input === 'string') {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    return new Uint8Array(await fs.readFile(input, signal ? { signal } : undefined));
  }
  if (input instanceof Uint8Array) return new Uint8Array(input);
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) return new Uint8Array(await input.arrayBuffer());
  const chunks: Uint8Array[] = [];
  if (isReadableStream(input)) {
    const reader = input.getReader();
    try {
      while (true) {
        throwIfAborted(signal);
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(normalizeByteChunk(value));
      }
    } finally {
      reader.releaseLock();
    }
    return concatenateBytes(chunks);
  }
  if (!isAsyncIterable(input)) throw new TypeError('Unsupported PPTX input type');
  for await (const chunk of input) {
    if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
    chunks.push(normalizeByteChunk(chunk));
  }
  return concatenateBytes(chunks);
}

interface NodeFsPromises {
  readFile(path: string, options?: { readonly signal?: AbortSignal }): Promise<Uint8Array>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
}

interface NodeFs {
  createReadStream(path: string): AsyncIterable<unknown>;
}

interface NodeStreamModule {
  readonly Readable: {
    from(
      iterable: Iterable<Uint8Array>,
      options: { readonly objectMode: false },
    ): unknown;
  };
}

const NODE_STREAM_CHUNK_SIZE = 64 * 1024;

function resolveOutputCompression(value: unknown): boolean | undefined {
  if (value !== undefined && typeof value !== 'boolean') {
    throw new TypeError('PptxDocument output compression must be a boolean');
  }
  return value;
}

function isNodeRuntime(): boolean {
  return typeof process !== 'undefined' && Boolean(process.versions?.node);
}

function* chunkPptxBytes(bytes: Uint8Array): Iterable<Uint8Array> {
  for (let offset = 0; offset < bytes.byteLength; offset += NODE_STREAM_CHUNK_SIZE) {
    yield bytes.subarray(offset, Math.min(offset + NODE_STREAM_CHUNK_SIZE, bytes.byteLength));
  }
}

async function loadNodeModule<T>(specifier: string): Promise<T> {
  if (typeof process === 'undefined' || !process.versions?.node) {
    throw new Error('Local file paths are only supported in Node.js; pass a Blob, File, or byte stream');
  }
  return import(specifier) as Promise<T>;
}

function isReadableStream(value: unknown): value is ReadableStream<unknown> {
  return Boolean(value && typeof (value as { getReader?: unknown }).getReader === 'function');
}

function isAsyncIterable(value: unknown): value is AsyncIterable<unknown> {
  return Boolean(value && typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === 'function');
}

function normalizeByteChunk(chunk: unknown): Uint8Array {
  if (typeof chunk === 'number' && Number.isInteger(chunk) && chunk >= 0 && chunk <= 255) {
    return Uint8Array.of(chunk);
  }
  if (chunk instanceof Uint8Array) return new Uint8Array(chunk);
  if (chunk instanceof ArrayBuffer) return new Uint8Array(chunk);
  if (ArrayBuffer.isView(chunk)) {
    return new Uint8Array(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
  }
  throw new TypeError('PPTX streams must yield byte numbers, Uint8Array, ArrayBuffer, or ArrayBufferView chunks');
}

function concatenateBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.byteLength, 0));
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
  }
}
