import type {
  LayoutModel as RawLayoutModel,
  MasterModel as RawMasterModel,
  SlideNumber,
  SlideNumberOptions,
  ThemeModel,
} from '@pptx/codecs';
import { normalizeSlideNumberOptions } from '@pptx/codecs';
import {
  PLACEHOLDER_TYPES,
  ShapeModel,
  SlideModel,
  type AddChartOptions,
  type AddImageOptions,
  type AddPlaceholderOptions,
  type AddShapeOptions,
  type AddTextOptions,
  type ChartGroupInput,
  type ChartModel,
  type ChartSeriesInput,
  type ChartType,
  type ImageModel,
  type PresentationModel,
  type PresetShapeType,
  type RichTextParagraph,
  type SemanticShape,
  type SlideBackground,
  type SlideSize,
  type Emu,
} from '@pptx/model';

const DEFINE_KEYS = new Set([
  'title',
  'master',
  'background',
  'margin',
  'slideNumber',
  'objects',
]);
const SHAPE_OPTION_KEYS = new Set([
  'name',
  'placeholder',
  'adjustments',
  'fill',
  'line',
  'arrows',
  'hyperlink',
  'shadow',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
]);
const TEXT_OPTION_KEYS = new Set([
  'name',
  'placeholder',
  'align',
  'bullet',
  'fit',
  'lang',
  'level',
  'margin',
  'paragraphIndent',
  'paragraphMarginLeft',
  'paragraphMarginRight',
  'rtlMode',
  'spacing',
  'tabStops',
  'valign',
  'vert',
  'wrap',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
]);
const PLACEHOLDER_OPTION_KEYS = new Set([
  ...TEXT_OPTION_KEYS,
  'type',
  'index',
]);

export interface SlideMasterMargin {
  readonly top: Emu;
  readonly right: Emu;
  readonly bottom: Emu;
  readonly left: Emu;
}

export type SlideMasterMarginInput =
  | Emu
  | readonly [Emu, Emu, Emu, Emu];

export type SlideMasterObject =
  | { readonly kind: 'rect'; readonly options?: AddShapeOptions }
  | { readonly kind: 'line'; readonly options?: AddShapeOptions }
  | {
      readonly kind: 'text';
      readonly text: string | readonly RichTextParagraph[];
      readonly options?: AddTextOptions;
    }
  | {
      readonly kind: 'placeholder';
      readonly text?: string | readonly RichTextParagraph[];
      readonly options: AddPlaceholderOptions;
    };

export interface DefineSlideMasterOptions {
  readonly title: string;
  readonly master?: SlideMasterModel;
  readonly background?: SlideBackground;
  readonly margin?: SlideMasterMarginInput;
  readonly slideNumber?: SlideNumberOptions;
  readonly objects?: readonly SlideMasterObject[];
}

export type NormalizedSlideMasterObject =
  | { readonly kind: 'rect' | 'line'; readonly options: AddShapeOptions }
  | {
      readonly kind: 'text';
      readonly text: string | readonly RichTextParagraph[];
      readonly options: AddTextOptions;
    }
  | {
      readonly kind: 'placeholder';
      readonly text: string | readonly RichTextParagraph[];
      readonly options: AddPlaceholderOptions;
    };

export interface NormalizedDefineSlideMasterOptions {
  readonly title: string;
  readonly master?: SlideMasterModel;
  readonly background?: SlideBackground;
  readonly margin?: Readonly<SlideMasterMargin>;
  readonly slideNumber?: Readonly<SlideNumber>;
  readonly objects: readonly NormalizedSlideMasterObject[];
}

export function normalizeDefineSlideMasterOptions(
  value: unknown,
  slideSize: Readonly<SlideSize>,
): Readonly<NormalizedDefineSlideMasterOptions> {
  const input = readDataObject(value, DEFINE_KEYS, 'Slide master definition');
  const title = normalizeTitle(input.title);
  if (input.master !== undefined && !(input.master instanceof SlideMasterModel)) {
    throw new TypeError('Slide master definition master must be a SlideMasterModel');
  }
  const background = input.background === undefined
    ? undefined
    : cloneDataValue(input.background, 'Slide master background') as SlideBackground;
  if (
    background !== undefined
    && !['none', 'solid', 'image', 'linear-gradient', 'path-gradient'].includes(background.kind)
  ) {
    throw new TypeError('Slide master background kind is unsupported');
  }
  const margin = input.margin === undefined
    ? undefined
    : normalizeMargin(input.margin, slideSize);
  const slideNumber = input.slideNumber === undefined
    ? undefined
    : normalizeSlideNumberOptions(input.slideNumber);
  const objects = input.objects === undefined
    ? []
    : readDenseArray(input.objects, 'Slide master objects').map(
        (object, index) => normalizeObject(object, index),
      );
  return Object.freeze({
    title,
    ...(input.master === undefined ? {} : { master: input.master as SlideMasterModel }),
    ...(background === undefined ? {} : { background }),
    ...(margin === undefined ? {} : { margin }),
    ...(slideNumber === undefined ? {} : { slideNumber }),
    objects: Object.freeze(objects),
  });
}

abstract class CommonSlideOwnerModel {
  protected readonly content: SlideModel;

  protected constructor(
    document: PresentationModel,
    readonly partUri: string,
    ownerKind: 'layout' | 'master',
  ) {
    this.content = new SlideModel(document, partUri, '', 0, ownerKind);
  }

  get background(): SlideBackground | undefined {
    return this.content.background;
  }

  set background(value: SlideBackground | undefined) {
    this.content.background = value;
  }

  get shapes(): readonly SemanticShape[] {
    return this.content.shapes;
  }

  get placeholders(): readonly ShapeModel[] {
    return this.content.placeholders.filter(
      (shape): shape is ShapeModel => shape instanceof ShapeModel,
    );
  }

  addPlaceholder(
    value: string | readonly RichTextParagraph[],
    options: AddPlaceholderOptions,
  ): ShapeModel {
    return this.content.addPlaceholder(value, options);
  }

  addText(value: string, options: AddTextOptions = {}): ShapeModel {
    return this.content.addText(value, options);
  }

  addRichText(
    value: readonly RichTextParagraph[],
    options: AddTextOptions = {},
  ): ShapeModel {
    return this.content.addRichText(value, options);
  }

  addShape(
    type: PresetShapeType,
    options: AddShapeOptions = {},
  ): ShapeModel {
    return this.content.addShape(type, options);
  }

  addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel {
    return this.content.addImage(bytes, options);
  }

  addChart(
    type: ChartType,
    series: readonly ChartSeriesInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  addChart(
    groups: readonly ChartGroupInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  addChart(
    typeOrGroups: ChartType | readonly ChartGroupInput[],
    seriesOrOptions?: readonly ChartSeriesInput[] | AddChartOptions,
    options: AddChartOptions = {},
  ): Promise<ChartModel> {
    if (Array.isArray(typeOrGroups)) {
      return this.content.addChart(
        typeOrGroups,
        seriesOrOptions as AddChartOptions | undefined,
      );
    }
    return this.content.addChart(
      typeOrGroups as ChartType,
      seriesOrOptions as readonly ChartSeriesInput[],
      options,
    );
  }
}

export class SlideLayoutModel extends CommonSlideOwnerModel {
  /** @internal */
  constructor(
    document: PresentationModel,
    private readonly raw: RawLayoutModel,
    private readonly readMargin: (
      partUri: string,
    ) => Readonly<SlideMasterMargin> | undefined = () => undefined,
  ) {
    super(document, raw.partUri, 'layout');
  }

  get name(): string {
    return this.raw.name;
  }

  get masterPartUri(): string | undefined {
    return this.raw.masterPartUri;
  }

  get margin(): Readonly<SlideMasterMargin> | undefined {
    return this.readMargin(this.partUri);
  }

  get slideNumber(): Readonly<SlideNumber> | undefined {
    return this.raw.slideNumber;
  }

  set slideNumber(value: SlideNumberOptions | undefined) {
    this.raw.slideNumber = value;
  }
}

export class SlideMasterModel extends CommonSlideOwnerModel {
  /** @internal */
  constructor(
    document: PresentationModel,
    private readonly raw: RawMasterModel,
    private readonly layoutModel: (raw: RawLayoutModel) => SlideLayoutModel,
  ) {
    super(document, raw.partUri, 'master');
  }

  get layouts(): readonly SlideLayoutModel[] {
    return this.raw.layouts.map((layout) => this.layoutModel(layout));
  }

  get theme(): ThemeModel | undefined {
    return this.raw.theme;
  }

  get slideNumber(): Readonly<SlideNumber> | undefined {
    return this.raw.slideNumber;
  }

  set slideNumber(value: SlideNumberOptions | undefined) {
    this.raw.slideNumber = value;
  }
}

function normalizeObject(value: unknown, index: number): NormalizedSlideMasterObject {
  const context = `Slide master object ${index}`;
  const kind = readKind(value, context);
  if (kind === 'rect' || kind === 'line') {
    const input = readDataObject(value, new Set(['kind', 'options']), context);
    const options = readOptions(input.options, SHAPE_OPTION_KEYS, `${context} options`);
    return Object.freeze({ kind, options: options as AddShapeOptions });
  }
  if (kind === 'text') {
    const input = readDataObject(value, new Set(['kind', 'text', 'options']), context);
    const text = normalizeText(input.text, `${context} text`);
    const options = readOptions(input.options, TEXT_OPTION_KEYS, `${context} options`);
    return Object.freeze({ kind, text, options: options as AddTextOptions });
  }
  const input = readDataObject(value, new Set(['kind', 'text', 'options']), context);
  const text = input.text === undefined ? '' : normalizeText(input.text, `${context} text`);
  const options = readOptions(
    input.options,
    PLACEHOLDER_OPTION_KEYS,
    `${context} options`,
    false,
  ) as Record<string, unknown>;
  if (typeof options.name !== 'string' || options.name.length === 0 || !isValidXmlString(options.name)) {
    throw new TypeError(`${context} placeholder name must be a non-empty XML string`);
  }
  if (typeof options.type !== 'string' || !PLACEHOLDER_TYPES.includes(
    options.type as typeof PLACEHOLDER_TYPES[number],
  )) {
    throw new TypeError(`${context} placeholder type is invalid`);
  }
  const identityIndex = options.index === undefined ? 100 + index : options.index;
  if (
    typeof identityIndex !== 'number'
    || !Number.isSafeInteger(identityIndex)
    || identityIndex < 0
    || identityIndex > 4_294_967_294
  ) {
    throw new RangeError(`${context} placeholder index is out of range`);
  }
  return Object.freeze({
    kind,
    text,
    options: Object.freeze({
      ...options,
      type: options.type,
      index: identityIndex,
    }) as unknown as AddPlaceholderOptions,
  });
}

function readKind(
  value: unknown,
  context: string,
): 'rect' | 'line' | 'text' | 'placeholder' {
  const input = readDataObject(value, new Set(['kind', 'text', 'options']), context);
  if (
    typeof input.kind !== 'string'
    || !['rect', 'line', 'text', 'placeholder'].includes(input.kind)
  ) {
    throw new TypeError(`${context} kind must be rect, line, text, or placeholder`);
  }
  return input.kind as 'rect' | 'line' | 'text' | 'placeholder';
}

function normalizeText(
  value: unknown,
  context: string,
): string | readonly RichTextParagraph[] {
  if (typeof value === 'string') {
    if (!isValidXmlString(value)) throw new TypeError(`${context} contains invalid XML characters`);
    return value;
  }
  if (!Array.isArray(value)) throw new TypeError(`${context} must be text or rich text paragraphs`);
  return cloneDataValue(value, context) as unknown as readonly RichTextParagraph[];
}

function normalizeMargin(
  value: unknown,
  slideSize: Readonly<SlideSize>,
): Readonly<SlideMasterMargin> {
  const sides = typeof value === 'number'
    ? [value, value, value, value]
    : readDenseArray(value, 'Slide master margin');
  if (sides.length !== 4) {
    throw new RangeError('Slide master margin must contain exactly four values');
  }
  const normalized = sides.map((side, index) => {
    if (typeof side !== 'number' || !Number.isSafeInteger(side) || side < 0) {
      throw new TypeError(
        `Slide master margin ${['top', 'right', 'bottom', 'left'][index]} must be a non-negative safe integer`,
      );
    }
    return side;
  });
  if (normalized[1]! + normalized[3]! >= slideSize.width) {
    throw new RangeError('Slide master horizontal margins must be smaller than slide width');
  }
  if (normalized[0]! + normalized[2]! >= slideSize.height) {
    throw new RangeError('Slide master vertical margins must be smaller than slide height');
  }
  return Object.freeze({
    top: normalized[0] as Emu,
    right: normalized[1] as Emu,
    bottom: normalized[2] as Emu,
    left: normalized[3] as Emu,
  });
}

function normalizeTitle(value: unknown): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError('Slide master title must be a non-whitespace string');
  }
  if (!isValidXmlString(value)) {
    throw new TypeError('Slide master title contains invalid XML characters');
  }
  return value;
}

function readOptions(
  value: unknown,
  keys: ReadonlySet<string>,
  context: string,
  optional = true,
): Readonly<Record<string, unknown>> {
  if (value === undefined && optional) return Object.freeze({});
  const input = readDataObject(value, keys, context);
  const cloned: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(input)) {
    cloned[key] = cloneDataValue(entry, `${context} ${key}`);
  }
  return Object.freeze(cloned);
}

function readDataObject(
  value: unknown,
  keys: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !keys.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function readDenseArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value)) throw new TypeError(`${context} must be an array`);
  const keys = new Set(['length', ...Array.from({ length: value.length }, (_, index) => String(index))]);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !keys.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
  }
  return Array.from({ length: value.length }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} must be dense and contain only data items`);
    }
    return descriptor.value;
  });
}

function cloneDataValue(value: unknown, context: string): unknown {
  if (
    value === undefined
    || value === null
    || typeof value === 'string'
    || typeof value === 'number'
    || typeof value === 'boolean'
  ) return value;
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (Array.isArray(value)) {
    return Object.freeze(readDenseArray(value, context).map(
      (entry, index) => cloneDataValue(entry, `${context} ${index}`),
    ));
  }
  if (typeof value !== 'object') throw new TypeError(`${context} contains an unsupported value`);
  const input = readDataObject(
    value,
    new Set(Reflect.ownKeys(value).filter((key): key is string => typeof key === 'string')),
    context,
  );
  const result: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(input)) {
    result[key] = cloneDataValue(entry, `${context} ${key}`);
  }
  return Object.freeze(result);
}

function isValidXmlString(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x09
      || codePoint === 0x0A
      || codePoint === 0x0D
      || (codePoint >= 0x20 && codePoint <= 0xD7FF)
      || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
      || (codePoint >= 0x10000 && codePoint <= 0x10FFFF)
    ) continue;
    return false;
  }
  return true;
}
