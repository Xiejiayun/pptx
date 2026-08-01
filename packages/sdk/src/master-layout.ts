import type {
  LayoutModel as RawLayoutModel,
  MasterModel as RawMasterModel,
  SlideNumber,
  SlideNumberOptions,
  ThemeModel,
} from '@pptx/codecs';
import {
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
} from '@pptx/model';

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
  ) {
    super(document, raw.partUri, 'layout');
  }

  get name(): string {
    return this.raw.name;
  }

  get masterPartUri(): string | undefined {
    return this.raw.masterPartUri;
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
