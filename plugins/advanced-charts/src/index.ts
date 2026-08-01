import type { CodecDiagnostic, CodecRegistry } from '@pptx/codecs';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  ChartModel,
  PresentationModel,
  chartWorkbookMatches,
  readChartState,
  type ChartGroupInput,
} from '@pptx/model';
import type { OpcPackage } from '@pptx/opc';

export interface AdvancedChartSeries {
  readonly index: number;
  readonly name: string;
  readonly values: readonly number[];
  readonly categories: readonly string[];
  readonly hasTrendline: boolean;
  readonly hasErrorBars: boolean;
  readonly hasDataLabels: boolean;
}

export interface AdvancedChartModel {
  readonly partUri: string;
  readonly chartTypes: readonly string[];
  readonly combination: boolean;
  readonly modern: boolean;
  readonly axisIds: readonly number[];
  readonly embeddedWorkbookPartUri?: string;
  readonly series: readonly AdvancedChartSeries[];
}

export interface PluginHost {
  readonly opcPackage: OpcPackage;
  readonly codecRegistry: CodecRegistry;
}

export class AdvancedChartCodec {
  readonly id = 'plugin.advanced-charts';
  readonly priority = 200;
  readonly ownership = {
    elements: ['c:trendline', 'c:errBars', 'c:dLbls', 'cx:chart'],
    contentTypes: [
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      'application/vnd.ms-office.chartex+xml',
    ],
  } as const;

  constructor(readonly pkg: OpcPackage) {}

  list(): readonly AdvancedChartModel[] {
    return this.pkg.parts
      .filter(({ contentType }) => contentType.includes('drawingml.chart') || contentType.includes('chartex'))
      .map(({ uri }) => this.inspect(uri));
  }

  inspect(chartPartUri: string): AdvancedChartModel {
    const part = this.pkg.requirePart(chartPartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const plotArea = xml.elements('plotArea')[0];
    const rawChartTypes = plotArea
      ? plotArea.children
          .filter(
            (child): child is XmlElement =>
              child.type === 'element' && (child.localName.endsWith('Chart') || child.localName === 'plotAreaRegion'),
          )
          .map(({ localName }) => localName)
      : [];
    const state = readChartState(this.pkg, chartPartUri);
    const chartTypes = state.definition
      ? state.definition.groups.map(({ type }) => chartElementName(type))
      : rawChartTypes;
    const workbook = state.workbookPartUri ?? this.pkg
      .relationships(chartPartUri)
      .find(({ type }) => type.endsWith('/package'))?.resolvedTarget;
    const modern = state.status === 'modern'
      || part.contentType.includes('chartex')
      || xml.elements().some(({ name }) => name.startsWith('cx:'));
    return {
      partUri: chartPartUri,
      chartTypes,
      combination: chartTypes.length > 1,
      modern,
      axisIds: xml.elements('axId').map((element) => Number(xml.attribute(element, 'val')?.value ?? 0)),
      ...(workbook ? { embeddedWorkbookPartUri: workbook } : {}),
      series: state.definition
        ? projectStrictSeries(xml, state.definition.groups)
        : modern
          ? xml.elements('ser').map((series, index) => decodeSeries(xml, series, index))
          : [],
    };
  }

  async setSeriesValues(
    chartPartUri: string,
    seriesIndex: number,
    values: readonly number[],
  ): Promise<void> {
    const chart = this.resolveChart(chartPartUri);
    const definition = chart.definition;
    if (!definition) throw new Error(`Chart ${chartPartUri} is not semantically editable`);
    const normalizedValues = values.map(finite);
    let currentIndex = 0;
    let replaced = false;
    const groups = definition.groups.map((group) => ({
      type: group.type,
      series: group.series.map((series) => {
        if (currentIndex !== seriesIndex) {
          currentIndex += 1;
          return series;
        }
        currentIndex += 1;
        replaced = true;
        return { ...series, values: normalizedValues };
      }),
      ...(group.axis === undefined ? {} : { axis: group.axis }),
      ...(group.options === undefined ? {} : { options: group.options }),
    })) as readonly ChartGroupInput[];
    if (!replaced) throw new RangeError(`Chart series ${seriesIndex} was not found`);
    await chart.replaceDefinition({ groups, options: definition.options });
  }

  addTrendline(chartPartUri: string, seriesIndex: number, type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg' = 'linear'): void {
    this.appendToSeries(chartPartUri, seriesIndex, `<c:trendline><c:trendlineType val="${type}"/></c:trendline>`);
  }

  addErrorBars(chartPartUri: string, seriesIndex: number, value: number): void {
    this.appendToSeries(
      chartPartUri,
      seriesIndex,
      `<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="${finite(
        value,
      )}"/></c:errBars>`,
    );
  }

  enableDataLabels(chartPartUri: string, seriesIndex: number): void {
    this.appendToSeries(chartPartUri, seriesIndex, '<c:dLbls><c:showVal val="1"/></c:dLbls>');
  }

  replaceEmbeddedWorkbook(chartPartUri: string, bytes: Uint8Array): void {
    const workbook = this.inspect(chartPartUri).embeddedWorkbookPartUri;
    if (!workbook) throw new Error(`Chart ${chartPartUri} has no embedded workbook`);
    this.pkg.setPart(
      workbook,
      bytes,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
  }

  createImageFallback(chartPartUri: string, bytes: Uint8Array, contentType = 'image/png'): string {
    this.pkg.requirePart(chartPartUri);
    const extension = contentType === 'image/jpeg' ? '.jpg' : '.png';
    const uri = this.pkg.allocatePartUri('/ppt/media', 'chartFallback', extension);
    this.pkg.setPart(uri, bytes, contentType);
    return uri;
  }

  async diagnostics(chartPartUri: string): Promise<CodecDiagnostic[]> {
    const chart = this.inspect(chartPartUri);
    const state = readChartState(this.pkg, chartPartUri);
    if (
      state.status === 'recognized'
      && state.definition
      && state.workbookPartUri
      && !await chartWorkbookMatches(
        this.pkg.requirePart(state.workbookPartUri).bytes,
        state.definition,
      )
    ) {
      return [{
        severity: 'warning',
        code: 'CHART_WORKBOOK_CACHE_DIVERGENCE',
        message: 'Chart caches/formulas and embedded workbook cells disagree',
        partUri: chartPartUri,
      }];
    }
    if (chart.modern) {
      return [{
        severity: 'info',
        code: 'MODERN_CHART_EXTENSION',
        message: 'Modern chart extensions may degrade in older clients',
        partUri: chartPartUri,
      }];
    }
    return [];
  }

  private resolveChart(chartPartUri: string): ChartModel {
    const presentation = new PresentationModel(this.pkg);
    const matches = presentation.slides.flatMap(({ shapes }) => shapes).filter(
      (shape): shape is ChartModel =>
        shape instanceof ChartModel && shape.chartPartUri === chartPartUri,
    );
    if (matches.length !== 1) {
      throw new Error(
        `Chart ${chartPartUri} must have exactly one slide reference for semantic editing`,
      );
    }
    return matches[0]!;
  }

  private appendToSeries(chartPartUri: string, seriesIndex: number, childXml: string): void {
    const part = this.pkg.requirePart(chartPartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const series = xml.elements('ser')[seriesIndex];
    if (!series) throw new RangeError(`Chart series ${seriesIndex} was not found`);
    xml.appendChildXml(series, childXml);
    this.pkg.setPart(chartPartUri, xml.serialize(), part.contentType);
  }
}

export function installAdvancedChartPlugin(host: PluginHost): AdvancedChartCodec {
  const codec = new AdvancedChartCodec(host.opcPackage);
  host.codecRegistry.register(codec);
  return codec;
}

function decodeSeries(xml: LosslessXmlDocument, series: XmlElement, index: number): AdvancedChartSeries {
  const tx = xml.descendants(series, 'tx')[0];
  const categories = xml.descendants(series, 'cat')[0];
  const values = xml.descendants(series, 'val')[0];
  return {
    index,
    name: tx ? lastValue(xml, tx) : '',
    categories: categories ? pointValues(xml, categories) : [],
    values: values ? pointValues(xml, values).map(Number).filter(Number.isFinite) : [],
    hasTrendline: xml.descendants(series, 'trendline').length > 0,
    hasErrorBars: xml.descendants(series, 'errBars').length > 0,
    hasDataLabels: xml.descendants(series, 'dLbls').length > 0,
  };
}

function projectStrictSeries(
  xml: LosslessXmlDocument,
  groups: readonly ChartGroupInput[],
): readonly AdvancedChartSeries[] {
  const seriesElements = xml.elements('ser');
  let index = 0;
  return groups.flatMap((group) => group.series.map((series): AdvancedChartSeries => {
    const element = seriesElements[index];
    const categories = series.categories;
    const projectedCategories = categories === undefined
      ? []
      : Array.isArray(categories[0])
        ? (categories.at(-1) as readonly string[]).map(String)
        : (categories as readonly (string | number)[]).map(String);
    const result = {
      index,
      name: series.name,
      values: series.values,
      categories: projectedCategories,
      hasTrendline: Boolean(element && xml.descendants(element, 'trendline').length > 0),
      hasErrorBars: Boolean(element && xml.descendants(element, 'errBars').length > 0),
      hasDataLabels: Boolean(element && xml.descendants(element, 'dLbls').length > 0),
    };
    index += 1;
    return result;
  }));
}

function chartElementName(type: ChartGroupInput['type']): string {
  return `${type === 'bar3D' ? 'bar3D' : type}Chart`;
}

function pointValues(xml: LosslessXmlDocument, element: XmlElement): string[] {
  return xml
    .descendants(element, 'pt')
    .sort((left, right) => Number(xml.attribute(left, 'idx')?.value ?? 0) - Number(xml.attribute(right, 'idx')?.value ?? 0))
    .map((point) => lastValue(xml, point));
}

function lastValue(xml: LosslessXmlDocument, element: XmlElement): string {
  const value = xml.descendants(element, 'v').at(-1);
  return value ? xml.text(value) : '';
}

function finite(value: number): number {
  if (!Number.isFinite(value)) throw new Error('Chart values must be finite numbers');
  return value;
}
