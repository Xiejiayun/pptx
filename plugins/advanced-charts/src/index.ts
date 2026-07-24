import type { CodecDiagnostic, CodecRegistry } from '@pptx/codecs';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
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
  readonly #cacheMutations = new Set<string>();

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
    const chartTypes = plotArea
      ? plotArea.children
          .filter(
            (child): child is XmlElement =>
              child.type === 'element' && (child.localName.endsWith('Chart') || child.localName === 'plotAreaRegion'),
          )
          .map(({ localName }) => localName)
      : [];
    const workbook = this.pkg
      .relationships(chartPartUri)
      .find(({ type }) => type.endsWith('/package'))?.resolvedTarget;
    return {
      partUri: chartPartUri,
      chartTypes,
      combination: chartTypes.length > 1,
      modern: part.contentType.includes('chartex') || xml.elements().some(({ name }) => name.startsWith('cx:')),
      axisIds: xml.elements('axId').map((element) => Number(xml.attribute(element, 'val')?.value ?? 0)),
      ...(workbook ? { embeddedWorkbookPartUri: workbook } : {}),
      series: xml.elements('ser').map((series, index) => decodeSeries(xml, series, index)),
    };
  }

  setSeriesValues(chartPartUri: string, seriesIndex: number, values: readonly number[]): void {
    const part = this.pkg.requirePart(chartPartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const series = xml.elements('ser')[seriesIndex];
    const valueContainer = series ? xml.descendants(series, 'val')[0] : undefined;
    const cache = valueContainer ? xml.descendants(valueContainer, 'numCache')[0] : undefined;
    if (!cache) throw new Error(`Chart series ${seriesIndex} has no numeric cache`);
    for (const point of xml.descendants(cache, 'pt')) xml.removeElement(point);
    const pointCount = xml.descendants(cache, 'ptCount')[0];
    if (pointCount) {
      const value = xml.attribute(pointCount, 'val');
      if (value) xml.replaceAttribute(value, String(values.length));
    }
    xml.appendChildXml(
      cache,
      values.map((value, index) => `<c:pt idx="${index}"><c:v>${finite(value)}</c:v></c:pt>`).join(''),
    );
    this.pkg.setPart(chartPartUri, xml.serialize(), part.contentType);
    this.#cacheMutations.add(chartPartUri);
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
    this.#cacheMutations.delete(chartPartUri);
  }

  createImageFallback(chartPartUri: string, bytes: Uint8Array, contentType = 'image/png'): string {
    this.pkg.requirePart(chartPartUri);
    const extension = contentType === 'image/jpeg' ? '.jpg' : '.png';
    const uri = this.pkg.allocatePartUri('/ppt/media', 'chartFallback', extension);
    this.pkg.setPart(uri, bytes, contentType);
    return uri;
  }

  diagnostics(chartPartUri: string): CodecDiagnostic[] {
    const chart = this.inspect(chartPartUri);
    if (this.#cacheMutations.has(chartPartUri) && chart.embeddedWorkbookPartUri) {
      return [{
        severity: 'warning',
        code: 'CHART_WORKBOOK_CACHE_DIVERGENCE',
        message: 'Chart cache changed without replacing the embedded workbook',
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
