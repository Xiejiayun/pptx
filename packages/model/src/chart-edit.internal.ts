import {
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import {
  relativeRelationshipTarget,
  type OpcPackage,
  type Relationship,
} from '@pptx/opc';
import type {
  ChartCategories,
  ChartDefinition,
  ChartSeries,
  ChartState,
} from './chart.js';
import { cloneOwnedPartForMutation } from './dependency.internal.js';
import { renderChartPart } from './chart-render.internal.js';
import { readChartState } from './chart-state.internal.js';
import { planChartWorkbook } from './chart-workbook.internal.js';
import type { SlideModel } from './slide.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/chart`;
const PACKAGE_RELATIONSHIP =
  `${RELATIONSHIP_NAMESPACE}/package`;
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const GROUP_NAMES = new Set([
  'areaChart',
  'barChart',
  'bar3DChart',
  'bubbleChart',
  'doughnutChart',
  'lineChart',
  'pieChart',
  'radarChart',
  'scatterChart',
]);
const AXIS_NAMES = new Set(['catAx', 'dateAx', 'serAx', 'valAx']);

export function replaceChartDefinition(
  pkg: OpcPackage,
  slide: SlideModel,
  shapeId: number,
  current: Readonly<ChartState>,
  next: Readonly<ChartDefinition>,
  workbookBytes: Uint8Array,
): void {
  pkg.transaction(() => {
    const resolved = resolveChartReference(pkg, slide, shapeId);
    const live = readChartState(pkg, resolved.chartPartUri);
    requireEditableState(live);
    if (
      !current.definition
      || !chartDefinitionsEqual(current.definition, live.definition!)
      || current.status !== live.status
      || current.workbookPartUri !== live.workbookPartUri
    ) {
      throw new Error('Chart changed while semantic replacement was being prepared');
    }

    const isolated = isolateSharedChart(pkg, slide, resolved);
    const isolatedState = readChartState(pkg, isolated.chartPartUri);
    requireEditableState(isolatedState);
    const workbook = replaceWorkbook(
      pkg,
      isolated.chartPartUri,
      isolatedState,
      workbookBytes,
    );
    const plan = planChartWorkbook(next);
    const canonical = renderChartPart(next, plan.formulas, workbook.relationshipId);
    const chartPart = pkg.requirePart(isolated.chartPartUri);
    const patched = patchChartPart(
      chartPart.bytes,
      canonical,
      isolatedState.definition!,
      next,
      isolatedState.status === 'cache-only',
    );
    pkg.setPart(isolated.chartPartUri, patched, chartPart.contentType);

    const result = readChartState(pkg, isolated.chartPartUri);
    if (
      result.status !== 'recognized'
      || !result.definition
      || !chartDefinitionsEqual(result.definition, next)
      || result.workbookPartUri !== workbook.partUri
    ) {
      throw new Error('Chart semantic replacement did not produce the requested state');
    }
  });
}

export function chartDefinitionsEqual(
  left: Readonly<ChartDefinition>,
  right: Readonly<ChartDefinition>,
): boolean {
  return left.groups.length === right.groups.length
    && left.groups.every((group, groupIndex) => {
      const other = right.groups[groupIndex];
      return other !== undefined
        && group.type === other.type
        && (group.axis ?? 'primary') === (other.axis ?? 'primary')
        && group.series.length === other.series.length
        && group.series.every((series, seriesIndex) =>
          seriesEqual(series, other.series[seriesIndex]));
    });
}

interface ResolvedChartReference {
  readonly xml: LosslessXmlDocument;
  readonly reference: XmlAttribute;
  readonly relationship: Relationship;
  readonly chartPartUri: string;
}

function resolveChartReference(
  pkg: OpcPackage,
  slide: SlideModel,
  shapeId: number,
): ResolvedChartReference {
  const { xml, element } = slide.resolveShape(shapeId);
  const chart = xml.descendants(element, 'chart')[0];
  const reference = chart ? xml.attribute(chart, 'r:id') : undefined;
  const relationship = reference
    ? pkg.relationships(slide.partUri).find(({ id }) => id === reference.value)
    : undefined;
  if (
    !reference
    || !relationship
    || namespaceUri(chart!) !== CHART_NAMESPACE
    || attributeNamespaceUri(chart!, reference) !== RELATIONSHIP_NAMESPACE
    || relationship.type !== CHART_RELATIONSHIP
    || relationship.targetMode !== 'Internal'
    || !relationship.resolvedTarget
  ) {
    throw new Error(`Chart ${shapeId} has no editable chart relationship`);
  }
  return {
    xml,
    reference,
    relationship,
    chartPartUri: relationship.resolvedTarget,
  };
}

function isolateSharedChart(
  pkg: OpcPackage,
  slide: SlideModel,
  resolved: ResolvedChartReference,
): ResolvedChartReference {
  const incoming = pkg.graph.find(({ uri }) => uri === resolved.chartPartUri)?.incoming.length ?? 0;
  const referenceCount = chartReferenceCount(resolved.xml, resolved.relationship.id);
  if (incoming <= 1 && referenceCount <= 1) return resolved;
  const cloneUri = cloneOwnedPartForMutation(pkg, resolved.chartPartUri);
  if (referenceCount > 1) {
    const relationship = pkg.addRelationship(slide.partUri, {
      type: resolved.relationship.type,
      target: relativeRelationshipTarget(slide.partUri, cloneUri),
      targetMode: 'Internal',
    });
    resolved.xml.replaceAttribute(resolved.reference, relationship.id);
    slide.setXml(resolved.xml.serialize());
    return { ...resolved, relationship, chartPartUri: cloneUri };
  }
  const relationship = pkg.updateRelationship(slide.partUri, resolved.relationship.id, {
    target: relativeRelationshipTarget(slide.partUri, cloneUri),
    targetMode: 'Internal',
  });
  return { ...resolved, relationship, chartPartUri: cloneUri };
}

function replaceWorkbook(
  pkg: OpcPackage,
  chartPartUri: string,
  state: Readonly<ChartState>,
  bytes: Uint8Array,
): { readonly relationshipId: string; readonly partUri: string } {
  const packageRelationships = pkg.relationships(chartPartUri).filter(
    ({ type }) => type === PACKAGE_RELATIONSHIP,
  );
  if (packageRelationships.length > 1) {
    throw new Error('Chart has multiple workbook package relationships');
  }
  const existingRelationship = packageRelationships[0];
  if (existingRelationship) {
    if (
      existingRelationship.targetMode !== 'Internal'
      || !existingRelationship.resolvedTarget
      || !pkg.hasPart(existingRelationship.resolvedTarget)
    ) {
      throw new Error('Chart workbook package relationship is invalid');
    }
    const existingUri = existingRelationship.resolvedTarget;
    const incoming = pkg.graph.find(({ uri }) => uri === existingUri)?.incoming.length ?? 0;
    if (incoming > 1) {
      const partUri = pkg.allocatePartUri(
        '/ppt/embeddings',
        'Microsoft_Excel_Worksheet',
        '.xlsx',
      );
      pkg.setPart(partUri, bytes, WORKBOOK_CONTENT_TYPE);
      pkg.updateRelationship(chartPartUri, existingRelationship.id, {
        target: relativeRelationshipTarget(chartPartUri, partUri),
        targetMode: 'Internal',
      });
      return { relationshipId: existingRelationship.id, partUri };
    }
    pkg.setPart(existingUri, bytes, WORKBOOK_CONTENT_TYPE);
    return { relationshipId: existingRelationship.id, partUri: existingUri };
  }
  if (state.status !== 'cache-only') throw new Error('Recognized chart workbook relationship is missing');
  const partUri = pkg.allocatePartUri(
    '/ppt/embeddings',
    'Microsoft_Excel_Worksheet',
    '.xlsx',
  );
  pkg.setPart(partUri, bytes, WORKBOOK_CONTENT_TYPE);
  const relationship = pkg.addRelationship(chartPartUri, {
    type: PACKAGE_RELATIONSHIP,
    target: relativeRelationshipTarget(chartPartUri, partUri),
    targetMode: 'Internal',
  });
  return { relationshipId: relationship.id, partUri };
}

function patchChartPart(
  source: Uint8Array,
  canonicalSource: string,
  current: Readonly<ChartDefinition>,
  next: Readonly<ChartDefinition>,
  addExternalData: boolean,
): string {
  const xml = LosslessXmlDocument.parse(source);
  const canonical = LosslessXmlDocument.parse(canonicalSource);
  const plotArea = requirePlotArea(xml);
  const canonicalPlotArea = requirePlotArea(canonical);
  if (sameStructure(current, next)) {
    patchSeriesData(xml, plotArea, canonical, canonicalPlotArea);
  } else {
    replaceOwnedPlotSpans(xml, plotArea, canonical, canonicalPlotArea);
  }
  if (addExternalData) insertExternalData(xml, canonical);
  return xml.serialize();
}

function patchSeriesData(
  xml: LosslessXmlDocument,
  plotArea: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalPlotArea: XmlElement,
): void {
  const groups = directChartChildren(plotArea).filter(({ localName }) => GROUP_NAMES.has(localName));
  const canonicalGroups = directChartChildren(canonicalPlotArea)
    .filter(({ localName }) => GROUP_NAMES.has(localName));
  groups.forEach((group, groupIndex) => {
    const canonicalGroup = canonicalGroups[groupIndex]!;
    const series = directChartChildren(group).filter(({ localName }) => localName === 'ser');
    const canonicalSeries = directChartChildren(canonicalGroup).filter(({ localName }) => localName === 'ser');
    series.forEach((entry, seriesIndex) => {
      const canonicalEntry = canonicalSeries[seriesIndex]!;
      const ownedNames = group.localName === 'scatterChart'
        ? ['idx', 'order', 'tx', 'xVal', 'yVal']
        : group.localName === 'bubbleChart'
          ? ['idx', 'order', 'tx', 'xVal', 'yVal', 'bubbleSize']
          : ['idx', 'order', 'tx', 'cat', 'val'];
      for (const localName of ownedNames) {
        const existing = exactlyOneDirectChartChild(entry, localName);
        const replacement = exactlyOneDirectChartChild(canonicalEntry, localName);
        xml.replaceElement(existing, canonical.original(replacement));
      }
    });
  });
}

function replaceOwnedPlotSpans(
  xml: LosslessXmlDocument,
  plotArea: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalPlotArea: XmlElement,
): void {
  const owned = directChartChildren(plotArea).filter(isOwnedPlotElement);
  const canonicalOwned = directChartChildren(canonicalPlotArea).filter(isOwnedPlotElement);
  if (owned.length === 0 || canonicalOwned.length === 0) {
    throw new Error('Chart plotArea has no replaceable owned spans');
  }
  xml.replaceElement(
    owned[0]!,
    canonicalOwned.map((element) => canonical.original(element)).join(''),
  );
  for (const element of owned.slice(1)) xml.removeElement(element);
}

function insertExternalData(
  xml: LosslessXmlDocument,
  canonical: LosslessXmlDocument,
): void {
  const root = xml.roots[0]!;
  const canonicalRoot = canonical.roots[0]!;
  const externalData = exactlyOneDirectChartChild(canonicalRoot, 'externalData');
  const rendered = canonical.original(externalData);
  const following = directChartChildren(root).find(({ localName }) =>
    localName === 'printSettings'
      || localName === 'userShapes'
      || localName === 'extLst');
  if (following) xml.replace(following.start, following.start, rendered);
  else xml.appendChildXml(root, rendered);
}

function requirePlotArea(xml: LosslessXmlDocument): XmlElement {
  if (xml.roots.length !== 1) throw new Error('Chart root is ambiguous');
  const chart = exactlyOneDirectChartChild(xml.roots[0]!, 'chart');
  return exactlyOneDirectChartChild(chart, 'plotArea');
}

function exactlyOneDirectChartChild(parent: XmlElement, localName: string): XmlElement {
  const children = directChartChildren(parent).filter((child) => child.localName === localName);
  if (children.length !== 1) throw new Error(`Chart ${localName} must occur exactly once`);
  return children[0]!;
}

function directChartChildren(parent: XmlElement): XmlElement[] {
  return parent.children.filter((child): child is XmlElement =>
    child.type === 'element' && namespaceUri(child) === CHART_NAMESPACE);
}

function isOwnedPlotElement(element: XmlElement): boolean {
  return GROUP_NAMES.has(element.localName) || AXIS_NAMES.has(element.localName);
}

function sameStructure(
  left: Readonly<ChartDefinition>,
  right: Readonly<ChartDefinition>,
): boolean {
  return left.groups.length === right.groups.length
    && left.groups.every((group, index) => {
      const other = right.groups[index];
      return other !== undefined
        && group.type === other.type
        && (group.axis ?? 'primary') === (other.axis ?? 'primary')
        && group.series.length === other.series.length;
    });
}

function seriesEqual(
  left: Readonly<ChartSeries>,
  right: Readonly<ChartSeries> | undefined,
): boolean {
  return right !== undefined
    && left.name === right.name
    && categoriesEqual(left.categories, right.categories)
    && valuesEqual(left.values, right.values)
    && valuesEqual(left.xValues, right.xValues)
    && valuesEqual(left.sizes, right.sizes);
}

function categoriesEqual(
  left: ChartCategories | undefined,
  right: ChartCategories | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  if (left.length !== right.length) return false;
  return left.every((value, index) => {
    const other = right[index];
    return Array.isArray(value) && Array.isArray(other)
      ? valuesEqual(value, other)
      : value === other;
  });
}

function valuesEqual<T>(left: readonly T[] | undefined, right: readonly T[] | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function requireEditableState(state: Readonly<ChartState>): void {
  if (
    (state.status !== 'recognized' && state.status !== 'cache-only')
    || !state.definition
  ) {
    throw new Error(`Chart semantic editing is unavailable for ${state.status} state`);
  }
}

function chartReferenceCount(xml: LosslessXmlDocument, relationshipId: string): number {
  return xml.elements('chart').filter((chart) => {
    const reference = xml.attribute(chart, 'r:id');
    return namespaceUri(chart) === CHART_NAMESPACE
      && reference?.value === relationshipId
      && attributeNamespaceUri(chart, reference) === RELATIONSHIP_NAMESPACE;
  }).length;
}

function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
  const separator = attribute.name.indexOf(':');
  if (separator < 0) return undefined;
  const declaration = `xmlns:${attribute.name.slice(0, separator)}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const match = current.attributes.find(({ name }) => name === declaration);
    if (match) return match.value;
  }
  return undefined;
}

function namespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const attribute = current.attributes.find(({ name }) => name === declaration);
    if (attribute) return attribute.value;
  }
  return undefined;
}
