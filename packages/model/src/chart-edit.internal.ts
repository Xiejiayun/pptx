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
import { chartOptionValuesEqual } from './chart-options.internal.js';
import {
  simpleLinesEqual,
  type NormalizedSimpleLine,
} from './simple-line.internal.js';
import { cloneOwnedPartForMutation } from './dependency.internal.js';
import { renderChartPart } from './chart-render.internal.js';
import {
  readChartState,
  readPromotableSeriesDataLabels,
} from './chart-state.internal.js';
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
  workbookBytes: Uint8Array | undefined,
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
    const workbook = workbookBytes === undefined
      ? retainWorkbook(pkg, isolated.chartPartUri, isolatedState)
      : replaceWorkbook(pkg, isolated.chartPartUri, isolatedState, workbookBytes);
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
      throw new Error(
        `Chart semantic replacement did not produce the requested state: ${result.status} ${result.reason ?? ''}`.trim(),
      );
    }
  });
}

export function chartDefinitionsEqual(
  left: Readonly<ChartDefinition>,
  right: Readonly<ChartDefinition>,
): boolean {
  return chartDefinitionDataEqual(left, right)
    && renderDefinitionForComparison(left) === renderDefinitionForComparison(right);
}

export function chartDefinitionDataEqual(
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

function retainWorkbook(
  pkg: OpcPackage,
  chartPartUri: string,
  state: Readonly<ChartState>,
): { readonly relationshipId: string; readonly partUri: string } {
  if (state.status !== 'recognized' || !state.workbookPartUri) {
    throw new Error('Option-only chart editing requires a synchronized workbook');
  }
  const relationships = pkg.relationships(chartPartUri).filter(
    ({ type, targetMode, resolvedTarget }) =>
      type === PACKAGE_RELATIONSHIP
      && targetMode === 'Internal'
      && resolvedTarget === state.workbookPartUri,
  );
  if (relationships.length !== 1) throw new Error('Chart workbook relationship is ambiguous');
  return { relationshipId: relationships[0]!.id, partUri: state.workbookPartUri };
}

function patchChartPart(
  source: Uint8Array,
  canonicalSource: string,
  current: Readonly<ChartDefinition>,
  next: Readonly<ChartDefinition>,
  addExternalData: boolean,
): string {
  let xml = LosslessXmlDocument.parse(source);
  const canonical = LosslessXmlDocument.parse(canonicalSource);
  const canonicalPlotArea = requirePlotArea(canonical);
  const structureMatches = sameStructure(current, next);
  const optionsMatch = chartOptionsEqual(current, next);
  if (structureMatches) {
    if (!optionsMatch) {
      patchChartOptions(xml, canonical, true, current, next);
      xml = LosslessXmlDocument.parse(xml.serialize());
    }
    patchSeriesData(xml, requirePlotArea(xml), canonical, canonicalPlotArea);
  } else {
    replaceOwnedPlotSpans(xml, requirePlotArea(xml), canonical, canonicalPlotArea);
    if (!optionsMatch) {
      xml = LosslessXmlDocument.parse(xml.serialize());
      patchChartOptions(xml, canonical, false, current, next);
    }
  }
  if (addExternalData) {
    xml = LosslessXmlDocument.parse(xml.serialize());
    insertExternalData(xml, canonical);
  }
  return xml.serialize();
}

function patchChartOptions(
  xml: LosslessXmlDocument,
  canonical: LosslessXmlDocument,
  patchPlotStructures: boolean,
  current: Readonly<ChartDefinition>,
  next: Readonly<ChartDefinition>,
): void {
  const root = xml.roots[0]!;
  const canonicalRoot = canonical.roots[0]!;
  const currentOptions = current.options;
  const nextOptions = next.options;
  const rootNames = [
    ...changedOptionName(currentOptions.language, nextOptions.language, 'lang'),
    ...changedOptionName(currentOptions.roundedCorners, nextOptions.roundedCorners, 'roundedCorners'),
    ...changedOptionName(currentOptions.style, nextOptions.style, 'style'),
    ...changedOptionName(currentOptions.chartArea, nextOptions.chartArea, 'spPr'),
  ];
  syncChartChildren(xml, root, canonical, canonicalRoot, rootNames);
  const chart = exactlyOneDirectChartChild(root, 'chart');
  const canonicalChart = exactlyOneDirectChartChild(canonicalRoot, 'chart');
  const chartNames = [
    ...changedOptionNames(currentOptions.title, nextOptions.title, ['title', 'autoTitleDeleted']),
    ...changedOptionName(currentOptions.legend, nextOptions.legend, 'legend'),
    ...changedOptionName(currentOptions.displayBlanksAs, nextOptions.displayBlanksAs, 'dispBlanksAs'),
    ...changedOptionNames(
      [
        currentOptions.rotationX,
        currentOptions.rotationY,
        currentOptions.rightAngleAxes,
        currentOptions.perspective,
      ],
      [
        nextOptions.rotationX,
        nextOptions.rotationY,
        nextOptions.rightAngleAxes,
        nextOptions.perspective,
      ],
      ['view3D'],
    ),
  ];
  syncChartChildren(xml, chart, canonical, canonicalChart, chartNames);
  const plotArea = exactlyOneDirectChartChild(chart, 'plotArea');
  const canonicalPlotArea = exactlyOneDirectChartChild(canonicalChart, 'plotArea');
  const plotNames = [
    ...changedOptionName(currentOptions.layout, nextOptions.layout, 'layout'),
    ...changedOptionName(currentOptions.dataTable, nextOptions.dataTable, 'dTable'),
    ...changedOptionName(currentOptions.plotArea, nextOptions.plotArea, 'spPr'),
  ];
  syncChartChildren(xml, plotArea, canonical, canonicalPlotArea, plotNames);
  if (!patchPlotStructures) return;

  const groups = directChartChildren(plotArea).filter(({ localName }) => GROUP_NAMES.has(localName));
  const canonicalGroups = directChartChildren(canonicalPlotArea)
    .filter(({ localName }) => GROUP_NAMES.has(localName));
  groups.forEach((group, groupIndex) => {
    const canonicalGroup = canonicalGroups[groupIndex]!;
    const currentGroup = current.groups[groupIndex]!;
    const nextGroup = next.groups[groupIndex]!;
    const dataLabelsChanged = !dataLabelOptionsEqual(
      currentGroup.options?.dataLabels,
      nextGroup.options?.dataLabels,
    );
    const optionNames = changedGroupOptionElementNames(
      group.localName,
      currentGroup.options,
      nextGroup.options,
    );
    syncChartChildren(xml, group, canonical, canonicalGroup, optionNames);
    const series = directChartChildren(group).filter(({ localName }) => localName === 'ser');
    const canonicalSeries = directChartChildren(canonicalGroup).filter(({ localName }) => localName === 'ser');
    series.forEach((entry, seriesIndex) => {
      const currentSeriesOptions = currentGroup.options?.series?.[seriesIndex];
      const nextSeriesOptions = nextGroup.options?.series?.[seriesIndex];
      const seriesCapPatched = patchChartLineCap(
        xml,
        entry,
        currentSeriesOptions?.line,
        nextSeriesOptions?.line,
      );
      const seriesShapeChanged = !chartOptionValuesEqual(
        [currentSeriesOptions?.fill, currentSeriesOptions?.shadow, currentOptions.colors],
        [nextSeriesOptions?.fill, nextSeriesOptions?.shadow, nextOptions.colors],
      ) || (!lineValuesEqual(currentSeriesOptions?.line, nextSeriesOptions?.line)
        && !seriesCapPatched);
      const seriesNames = [
        ...(seriesShapeChanged ? ['spPr'] : []),
        ...changedOptionNames(
          [currentSeriesOptions?.marker, optionProperty(currentGroup.options, 'marker')],
          [nextSeriesOptions?.marker, optionProperty(nextGroup.options, 'marker')],
          ['marker'],
        ),
        ...changedOptionNames(
          [
            optionProperty(currentGroup.options, 'smooth'),
            optionProperty(currentGroup.options, 'style'),
          ],
          [
            optionProperty(nextGroup.options, 'smooth'),
            optionProperty(nextGroup.options, 'style'),
          ],
          ['smooth'],
        ),
      ];
      syncChartChildren(xml, entry, canonical, canonicalSeries[seriesIndex]!, seriesNames);
      patchSeriesPointStyles(
        xml,
        entry,
        canonical,
        canonicalSeries[seriesIndex]!,
        currentSeriesOptions?.points,
        nextSeriesOptions?.points,
      );
      patchSeriesDataLabels(
        xml,
        entry,
        canonical,
        canonicalSeries[seriesIndex]!,
        currentSeriesOptions?.dataLabels,
        nextSeriesOptions?.dataLabels,
      );
      if (
        dataLabelsChanged
        && readPromotableSeriesDataLabels(
          entry,
          currentGroup.series[seriesIndex]!.values.length,
          currentGroup.options?.dataLabels,
        ) !== undefined
      ) {
        const labels = directChartChildren(entry).filter(({ localName }) => localName === 'dLbls');
        if (labels.length === 1) xml.removeElement(labels[0]!);
      }
    });
  });

  const axes = directChartChildren(plotArea).filter(({ localName }) => AXIS_NAMES.has(localName));
  const canonicalAxes = directChartChildren(canonicalPlotArea)
    .filter(({ localName }) => AXIS_NAMES.has(localName));
  axes.forEach((axis, index) => {
    const canonicalAxis = canonicalAxes[index]!;
    const currentAxisOptions = axisSemanticOptions(current, index);
    const nextAxisOptions = axisSemanticOptions(next, index);
    const locallyPatched = new Set<string>();
    for (const [key, ownerName] of [
      ['line', 'spPr'],
      ['majorGridLine', 'majorGridlines'],
      ['minorGridLine', 'minorGridlines'],
    ] as const) {
      const owner = ownerName === 'spPr'
        ? axis
        : directChartChildren(axis).find(({ localName }) => localName === ownerName);
      if (
        owner
        && patchChartLineCap(
          xml,
          owner,
          optionProperty(currentAxisOptions, key),
          optionProperty(nextAxisOptions, key),
        )
      ) locallyPatched.add(ownerName);
    }
    const optionNames = changedAxisElementNames(currentAxisOptions, nextAxisOptions)
      .filter((name) => !locallyPatched.has(name));
    syncChartChildren(
      xml,
      axis,
      canonical,
      canonicalAxis,
      optionNames,
    );
    if (axis.localName !== canonicalAxis.localName) {
      renameChartElement(xml, axis, canonicalAxis.name);
    }
  });
}

function dataLabelOptionsEqual(left: unknown, right: unknown): boolean {
  if (left === undefined || right === undefined) return left === right;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  const canonical = (value: object): Record<string, unknown> => Object.fromEntries(
    Object.entries(value).filter(([, entry]) => entry !== false),
  );
  return chartOptionValuesEqual(canonical(left), canonical(right));
}

function changedOptionName(left: unknown, right: unknown, name: string): string[] {
  return chartOptionValuesEqual(left, right) ? [] : [name];
}

function lineValuesEqual(left: unknown, right: unknown): boolean {
  const leftLine = left as NormalizedSimpleLine | undefined;
  const rightLine = right as NormalizedSimpleLine | undefined;
  return rightLine === undefined ? leftLine === undefined : simpleLinesEqual(leftLine, rightLine);
}

function patchChartLineCap(
  xml: LosslessXmlDocument,
  owner: XmlElement,
  currentValue: unknown,
  nextValue: unknown,
): boolean {
  const current = currentValue as NormalizedSimpleLine | undefined;
  const next = nextValue as NormalizedSimpleLine | undefined;
  if (
    current?.kind !== 'line'
    || next?.kind !== 'line'
    || simpleLinesEqual(current, next)
    || !simpleLinesEqual(withoutLineCap(current), withoutLineCap(next))
  ) return false;
  const properties = owner.localName === 'spPr'
    ? owner
    : directChartChildren(owner).filter(({ localName }) => localName === 'spPr')[0];
  if (!properties) return false;
  const lines = properties.children.filter((child): child is XmlElement =>
    child.type === 'element'
      && namespaceUri(child) === 'http://schemas.openxmlformats.org/drawingml/2006/main'
      && child.localName === 'ln');
  if (lines.length !== 1) return false;
  const line = lines[0]!;
  const caps = line.attributes.filter(({ name }) => name === 'cap');
  if (caps.length > 1) return false;
  const target = next.cap === undefined
    ? undefined
    : ({ flat: 'flat', round: 'rnd', square: 'sq' } as const)[next.cap];
  const cap = caps[0];
  if (cap && target !== undefined) {
    if (cap.value === target) return true;
    xml.replace(cap.valueStart, cap.valueEnd, target);
    return true;
  }
  if (cap) {
    let start = cap.start;
    while (start > line.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
    xml.replace(start, cap.end, '');
    return true;
  }
  if (target === undefined) return true;
  const position = line.selfClosing
    ? xml.source.lastIndexOf('/', line.startTagEnd - 1)
    : line.startTagEnd - 1;
  if (position <= line.start) return false;
  xml.replace(position, position, ` cap="${target}"`);
  return true;
}

function withoutLineCap(line: Extract<NormalizedSimpleLine, { readonly kind: 'line' }>): NormalizedSimpleLine {
  const { cap: _cap, ...rest } = line;
  return rest;
}

function patchSeriesPointStyles(
  xml: LosslessXmlDocument,
  series: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalSeries: XmlElement,
  current: readonly Readonly<{ readonly index: number }>[] | undefined,
  next: readonly Readonly<{ readonly index: number }>[] | undefined,
): void {
  if (chartOptionValuesEqual(current, next)) return;
  const existing = pointElementMap(series);
  const replacements = pointElementMap(canonicalSeries);
  const currentByIndex = new Map((current ?? []).map((point) => [point.index, point]));
  const nextByIndex = new Map((next ?? []).map((point) => [point.index, point]));
  const indexes = new Set([...currentByIndex.keys(), ...nextByIndex.keys()]);
  for (const index of [...indexes].sort((left, right) => left - right)) {
    if (chartOptionValuesEqual(currentByIndex.get(index), nextByIndex.get(index))) continue;
    const target = existing.get(index);
    const replacement = replacements.get(index);
    if (target && replacement) {
      xml.replaceElement(target, canonical.original(replacement));
    } else if (target) {
      xml.removeElement(target);
    } else if (replacement) {
      insertPointStyle(xml, series, canonical, replacement, index);
    }
  }
}

function pointElementMap(series: XmlElement): Map<number, XmlElement> {
  const result = new Map<number, XmlElement>();
  for (const point of directChartChildren(series).filter(({ localName }) => localName === 'dPt')) {
    const indexes = directChartChildren(point).filter(({ localName }) => localName === 'idx');
    const raw = indexes.length === 1
      ? indexes[0]!.attributes.find(({ name }) => name === 'val')?.value
      : undefined;
    const index = raw === undefined ? Number.NaN : Number(raw);
    if (!Number.isSafeInteger(index) || index < 0 || result.has(index)) {
      throw new Error('Chart point style indexes are ambiguous');
    }
    result.set(index, point);
  }
  return result;
}

function insertPointStyle(
  xml: LosslessXmlDocument,
  series: XmlElement,
  canonical: LosslessXmlDocument,
  replacement: XmlElement,
  index: number,
): void {
  const rendered = canonical.original(replacement);
  const followingPoint = [...pointElementMap(series)]
    .filter(([candidate]) => candidate > index)
    .sort(([left], [right]) => left - right)[0]?.[1];
  if (followingPoint) {
    xml.replace(followingPoint.start, followingPoint.start, rendered);
    return;
  }
  const canonicalSiblings = directChartChildren(replacement.parent!);
  const replacementIndex = canonicalSiblings.indexOf(replacement);
  for (const sibling of canonicalSiblings.slice(replacementIndex + 1)) {
    if (sibling.localName === 'dPt') continue;
    const anchor = directChartChildren(series).find(({ localName }) => localName === sibling.localName);
    if (anchor) {
      xml.replace(anchor.start, anchor.start, rendered);
      return;
    }
  }
  const extension = directChartChildren(series).find(({ localName }) => localName === 'extLst');
  if (extension) xml.replace(extension.start, extension.start, rendered);
  else xml.appendChildXml(series, rendered);
}

function patchSeriesDataLabels(
  xml: LosslessXmlDocument,
  series: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalSeries: XmlElement,
  current: Readonly<object> | undefined,
  next: Readonly<object> | undefined,
): void {
  if (dataLabelOptionsEqual(current, next)) return;
  const existing = directChartChildren(series).filter(({ localName }) => localName === 'dLbls');
  const replacements = directChartChildren(canonicalSeries)
    .filter(({ localName }) => localName === 'dLbls');
  if (existing.length !== 1 || replacements.length !== 1) {
    syncChartChildren(xml, series, canonical, canonicalSeries, ['dLbls']);
    return;
  }
  patchSeriesPointLabels(
    xml,
    existing[0]!,
    canonical,
    replacements[0]!,
    optionProperty(current, 'pointLabels') as readonly Readonly<{ readonly index: number }>[] | undefined,
    optionProperty(next, 'pointLabels') as readonly Readonly<{ readonly index: number }>[] | undefined,
  );
  const mappings: readonly [readonly string[], string][] = [
    [['numberFormat'], 'numFmt'],
    [['fill'], 'spPr'],
    [['face', 'size', 'bold', 'italic', 'color'], 'txPr'],
    [['position'], 'dLblPos'],
    [['showValue'], 'showVal'],
    [['showCategoryName'], 'showCatName'],
    [['showSeriesName'], 'showSerName'],
    [['showPercent'], 'showPercent'],
    [['showBubbleSize'], 'showBubbleSize'],
    [['showLeaderLines'], 'showLeaderLines'],
  ];
  const names = mappings.flatMap(([keys, element]) =>
    chartOptionValuesEqual(
      keys.map((key) => optionProperty(current, key)),
      keys.map((key) => optionProperty(next, key)),
    ) ? [] : [element]);
  syncChartChildren(xml, existing[0]!, canonical, replacements[0]!, names);
}

function patchSeriesPointLabels(
  xml: LosslessXmlDocument,
  labels: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalLabels: XmlElement,
  current: readonly Readonly<{ readonly index: number }>[] | undefined,
  next: readonly Readonly<{ readonly index: number }>[] | undefined,
): void {
  if (chartOptionValuesEqual(current, next)) return;
  const existing = indexedChartChildMap(labels, 'dLbl');
  const replacements = indexedChartChildMap(canonicalLabels, 'dLbl');
  const currentByIndex = new Map((current ?? []).map((label) => [label.index, label]));
  const nextByIndex = new Map((next ?? []).map((label) => [label.index, label]));
  const indexes = new Set([...currentByIndex.keys(), ...nextByIndex.keys()]);
  for (const index of [...indexes].sort((left, right) => left - right)) {
    if (chartOptionValuesEqual(currentByIndex.get(index), nextByIndex.get(index))) continue;
    const target = existing.get(index);
    const replacement = replacements.get(index);
    if (target && replacement) {
      xml.replaceElement(target, canonical.original(replacement));
    } else if (target) {
      xml.removeElement(target);
    } else if (replacement) {
      insertIndexedChartChild(xml, labels, canonical, replacement, 'dLbl', index);
    }
  }
}

function indexedChartChildMap(parent: XmlElement, localName: string): Map<number, XmlElement> {
  const result = new Map<number, XmlElement>();
  for (const child of directChartChildren(parent).filter((entry) => entry.localName === localName)) {
    const indexes = directChartChildren(child).filter(({ localName: name }) => name === 'idx');
    const raw = indexes.length === 1
      ? indexes[0]!.attributes.find(({ name }) => name === 'val')?.value
      : undefined;
    const index = raw === undefined ? Number.NaN : Number(raw);
    if (!Number.isSafeInteger(index) || index < 0 || result.has(index)) {
      throw new Error(`Chart ${localName} indexes are ambiguous`);
    }
    result.set(index, child);
  }
  return result;
}

function insertIndexedChartChild(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  canonical: LosslessXmlDocument,
  replacement: XmlElement,
  localName: string,
  index: number,
): void {
  const rendered = canonical.original(replacement);
  const following = [...indexedChartChildMap(parent, localName)]
    .filter(([candidate]) => candidate > index)
    .sort(([left], [right]) => left - right)[0]?.[1];
  if (following) {
    xml.replace(following.start, following.start, rendered);
    return;
  }
  const canonicalSiblings = directChartChildren(replacement.parent!);
  const replacementIndex = canonicalSiblings.indexOf(replacement);
  for (const sibling of canonicalSiblings.slice(replacementIndex + 1)) {
    if (sibling.localName === localName) continue;
    const anchor = directChartChildren(parent).find((entry) => entry.localName === sibling.localName);
    if (anchor) {
      xml.replace(anchor.start, anchor.start, rendered);
      return;
    }
  }
  const extension = directChartChildren(parent).find(({ localName: name }) => name === 'extLst');
  if (extension) xml.replace(extension.start, extension.start, rendered);
  else xml.appendChildXml(parent, rendered);
}

function optionProperty(value: unknown, key: string): unknown {
  return value && typeof value === 'object'
    ? (value as Readonly<Record<string, unknown>>)[key]
    : undefined;
}

function changedOptionNames(
  left: unknown,
  right: unknown,
  names: readonly string[],
): string[] {
  return chartOptionValuesEqual(left, right) ? [] : [...names];
}

function changedGroupOptionElementNames(
  localName: string,
  current: Readonly<Record<string, unknown>> | undefined,
  next: Readonly<Record<string, unknown>> | undefined,
): string[] {
  const mapping: Readonly<Record<string, readonly [string, string][]>> = {
    areaChart: [['grouping', 'grouping'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls']],
    barChart: [['direction', 'barDir'], ['grouping', 'grouping'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls'], ['gapWidth', 'gapWidth'], ['overlap', 'overlap']],
    bar3DChart: [['direction', 'barDir'], ['grouping', 'grouping'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls'], ['gapWidth', 'gapWidth'], ['gapDepth', 'gapDepth'], ['shape', 'shape']],
    bubbleChart: [['varyColors', 'varyColors'], ['dataLabels', 'dLbls'], ['scale', 'bubbleScale'], ['showNegativeBubbles', 'showNegBubbles'], ['sizeRepresents', 'sizeRepresents']],
    doughnutChart: [['varyColors', 'varyColors'], ['dataLabels', 'dLbls'], ['firstSliceAngle', 'firstSliceAng'], ['holeSize', 'holeSize']],
    lineChart: [['grouping', 'grouping'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls']],
    pieChart: [['varyColors', 'varyColors'], ['dataLabels', 'dLbls'], ['firstSliceAngle', 'firstSliceAng']],
    radarChart: [['style', 'radarStyle'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls']],
    scatterChart: [['style', 'scatterStyle'], ['varyColors', 'varyColors'], ['dataLabels', 'dLbls']],
  };
  return (mapping[localName] ?? []).flatMap(([key, element]) => {
    const equal = key === 'dataLabels'
      ? dataLabelOptionsEqual(current?.[key], next?.[key])
      : chartOptionValuesEqual(current?.[key], next?.[key]);
    return equal ? [] : [element];
  });
}

function axisSemanticOptions(
  definition: Readonly<ChartDefinition>,
  index: number,
): unknown {
  const options = definition.options;
  if (definition.groups[0]?.type === 'bar3D') {
    return [options.categoryAxis, options.valueAxis, options.seriesAxis][index];
  }
  return [
    options.categoryAxis,
    options.valueAxis,
    options.secondaryCategoryAxis,
    options.secondaryValueAxis,
  ][index];
}

function changedAxisElementNames(current: unknown, next: unknown): string[] {
  const mappings: readonly [readonly string[], readonly string[]][] = [
    [['kind'], ['scaling', 'baseTimeUnit', 'majorTimeUnit', 'minorTimeUnit', 'crossBetween']],
    [['minimum', 'maximum', 'logarithmicBase', 'orientation'], ['scaling']],
    [['visible'], ['delete']],
    [['position'], ['axPos']],
    [['majorGridLine'], ['majorGridlines']],
    [['minorGridLine'], ['minorGridlines']],
    [['title'], ['title']],
    [['numberFormat'], ['numFmt']],
    [['majorTickMark'], ['majorTickMark']],
    [['minorTickMark'], ['minorTickMark']],
    [['labelPosition'], ['tickLblPos']],
    [['line'], ['spPr']],
    [['face', 'size', 'bold', 'italic', 'color', 'labelRotation'], ['txPr']],
    [['crossesAt'], ['crosses', 'crossesAt']],
    [['baseTimeUnit'], ['baseTimeUnit']],
    [['majorUnit'], ['majorUnit']],
    [['majorTimeUnit'], ['majorTimeUnit']],
    [['minorUnit'], ['minorUnit']],
    [['minorTimeUnit'], ['minorTimeUnit']],
    [['labelFrequency'], ['tickLblSkip']],
    [['multiLevelLabels'], ['noMultiLvlLbl']],
    [['displayUnit', 'displayUnitLabel'], ['dispUnits']],
  ];
  const names = new Set<string>();
  for (const [keys, elements] of mappings) {
    const left = keys.map((key) => optionProperty(current, key));
    const right = keys.map((key) => optionProperty(next, key));
    const equal = keys.length === 1 && ['line', 'majorGridLine', 'minorGridLine'].includes(keys[0]!)
      ? lineValuesEqual(left[0], right[0])
      : chartOptionValuesEqual(left, right);
    if (!equal) elements.forEach((name) => names.add(name));
  }
  return [...names];
}

function renameChartElement(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
): void {
  if (element.selfClosing || element.endTagStart < element.startTagEnd) {
    throw new Error(`Chart ${element.localName} cannot be renamed safely`);
  }
  const opening = xml.source.slice(element.start, element.startTagEnd);
  if (!opening.startsWith(`<${element.name}`)) {
    throw new Error(`Chart ${element.localName} opening tag is malformed`);
  }
  xml.replace(
    element.start,
    element.startTagEnd,
    `<${name}${opening.slice(element.name.length + 1)}`,
  );
  xml.replace(element.endTagStart, element.end, `</${name}>`);
}

function syncChartChildren(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalParent: XmlElement,
  localNames: readonly string[],
): void {
  for (const localName of localNames) {
    const existing = directChartChildren(parent).filter((child) => child.localName === localName);
    const replacements = directChartChildren(canonicalParent)
      .filter((child) => child.localName === localName);
    if (replacements.length > 1) throw new Error(`Canonical chart ${localName} is ambiguous`);
    const replacement = replacements[0];
    if (replacement && existing[0]) {
      xml.replaceElement(existing[0], canonical.original(replacement));
      for (const duplicate of existing.slice(1)) xml.removeElement(duplicate);
      continue;
    }
    if (!replacement) {
      for (const entry of existing) xml.removeElement(entry);
      continue;
    }
    insertCanonicalChild(xml, parent, canonical, canonicalParent, replacement);
  }
}

function insertCanonicalChild(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  canonical: LosslessXmlDocument,
  canonicalParent: XmlElement,
  replacement: XmlElement,
): void {
  const canonicalSiblings = directChartChildren(canonicalParent);
  const replacementIndex = canonicalSiblings.indexOf(replacement);
  for (const sibling of canonicalSiblings.slice(replacementIndex + 1)) {
    const anchor = directChartChildren(parent).find(({ localName }) =>
      localName === sibling.localName);
    if (anchor) {
      xml.replace(anchor.start, anchor.start, canonical.original(replacement));
      return;
    }
  }
  const extension = directChartChildren(parent).find(({ localName }) => localName === 'extLst');
  if (extension) {
    xml.replace(extension.start, extension.start, canonical.original(replacement));
    return;
  }
  xml.appendChildXml(parent, canonical.original(replacement));
}

function groupOptionElementNames(localName: string): readonly string[] {
  switch (localName) {
    case 'areaChart': return ['grouping', 'varyColors', 'dLbls'];
    case 'barChart': return ['barDir', 'grouping', 'varyColors', 'dLbls', 'gapWidth', 'overlap'];
    case 'bar3DChart': return ['barDir', 'grouping', 'varyColors', 'dLbls', 'gapWidth', 'gapDepth', 'shape'];
    case 'bubbleChart': return ['varyColors', 'dLbls', 'bubbleScale', 'showNegBubbles', 'sizeRepresents'];
    case 'doughnutChart': return ['varyColors', 'dLbls', 'firstSliceAng', 'holeSize'];
    case 'lineChart': return ['grouping', 'varyColors', 'dLbls'];
    case 'pieChart': return ['varyColors', 'dLbls', 'firstSliceAng'];
    case 'radarChart': return ['radarStyle', 'varyColors', 'dLbls'];
    case 'scatterChart': return ['scatterStyle', 'varyColors', 'dLbls'];
    default: return [];
  }
}

function chartOptionsEqual(
  left: Readonly<ChartDefinition>,
  right: Readonly<ChartDefinition>,
): boolean {
  if (chartDefinitionDataEqual(left, right)) {
    return renderDefinitionForComparison(left) === renderDefinitionForComparison(right);
  }
  return chartOptionValuesEqual(left.options, right.options)
    && left.groups.length === right.groups.length
    && left.groups.every((group, index) =>
      chartOptionValuesEqual(group.options, right.groups[index]?.options));
}

function renderDefinitionForComparison(definition: Readonly<ChartDefinition>): string {
  const plan = planChartWorkbook(definition);
  return renderChartPart(definition, plan.formulas, 'rIdWorkbook');
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
