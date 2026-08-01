import { LosslessXmlDocument, type XmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage, Relationship } from '@pptx/opc';
import type { ChartDiagnostic, ChartDiagnosticCode, ChartState } from './chart.js';
import { readChartState } from './chart-state.internal.js';
import { chartWorkbookMatches } from './chart-workbook.internal.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MODERN_CHART_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/chart`;
const MODERN_CHART_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/chartEx`;

export async function chartDiagnostics(
  pkg: OpcPackage,
  slidePartUri: string,
): Promise<readonly ChartDiagnostic[]> {
  const part = pkg.getPart(slidePartUri);
  if (!part) return Object.freeze([]);
  let xml: LosslessXmlDocument;
  try {
    xml = LosslessXmlDocument.parse(part.bytes);
  } catch {
    return Object.freeze([]);
  }
  const relationships = pkg.relationships(slidePartUri);
  const diagnostics: ChartDiagnostic[] = [];
  for (const frame of xml.elements('graphicFrame')) {
    const references = xml.descendants(frame, 'chart').filter((element) => {
      const namespace = elementNamespaceUri(element);
      return namespace === CHART_NAMESPACE || namespace === MODERN_CHART_NAMESPACE;
    });
    if (references.length === 0) continue;
    const objectId = readShapeId(xml, frame);
    if (references.length !== 1) {
      diagnostics.push(diagnostic(
        'error',
        'CHART_STRUCTURE_AMBIGUOUS',
        'Chart graphic frame must contain exactly one chart reference',
        slidePartUri,
        objectId,
      ));
      continue;
    }
    const reference = references[0]!;
    const ids = relationshipAttributes(reference);
    if (ids.length !== 1) {
      diagnostics.push(diagnostic(
        'error',
        'CHART_RELATIONSHIP_INVALID',
        'Chart reference must contain exactly one relationship id',
        slidePartUri,
        objectId,
      ));
      continue;
    }
    const matches = relationships.filter(({ id }) => id === ids[0]!.value);
    const relationship = matches.length === 1 ? matches[0] : undefined;
    if (!relationship || !validChartRelationship(pkg, relationship)) {
      diagnostics.push(diagnostic(
        'error',
        'CHART_RELATIONSHIP_INVALID',
        `Chart relationship ${ids[0]!.value} is missing, repeated, external, wrong-type, or dangling`,
        slidePartUri,
        objectId,
      ));
      continue;
    }
    diagnostics.push(...await diagnosticsForState(
      pkg,
      relationship.resolvedTarget!,
      readChartState(pkg, relationship.resolvedTarget!),
      objectId,
    ));
  }
  return Object.freeze(diagnostics.map((entry) => Object.freeze(entry)));
}

async function diagnosticsForState(
  pkg: OpcPackage,
  chartPartUri: string,
  state: Readonly<ChartState>,
  objectId: string | undefined,
): Promise<readonly ChartDiagnostic[]> {
  if (state.status === 'modern') {
    return [diagnostic(
      'info',
      'MODERN_CHART_EXTENSION',
      state.reason ?? 'Modern chart extensions may degrade in older clients',
      chartPartUri,
      objectId,
    )];
  }
  if (state.status === 'ambiguous') {
    return [diagnostic(
      'error',
      'CHART_STRUCTURE_AMBIGUOUS',
      state.reason ?? 'Chart structure is ambiguous',
      chartPartUri,
      objectId,
    )];
  }
  if (state.status === 'unsupported') {
    const code = classifyUnsupportedReason(state.reason);
    return [diagnostic(
      'error',
      code,
      state.reason ?? 'Chart structure is unsupported',
      chartPartUri,
      objectId,
    )];
  }
  if (state.status === 'cache-only') {
    return [diagnostic(
      'warning',
      'CHART_WORKBOOK_MISSING',
      'Chart has valid display caches but no embedded workbook',
      chartPartUri,
      objectId,
    )];
  }
  const matches = state.definition
    && state.workbookPartUri
    && pkg.hasPart(state.workbookPartUri)
    && await chartWorkbookMatches(pkg.requirePart(state.workbookPartUri).bytes, state.definition);
  return matches
    ? []
    : [diagnostic(
        'error',
        'CHART_WORKBOOK_CACHE_DIVERGENCE',
        'Chart caches/formulas and embedded workbook cells disagree',
        chartPartUri,
        objectId,
      )];
}

function classifyUnsupportedReason(reason: string | undefined): ChartDiagnosticCode {
  const value = reason?.toLowerCase() ?? '';
  if (value.includes('axis') || value.includes('axes')) return 'CHART_AXIS_INVALID';
  if (
    value.includes('cache')
    || value.includes('formula')
    || value.includes('point')
    || value.includes('vector')
    || value.includes('finite decimal')
    || value.includes('categories and values')
    || value.includes('xvalues and values')
    || value.includes('sizes and values')
  ) return 'CHART_CACHE_INVALID';
  if (
    value.includes('relationship')
    || value.includes('externaldata')
    || value.includes('workbook')
    || value.includes('xlsx')
  ) return 'CHART_RELATIONSHIP_INVALID';
  return 'CHART_STRUCTURE_UNSUPPORTED';
}

function validChartRelationship(pkg: OpcPackage, relationship: Relationship): boolean {
  return relationship.targetMode === 'Internal'
    && relationship.resolvedTarget !== undefined
    && pkg.hasPart(relationship.resolvedTarget)
    && (
      relationship.type === CHART_RELATIONSHIP
      || relationship.type === MODERN_CHART_RELATIONSHIP
      || relationship.type.endsWith('/chartEx')
    );
}

function diagnostic(
  severity: ChartDiagnostic['severity'],
  code: ChartDiagnosticCode,
  message: string,
  partUri: string,
  objectId: string | undefined,
): ChartDiagnostic {
  return {
    severity,
    code,
    message,
    partUri,
    ...(objectId === undefined ? {} : { objectId }),
  };
}

function readShapeId(xml: LosslessXmlDocument, frame: XmlElement): string | undefined {
  const properties = xml.descendants(frame, 'cNvPr')[0];
  return properties ? xml.attribute(properties, 'id')?.value : undefined;
}

function relationshipAttributes(element: XmlElement): readonly XmlAttribute[] {
  return element.attributes.filter((attribute) =>
    localName(attribute.name) === 'id'
    && attributeNamespaceUri(element, attribute) === RELATIONSHIP_NAMESPACE);
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(attribute.name));
}

function namespaceUriForPrefix(element: XmlElement, prefix: string): string | undefined {
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    for (const attribute of scope.attributes) {
      if (attribute.name === 'xmlns' && prefix === '') return attribute.value;
      if (attribute.name === `xmlns:${prefix}`) return attribute.value;
    }
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function localName(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}
