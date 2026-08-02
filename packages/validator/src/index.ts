import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage, Relationship } from '@pptx/opc';

export type DiagnosticSeverity = 'info' | 'warning' | 'error';
export type CompatibilityProfile =
  | 'powerpoint-2010'
  | 'powerpoint-current'
  | 'keynote-current'
  | 'libreoffice-current'
  | 'google-slides-import';

export interface Diagnostic {
  readonly severity: DiagnosticSeverity;
  readonly code: string;
  readonly message: string;
  readonly partUri?: string;
  readonly xmlPath?: string;
  readonly objectId?: string;
  readonly compatibility?: CompatibilityProfile;
  readonly suggestion?: string;
}

export class ValidationError extends Error {
  constructor(readonly diagnostics: readonly Diagnostic[]) {
    super(`Validation failed with ${diagnostics.filter(({ severity }) => severity === 'error').length} error(s)`);
    this.name = 'ValidationError';
  }
}

export type MasterLayoutDiagnosticCode =
  | 'LAYOUT_NAME_DUPLICATE'
  | 'LAYOUT_RELATIONSHIP_INVALID'
  | 'PLACEHOLDER_IDENTITY_AMBIGUOUS'
  | 'PLACEHOLDER_OWNER_MISSING'
  | 'PLACEHOLDER_DOMAIN_MISMATCH';

const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const CHART_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/chart';
const POWERPOINT_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/powerpoint/2010/main';
const PRESENTATION_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
]);
const MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const MASTER_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}slideMaster`;
const LAYOUT_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}slideLayout`;
const OFFICE_DOCUMENT_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}officeDocument`;
const PLACEHOLDER_TYPES = new Set(['title', 'body', 'pic', 'chart', 'tbl', 'media']);

type PlaceholderDomain = 'empty' | 'text-shape' | 'image' | 'chart' | 'table' | 'media' | 'unknown';

interface PlaceholderRecord {
  readonly identity: string;
  readonly type: string;
  readonly index: number;
  readonly name?: string;
  readonly domain: PlaceholderDomain;
}

interface LayoutRecord {
  readonly partUri: string;
  readonly name?: string;
  readonly placeholders: readonly PlaceholderRecord[];
}

type MasterLayoutDiagnosticEmitter = (
  severity: DiagnosticSeverity,
  code: MasterLayoutDiagnosticCode,
  message: string,
  partUri: string,
  objectId: string,
  suggestion: string,
) => void;

export function validateMasterLayoutPlaceholders(
  pkg: OpcPackage,
  compatibility?: CompatibilityProfile,
): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const emitted = new Set<string>();
  const emit = (
    severity: DiagnosticSeverity,
    code: MasterLayoutDiagnosticCode,
    message: string,
    partUri: string,
    objectId: string,
    suggestion: string,
  ): void => {
    const key = `${code}\u0000${partUri}\u0000${objectId}`;
    if (emitted.has(key)) return;
    emitted.add(key);
    diagnostics.push({
      severity,
      code,
      message,
      partUri,
      objectId,
      ...(compatibility === undefined ? {} : { compatibility }),
      suggestion,
    });
  };

  const presentationRelationships = pkg.relationships('/').filter(
    ({ type }) => type === OFFICE_DOCUMENT_RELATIONSHIP,
  );
  const presentationRelationship = presentationRelationships.find((relationship) =>
    isInternalTargetWithContentTypes(pkg, relationship, PRESENTATION_CONTENT_TYPES));
  if (!presentationRelationship?.resolvedTarget) return diagnostics;
  const presentationPartUri = presentationRelationship.resolvedTarget;

  const masterPartUris: string[] = [];
  for (const relationship of pkg.relationships(presentationPartUri).filter(
    ({ type }) => type === MASTER_RELATIONSHIP,
  )) {
    if (!isInternalTarget(pkg, relationship, MASTER_CONTENT_TYPE)) {
      emit(
        'error',
        'LAYOUT_RELATIONSHIP_INVALID',
        `Slide master relationship ${relationship.id} is not a valid internal master target`,
        presentationPartUri,
        relationship.id,
        'Retarget the relationship to one attached slideMaster part.',
      );
      continue;
    }
    masterPartUris.push(relationship.resolvedTarget!);
  }

  const layouts = new Map<string, LayoutRecord>();
  for (const masterPartUri of masterPartUris) {
    for (const relationship of pkg.relationships(masterPartUri).filter(
      ({ type }) => type === LAYOUT_RELATIONSHIP,
    )) {
      if (!isInternalTarget(pkg, relationship, LAYOUT_CONTENT_TYPE)) {
        emit(
          'error',
          'LAYOUT_RELATIONSHIP_INVALID',
          `Slide layout relationship ${relationship.id} is not a valid internal layout target`,
          masterPartUri,
          relationship.id,
          'Retarget the relationship to one slideLayout part.',
        );
        continue;
      }
      const layoutPartUri = relationship.resolvedTarget!;
      const backlinks = pkg.relationships(layoutPartUri).filter(
        ({ type }) => type === MASTER_RELATIONSHIP,
      );
      const backlink = backlinks[0];
      if (
        backlinks.length !== 1
        || !backlink
        || !isInternalTarget(pkg, backlink, MASTER_CONTENT_TYPE)
        || backlink.resolvedTarget !== masterPartUri
      ) {
        emit(
          'error',
          'LAYOUT_RELATIONSHIP_INVALID',
          'Slide layout does not have one internal backlink to its attached master',
          layoutPartUri,
          backlink?.id ?? 'slideMaster',
          'Keep exactly one slideMaster relationship pointing to the parent master.',
        );
        continue;
      }
      if (layouts.has(layoutPartUri)) {
        emit(
          'error',
          'LAYOUT_RELATIONSHIP_INVALID',
          'Slide layout is attached more than once',
          layoutPartUri,
          relationship.id,
          'Attach each slide layout exactly once.',
        );
        continue;
      }
      const xml = parsePart(pkg, layoutPartUri);
      if (!xml) continue;
      const name = directLayoutName(xml);
      layouts.set(layoutPartUri, {
        partUri: layoutPartUri,
        ...(name === undefined ? {} : { name }),
        placeholders: placeholderRecords(xml, 'sldLayout'),
      });
    }
  }

  const firstLayoutByName = new Map<string, string>();
  for (const layout of layouts.values()) {
    if (layout.name === undefined) continue;
    const first = firstLayoutByName.get(layout.name);
    if (first === undefined) {
      firstLayoutByName.set(layout.name, layout.partUri);
      continue;
    }
    emit(
      'error',
      'LAYOUT_NAME_DUPLICATE',
      `Slide layout name ${layout.name} is also used by ${first}`,
      layout.partUri,
      layout.name,
      'Give every attached slide layout a unique name.',
    );
  }

  const ambiguousLayoutIdentities = new Map<string, ReadonlySet<string>>();
  for (const layout of layouts.values()) {
    ambiguousLayoutIdentities.set(
      layout.partUri,
      reportAmbiguousPlaceholders(layout.placeholders, layout.partUri, emit),
    );
  }

  for (const slide of pkg.parts.filter(({ contentType }) => contentType === SLIDE_CONTENT_TYPE)) {
    const relationships = pkg.relationships(slide.uri).filter(
      ({ type }) => type === LAYOUT_RELATIONSHIP,
    );
    if (relationships.length === 0 && layouts.size === 0) continue;
    const relationship = relationships[0];
    if (
      relationships.length !== 1
      || !relationship
      || !isInternalTarget(pkg, relationship, LAYOUT_CONTENT_TYPE)
      || !relationship.resolvedTarget
      || !layouts.has(relationship.resolvedTarget)
    ) {
      emit(
        'error',
        'LAYOUT_RELATIONSHIP_INVALID',
        'Slide does not have one internal relationship to an attached slide layout',
        slide.uri,
        relationship?.id ?? 'slideLayout',
        'Keep exactly one slideLayout relationship pointing to an attached layout.',
      );
      continue;
    }
    const xml = parsePart(pkg, slide.uri);
    if (!xml) continue;
    const owners = placeholderRecords(xml, 'sld');
    const ambiguousOwners = reportAmbiguousPlaceholders(owners, slide.uri, emit);
    const ownerByIdentity = groupByIdentity(owners);
    const layout = layouts.get(relationship.resolvedTarget)!;
    const ambiguousLayout = ambiguousLayoutIdentities.get(layout.partUri) ?? new Set();
    for (const placeholder of layout.placeholders) {
      if (ambiguousLayout.has(placeholder.identity)) continue;
      const matches = ownerByIdentity.get(placeholder.identity) ?? [];
      if (matches.length === 0) {
        emit(
          'warning',
          'PLACEHOLDER_OWNER_MISSING',
          `Slide is missing placeholder owner ${placeholder.identity}`,
          slide.uri,
          placeholder.identity,
          'Materialize one slide placeholder owner with the layout identity.',
        );
        continue;
      }
      if (ambiguousOwners.has(placeholder.identity) || matches.length !== 1) continue;
      if (!domainAccepts(matches[0]!.domain, placeholder.type)) {
        emit(
          'error',
          'PLACEHOLDER_DOMAIN_MISMATCH',
          `Placeholder ${placeholder.identity} is populated by ${matches[0]!.domain}`,
          slide.uri,
          placeholder.identity,
          'Replace the owner with an object matching the placeholder domain.',
        );
      }
    }
  }
  return diagnostics;
}

export function validatePackage(pkg: OpcPackage): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const officeDocumentRelationships = pkg.relationships('/').filter(({ type }) => type.endsWith('/officeDocument'));
  if (officeDocumentRelationships.length !== 1) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_OFFICE_DOCUMENT_CARDINALITY',
      message: `Expected exactly one root officeDocument relationship; found ${officeDocumentRelationships.length}`,
      partUri: '/_rels/.rels',
      suggestion: 'Add one root relationship targeting the presentation part.',
    });
  }
  const idsByRelationshipPart = new Map<string, Set<string>>();
  for (const part of pkg.parts.filter(({ uri }) => uri.endsWith('.rels'))) {
    const ids = idsByRelationshipPart.get(part.uri) ?? new Set<string>();
    idsByRelationshipPart.set(part.uri, ids);
    for (const relationship of part.relationships) {
      validateRelationship(pkg, part.uri, relationship, ids, diagnostics);
    }
  }
  return diagnostics;
}

function validateRelationship(
  pkg: OpcPackage,
  relationshipPartUri: string,
  relationship: Relationship,
  ids: Set<string>,
  diagnostics: Diagnostic[],
): void {
  if (ids.has(relationship.id)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_DUPLICATE_RELATIONSHIP_ID',
      message: `Duplicate relationship id ${relationship.id}`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Allocate a unique relationship id in the source part.',
    });
  }
  ids.add(relationship.id);
  if (!/^[A-Za-z_][\w.-]*$/.test(relationship.id)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_INVALID_RELATIONSHIP_ID',
      message: `Relationship id ${relationship.id} is not a valid XML ID`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Use an id such as rId1.',
    });
  }
  if (relationship.targetMode === 'Internal' && relationship.resolvedTarget && !pkg.hasPart(relationship.resolvedTarget)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_DANGLING_RELATIONSHIP',
      message: `Relationship ${relationship.id} targets missing part ${relationship.resolvedTarget}`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Add the target part or remove the relationship.',
    });
  }
  if (relationship.targetMode === 'External') {
    diagnostics.push({
      severity: 'warning',
      code: 'OPC_EXTERNAL_RELATIONSHIP',
      message: `External relationship ${relationship.id} is not portable`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Embed the resource when the presentation must be self-contained.',
    });
  }
}

function isInternalTarget(
  pkg: OpcPackage,
  relationship: Relationship,
  contentType: string,
): boolean {
  return relationship.targetMode === 'Internal'
    && relationship.resolvedTarget !== undefined
    && pkg.getPart(relationship.resolvedTarget)?.contentType === contentType;
}

function isInternalTargetWithContentTypes(
  pkg: OpcPackage,
  relationship: Relationship,
  contentTypes: ReadonlySet<string>,
): boolean {
  return relationship.targetMode === 'Internal'
    && relationship.resolvedTarget !== undefined
    && contentTypes.has(pkg.getPart(relationship.resolvedTarget)?.contentType ?? '');
}

function parsePart(pkg: OpcPackage, partUri: string): LosslessXmlDocument | undefined {
  try {
    return LosslessXmlDocument.parse(pkg.requirePart(partUri).bytes);
  } catch {
    return undefined;
  }
}

function directLayoutName(xml: LosslessXmlDocument): string | undefined {
  const root = uniqueRoot(xml, 'sldLayout');
  const commonSlide = root ? uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE) : undefined;
  return commonSlide ? strictUnqualifiedAttribute(commonSlide, 'name') : undefined;
}

function placeholderRecords(
  xml: LosslessXmlDocument,
  rootName: 'sld' | 'sldLayout',
): PlaceholderRecord[] {
  const root = uniqueRoot(xml, rootName);
  const commonSlide = root ? uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE) : undefined;
  const shapeTree = commonSlide
    ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
    : undefined;
  if (!shapeTree) return [];
  return directElementChildren(shapeTree).flatMap((shape) => {
    const nonVisualName = {
      sp: 'nvSpPr',
      pic: 'nvPicPr',
      graphicFrame: 'nvGraphicFramePr',
      grpSp: 'nvGrpSpPr',
    }[shape.localName];
    if (!nonVisualName || elementNamespaceUri(shape) !== PRESENTATION_NAMESPACE) return [];
    const nonVisual = uniqueDirectChild(shape, nonVisualName, PRESENTATION_NAMESPACE);
    const application = nonVisual
      ? uniqueDirectChild(nonVisual, 'nvPr', PRESENTATION_NAMESPACE)
      : undefined;
    const placeholder = application
      ? uniqueDirectChild(application, 'ph', PRESENTATION_NAMESPACE)
      : undefined;
    if (!nonVisual || !application || !placeholder) return [];
    const type = strictUnqualifiedAttribute(placeholder, 'type') ?? 'body';
    const indexValue = strictUnqualifiedAttribute(placeholder, 'idx') ?? '0';
    if (!PLACEHOLDER_TYPES.has(type) || !/^(0|[1-9]\d*)$/.test(indexValue)) return [];
    const index = Number(indexValue);
    if (!Number.isSafeInteger(index) || index > 4_294_967_294) return [];
    const properties = uniqueDirectChild(nonVisual, 'cNvPr', PRESENTATION_NAMESPACE);
    const name = properties ? strictUnqualifiedAttribute(properties, 'name') : undefined;
    return [{
      identity: `${type}:${index}`,
      type,
      index,
      ...(name === undefined ? {} : { name }),
      domain: placeholderDomain(xml, shape),
    }];
  });
}

function reportAmbiguousPlaceholders(
  placeholders: readonly PlaceholderRecord[],
  partUri: string,
  emit: MasterLayoutDiagnosticEmitter,
): ReadonlySet<string> {
  const firstByIdentity = new Map<string, PlaceholderRecord>();
  const firstByName = new Map<string, PlaceholderRecord>();
  const ambiguous = new Set<string>();
  for (const placeholder of placeholders) {
    const firstIdentity = firstByIdentity.get(placeholder.identity);
    if (firstIdentity) {
      ambiguous.add(firstIdentity.identity);
      ambiguous.add(placeholder.identity);
      emit(
        'error',
        'PLACEHOLDER_IDENTITY_AMBIGUOUS',
        `Placeholder identity ${placeholder.identity} is used more than once`,
        partUri,
        placeholder.identity,
        'Give every placeholder in one owner a unique type and index.',
      );
    } else {
      firstByIdentity.set(placeholder.identity, placeholder);
    }
    if (placeholder.name === undefined) continue;
    const firstName = firstByName.get(placeholder.name);
    if (firstName) {
      ambiguous.add(firstName.identity);
      ambiguous.add(placeholder.identity);
      emit(
        'error',
        'PLACEHOLDER_IDENTITY_AMBIGUOUS',
        `Placeholder name ${placeholder.name} is used more than once`,
        partUri,
        `name:${placeholder.name}`,
        'Give every placeholder in one owner a unique name.',
      );
    } else {
      firstByName.set(placeholder.name, placeholder);
    }
  }
  return ambiguous;
}

function groupByIdentity(
  placeholders: readonly PlaceholderRecord[],
): ReadonlyMap<string, readonly PlaceholderRecord[]> {
  const groups = new Map<string, PlaceholderRecord[]>();
  for (const placeholder of placeholders) {
    const group = groups.get(placeholder.identity) ?? [];
    group.push(placeholder);
    groups.set(placeholder.identity, group);
  }
  return groups;
}

function placeholderDomain(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): PlaceholderDomain {
  if (shape.localName === 'sp') {
    const text = descendants(shape)
      .filter((element) =>
        element.localName === 't' && elementNamespaceUri(element) === DRAWING_NAMESPACE)
      .map((element) => xml.text(element))
      .join('');
    return text.length === 0 ? 'empty' : 'text-shape';
  }
  if (shape.localName === 'pic') {
    return descendants(shape).some((element) => (
      ['audioFile', 'videoFile'].includes(element.localName)
        && elementNamespaceUri(element) === DRAWING_NAMESPACE
    ) || (
      element.localName === 'media'
        && elementNamespaceUri(element) === POWERPOINT_2010_NAMESPACE
    ))
      ? 'media'
      : 'image';
  }
  if (shape.localName !== 'graphicFrame') return 'unknown';
  if (descendants(shape).some((element) =>
    element.localName === 'tbl' && elementNamespaceUri(element) === DRAWING_NAMESPACE)) {
    return 'table';
  }
  if (descendants(shape).some((element) =>
    element.localName === 'chart' && elementNamespaceUri(element) === CHART_NAMESPACE)) {
    return 'chart';
  }
  return 'unknown';
}

function domainAccepts(domain: PlaceholderDomain, type: string): boolean {
  if (domain === 'empty') return true;
  if (domain === 'text-shape') return type === 'title' || type === 'body';
  if (domain === 'image') return type === 'pic';
  if (domain === 'chart') return type === 'chart';
  if (domain === 'table') return type === 'tbl';
  return domain === 'media' && type === 'media';
}

function uniqueRoot(
  xml: LosslessXmlDocument,
  localName: string,
): XmlElement | undefined {
  const roots = xml.roots.filter((root) =>
    root.localName === localName && elementNamespaceUri(root) === PRESENTATION_NAMESPACE);
  return roots.length === 1 ? roots[0] : undefined;
}

function uniqueDirectChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const matches = directElementChildren(parent).filter((child) =>
    child.localName === localName && elementNamespaceUri(child) === namespace);
  return matches.length === 1 ? matches[0] : undefined;
}

function directElementChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function descendants(element: XmlElement): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of directElementChildren(element)) {
    result.push(child, ...descendants(child));
  }
  return result;
}

function strictUnqualifiedAttribute(
  element: XmlElement,
  localName: string,
): string | undefined {
  const matches = element.attributes.filter((attribute) =>
    attribute.localName === localName
    && attribute.name !== 'xmlns'
    && !attribute.name.startsWith('xmlns:'));
  return matches.length === 1 && matches[0]!.name === localName
    ? matches[0]!.value
    : undefined;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declaration = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const matches = current.attributes.filter(({ name }) => name === declaration);
    if (matches.length > 1) return undefined;
    if (matches[0]) return matches[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}
