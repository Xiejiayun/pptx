import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import {
  relativeRelationshipTarget,
  type OpcPackage,
  type PackagePart,
  type Relationship,
} from '@pptx/opc';
import { garbageCollectOwnedDependencies } from './dependency.internal.js';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NOTES_SLIDE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/notesSlide`;
const NOTES_MASTER_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/notesMaster`;
const SLIDE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/slide`;
const NOTES_SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const NOTES_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';
const MAX_DRAWING_ELEMENT_ID = 4_294_967_295;

interface NotesBodyState {
  readonly root: XmlElement;
  readonly shapeTree: XmlElement;
  readonly bodyShape?: XmlElement;
  readonly textBody?: XmlElement;
  readonly paragraphs: readonly XmlElement[];
}

type BodyShapeResolution =
  | { readonly kind: 'none' }
  | { readonly kind: 'body'; readonly shape: XmlElement }
  | { readonly kind: 'unsafe' };

interface DrawingPrefix {
  readonly qualified: string;
  readonly declaration: string;
}

interface ValidNotesSlideState {
  readonly kind: 'valid';
  readonly relationship: Relationship;
  readonly part: PackagePart;
  readonly xml: LosslessXmlDocument;
}

type NotesSlideState =
  | { readonly kind: 'absent' }
  | { readonly kind: 'unsafe' }
  | ValidNotesSlideState;

type NotesMasterState =
  | { readonly kind: 'absent' }
  | { readonly kind: 'unsafe' }
  | { readonly kind: 'valid'; readonly partUri: string };

export function normalizeSlideNotes(value: unknown, context: string): string {
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value.replace(/\r\n?/g, '\n');
}

export function readNotesBody(
  xml: LosslessXmlDocument,
): string | undefined {
  const state = resolveNotesBodyState(xml);
  if (!state?.bodyShape) return undefined;
  return readResolvedNotesBody(xml, state);
}

export function replaceNotesBody(
  xml: LosslessXmlDocument,
  value: unknown,
  partUri: string,
): boolean {
  const normalized = normalizeSlideNotes(value, 'Slide notes');
  const state = resolveNotesBodyState(xml);
  if (!state) {
    throw new ModelParseError('Speaker notes body is not safely editable', partUri);
  }

  if (state.bodyShape && readResolvedNotesBody(xml, state) === normalized) {
    return false;
  }

  if (!state.bodyShape) {
    const id = allocateNextShapeId(state.shapeTree, partUri);
    xml.appendChildXml(
      state.shapeTree,
      renderBodyPlaceholderShape(state.shapeTree, id, normalized),
    );
    return true;
  }

  if (!state.textBody) {
    xml.appendChildXml(
      state.bodyShape,
      renderTextBody(state.bodyShape, normalized),
    );
    return true;
  }

  const paragraph = renderNotesParagraph(
    normalized,
    drawingPrefixFor(state.textBody),
  );
  const first = state.paragraphs[0];
  if (first) {
    xml.replaceElement(first, paragraph);
    for (const extra of state.paragraphs.slice(1)) xml.removeElement(extra);
  } else if (state.textBody.selfClosing) {
    const drawing = drawingPrefixFor(state.textBody);
    xml.appendChildXml(
      state.textBody,
      renderEmptyTextBodyChildren(drawing) + renderNotesParagraph(normalized, drawing),
    );
  } else {
    xml.replace(state.textBody.endTagStart, state.textBody.endTagStart, paragraph);
  }
  return true;
}

export function createNotesSlideXml(value: unknown): string {
  const normalized = normalizeSlideNotes(value, 'Slide notes');
  const drawing: DrawingPrefix = { qualified: 'a:', declaration: '' };
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<p:notes xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">`
    + '<p:cSld><p:spTree>'
    + renderGroupShapeProperties('p:', drawing)
    + renderBodyPlaceholderShapeWithPrefixes('p:', drawing, 2, normalized)
    + '</p:spTree></p:cSld>'
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    + '</p:notes>';
}

export function readSlideNotes(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
): string | undefined {
  const state = resolveNotesSlideState(
    pkg,
    presentationPartUri,
    slidePartUri,
  );
  if (state.kind !== 'valid') return undefined;
  return readNotesBody(state.xml);
}

export function replaceSlideNotes(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
  value: string | undefined,
): boolean {
  const normalized = value === undefined
    ? undefined
    : normalizeSlideNotes(value, 'Slide notes');
  const state = resolveNotesSlideState(
    pkg,
    presentationPartUri,
    slidePartUri,
  );

  if (state.kind === 'unsafe') {
    throw new ModelParseError('Slide speaker notes relationship is not safely editable', slidePartUri);
  }
  if (normalized === undefined) {
    if (state.kind === 'absent') return false;
    return pkg.transaction(() => {
      pkg.removeRelationship(slidePartUri, state.relationship.id);
      garbageCollectOwnedDependencies(pkg, [state.part.uri]);
      return true;
    });
  }

  if (state.kind === 'valid') {
    if (!replaceNotesBody(state.xml, normalized, state.part.uri)) return false;
    return pkg.transaction(() => {
      pkg.setPart(state.part.uri, state.xml.serialize(), state.part.contentType);
      return true;
    });
  }

  const master = resolveNotesMasterState(pkg, presentationPartUri);
  if (master.kind !== 'valid') {
    throw new ModelParseError('Presentation does not contain one safe notes master', presentationPartUri);
  }
  const notesPartUri = pkg.allocatePartUri(
    '/ppt/notesSlides',
    'notesSlide',
    '.xml',
  );
  const xml = createNotesSlideXml(normalized);
  const notesMasterTarget = relativeRelationshipTarget(
    notesPartUri,
    master.partUri,
  );
  const slideTarget = relativeRelationshipTarget(notesPartUri, slidePartUri);
  const notesTarget = relativeRelationshipTarget(slidePartUri, notesPartUri);
  return pkg.transaction(() => {
    pkg.setPart(notesPartUri, xml, NOTES_SLIDE_CONTENT_TYPE);
    pkg.addRelationship(notesPartUri, {
      type: NOTES_MASTER_RELATIONSHIP,
      target: notesMasterTarget,
    });
    pkg.addRelationship(notesPartUri, {
      type: SLIDE_RELATIONSHIP,
      target: slideTarget,
    });
    pkg.addRelationship(slidePartUri, {
      type: NOTES_SLIDE_RELATIONSHIP,
      target: notesTarget,
    });
    return true;
  });
}

function resolveNotesSlideState(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
): NotesSlideState {
  const relationships = pkg.relationships(slidePartUri).filter(
    ({ type }) => type === NOTES_SLIDE_RELATIONSHIP,
  );
  if (relationships.length === 0) return { kind: 'absent' };
  if (relationships.length !== 1) return { kind: 'unsafe' };
  const relationship = relationships[0]!;
  if (
    relationship.targetMode !== 'Internal'
    || !relationship.resolvedTarget
  ) return { kind: 'unsafe' };
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part || part.contentType !== NOTES_SLIDE_CONTENT_TYPE) {
    return { kind: 'unsafe' };
  }
  const incoming = pkg.graph.find(({ uri }) => uri === part.uri)?.incoming ?? [];
  if (
    incoming.length !== 1
    || incoming[0]?.sourceUri !== slidePartUri
    || incoming[0]?.relationship.id !== relationship.id
    || incoming[0]?.relationship.type !== NOTES_SLIDE_RELATIONSHIP
  ) return { kind: 'unsafe' };

  const xml = LosslessXmlDocument.parse(part.bytes);
  if (!resolveNotesBodyState(xml)) return { kind: 'unsafe' };

  const slideRelationships = pkg.relationships(part.uri).filter(
    ({ type }) => type === SLIDE_RELATIONSHIP,
  );
  if (
    slideRelationships.length !== 1
    || slideRelationships[0]?.targetMode !== 'Internal'
    || slideRelationships[0]?.resolvedTarget !== slidePartUri
  ) return { kind: 'unsafe' };

  const master = resolveNotesMasterState(pkg, presentationPartUri);
  if (master.kind !== 'valid') return { kind: 'unsafe' };
  const masterRelationships = pkg.relationships(part.uri).filter(
    ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
  );
  if (
    masterRelationships.length !== 1
    || masterRelationships[0]?.targetMode !== 'Internal'
    || masterRelationships[0]?.resolvedTarget !== master.partUri
  ) return { kind: 'unsafe' };

  return { kind: 'valid', relationship, part, xml };
}

function resolveNotesMasterState(
  pkg: OpcPackage,
  presentationPartUri: string,
): NotesMasterState {
  const relationships = pkg.relationships(presentationPartUri).filter(
    ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
  );
  if (relationships.length === 0) return { kind: 'absent' };
  if (relationships.length !== 1) return { kind: 'unsafe' };
  const relationship = relationships[0]!;
  if (
    relationship.targetMode !== 'Internal'
    || !relationship.resolvedTarget
  ) return { kind: 'unsafe' };
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part || part.contentType !== NOTES_MASTER_CONTENT_TYPE) {
    return { kind: 'unsafe' };
  }
  const xml = LosslessXmlDocument.parse(part.bytes);
  if (
    xml.roots.length !== 1
    || !xml.roots[0]
    || !isElement(xml.roots[0], 'notesMaster', PRESENTATION_NAMESPACE)
  ) return { kind: 'unsafe' };
  return { kind: 'valid', partUri: part.uri };
}

function resolveNotesBodyState(
  xml: LosslessXmlDocument,
): NotesBodyState | undefined {
  if (xml.roots.length !== 1) return undefined;
  const root = xml.roots[0];
  if (!root || !isElement(root, 'notes', PRESENTATION_NAMESPACE)) return undefined;

  const commonSlides = directChildren(root, 'cSld', PRESENTATION_NAMESPACE);
  if (commonSlides.length !== 1) return undefined;
  const shapeTrees = directChildren(commonSlides[0]!, 'spTree', PRESENTATION_NAMESPACE);
  if (shapeTrees.length !== 1) return undefined;
  const shapeTree = shapeTrees[0]!;

  let bodyShape: XmlElement | undefined;
  for (const shape of directChildren(shapeTree, 'sp', PRESENTATION_NAMESPACE)) {
    const candidate = resolveBodyShape(shape);
    if (candidate.kind === 'unsafe') return undefined;
    if (candidate.kind === 'body') {
      if (bodyShape) return undefined;
      bodyShape = candidate.shape;
    }
  }

  if (!bodyShape) return { root, shapeTree, paragraphs: [] };
  const textBodies = directChildren(bodyShape, 'txBody', PRESENTATION_NAMESPACE);
  if (textBodies.length > 1) return undefined;
  const textBody = textBodies[0];
  if (!textBody) return { root, shapeTree, bodyShape, paragraphs: [] };
  return {
    root,
    shapeTree,
    bodyShape,
    textBody,
    paragraphs: directChildren(textBody, 'p', DRAWING_NAMESPACE),
  };
}

function resolveBodyShape(shape: XmlElement): BodyShapeResolution {
  const nonVisualProperties = directChildren(
    shape,
    'nvSpPr',
    PRESENTATION_NAMESPACE,
  );
  let containsDirectBodyPlaceholder = false;
  for (const nonVisual of nonVisualProperties) {
    for (const applicationProperties of directChildren(
      nonVisual,
      'nvPr',
      PRESENTATION_NAMESPACE,
    )) {
      for (const placeholder of directChildren(
        applicationProperties,
        'ph',
        PRESENTATION_NAMESPACE,
      )) {
        if (placeholder.attributes.some(
          (attribute) => attribute.name === 'type' && attribute.value === 'body',
        )) {
          containsDirectBodyPlaceholder = true;
        }
      }
    }
  }
  if (!containsDirectBodyPlaceholder) return { kind: 'none' };
  if (nonVisualProperties.length !== 1) return { kind: 'unsafe' };

  const applicationProperties = directChildren(
    nonVisualProperties[0]!,
    'nvPr',
    PRESENTATION_NAMESPACE,
  );
  if (applicationProperties.length !== 1) return { kind: 'unsafe' };
  const placeholders = directChildren(
    applicationProperties[0]!,
    'ph',
    PRESENTATION_NAMESPACE,
  );
  if (placeholders.length !== 1) return { kind: 'unsafe' };
  const typeAttributes = placeholders[0]!.attributes.filter(
    (attribute) => attribute.name === 'type',
  );
  if (typeAttributes.length !== 1 || typeAttributes[0]!.value !== 'body') {
    return { kind: 'unsafe' };
  }
  return { kind: 'body', shape };
}

function readResolvedNotesBody(
  xml: LosslessXmlDocument,
  state: NotesBodyState,
): string {
  if (!state.textBody) return '';
  return state.paragraphs
    .map((paragraph) => readParagraph(xml, paragraph))
    .join('\n')
    .replace(/\r\n?/g, '\n');
}

function readParagraph(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): string {
  let value = '';
  const visit = (element: XmlElement): void => {
    for (const child of element.children) {
      if (child.type !== 'element') continue;
      if (isElement(child, 't', DRAWING_NAMESPACE)) {
        value += xml.text(child);
      } else if (isElement(child, 'br', DRAWING_NAMESPACE)) {
        value += '\n';
      } else {
        visit(child);
      }
    }
  };
  visit(paragraph);
  return value;
}

function allocateNextShapeId(
  shapeTree: XmlElement,
  partUri: string,
): number {
  const ids = new Set<number>();
  let maximum = 0;
  const visit = (element: XmlElement): void => {
    if (isElement(element, 'cNvPr', PRESENTATION_NAMESPACE)) {
      const attributes = element.attributes.filter(({ name }) => name === 'id');
      const raw = attributes[0]?.value;
      if (
        attributes.length !== 1
        || raw === undefined
        || !/^\d+$/.test(raw)
      ) {
        throw new ModelParseError('Speaker notes shape tree contains an unsafe shape id', partUri);
      }
      const id = Number(raw);
      if (!Number.isSafeInteger(id) || id < 0 || id > MAX_DRAWING_ELEMENT_ID || ids.has(id)) {
        throw new ModelParseError('Speaker notes shape tree contains an unsafe shape id', partUri);
      }
      ids.add(id);
      maximum = Math.max(maximum, id);
    }
    for (const child of element.children) {
      if (child.type === 'element') visit(child);
    }
  };
  visit(shapeTree);
  if (maximum >= MAX_DRAWING_ELEMENT_ID) {
    throw new ModelParseError('Speaker notes shape ids are exhausted', partUri);
  }
  return maximum + 1;
}

function renderBodyPlaceholderShape(
  shapeTree: XmlElement,
  id: number,
  value: string,
): string {
  return renderBodyPlaceholderShapeWithPrefixes(
    qualifiedPrefix(shapeTree.name),
    drawingPrefixFor(shapeTree),
    id,
    value,
  );
}

function renderBodyPlaceholderShapeWithPrefixes(
  presentationPrefix: string,
  drawing: DrawingPrefix,
  id: number,
  value: string,
): string {
  const p = presentationPrefix;
  const a = drawing.qualified;
  const drawingDeclaration = drawing.declaration;
  const name = escapeXmlAttribute(`Notes Placeholder ${id}`);
  return `<${p}sp${drawingDeclaration}><${p}nvSpPr><${p}cNvPr id="${id}" name="${name}"/>`
    + `<${p}cNvSpPr><${a}spLocks noGrp="1"/></${p}cNvSpPr>`
    + `<${p}nvPr><${p}ph type="body" idx="1"/></${p}nvPr></${p}nvSpPr>`
    + `<${p}spPr/><${p}txBody><${a}bodyPr/><${a}lstStyle/>`
    + `${renderNotesParagraph(value, { qualified: a, declaration: '' })}`
    + `</${p}txBody></${p}sp>`;
}

function renderGroupShapeProperties(
  presentationPrefix: string,
  drawing: DrawingPrefix,
): string {
  const p = presentationPrefix;
  const a = drawing.qualified;
  return `<${p}nvGrpSpPr><${p}cNvPr id="1" name=""/><${p}cNvGrpSpPr/><${p}nvPr/></${p}nvGrpSpPr>`
    + `<${p}grpSpPr><${a}xfrm><${a}off x="0" y="0"/><${a}ext cx="0" cy="0"/>`
    + `<${a}chOff x="0" y="0"/><${a}chExt cx="0" cy="0"/></${a}xfrm></${p}grpSpPr>`;
}

function renderTextBody(owner: XmlElement, value: string): string {
  const p = qualifiedPrefix(owner.name);
  const drawing = drawingPrefixFor(owner);
  const declaration = drawing.declaration;
  return `<${p}txBody${declaration}><${drawing.qualified}bodyPr/>`
    + `<${drawing.qualified}lstStyle/>`
    + renderNotesParagraph(value, { qualified: drawing.qualified, declaration: '' })
    + `</${p}txBody>`;
}

function renderEmptyTextBodyChildren(drawing: DrawingPrefix): string {
  const declaration = drawing.declaration;
  return `<${drawing.qualified}bodyPr${declaration}/>`
    + `<${drawing.qualified}lstStyle${declaration}/>`;
}

function renderNotesParagraph(
  value: string,
  drawing: DrawingPrefix,
): string {
  const a = drawing.qualified;
  return `<${a}p${drawing.declaration}><${a}r><${a}rPr lang="en-US" dirty="0"/>`
    + `<${a}t xml:space="preserve">${escapeXmlText(value)}</${a}t></${a}r>`
    + `<${a}endParaRPr lang="en-US" dirty="0"/></${a}p>`;
}

function drawingPrefixFor(element: XmlElement): DrawingPrefix {
  const inScope = inScopePrefixForNamespace(element, DRAWING_NAMESPACE);
  if (inScope !== undefined && inScope !== '') {
    return { qualified: `${inScope}:`, declaration: '' };
  }
  const presentationPrefix = lexicalPrefix(element.name);
  const prefix = presentationPrefix === 'a' ? 'd' : 'a';
  return {
    qualified: `${prefix}:`,
    declaration: ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`,
  };
}

function inScopePrefixForNamespace(
  element: XmlElement,
  namespace: string,
): string | undefined {
  const seen = new Set<string>();
  let current: XmlElement | undefined = element;
  while (current) {
    const declarations = new Map<string, string[]>();
    for (const attribute of current.attributes) {
      const prefix = attribute.name === 'xmlns'
        ? ''
        : attribute.name.startsWith('xmlns:')
          ? attribute.name.slice('xmlns:'.length)
          : undefined;
      if (prefix === undefined) continue;
      const values = declarations.get(prefix) ?? [];
      values.push(attribute.value);
      declarations.set(prefix, values);
    }
    for (const [prefix, values] of declarations) {
      if (seen.has(prefix)) continue;
      seen.add(prefix);
      if (values.length === 1 && values[0] === namespace) return prefix;
    }
    current = current.parent;
  }
  return undefined;
}

function isElement(
  element: XmlElement,
  localName: string,
  namespace: string,
): boolean {
  return element.localName === localName
    && namespaceUriForPrefix(element, lexicalPrefix(element.name)) === namespace;
}

function directChildren(
  element: XmlElement,
  localName: string,
  namespace: string,
): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element'
      && isElement(child, localName, namespace),
  );
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  let current: XmlElement | undefined = element;
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  while (current) {
    const declarations = current.attributes.filter(
      ({ name }) => name === declarationName,
    );
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function qualifiedPrefix(name: string): string {
  const prefix = lexicalPrefix(name);
  return prefix === '' ? '' : `${prefix}:`;
}
