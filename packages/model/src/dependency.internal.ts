import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { readMediaState } from '@pptx/codecs';
import { slideBackgroundMediaTargets } from './slide-background.internal.js';
import {
  partUriBasename,
  partUriDirname,
  partUriExtension,
  relativeRelationshipTarget,
  type OpcPackage,
  type Relationship,
} from '@pptx/opc';

type RelationshipLifecycle = 'owned' | 'shared' | 'opaque';

const OWNED_RELATIONSHIPS = new Set([
  'chart',
  'comments',
  'control',
  'diagramColors',
  'diagramData',
  'diagramDrawing',
  'diagramLayout',
  'diagramQuickStyle',
  'notesSlide',
  'oleObject',
  'package',
  'tags',
]);

const SHARED_RELATIONSHIPS = new Set([
  'audio',
  'commentAuthors',
  'hyperlink',
  'image',
  'media',
  'notesMaster',
  'person',
  'slide',
  'slideLayout',
  'slideMaster',
  'theme',
  'video',
]);

const MEDIA_RELATIONSHIPS = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
  'http://schemas.microsoft.com/office/2007/relationships/media',
]);

export function cloneSlideDependencies(pkg: OpcPackage, sourceSlideUri: string, cloneSlideUri: string): void {
  const clones = new Map<string, string>([[sourceSlideUri, cloneSlideUri]]);
  for (const relationship of pkg.relationships(sourceSlideUri)) {
    const target = cloneRootRelationshipTarget(pkg, relationship, clones, cloneSlideUri);
    pkg.addRelationship(cloneSlideUri, {
      id: relationship.id,
      type: relationship.type,
      target,
      targetMode: relationship.targetMode,
    });
  }
}

export function ownedSlideDependencyRoots(pkg: OpcPackage, slidePartUri: string): readonly string[] {
  return pkg
    .relationships(slidePartUri)
    .filter(
      (relationship) =>
        relationship.targetMode === 'Internal' &&
        Boolean(relationship.resolvedTarget) &&
        lifecycleOf(relationship, false) === 'owned',
    )
    .map(({ resolvedTarget }) => resolvedTarget!);
}

export function garbageCollectOwnedDependencies(pkg: OpcPackage, roots: readonly string[]): void {
  for (const root of new Set(roots)) garbageCollectOwnedRoot(pkg, root);
}

export function mediaSlideDependencyTargets(pkg: OpcPackage, slidePartUri: string): readonly string[] {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(slidePartUri).bytes);
  const relationships = new Map(pkg.relationships(slidePartUri).map((relationship) => [relationship.id, relationship]));
  const targets = new Set<string>();
  for (const target of slideBackgroundMediaTargets(pkg, slidePartUri)) {
    targets.add(target);
  }
  for (const picture of xml.elements('pic')) {
    if (!readMediaState(pkg, slidePartUri, xml, picture)) continue;
    for (const { element, attribute } of mediaRelationshipReferences(xml, picture)) {
      const id = xml.attribute(element, attribute)?.value;
      const relationship = id ? relationships.get(id) : undefined;
      if (
        relationship?.targetMode === 'Internal'
        && relationship.resolvedTarget?.startsWith('/ppt/media/')
        && MEDIA_RELATIONSHIPS.has(relationship.type)
      ) {
        targets.add(relationship.resolvedTarget);
      }
    }
  }
  return [...targets];
}

export function garbageCollectMediaDependencies(pkg: OpcPackage, targets: readonly string[]): void {
  for (const target of new Set(targets)) {
    if (!pkg.hasPart(target) || !target.startsWith('/ppt/media/')) continue;
    const incoming = pkg.graph.find(({ uri }) => uri === target)?.incoming ?? [];
    if (incoming.length === 0) pkg.deletePart(target);
  }
}

export function cloneOwnedPartForMutation(pkg: OpcPackage, sourcePartUri: string): string {
  return cloneOwnedPart(pkg, sourcePartUri, new Map());
}

function cloneRootRelationshipTarget(
  pkg: OpcPackage,
  relationship: Relationship,
  clones: Map<string, string>,
  cloneSourceUri: string,
): string {
  if (relationship.targetMode === 'External') return relationship.target;
  const sourceTarget = relationship.resolvedTarget;
  if (!sourceTarget) return relationship.target;
  const mapped = clones.get(sourceTarget);
  if (mapped) return relativeRelationshipTarget(cloneSourceUri, mapped);
  const lifecycle = lifecycleOf(relationship, false);
  const target = lifecycle === 'owned' ? cloneOwnedPart(pkg, sourceTarget, clones) : sourceTarget;
  return relativeRelationshipTarget(cloneSourceUri, target);
}

function cloneOwnedPart(pkg: OpcPackage, sourcePartUri: string, clones: Map<string, string>): string {
  const existing = clones.get(sourcePartUri);
  if (existing) return existing;
  const source = pkg.requirePart(sourcePartUri);
  const cloneUri = allocateCloneUri(pkg, sourcePartUri);
  clones.set(sourcePartUri, cloneUri);
  pkg.setPart(cloneUri, source.bytes, source.contentType);
  for (const relationship of pkg.relationships(sourcePartUri)) {
    let target = relationship.target;
    if (relationship.targetMode === 'Internal' && relationship.resolvedTarget) {
      const mapped = clones.get(relationship.resolvedTarget);
      const targetUri =
        mapped ??
        (lifecycleOf(relationship, true) === 'owned'
          ? cloneOwnedPart(pkg, relationship.resolvedTarget, clones)
          : relationship.resolvedTarget);
      target = relativeRelationshipTarget(cloneUri, targetUri);
    }
    pkg.addRelationship(cloneUri, {
      id: relationship.id,
      type: relationship.type,
      target,
      targetMode: relationship.targetMode,
    });
  }
  return cloneUri;
}

function garbageCollectOwnedRoot(pkg: OpcPackage, root: string): void {
  if (!pkg.hasPart(root)) return;
  const closure = collectOwnedClosure(pkg, root);
  const retained = new Set<string>();
  for (const uri of closure) {
    const incoming = pkg.graph.find((node) => node.uri === uri)?.incoming ?? [];
    if (incoming.some(({ sourceUri }) => !closure.has(sourceUri))) retainOwnedClosure(pkg, uri, closure, retained);
  }
  for (const uri of closure) {
    if (!retained.has(uri)) pkg.deletePart(uri);
  }
}

function collectOwnedClosure(pkg: OpcPackage, root: string): Set<string> {
  const closure = new Set<string>();
  const visit = (uri: string): void => {
    if (closure.has(uri) || !pkg.hasPart(uri)) return;
    closure.add(uri);
    for (const relationship of pkg.relationships(uri)) {
      if (
        relationship.targetMode === 'Internal' &&
        relationship.resolvedTarget &&
        lifecycleOf(relationship, true) === 'owned'
      ) {
        visit(relationship.resolvedTarget);
      }
    }
  };
  visit(root);
  return closure;
}

function retainOwnedClosure(
  pkg: OpcPackage,
  uri: string,
  candidates: ReadonlySet<string>,
  retained: Set<string>,
): void {
  if (retained.has(uri)) return;
  retained.add(uri);
  for (const relationship of pkg.relationships(uri)) {
    if (
      relationship.targetMode === 'Internal' &&
      relationship.resolvedTarget &&
      candidates.has(relationship.resolvedTarget) &&
      lifecycleOf(relationship, true) === 'owned'
    ) {
      retainOwnedClosure(pkg, relationship.resolvedTarget, candidates, retained);
    }
  }
}

function lifecycleOf(relationship: Relationship, insideOwnedSubgraph: boolean): RelationshipLifecycle {
  const suffix = relationship.type.slice(relationship.type.lastIndexOf('/') + 1);
  if (SHARED_RELATIONSHIPS.has(suffix)) return 'shared';
  if (OWNED_RELATIONSHIPS.has(suffix) || insideOwnedSubgraph) return 'owned';
  return 'opaque';
}

function mediaRelationshipReferences(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): readonly { readonly element: XmlElement; readonly attribute: 'r:embed' | 'r:link' }[] {
  return xml.descendants(picture)
    .filter(({ localName, parent }) =>
      ['audioFile', 'videoFile', 'media'].includes(localName)
      || (localName === 'blip' && parent?.localName === 'blipFill'))
    .map((element) => ({
      element,
      attribute: element.localName === 'audioFile' || element.localName === 'videoFile'
        ? 'r:link' as const
        : 'r:embed' as const,
    }));
}

function allocateCloneUri(pkg: OpcPackage, sourcePartUri: string): string {
  const extension = partUriExtension(sourcePartUri) || '.bin';
  const basename = partUriBasename(sourcePartUri, partUriExtension(sourcePartUri));
  const stem = basename.replace(/\d+$/, '') || 'part';
  return pkg.allocatePartUri(partUriDirname(sourcePartUri), stem, extension);
}
