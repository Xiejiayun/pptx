import JSZip from 'jszip';
import { LosslessXmlDocument } from '@pptx/lossless-xml';

export const RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface ResourceLimits {
  readonly maxEntries: number;
  readonly maxPartBytes: number;
  readonly maxTotalUncompressedBytes: number;
  readonly maxCompressionRatio: number;
}

export const DEFAULT_RESOURCE_LIMITS: ResourceLimits = {
  maxEntries: 10_000,
  maxPartBytes: 256 * 1024 * 1024,
  maxTotalUncompressedBytes: 1024 * 1024 * 1024,
  maxCompressionRatio: 1_000,
};

export interface PackageOpenOptions {
  readonly limits?: Partial<ResourceLimits>;
  readonly signal?: AbortSignal;
}

export interface PackageCreateOptions {
  readonly entryDate?: Date;
}

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
  readonly resolvedTarget?: string;
}

export interface RelationshipInput {
  readonly id?: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: 'Internal' | 'External';
}

export interface PackageGraphNode {
  readonly uri: string;
  readonly contentType: string;
  readonly outgoing: readonly Relationship[];
  readonly incoming: readonly { sourceUri: string; relationship: Relationship }[];
}

export interface PackagePart {
  readonly uri: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly relationships: readonly Relationship[];
}

export interface MutationRecord {
  readonly kind: 'add' | 'update' | 'delete';
  readonly uri: string;
}

export class PackageError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'PackageError';
  }
}

interface MutablePart {
  uri: string;
  contentType: string;
  bytes: Uint8Array;
  relationships: Relationship[];
}

interface PackageSavepoint {
  readonly parts: readonly MutablePart[];
  readonly journal: readonly MutationRecord[];
  readonly defaults: readonly (readonly [string, string])[];
  readonly overrides: readonly (readonly [string, string])[];
}

export class OpcPackage {
  readonly #original: Uint8Array;
  readonly #zip: JSZip;
  readonly #entryDate: Date | undefined;
  readonly #parts = new Map<string, MutablePart>();
  readonly #journal: MutationRecord[] = [];
  readonly #defaults = new Map<string, string>();
  readonly #overrides = new Map<string, string>();
  #transactionDepth = 0;

  private constructor(original: Uint8Array, zip: JSZip, entryDate?: Date) {
    this.#original = original;
    this.#zip = zip;
    this.#entryDate = entryDate;
  }

  static create(options: PackageCreateOptions = {}): OpcPackage {
    const zip = new JSZip();
    const entryDate = normalizeEntryDate(options.entryDate);
    const pkg = new OpcPackage(new Uint8Array(), zip, entryDate);
    const uri = '/[Content_Types].xml';
    const bytes = new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`,
    );
    pkg.#defaults.set('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    pkg.#defaults.set('xml', 'application/xml');
    pkg.#parts.set(uri, { uri, contentType: 'application/xml', bytes, relationships: [] });
    pkg.#writeZipFile('[Content_Types].xml', bytes);
    pkg.#journal.push({ kind: 'add', uri });
    return pkg;
  }

  static async open(input: Uint8Array | ArrayBuffer, options: PackageOpenOptions = {}): Promise<OpcPackage> {
    throwIfAborted(options.signal);
    const original = input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(input);
    const limits: ResourceLimits = { ...DEFAULT_RESOURCE_LIMITS, ...options.limits };
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(original, { checkCRC32: true, createFolders: false });
    } catch (error) {
      throw new PackageError(`Invalid ZIP package: ${messageOf(error)}`);
    }

    const files = Object.values(zip.files).filter((entry) => !entry.dir);
    if (files.length > limits.maxEntries) {
      throw new PackageError(`Package has ${files.length} entries; limit is ${limits.maxEntries}`);
    }
    for (const entry of files) validateEntryName(entry.unsafeOriginalName ?? entry.name);

    const pkg = new OpcPackage(original, zip);
    const contentTypesEntry = zip.file('[Content_Types].xml');
    if (!contentTypesEntry) throw new PackageError('Missing [Content_Types].xml');
    const contentTypesBytes = await contentTypesEntry.async('uint8array');
    pkg.#parseContentTypes(contentTypesBytes);

    let totalBytes = 0;
    for (const entry of files) {
      throwIfAborted(options.signal);
      const bytes = await entry.async('uint8array');
      totalBytes += bytes.byteLength;
      if (bytes.byteLength > limits.maxPartBytes) {
        throw new PackageError(`Part exceeds ${limits.maxPartBytes} bytes`, `/${entry.name}`);
      }
      if (totalBytes > limits.maxTotalUncompressedBytes) {
        throw new PackageError(`Package exceeds ${limits.maxTotalUncompressedBytes} uncompressed bytes`);
      }
      const compressedSize = compressedSizeOf(entry);
      if (compressedSize > 0 && bytes.byteLength / compressedSize > limits.maxCompressionRatio) {
        throw new PackageError(`Part exceeds compression ratio limit ${limits.maxCompressionRatio}`, `/${entry.name}`);
      }
      const uri = normalizePartUri(`/${entry.name}`);
      pkg.#parts.set(uri, {
        uri,
        contentType: pkg.#contentTypeFor(uri),
        bytes,
        relationships: [],
      });
    }
    pkg.#loadRelationships();
    return pkg;
  }

  get parts(): readonly PackagePart[] {
    return [...this.#parts.values()].map(publicPart);
  }

  get mutations(): readonly MutationRecord[] {
    return this.#journal;
  }

  get graph(): readonly PackageGraphNode[] {
    const incoming = new Map<string, { sourceUri: string; relationship: Relationship }[]>();
    const sources = ['/', ...[...this.#parts.keys()].filter((uri) => !uri.endsWith('.rels'))];
    for (const sourceUri of sources) {
      for (const relationship of this.relationships(sourceUri)) {
        if (!relationship.resolvedTarget) continue;
        const references = incoming.get(relationship.resolvedTarget) ?? [];
        references.push({ sourceUri, relationship: { ...relationship } });
        incoming.set(relationship.resolvedTarget, references);
      }
    }
    return [...this.#parts.values()]
      .filter(({ uri }) => !uri.endsWith('.rels'))
      .map((part) => ({
        uri: part.uri,
        contentType: part.contentType,
        outgoing: this.relationships(part.uri).map((relationship) => ({ ...relationship })),
        incoming: incoming.get(part.uri) ?? [],
      }));
  }

  get changed(): boolean {
    return this.#journal.length > 0;
  }

  transaction<T>(operation: () => T): T {
    const savepoint = this.#createSavepoint();
    this.#transactionDepth += 1;
    try {
      const result = operation();
      if (isThenable(result)) {
        throw new TypeError('OpcPackage.transaction() only supports synchronous callbacks');
      }
      return result;
    } catch (error) {
      this.#restoreSavepoint(savepoint);
      throw error;
    } finally {
      this.#transactionDepth -= 1;
    }
  }

  hasPart(uri: string): boolean {
    return this.#parts.has(normalizePartUri(uri));
  }

  getPart(uri: string): PackagePart | undefined {
    const part = this.#parts.get(normalizePartUri(uri));
    return part ? publicPart(part) : undefined;
  }

  requirePart(uri: string): PackagePart {
    const part = this.getPart(uri);
    if (!part) throw new PackageError('Missing package part', normalizePartUri(uri));
    return part;
  }

  relationships(uri = '/'): readonly Relationship[] {
    const relationshipUri = relationshipPartUri(uri);
    return this.#parts.get(relationshipUri)?.relationships ?? [];
  }

  setPart(uri: string, bytes: Uint8Array | string, contentType?: string): void {
    if (this.#transactionDepth === 0) {
      this.transaction(() => this.setPart(uri, bytes, contentType));
      return;
    }
    const normalized = normalizePartUri(uri);
    const existing = this.#parts.get(normalized);
    const encoded = typeof bytes === 'string' ? new TextEncoder().encode(bytes) : new Uint8Array(bytes);
    if (existing && equalBytes(existing.bytes, encoded) && (!contentType || contentType === existing.contentType)) return;
    const type = contentType ?? existing?.contentType;
    if (!type) throw new PackageError('Content type is required for a new part', normalized);
    const contentTypeChanged = Boolean(existing && existing.contentType !== type);
    this.#parts.set(normalized, {
      uri: normalized,
      contentType: type,
      bytes: encoded,
      relationships: existing?.relationships ?? [],
    });
    this.#writeZipFile(normalized.slice(1), encoded);
    this.#journal.push({ kind: existing ? 'update' : 'add', uri: normalized });
    if (!existing || contentTypeChanged) {
      this.#overrides.set(normalized, type);
      this.#writeContentTypes();
    }
    if (normalized.endsWith('.rels')) this.#loadRelationshipsFor(normalized);
  }

  deletePart(uri: string): void {
    if (this.#transactionDepth === 0) {
      this.transaction(() => this.deletePart(uri));
      return;
    }
    const normalized = normalizePartUri(uri);
    if (!this.#parts.has(normalized)) return;
    for (const source of ['/', ...[...this.#parts.keys()].filter((candidate) => !candidate.endsWith('.rels'))]) {
      for (const relationship of [...this.relationships(source)]) {
        if (relationship.targetMode === 'Internal' && relationship.resolvedTarget === normalized) {
          this.removeRelationship(source, relationship.id);
        }
      }
    }
    const ownRelationships = relationshipPartUri(normalized);
    if (this.#parts.has(ownRelationships)) this.#deletePartRecord(ownRelationships);
    this.#deletePartRecord(normalized);
    this.#overrides.delete(normalized);
    this.#writeContentTypes();
  }

  addRelationship(sourceUri: string, input: RelationshipInput): Relationship {
    if (this.#transactionDepth === 0) {
      return this.transaction(() => this.addRelationship(sourceUri, input));
    }
    const source = sourceUri === '/' ? '/' : normalizePartUri(sourceUri);
    if (source !== '/' && !this.#parts.has(source)) throw new PackageError('Relationship source part is missing', source);
    const targetMode = input.targetMode ?? 'Internal';
    const resolvedTarget = targetMode === 'Internal' ? resolveRelationshipTarget(source, input.target) : undefined;
    if (resolvedTarget && !this.#parts.has(resolvedTarget)) {
      throw new PackageError('Relationship target part is missing', resolvedTarget);
    }
    const id = input.id ?? this.allocateRelationshipId(source);
    if (this.relationships(source).some((relationship) => relationship.id === id)) {
      throw new PackageError(`Relationship id ${id} already exists`, source);
    }
    const relationshipUri = relationshipPartUri(source);
    const existing = this.#parts.get(relationshipUri);
    const xml = existing
      ? LosslessXmlDocument.parse(existing.bytes)
      : LosslessXmlDocument.parse(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELATIONSHIP_NS}"></Relationships>`);
    const root = xml.elements('Relationships')[0];
    if (!root) throw new PackageError('Invalid relationship part', relationshipUri);
    const mode = targetMode === 'External' ? ' TargetMode="External"' : '';
    xml.appendChildXml(
      root,
      `<Relationship Id="${xmlAttribute(id)}" Type="${xmlAttribute(input.type)}" Target="${xmlAttribute(input.target)}"${mode}/>`
    );
    this.setPart(
      relationshipUri,
      xml.serialize(),
      'application/vnd.openxmlformats-package.relationships+xml',
    );
    const relationship: Relationship = {
      id,
      type: input.type,
      target: input.target,
      targetMode,
      ...(resolvedTarget ? { resolvedTarget } : {}),
    };
    return relationship;
  }

  removeRelationship(sourceUri: string, id: string): boolean {
    if (this.#transactionDepth === 0) {
      return this.transaction(() => this.removeRelationship(sourceUri, id));
    }
    const relationshipUri = relationshipPartUri(sourceUri);
    const part = this.#parts.get(relationshipUri);
    if (!part) return false;
    const xml = LosslessXmlDocument.parse(part.bytes);
    const element = xml
      .elements('Relationship')
      .find((candidate) => xml.attribute(candidate, 'Id')?.value === id);
    if (!element) return false;
    xml.removeElement(element);
    this.setPart(relationshipUri, xml.serialize(), part.contentType);
    return true;
  }

  updateRelationship(
    sourceUri: string,
    id: string,
    changes: Partial<Pick<RelationshipInput, 'type' | 'target' | 'targetMode'>>,
  ): Relationship {
    if (this.#transactionDepth === 0) {
      return this.transaction(() => this.updateRelationship(sourceUri, id, changes));
    }
    const relationshipUri = relationshipPartUri(sourceUri);
    const part = this.#parts.get(relationshipUri);
    if (!part) throw new PackageError(`Relationship ${id} was not found`, sourceUri);
    const current = part.relationships.find((relationship) => relationship.id === id);
    if (!current) throw new PackageError(`Relationship ${id} was not found`, sourceUri);
    const nextType = changes.type ?? current.type;
    const nextTarget = changes.target ?? current.target;
    const nextMode = changes.targetMode ?? current.targetMode;
    const resolvedTarget = nextMode === 'Internal' ? resolveRelationshipTarget(sourceUri, nextTarget) : undefined;
    if (resolvedTarget && !this.#parts.has(resolvedTarget)) {
      throw new PackageError('Relationship target part is missing', resolvedTarget);
    }
    const xml = LosslessXmlDocument.parse(part.bytes);
    const element = xml
      .elements('Relationship')
      .find((candidate) => xml.attribute(candidate, 'Id')?.value === id);
    if (!element) throw new PackageError(`Relationship ${id} was not found`, sourceUri);
    const typeAttribute = xml.attribute(element, 'Type');
    const targetAttribute = xml.attribute(element, 'Target');
    const modeAttribute = xml.attribute(element, 'TargetMode');
    if (!typeAttribute || !targetAttribute) throw new PackageError(`Relationship ${id} is malformed`, relationshipUri);
    if (typeAttribute.value !== nextType) xml.replaceAttribute(typeAttribute, nextType);
    if (targetAttribute.value !== nextTarget) xml.replaceAttribute(targetAttribute, nextTarget);
    if (modeAttribute) {
      if (modeAttribute.value !== nextMode) xml.replaceAttribute(modeAttribute, nextMode);
    } else if (nextMode === 'External') {
      const insertionPoint = element.startTagEnd - (element.selfClosing ? 2 : 1);
      xml.replace(insertionPoint, insertionPoint, ' TargetMode="External"');
    }
    this.setPart(relationshipUri, xml.serialize(), part.contentType);
    return {
      id,
      type: nextType,
      target: nextTarget,
      targetMode: nextMode,
      ...(resolvedTarget ? { resolvedTarget } : {}),
    };
  }

  allocateRelationshipId(sourceUri = '/'): string {
    const used = new Set(this.relationships(sourceUri).map(({ id }) => id));
    let number = 1;
    while (used.has(`rId${number}`)) number += 1;
    return `rId${number}`;
  }

  allocatePartUri(directory: string, stem: string, extension: string): string {
    const normalizedDirectory = normalizePartUri(directory);
    const suffix = extension.startsWith('.') ? extension : `.${extension}`;
    let number = 1;
    while (this.hasPart(`${normalizedDirectory}/${stem}${number}${suffix}`)) number += 1;
    return normalizePartUri(`${normalizedDirectory}/${stem}${number}${suffix}`);
  }

  async write(): Promise<Uint8Array> {
    if (!this.changed) return new Uint8Array(this.#original);
    return this.#zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      platform: 'DOS',
    });
  }

  #parseContentTypes(bytes: Uint8Array): void {
    const document = LosslessXmlDocument.parse(bytes);
    for (const element of document.elements()) {
      if (element.localName === 'Default') {
        const extension = document.attribute(element, 'Extension')?.value.toLowerCase();
        const contentType = document.attribute(element, 'ContentType')?.value;
        if (extension && contentType) this.#defaults.set(extension, contentType);
      } else if (element.localName === 'Override') {
        const partName = document.attribute(element, 'PartName')?.value;
        const contentType = document.attribute(element, 'ContentType')?.value;
        if (partName && contentType) this.#overrides.set(normalizePartUri(partName), contentType);
      }
    }
  }

  #contentTypeFor(uri: string): string {
    const override = this.#overrides.get(uri);
    if (override) return override;
    const extension = partUriExtension(uri).slice(1).toLowerCase();
    return this.#defaults.get(extension) ?? 'application/octet-stream';
  }

  #loadRelationships(): void {
    for (const uri of this.#parts.keys()) if (uri.endsWith('.rels')) this.#loadRelationshipsFor(uri);
  }

  #loadRelationshipsFor(relationshipUri: string): void {
    const relationshipPart = this.#parts.get(relationshipUri);
    if (!relationshipPart) return;
    const sourceUri = sourcePartUri(relationshipUri);
    const document = LosslessXmlDocument.parse(relationshipPart.bytes);
    const relationships: Relationship[] = [];
    for (const element of document.elements('Relationship')) {
      const id = document.attribute(element, 'Id')?.value;
      const type = document.attribute(element, 'Type')?.value;
      const target = document.attribute(element, 'Target')?.value;
      const modeValue = document.attribute(element, 'TargetMode')?.value;
      if (!id || !type || !target) continue;
      const targetMode = modeValue === 'External' ? 'External' : 'Internal';
      relationships.push({
        id,
        type,
        target,
        targetMode,
        ...(targetMode === 'Internal' ? { resolvedTarget: resolveRelationshipTarget(sourceUri, target) } : {}),
      });
    }
    relationshipPart.relationships = relationships;
  }

  #writeContentTypes(): void {
    const part = this.#parts.get('/[Content_Types].xml');
    if (!part) throw new PackageError('Missing [Content_Types].xml');
    const document = LosslessXmlDocument.parse(part.bytes);
    const root = document.elements('Types')[0];
    if (!root) throw new PackageError('Invalid [Content_Types].xml');
    const existing = new Map<string, ReturnType<LosslessXmlDocument['elements']>[number]>();
    for (const element of document.elements('Override')) {
      const partName = document.attribute(element, 'PartName')?.value;
      if (partName) existing.set(normalizePartUri(partName), element);
    }
    for (const [partName, element] of existing) {
      const expected = this.#overrides.get(partName);
      if (!expected) {
        document.removeElement(element);
        continue;
      }
      const attribute = document.attribute(element, 'ContentType');
      if (attribute && attribute.value !== expected) document.replaceAttribute(attribute, expected);
    }
    for (const [partName, contentType] of this.#overrides) {
      if (existing.has(partName)) continue;
      document.appendChildXml(
        root,
        `<Override PartName="${xmlAttribute(partName)}" ContentType="${xmlAttribute(contentType)}"/>`,
      );
    }
    const bytes = new TextEncoder().encode(document.serialize());
    part.bytes = bytes;
    this.#writeZipFile('[Content_Types].xml', bytes);
    if (!this.#journal.some(({ uri }) => uri === '/[Content_Types].xml')) {
      this.#journal.push({ kind: 'update', uri: '/[Content_Types].xml' });
    }
  }

  #deletePartRecord(normalized: string): void {
    if (!this.#parts.delete(normalized)) return;
    this.#zip.remove(normalized.slice(1));
    this.#overrides.delete(normalized);
    this.#journal.push({ kind: 'delete', uri: normalized });
  }

  #createSavepoint(): PackageSavepoint {
    return {
      parts: [...this.#parts.values()].map(cloneMutablePart),
      journal: this.#journal.map((record) => ({ ...record })),
      defaults: [...this.#defaults.entries()],
      overrides: [...this.#overrides.entries()],
    };
  }

  #restoreSavepoint(savepoint: PackageSavepoint): void {
    for (const name of Object.keys(this.#zip.files)) this.#zip.remove(name);
    this.#parts.clear();
    for (const saved of savepoint.parts) {
      const part = cloneMutablePart(saved);
      this.#parts.set(part.uri, part);
      this.#writeZipFile(part.uri.slice(1), part.bytes);
    }
    this.#defaults.clear();
    for (const [extension, contentType] of savepoint.defaults) this.#defaults.set(extension, contentType);
    this.#overrides.clear();
    for (const [uri, contentType] of savepoint.overrides) this.#overrides.set(uri, contentType);
    this.#journal.splice(0, this.#journal.length, ...savepoint.journal.map((record) => ({ ...record })));
  }

  #writeZipFile(name: string, bytes: Uint8Array): void {
    if (this.#entryDate) {
      this.#zip.file(name, bytes, { date: new Date(this.#entryDate.getTime()) });
    } else {
      this.#zip.file(name, bytes);
    }
  }
}

export function relationshipPartUri(sourceUri: string): string {
  if (sourceUri === '/') return '/_rels/.rels';
  const normalized = normalizePartUri(sourceUri);
  return joinPartUri(partUriDirname(normalized), '_rels', `${partUriBasename(normalized)}.rels`);
}

export function sourcePartUri(relationshipUri: string): string {
  const normalized = normalizePartUri(relationshipUri);
  if (normalized === '/_rels/.rels') return '/';
  const directory = partUriDirname(normalized);
  if (partUriBasename(directory) !== '_rels' || !normalized.endsWith('.rels')) {
    throw new PackageError('Invalid relationship part URI', normalized);
  }
  return joinPartUri(partUriDirname(directory), partUriBasename(normalized, '.rels'));
}

export function resolveRelationshipTarget(sourceUri: string, target: string): string {
  if (target.startsWith('/')) return normalizePartUri(target);
  const base = sourceUri === '/' ? '/' : partUriDirname(sourceUri);
  return joinPartUri(base, target);
}

export function normalizePartUri(uri: string): string {
  if (uri.includes('\\')) throw new PackageError('Invalid part URI', uri);
  const segments: string[] = [];
  for (const segment of (uri.startsWith('/') ? uri : `/${uri}`).split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      if (segments.length === 0) throw new PackageError('Invalid part URI', uri);
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return `/${segments.join('/')}`;
}

export function partUriDirname(uri: string): string {
  const normalized = normalizePartUri(uri);
  if (normalized === '/') return '/';
  const separator = normalized.lastIndexOf('/');
  return separator === 0 ? '/' : normalized.slice(0, separator);
}

export function partUriBasename(uri: string, suffix?: string): string {
  const normalized = normalizePartUri(uri);
  if (normalized === '/') return '';
  const value = normalized.slice(normalized.lastIndexOf('/') + 1);
  return suffix && value.endsWith(suffix) ? value.slice(0, -suffix.length) : value;
}

export function partUriExtension(uri: string): string {
  const basename = partUriBasename(uri);
  const dot = basename.lastIndexOf('.');
  return dot <= 0 ? '' : basename.slice(dot);
}

export function joinPartUri(base: string, ...segments: readonly string[]): string {
  return normalizePartUri([base, ...segments].join('/'));
}

export function relativeRelationshipTarget(sourcePartUri: string, targetPartUri: string): string {
  const source = partUriDirname(sourcePartUri).split('/').filter(Boolean);
  const target = normalizePartUri(targetPartUri).split('/').filter(Boolean);
  let common = 0;
  while (common < source.length && common < target.length && source[common] === target[common]) {
    common += 1;
  }
  return `${'../'.repeat(source.length - common)}${target.slice(common).join('/')}`;
}

function validateEntryName(name: string): void {
  if (name.startsWith('/') || name.includes('\\') || name.split('/').includes('..')) {
    throw new PackageError('ZIP path traversal is not allowed', name);
  }
}

function normalizeEntryDate(value: Date | undefined): Date | undefined {
  if (value === undefined) return undefined;
  if (!(value instanceof Date)) throw new TypeError('Package entryDate must be a valid Date');
  const time = Date.prototype.getTime.call(value);
  if (!Number.isFinite(time)) throw new TypeError('Package entryDate must be a valid Date');
  return new Date(time);
}

function compressedSizeOf(entry: JSZip.JSZipObject): number {
  const metadata = entry as JSZip.JSZipObject & { _data?: { compressedSize?: number } };
  return metadata._data?.compressedSize ?? 0;
}

function publicPart(part: MutablePart): PackagePart {
  return {
    uri: part.uri,
    contentType: part.contentType,
    bytes: new Uint8Array(part.bytes),
    relationships: part.relationships.map((relationship) => ({ ...relationship })),
  };
}

function cloneMutablePart(part: MutablePart): MutablePart {
  return {
    uri: part.uri,
    contentType: part.contentType,
    bytes: new Uint8Array(part.bytes),
    relationships: part.relationships.map((relationship) => ({ ...relationship })),
  };
}

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return Boolean(value && (typeof value === 'object' || typeof value === 'function') && 'then' in value);
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
}

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((byte, index) => byte === right[index]);
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function xmlAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
