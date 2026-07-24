import { posix } from 'node:path';
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

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
  readonly resolvedTarget?: string;
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

export class OpcPackage {
  readonly #original: Uint8Array;
  readonly #zip: JSZip;
  readonly #parts = new Map<string, MutablePart>();
  readonly #journal: MutationRecord[] = [];
  readonly #defaults = new Map<string, string>();
  readonly #overrides = new Map<string, string>();

  private constructor(original: Uint8Array, zip: JSZip) {
    this.#original = original;
    this.#zip = zip;
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

  get changed(): boolean {
    return this.#journal.length > 0;
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
    const normalized = normalizePartUri(uri);
    const existing = this.#parts.get(normalized);
    const encoded = typeof bytes === 'string' ? new TextEncoder().encode(bytes) : new Uint8Array(bytes);
    if (existing && equalBytes(existing.bytes, encoded) && (!contentType || contentType === existing.contentType)) return;
    const type = contentType ?? existing?.contentType;
    if (!type) throw new PackageError('Content type is required for a new part', normalized);
    this.#parts.set(normalized, {
      uri: normalized,
      contentType: type,
      bytes: encoded,
      relationships: existing?.relationships ?? [],
    });
    this.#zip.file(normalized.slice(1), encoded);
    this.#journal.push({ kind: existing ? 'update' : 'add', uri: normalized });
    if (!existing) {
      this.#overrides.set(normalized, type);
      this.#writeContentTypes();
    }
    if (normalized.endsWith('.rels')) this.#loadRelationshipsFor(normalized);
  }

  deletePart(uri: string): void {
    const normalized = normalizePartUri(uri);
    if (!this.#parts.delete(normalized)) return;
    this.#zip.remove(normalized.slice(1));
    this.#overrides.delete(normalized);
    this.#journal.push({ kind: 'delete', uri: normalized });
    this.#writeContentTypes();
  }

  allocateRelationshipId(sourceUri = '/'): string {
    const used = new Set(this.relationships(sourceUri).map(({ id }) => id));
    let number = 1;
    while (used.has(`rId${number}`)) number += 1;
    return `rId${number}`;
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
    const extension = posix.extname(uri).slice(1).toLowerCase();
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
    const defaults = [...this.#defaults.entries()]
      .map(([extension, contentType]) => `<Default Extension="${xmlAttribute(extension)}" ContentType="${xmlAttribute(contentType)}"/>`)
      .join('');
    const overrides = [...this.#overrides.entries()]
      .map(([partName, contentType]) => `<Override PartName="${xmlAttribute(partName)}" ContentType="${xmlAttribute(contentType)}"/>`)
      .join('');
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
    const bytes = new TextEncoder().encode(xml);
    part.bytes = bytes;
    this.#zip.file('[Content_Types].xml', bytes);
    if (!this.#journal.some(({ uri }) => uri === '/[Content_Types].xml')) {
      this.#journal.push({ kind: 'update', uri: '/[Content_Types].xml' });
    }
  }
}

export function relationshipPartUri(sourceUri: string): string {
  if (sourceUri === '/') return '/_rels/.rels';
  const normalized = normalizePartUri(sourceUri);
  return `${posix.dirname(normalized)}/_rels/${posix.basename(normalized)}.rels`.replace('//', '/');
}

export function sourcePartUri(relationshipUri: string): string {
  const normalized = normalizePartUri(relationshipUri);
  if (normalized === '/_rels/.rels') return '/';
  const directory = posix.dirname(normalized);
  if (posix.basename(directory) !== '_rels' || !normalized.endsWith('.rels')) {
    throw new PackageError('Invalid relationship part URI', normalized);
  }
  return normalizePartUri(`${posix.dirname(directory)}/${posix.basename(normalized, '.rels')}`);
}

export function resolveRelationshipTarget(sourceUri: string, target: string): string {
  if (target.startsWith('/')) return normalizePartUri(target);
  const base = sourceUri === '/' ? '/' : posix.dirname(normalizePartUri(sourceUri));
  return normalizePartUri(posix.join(base, target));
}

export function normalizePartUri(uri: string): string {
  const withSlash = uri.startsWith('/') ? uri : `/${uri}`;
  const normalized = posix.normalize(withSlash);
  if (!normalized.startsWith('/') || normalized.includes('/../') || normalized === '/..') {
    throw new PackageError('Invalid part URI', uri);
  }
  return normalized;
}

function validateEntryName(name: string): void {
  if (name.startsWith('/') || name.includes('\\') || name.split('/').includes('..')) {
    throw new PackageError('ZIP path traversal is not allowed', name);
  }
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

