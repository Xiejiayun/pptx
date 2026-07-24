import { createHash } from 'node:crypto';
import { promises as fs } from 'node:fs';
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { OpcPackage, type PackageOpenOptions } from '@pptx/opc';

export interface PartFingerprint {
  readonly uri: string;
  readonly contentType: string;
  readonly bytes: number;
  readonly sha256: string;
}

export interface PackageInspection {
  readonly partCount: number;
  readonly relationshipCount: number;
  readonly totalUncompressedBytes: number;
  readonly contentTypes: Readonly<Record<string, number>>;
  readonly parts: readonly PartFingerprint[];
}

export interface PackageDiff {
  readonly added: readonly PartFingerprint[];
  readonly removed: readonly PartFingerprint[];
  readonly changed: readonly { readonly before: PartFingerprint; readonly after: PartFingerprint }[];
  readonly unchanged: readonly PartFingerprint[];
}

export async function inspectPackage(
  input: string | Uint8Array | ArrayBuffer,
  options: PackageOpenOptions = {},
): Promise<PackageInspection> {
  const pkg = await open(input, options);
  const parts = pkg.parts.map(fingerprint).sort((left, right) => left.uri.localeCompare(right.uri));
  const contentTypes: Record<string, number> = {};
  for (const part of parts) contentTypes[part.contentType] = (contentTypes[part.contentType] ?? 0) + 1;
  return {
    partCount: parts.length,
    relationshipCount: pkg.parts.reduce((sum, part) => sum + part.relationships.length, 0),
    totalUncompressedBytes: parts.reduce((sum, part) => sum + part.bytes, 0),
    contentTypes,
    parts,
  };
}

export async function diffPackages(
  before: string | Uint8Array | ArrayBuffer,
  after: string | Uint8Array | ArrayBuffer,
): Promise<PackageDiff> {
  const [left, right] = await Promise.all([inspectPackage(before), inspectPackage(after)]);
  const leftByUri = new Map(left.parts.map((part) => [part.uri, part]));
  const rightByUri = new Map(right.parts.map((part) => [part.uri, part]));
  const added = right.parts.filter(({ uri }) => !leftByUri.has(uri));
  const removed = left.parts.filter(({ uri }) => !rightByUri.has(uri));
  const changed: { before: PartFingerprint; after: PartFingerprint }[] = [];
  const unchanged: PartFingerprint[] = [];
  for (const beforePart of left.parts) {
    const afterPart = rightByUri.get(beforePart.uri);
    if (!afterPart) continue;
    if (beforePart.sha256 === afterPart.sha256 && beforePart.contentType === afterPart.contentType) unchanged.push(afterPart);
    else changed.push({ before: beforePart, after: afterPart });
  }
  return { added, removed, changed, unchanged };
}

export function assertMutationIsolation(diff: PackageDiff, allowedPartUris: readonly string[]): void {
  const allowed = new Set(allowedPartUris);
  const unexpected = [
    ...diff.added.map(({ uri }) => uri),
    ...diff.removed.map(({ uri }) => uri),
    ...diff.changed.map(({ after }) => after.uri),
  ].filter((uri) => !allowed.has(uri));
  if (unexpected.length > 0) throw new Error(`Unexpected package mutations: ${unexpected.join(', ')}`);
}

export interface RenderOptions {
  readonly soffice?: string;
  readonly outputDirectory: string;
}

export async function renderWithLibreOffice(inputPath: string, options: RenderOptions): Promise<string> {
  await fs.mkdir(options.outputDirectory, { recursive: true });
  const binary = options.soffice ?? 'soffice';
  await spawnChecked(binary, ['--headless', '--convert-to', 'pdf', '--outdir', options.outputDirectory, inputPath]);
  return resolve(options.outputDirectory, inputPath.replace(/^.*\//, '').replace(/\.pptx$/i, '.pdf'));
}

export async function createMinimalPptx(title = 'Fixture'): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>${xmlText(title)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function open(input: string | Uint8Array | ArrayBuffer, options: PackageOpenOptions): Promise<OpcPackage> {
  const bytes = typeof input === 'string' ? new Uint8Array(await fs.readFile(input)) : input;
  return OpcPackage.open(bytes, options);
}

function fingerprint(part: { uri: string; contentType: string; bytes: Uint8Array }): PartFingerprint {
  return {
    uri: part.uri,
    contentType: part.contentType,
    bytes: part.bytes.byteLength,
    sha256: createHash('sha256').update(part.bytes).digest('hex'),
  };
}

async function spawnChecked(command: string, args: readonly string[]): Promise<void> {
  await new Promise<void>((resolvePromise, reject) => {
    const child = spawn(command, [...args], { stdio: 'ignore' });
    child.once('error', reject);
    child.once('exit', (code) => (code === 0 ? resolvePromise() : reject(new Error(`${command} exited with ${code}`))));
  });
}

function xmlText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
