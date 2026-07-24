import { createReadStream, promises as fs } from 'node:fs';
import { Readable } from 'node:stream';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { OpcPackage, PackageError, type PackageOpenOptions } from '@pptx/opc';
import {
  ValidationError,
  validatePackage,
  type CompatibilityProfile,
  type Diagnostic,
} from '@pptx/validator';

export type PptxInput = string | Uint8Array | ArrayBuffer | Readable;

export interface WriteOptions {
  readonly compatibility?: CompatibilityProfile;
  readonly mode?: 'strict' | 'permissive';
}

export class ParseError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'ParseError';
  }
}

export class OpaqueMutationError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'OpaqueMutationError';
  }
}

export class SlideTitle {
  constructor(private readonly slide: Slide) {}

  get text(): string {
    return this.slide.readTitle();
  }

  set text(value: string) {
    this.slide.writeTitle(value);
  }
}

export class Slide {
  readonly title = new SlideTitle(this);

  constructor(private readonly document: PptxDocument, readonly partUri: string) {}

  readTitle(): string {
    const { xml } = this.parse();
    const shape = findTitleShape(xml);
    if (!shape) return '';
    return xml.descendants(shape, 't').map((node) => xml.text(node)).join('');
  }

  writeTitle(value: string): void {
    const { xml } = this.parse();
    const shape = findTitleShape(xml);
    if (!shape) throw new ParseError('Slide does not contain a title shape', this.partUri);
    const runs = xml.descendants(shape, 't');
    const first = runs[0];
    if (!first) throw new ParseError('Title shape does not contain text', this.partUri);
    xml.replaceText(first, value);
    for (const extra of runs.slice(1)) xml.replaceText(extra, '');
    this.document.setXmlPart(this.partUri, xml.serialize());
  }

  private parse(): { xml: LosslessXmlDocument } {
    const part = this.document.opcPackage.requirePart(this.partUri);
    try {
      return { xml: LosslessXmlDocument.parse(part.bytes) };
    } catch (error) {
      throw new ParseError(error instanceof Error ? error.message : String(error), this.partUri);
    }
  }
}

export class PptxDocument {
  readonly slides: readonly Slide[];
  readonly diagnostics: Diagnostic[] = [];

  private constructor(readonly opcPackage: OpcPackage) {
    this.slides = this.loadSlides();
  }

  static async open(input: PptxInput, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    const bytes = await readInput(input, options.signal);
    const pkg = await OpcPackage.open(bytes, options);
    return new PptxDocument(pkg);
  }

  static async fromBuffer(input: Uint8Array | ArrayBuffer, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    return this.open(input, options);
  }

  async write(options: WriteOptions = {}): Promise<Uint8Array> {
    const diagnostics = validatePackage(this.opcPackage);
    this.diagnostics.splice(0, this.diagnostics.length, ...diagnostics);
    if ((options.mode ?? 'strict') === 'strict' && diagnostics.some(({ severity }) => severity === 'error')) {
      throw new ValidationError(diagnostics);
    }
    return this.opcPackage.write();
  }

  async writeFile(path: string, options: WriteOptions = {}): Promise<void> {
    await fs.writeFile(path, await this.write(options));
  }

  setXmlPart(partUri: string, xml: string): void {
    const part = this.opcPackage.requirePart(partUri);
    this.opcPackage.setPart(partUri, xml, part.contentType);
  }

  private loadSlides(): Slide[] {
    const rootRelationship = this.opcPackage.relationships('/').find(({ type }) => type.endsWith('/officeDocument'));
    const presentationUri = rootRelationship?.resolvedTarget ?? '/ppt/presentation.xml';
    const presentation = this.opcPackage.getPart(presentationUri);
    if (!presentation) throw new PackageError('Presentation part was not found', presentationUri);
    const relationships = this.opcPackage.relationships(presentationUri);
    const xml = LosslessXmlDocument.parse(presentation.bytes);
    const orderedIds = xml
      .elements('sldId')
      .map((element) => xml.attribute(element, 'id')?.value)
      .filter((id): id is string => Boolean(id));
    const slideUris = orderedIds
      .map((id) => relationships.find((relationship) => relationship.id === id))
      .filter((relationship) => relationship?.type.endsWith('/slide') && relationship.resolvedTarget)
      .map((relationship) => relationship!.resolvedTarget!);

    if (slideUris.length === 0) {
      for (const relationship of relationships) {
        if (relationship.type.endsWith('/slide') && relationship.resolvedTarget) slideUris.push(relationship.resolvedTarget);
      }
    }
    return slideUris.map((uri) => new Slide(this, uri));
  }
}

export function openPptxStream(path: string): Readable {
  return createReadStream(path);
}

async function readInput(input: PptxInput, signal?: AbortSignal): Promise<Uint8Array> {
  if (typeof input === 'string') return new Uint8Array(await fs.readFile(input, { signal }));
  if (input instanceof Uint8Array) return new Uint8Array(input);
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  const chunks: Buffer[] = [];
  for await (const chunk of input) {
    if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
    if (typeof chunk === 'number') chunks.push(Buffer.from([chunk]));
    else chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
  }
  return new Uint8Array(Buffer.concat(chunks));
}

function findTitleShape(xml: LosslessXmlDocument): XmlElement | undefined {
  const shapes = xml.elements('sp');
  return (
    shapes.find((shape) =>
      xml.descendants(shape, 'ph').some((placeholder) => {
        const type = xml.attribute(placeholder, 'type')?.value;
        return type === 'title' || type === 'ctrTitle';
      }),
    ) ?? shapes.find((shape) => xml.descendants(shape, 't').length > 0)
  );
}
