import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  joinPartUri,
  normalizePartUri,
  OpcPackage,
  partUriBasename,
  partUriDirname,
  partUriExtension,
  relativeRelationshipTarget,
  relationshipPartUri,
  resolveRelationshipTarget,
  sourcePartUri,
  type PackageWriteOptions,
} from './index.js';

async function fixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><x:Unknown xmlns:x="urn:test" keep="yes"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p"><p:sldIdLst><p:sldId id="256" r:id="rId1" xmlns:r="r"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><x:unknown xmlns:x="x" keep="true"/></p:sld>');
  zip.file('docProps/opaque.bin', new Uint8Array([0, 1, 2, 3, 255]));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function compressionPackage(): OpcPackage {
  const pkg = OpcPackage.create({
    entryDate: new Date('1980-01-01T00:00:00.000Z'),
  });
  pkg.setPart(
    '/data.xml',
    `<data>${'compression-policy-'.repeat(8_192)}</data>`,
    'application/xml',
  );
  return pkg;
}

function zipCompressionMethods(bytes: Uint8Array): number[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.byteLength - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1;
  if (eocd < 0) throw new Error('ZIP EOCD not found');
  const entries = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const methods: number[] = [];
  for (let index = 0; index < entries; index += 1) {
    if (view.getUint32(offset, true) !== 0x0201_4b50) {
      throw new Error('ZIP central directory entry not found');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const name = new TextDecoder().decode(
      bytes.subarray(offset + 46, offset + 46 + nameLength),
    );
    if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return methods;
}

describe('OpcPackage', () => {
  it('creates an empty mutable package with content type bookkeeping', async () => {
    const pkg = OpcPackage.create();
    pkg.setPart('/example/data.xml', '<example/>', 'application/example+xml');

    const reopened = await OpcPackage.open(await pkg.write());
    expect(reopened.requirePart('/example/data.xml').contentType).toBe('application/example+xml');
    expect(new TextDecoder().decode(reopened.requirePart('/[Content_Types].xml').bytes)).toContain(
      'PartName="/example/data.xml"',
    );
  });

  it('writes byte-identical new packages with a copied fixed entry date', async () => {
    const entryDate = new Date('1980-01-01T00:00:00.000Z');
    const first = OpcPackage.create({ entryDate });
    first.setPart('/data.xml', '<data/>', 'application/xml');
    entryDate.setUTCFullYear(2026);

    const second = OpcPackage.create({ entryDate: new Date('1980-01-01T00:00:00.000Z') });
    second.setPart('/data.xml', '<data/>', 'application/xml');

    expect(await first.write()).toEqual(await second.write());
    expect(() => OpcPackage.create({ entryDate: new Date(Number.NaN) })).toThrow(/valid Date/);
  });

  it('returns the original bytes if there are no mutations', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    expect(await pkg.write()).toEqual(input);
  });

  it('selects deterministic STORE or DEFLATE output for changed packages', async () => {
    const pkg = compressionPackage();
    const defaultBytes = await pkg.write();
    const storedBytes = await pkg.write({ compression: false });
    const deflatedBytes = await pkg.write({ compression: true });

    expect(defaultBytes).toEqual(storedBytes);
    expect(new Set(zipCompressionMethods(storedBytes))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(deflatedBytes))).toEqual(new Set([8]));
    expect(deflatedBytes.byteLength).toBeLessThan(storedBytes.byteLength);
    expect(await pkg.write({ compression: false })).toEqual(storedBytes);
    expect(await pkg.write({ compression: true })).toEqual(deflatedBytes);
    await expect(OpcPackage.open(storedBytes)).resolves.toBeInstanceOf(OpcPackage);
    await expect(OpcPackage.open(deflatedBytes)).resolves.toBeInstanceOf(OpcPackage);

    if (false) {
      const options: PackageWriteOptions = { compression: true };
      pkg.write(options) satisfies Promise<Uint8Array>;
      // @ts-expect-error package compression is boolean-only
      pkg.write({ compression: 'DEFLATE' });
    }
  });

  it('preserves unchanged originals only when compression is omitted', async () => {
    const source = compressionPackage();
    const deflatedInput = await source.write({ compression: true });
    const fromDeflate = await OpcPackage.open(deflatedInput);
    const deflateJournal = [...fromDeflate.mutations];

    expect(await fromDeflate.write()).toEqual(deflatedInput);
    expect(await fromDeflate.write({ compression: undefined } as never)).toEqual(deflatedInput);
    expect(new Set(zipCompressionMethods(
      await fromDeflate.write({ compression: false }),
    ))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(
      await fromDeflate.write({ compression: true }),
    ))).toEqual(new Set([8]));
    expect(fromDeflate.mutations).toEqual(deflateJournal);

    const storedInput = await source.write({ compression: false });
    const fromStore = await OpcPackage.open(storedInput);
    expect(await fromStore.write()).toEqual(storedInput);
    expect(new Set(zipCompressionMethods(
      await fromStore.write({ compression: true }),
    ))).toEqual(new Set([8]));
    expect(fromStore.mutations).toHaveLength(0);
  });

  it('rejects non-boolean compression without package mutation', async () => {
    for (const compression of ['yes', 1, 0, null, {}, [], new Boolean(true)]) {
      const pkg = compressionPackage();
      const journal = [...pkg.mutations];
      await expect(pkg.write({ compression } as never)).rejects.toThrow(
        new TypeError('Package compression must be a boolean'),
      );
      expect(pkg.mutations).toEqual(journal);
    }
  });

  it('builds resolved relationships and preserves untouched payloads', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    expect(pkg.relationships('/ppt/presentation.xml')[0]?.resolvedTarget).toBe('/ppt/slides/slide1.xml');
    const opaqueHash = hash(pkg.requirePart('/docProps/opaque.bin').bytes);
    pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:changed/></p:sld>');
    const reopened = await OpcPackage.open(await pkg.write());
    expect(hash(reopened.requirePart('/docProps/opaque.bin').bytes)).toBe(opaqueHash);
  });

  it('resolves relationship targets relative to the source part', () => {
    expect(resolveRelationshipTarget('/ppt/slides/slide1.xml', '../media/image1.png')).toBe('/ppt/media/image1.png');
  });

  it('handles OPC part URIs without host path semantics', () => {
    expect(normalizePartUri('ppt//slides/./../media/image1.png')).toBe('/ppt/media/image1.png');
    expect(partUriDirname('/ppt/slides/slide1.xml')).toBe('/ppt/slides');
    expect(partUriBasename('/ppt/slides/slide1.xml')).toBe('slide1.xml');
    expect(partUriBasename('/_rels/.rels', '.rels')).toBe('');
    expect(partUriExtension('/ppt/slides/slide1.xml')).toBe('.xml');
    expect(partUriExtension('/_rels/.rels')).toBe('');
    expect(joinPartUri('/ppt/slideMasters', '../slideLayouts', 'slideLayout1.xml')).toBe(
      '/ppt/slideLayouts/slideLayout1.xml',
    );
    expect(relativeRelationshipTarget('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe(
      '../media/image1.png',
    );
    expect(relativeRelationshipTarget('/', '/ppt/presentation.xml')).toBe('ppt/presentation.xml');
    expect(relationshipPartUri('/')).toBe('/_rels/.rels');
    expect(sourcePartUri('/_rels/.rels')).toBe('/');
    expect(relationshipPartUri('/ppt/slides/slide1.xml')).toBe('/ppt/slides/_rels/slide1.xml.rels');
    expect(sourcePartUri('/ppt/slides/_rels/slide1.xml.rels')).toBe('/ppt/slides/slide1.xml');
    expect(() => normalizePartUri('../../outside.xml')).toThrow(/Invalid part URI/);
    expect(() => normalizePartUri('/ppt\\outside.xml')).toThrow(/Invalid part URI/);
  });

  it('adds and removes parts, content types, and relationships together', async () => {
    const pkg = await OpcPackage.open(await fixture());
    pkg.setPart('/ppt/media/image1.png', new Uint8Array([137, 80, 78, 71]), 'image/png');
    const relationship = pkg.addRelationship('/ppt/slides/slide1.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/image1.png',
    });
    expect(relationship.id).toBe('rId1');
    expect(pkg.graph.find(({ uri }) => uri === '/ppt/media/image1.png')?.incoming[0]?.sourceUri).toBe(
      '/ppt/slides/slide1.xml',
    );

    const reopened = await OpcPackage.open(await pkg.write());
    expect(reopened.requirePart('/ppt/media/image1.png').contentType).toBe('image/png');
    expect(reopened.relationships('/ppt/slides/slide1.xml')[0]?.resolvedTarget).toBe('/ppt/media/image1.png');
    expect(new TextDecoder().decode(reopened.requirePart('/[Content_Types].xml').bytes)).toContain(
      '<x:Unknown xmlns:x="urn:test" keep="yes"/>',
    );
    reopened.setPart('/ppt/media/image1.png', new Uint8Array([1, 2, 3]), 'image/jpeg');
    const retyped = await OpcPackage.open(await reopened.write());
    expect(retyped.requirePart('/ppt/media/image1.png').contentType).toBe('image/jpeg');
    const external = reopened.updateRelationship('/ppt/slides/slide1.xml', relationship.id, {
      target: 'https://example.com/image.png',
      targetMode: 'External',
    });
    expect(external.resolvedTarget).toBeUndefined();
    expect(external.targetMode).toBe('External');
    expect(reopened.removeRelationship('/ppt/slides/slide1.xml', relationship.id)).toBe(true);

    reopened.deletePart('/ppt/media/image1.png');
    const deleted = await OpcPackage.open(await reopened.write());
    expect(deleted.hasPart('/ppt/media/image1.png')).toBe(false);
    expect(deleted.relationships('/ppt/slides/slide1.xml')).toHaveLength(0);
  });

  it('commits and rolls back package transactions without leaking graph or journal changes', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    const originalSlide = pkg.requirePart('/ppt/slides/slide1.xml').bytes;

    expect(() =>
      pkg.transaction(() => {
        pkg.setPart('/ppt/media/image1.png', new Uint8Array([1, 2, 3]), 'image/png');
        pkg.addRelationship('/ppt/slides/slide1.xml', {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          target: '../media/image1.png',
        });
        pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:changed/></p:sld>');
        throw new Error('rollback');
      }),
    ).toThrow('rollback');

    expect(pkg.hasPart('/ppt/media/image1.png')).toBe(false);
    expect(pkg.relationships('/ppt/slides/slide1.xml')).toHaveLength(0);
    expect(pkg.requirePart('/ppt/slides/slide1.xml').bytes).toEqual(originalSlide);
    expect(pkg.mutations).toHaveLength(0);
    expect(await pkg.write()).toEqual(input);

    const relationship = pkg.transaction(() => {
      pkg.setPart('/ppt/media/image1.png', new Uint8Array([4, 5, 6]), 'image/png');
      return pkg.addRelationship('/ppt/slides/slide1.xml', {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: '../media/image1.png',
      });
    });
    expect(relationship.id).toBe('rId1');
    expect(pkg.graph.find(({ uri }) => uri === '/ppt/media/image1.png')?.incoming).toHaveLength(1);
  });

  it('uses nested savepoints and rejects asynchronous transaction callbacks', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);

    pkg.transaction(() => {
      pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:outer/></p:sld>');
      expect(() =>
        pkg.transaction(() => {
          pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:inner/></p:sld>');
          pkg.setPart('/ppt/media/inner.png', new Uint8Array([1]), 'image/png');
          throw new Error('inner rollback');
        }),
      ).toThrow('inner rollback');
      expect(new TextDecoder().decode(pkg.requirePart('/ppt/slides/slide1.xml').bytes)).toContain('<p:outer/>');
      expect(pkg.hasPart('/ppt/media/inner.png')).toBe(false);
    });

    const committedOuter = pkg.requirePart('/ppt/slides/slide1.xml').bytes;
    const committedJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:outer-rollback/></p:sld>');
        pkg.transaction(() => {
          pkg.setPart('/ppt/media/nested.png', new Uint8Array([2]), 'image/png');
        });
        throw new Error('outer rollback');
      }),
    ).toThrow('outer rollback');
    expect(pkg.requirePart('/ppt/slides/slide1.xml').bytes).toEqual(committedOuter);
    expect(pkg.hasPart('/ppt/media/nested.png')).toBe(false);
    expect(pkg.mutations).toEqual(committedJournal);

    expect(() =>
      pkg.transaction(() => {
        pkg.setPart('/ppt/media/async.png', new Uint8Array([3]), 'image/png');
        return Promise.resolve('unsupported');
      }),
    ).toThrow(/synchronous callbacks/);
    expect(pkg.hasPart('/ppt/media/async.png')).toBe(false);
    expect(pkg.hasPart('/ppt/media/inner.png')).toBe(false);

    const reopened = await OpcPackage.open(await pkg.write());
    expect(new TextDecoder().decode(reopened.requirePart('/ppt/slides/slide1.xml').bytes)).toContain('<p:outer/>');
    expect(new TextDecoder().decode(reopened.requirePart('/[Content_Types].xml').bytes)).not.toContain('inner.png');
  });

  it('makes individual package mutations atomic when dependent metadata is malformed', async () => {
    const pkg = await OpcPackage.open(await fixture());
    pkg.setPart('/[Content_Types].xml', '<invalid/>');
    const journal = [...pkg.mutations];

    expect(() => pkg.setPart('/ppt/media/image1.png', new Uint8Array([1]), 'image/png')).toThrow(
      /Invalid \[Content_Types\]\.xml/,
    );
    expect(pkg.hasPart('/ppt/media/image1.png')).toBe(false);
    expect(pkg.mutations).toEqual(journal);
  });

  it('enforces entry and part resource budgets', async () => {
    const input = await fixture();
    await expect(OpcPackage.open(input, { limits: { maxEntries: 2 } })).rejects.toThrow(/entries/);
    await expect(OpcPackage.open(input, { limits: { maxPartBytes: 3 } })).rejects.toThrow(/exceeds/);
  });

  it('rejects ZIP path traversal', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file('../outside.xml', '<x/>');
    const malicious = await zip.generateAsync({ type: 'uint8array' });
    await expect(OpcPackage.open(malicious)).rejects.toThrow(/path traversal/);
  });
});

function hash(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}
