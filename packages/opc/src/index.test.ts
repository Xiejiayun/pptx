import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage, resolveRelationshipTarget } from './index.js';

async function fixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p"><p:sldIdLst><p:sldId id="256" r:id="rId1" xmlns:r="r"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><x:unknown xmlns:x="x" keep="true"/></p:sld>');
  zip.file('docProps/opaque.bin', new Uint8Array([0, 1, 2, 3, 255]));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('OpcPackage', () => {
  it('returns the original bytes if there are no mutations', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    expect(await pkg.write()).toEqual(input);
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
});

function hash(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

