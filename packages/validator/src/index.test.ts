import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { validatePackage, type Diagnostic } from './index.js';

async function invalidFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="1 bad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/missing.xml"/><Relationship Id="1 bad" Type="https://example.test/external" Target="https://example.test/a" TargetMode="External"/></Relationships>');
  return zip.generateAsync({ type: 'uint8array' });
}

function validImageBackgroundPackage(): OpcPackage {
  const pkg = OpcPackage.create();
  return pkg.transaction(() => {
    pkg.setPart(
      '/ppt/presentation.xml',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    );
    pkg.setPart(
      '/ppt/slides/slide1.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + '<p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId1"/>'
        + '<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/>'
        + '</p:bgPr></p:bg><p:spTree/></p:cSld></p:sld>',
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    );
    pkg.setPart('/ppt/media/background1.png', Uint8Array.of(1, 2, 3), 'image/png');
    pkg.addRelationship('/', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'ppt/presentation.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: 'slides/slide1.xml',
    });
    pkg.addRelationship('/ppt/slides/slide1.xml', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/background1.png',
    });
    return pkg;
  });
}

function validSlideDefaultColorPackage(): OpcPackage {
  const pkg = OpcPackage.create();
  return pkg.transaction(() => {
    pkg.setPart(
      '/ppt/presentation.xml',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    );
    pkg.setPart(
      '/ppt/slides/slide1.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>'
        + '<p:sp><p:txBody><a:p><a:r><a:rPr><a:solidFill>'
        + '<a:srgbClr val="FF3399"/></a:solidFill></a:rPr><a:t>sRGB</a:t></a:r>'
        + '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/>'
        + '</a:schemeClr></a:solidFill></a:rPr><a:t>Theme alpha</a:t></a:r>'
        + '</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    );
    pkg.addRelationship('/', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'ppt/presentation.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: 'slides/slide1.xml',
    });
    return pkg;
  });
}

describe('validatePackage', () => {
  it('retains compatibility metadata for feature diagnostics', () => {
    const diagnostic: Diagnostic = {
      severity: 'warning',
      code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
      message: 'Expected cached slide-number text 1',
      partUri: '/ppt/slides/slide1.xml',
      compatibility: 'powerpoint-current',
    };
    expect(diagnostic.compatibility).toBe('powerpoint-current');
  });

  it('reports duplicate ids, invalid ids, dangling targets, and external resources', async () => {
    const diagnostics = validatePackage(await OpcPackage.open(await invalidFixture()));
    const codes = diagnostics.map(({ code }) => code);
    expect(codes).toContain('OPC_DUPLICATE_RELATIONSHIP_ID');
    expect(codes).toContain('OPC_INVALID_RELATIONSHIP_ID');
    expect(codes).toContain('OPC_DANGLING_RELATIONSHIP');
    expect(codes).toContain('OPC_EXTERNAL_RELATIONSHIP');
  });

  it('accepts a valid internal slide background image relationship graph', () => {
    expect(validatePackage(validImageBackgroundPackage())).toEqual([]);
  });

  it('accepts canonical materialized slide default color runs', () => {
    expect(validatePackage(validSlideDefaultColorPackage())).toEqual([]);
  });
});
