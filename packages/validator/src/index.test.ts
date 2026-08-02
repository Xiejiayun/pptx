import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  validateMasterLayoutPlaceholders,
  validatePackage,
  type CompatibilityProfile,
  type Diagnostic,
} from './index.js';

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

const PRESENTATION_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const OFFICE_DOCUMENT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const SLIDE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const LAYOUT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const MASTER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const COMPATIBILITY_PROFILES: readonly CompatibilityProfile[] = [
  'powerpoint-2010',
  'powerpoint-current',
  'keynote-current',
  'libreoffice-current',
  'google-slides-import',
];

interface MasterLayoutDiagnosticFixtureOptions {
  readonly domainMismatch?: boolean;
  readonly duplicateIdentity?: boolean;
  readonly duplicateLayoutName?: boolean;
  readonly invalidSlideLayoutRelationship?: boolean;
  readonly missingOwner?: boolean;
  readonly presentationContentType?: string;
}

function masterLayoutDiagnosticFixture(
  options: MasterLayoutDiagnosticFixtureOptions = {},
): OpcPackage {
  const pkg = OpcPackage.create();
  return pkg.transaction(() => {
    const presentationUri = '/ppt/presentation.xml';
    const masterUri = '/ppt/slideMasters/slideMaster1.xml';
    const layoutUri = '/ppt/slideLayouts/slideLayout1.xml';
    const slideUri = '/ppt/slides/slide1.xml';
    const layoutPlaceholder = (
      id: number,
      name: string,
      type: 'title' | 'pic',
      index: number,
    ) => '<p:sp><p:nvSpPr>'
      + `<p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>`
      + `<p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr>`
      + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
    const slideOwner = (
      id: number,
      name: string,
      type: 'title' | 'pic',
      index: number,
      text = '',
    ) => '<p:sp><p:nvSpPr>'
      + `<p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>`
      + `<p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr>`
      + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
      + (text.length === 0 ? '' : `<a:r><a:t>${text}</a:t></a:r>`)
      + '</a:p></p:txBody></p:sp>';

    pkg.setPart(
      presentationUri,
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/>'
        + '</p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/>'
        + '</p:sldIdLst></p:presentation>',
      options.presentationContentType ?? PRESENTATION_CONTENT_TYPE,
    );
    pkg.setPart(
      masterUri,
      '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + '<p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst>'
        + '<p:sldLayoutId id="1" r:id="rIdLayout"/>'
        + (options.duplicateLayoutName
          ? '<p:sldLayoutId id="2" r:id="rIdDuplicateLayout"/>'
          : '')
        + '</p:sldLayoutIdLst></p:sldMaster>',
      MASTER_CONTENT_TYPE,
    );
    pkg.setPart(
      layoutUri,
      '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + '<p:cSld name="DUPLICATE"><p:spTree>'
        + layoutPlaceholder(2, 'title_box', 'title', 7)
        + (options.duplicateIdentity
          ? layoutPlaceholder(3, 'duplicate_title', 'title', 7)
          : '')
        + layoutPlaceholder(4, 'picture_box', 'pic', 8)
        + '</p:spTree></p:cSld></p:sldLayout>',
      LAYOUT_CONTENT_TYPE,
    );
    pkg.setPart(
      slideUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + '<p:cSld><p:spTree>'
        + slideOwner(2, 'title_box', 'title', 7)
        + (options.missingOwner
          ? ''
          : slideOwner(
              3,
              'picture_box',
              'pic',
              8,
              options.domainMismatch ? 'Wrong domain' : '',
            ))
        + '</p:spTree></p:cSld></p:sld>',
      SLIDE_CONTENT_TYPE,
    );
    if (options.duplicateLayoutName) {
      pkg.setPart(
        '/ppt/slideLayouts/slideLayout2.xml',
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
          + '<p:cSld name="DUPLICATE"><p:spTree/></p:cSld></p:sldLayout>',
        LAYOUT_CONTENT_TYPE,
      );
    }
    pkg.addRelationship('/', {
      id: 'rIdOffice',
      type: OFFICE_DOCUMENT_RELATIONSHIP,
      target: 'ppt/presentation.xml',
    });
    pkg.addRelationship(presentationUri, {
      id: 'rIdMaster',
      type: MASTER_RELATIONSHIP,
      target: 'slideMasters/slideMaster1.xml',
    });
    pkg.addRelationship(presentationUri, {
      id: 'rIdSlide',
      type: SLIDE_RELATIONSHIP,
      target: 'slides/slide1.xml',
    });
    pkg.addRelationship(masterUri, {
      id: 'rIdLayout',
      type: LAYOUT_RELATIONSHIP,
      target: '../slideLayouts/slideLayout1.xml',
    });
    if (options.duplicateLayoutName) {
      pkg.addRelationship(masterUri, {
        id: 'rIdDuplicateLayout',
        type: LAYOUT_RELATIONSHIP,
        target: '../slideLayouts/slideLayout2.xml',
      });
      pkg.addRelationship('/ppt/slideLayouts/slideLayout2.xml', {
        id: 'rIdMaster',
        type: MASTER_RELATIONSHIP,
        target: '../slideMasters/slideMaster1.xml',
      });
    }
    pkg.addRelationship(layoutUri, {
      id: 'rIdMaster',
      type: MASTER_RELATIONSHIP,
      target: '../slideMasters/slideMaster1.xml',
    });
    pkg.addRelationship(slideUri, options.invalidSlideLayoutRelationship ? {
      id: 'rIdLayout',
      type: LAYOUT_RELATIONSHIP,
      target: 'https://example.test/layout.xml',
      targetMode: 'External',
    } : {
      id: 'rIdLayout',
      type: LAYOUT_RELATIONSHIP,
      target: '../slideLayouts/slideLayout1.xml',
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

  it('accepts a healthy master layout placeholder graph in every compatibility profile', () => {
    for (const compatibility of COMPATIBILITY_PROFILES) {
      expect(validateMasterLayoutPlaceholders(
        masterLayoutDiagnosticFixture(),
        compatibility,
      )).toEqual([]);
    }
  });

  it.each([
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
    'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
    'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
  ])('validates master layouts for presentation content type %s', (presentationContentType) => {
    expect(validateMasterLayoutPlaceholders(masterLayoutDiagnosticFixture({
      duplicateLayoutName: true,
      presentationContentType,
    }))).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'LAYOUT_NAME_DUPLICATE',
        partUri: '/ppt/slideLayouts/slideLayout2.xml',
      }),
    ]));
  });

  it.each([
    {
      name: 'duplicate layout name',
      options: { duplicateLayoutName: true },
      expected: {
        severity: 'error',
        code: 'LAYOUT_NAME_DUPLICATE',
        partUri: '/ppt/slideLayouts/slideLayout2.xml',
        objectId: 'DUPLICATE',
      },
    },
    {
      name: 'invalid slide layout relationship',
      options: { invalidSlideLayoutRelationship: true },
      expected: {
        severity: 'error',
        code: 'LAYOUT_RELATIONSHIP_INVALID',
        partUri: '/ppt/slides/slide1.xml',
        objectId: 'rIdLayout',
      },
    },
    {
      name: 'ambiguous placeholder identity',
      options: { duplicateIdentity: true },
      expected: {
        severity: 'error',
        code: 'PLACEHOLDER_IDENTITY_AMBIGUOUS',
        partUri: '/ppt/slideLayouts/slideLayout1.xml',
        objectId: 'title:7',
      },
    },
    {
      name: 'missing slide placeholder owner',
      options: { missingOwner: true },
      expected: {
        severity: 'warning',
        code: 'PLACEHOLDER_OWNER_MISSING',
        partUri: '/ppt/slides/slide1.xml',
        objectId: 'pic:8',
      },
    },
    {
      name: 'placeholder owner domain mismatch',
      options: { domainMismatch: true },
      expected: {
        severity: 'error',
        code: 'PLACEHOLDER_DOMAIN_MISMATCH',
        partUri: '/ppt/slides/slide1.xml',
        objectId: 'pic:8',
      },
    },
  ])('reports $name with stable profile metadata', ({ options, expected }) => {
    for (const compatibility of COMPATIBILITY_PROFILES) {
      const diagnostics = validateMasterLayoutPlaceholders(
        masterLayoutDiagnosticFixture(options),
        compatibility,
      );
      expect(diagnostics).toContainEqual({
        ...expected,
        compatibility,
        message: expect.any(String),
        suggestion: expect.any(String),
      });
    }
  });
});
