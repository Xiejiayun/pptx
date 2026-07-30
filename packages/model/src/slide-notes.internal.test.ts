import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  LosslessXmlError,
} from '@pptx/lossless-xml';
import { OpcPackage } from '@pptx/opc';
import { ModelParseError } from './errors.js';
import {
  createNotesSlideXml,
  normalizeSlideNotes,
  readNotesBody,
  readSlideNotes,
  replaceNotesBody,
  replaceSlideNotes,
} from './slide-notes.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PART_URI = '/ppt/notesSlides/notesSlide1.xml';
const PRESENTATION_URI = '/ppt/presentation.xml';
const SLIDE_ONE_URI = '/ppt/slides/slide1.xml';
const SLIDE_TWO_URI = '/ppt/slides/slide2.xml';
const NOTES_MASTER_URI = '/ppt/notesMasters/notesMaster1.xml';
const NOTES_SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const NOTES_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';
const PRESENTATION_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const RELATIONSHIP_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.relationships+xml';
const NOTES_SLIDE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const NOTES_MASTER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster';
const SLIDE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_MASTER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const THEME_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const THEME_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.theme+xml';

interface NotesXmlOptions {
  readonly paragraphs?: string;
  readonly bodyShape?: string;
  readonly extraShapes?: string;
  readonly cSldBefore?: string;
  readonly cSldAfter?: string;
  readonly rootAfter?: string;
}

function notesXml(options: NotesXmlOptions = {}): string {
  const bodyShape = options.bodyShape ?? bodyPlaceholderShape(
    3,
    options.paragraphs ?? paragraph('<a:r><a:t></a:t></a:r>'),
  );
  return `<?xml version="1.0" encoding="UTF-8"?>\n`
    + `<p:notes xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:x="urn:keep">`
    + `<p:cSld data="KEEP">${options.cSldBefore ?? ''}<p:spTree>`
    + groupShapeProperties()
    + placeholderShape(2, 'sldImg', 'Slide image text')
    + bodyShape
    + placeholderShape(4, 'sldNum', '42')
    + (options.extraShapes ?? '')
    + `</p:spTree>${options.cSldAfter ?? '<p:extLst><x:keep where="cSld"/></p:extLst>'}</p:cSld>`
    + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`
    + `${options.rootAfter ?? '<p:extLst><x:keep where="root"/></p:extLst>'}`
    + `</p:notes>`;
}

function groupShapeProperties(id = '1'): string {
  return `<p:nvGrpSpPr><p:cNvPr id="${id}" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
    + `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>`
    + `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
}

function paragraph(contents: string, attributes = ''): string {
  return `<a:p${attributes}>${contents}</a:p>`;
}

function bodyPlaceholderShape(
  id: number | string,
  paragraphs: string,
  options: {
    readonly placeholder?: string;
    readonly textBodies?: string;
    readonly afterTextBody?: string;
  } = {},
): string {
  const placeholder = options.placeholder ?? '<p:ph type="body" idx="1"/>';
  const textBodies = options.textBodies
    ?? `<p:txBody data="KEEP"><a:bodyPr data="KEEP"/><a:lstStyle><a:lvl1pPr/></a:lstStyle>`
      + `${paragraphs}</p:txBody>`;
  return `<p:sp data="KEEP"><p:nvSpPr><p:cNvPr id="${id}" name="Notes Placeholder"/>`
    + `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${placeholder}</p:nvPr>`
    + `</p:nvSpPr><p:spPr data="KEEP"/>${textBodies}`
    + `${options.afterTextBody ?? '<p:extLst><x:keep where="shape"/></p:extLst>'}</p:sp>`;
}

function placeholderShape(id: number, type: string, text: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${type}"/><p:cNvSpPr/>`
    + `<p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:spPr/>`
    + `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraph(`<a:r><a:t>${text}</a:t></a:r>`)}</p:txBody></p:sp>`;
}

function expectRejectedWithoutPatch(source: string): void {
  const xml = LosslessXmlDocument.parse(source);
  expect(() => replaceNotesBody(xml, 'Replacement', PART_URI)).toThrow(ModelParseError);
  expect(xml.changed).toBe(false);
  expect(xml.serialize()).toBe(source);
}

function packageWithNotesMaster(): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.setPart(
    PRESENTATION_URI,
    `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
      + '<p:notesMasterIdLst><p:notesMasterId r:id="rId1"/></p:notesMasterIdLst>'
      + '</p:presentation>',
    PRESENTATION_CONTENT_TYPE,
  );
  pkg.setPart(
    SLIDE_ONE_URI,
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    SLIDE_CONTENT_TYPE,
  );
  pkg.setPart(
    SLIDE_TWO_URI,
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    SLIDE_CONTENT_TYPE,
  );
  pkg.setPart(
    NOTES_MASTER_URI,
    `<p:notesMaster xmlns:p="${PRESENTATION_NAMESPACE}"><p:cSld><p:spTree/></p:cSld></p:notesMaster>`,
    NOTES_MASTER_CONTENT_TYPE,
  );
  pkg.addRelationship(PRESENTATION_URI, {
    type: NOTES_MASTER_RELATIONSHIP,
    target: 'notesMasters/notesMaster1.xml',
  });
  return pkg;
}

function packageWithoutNotesMaster(options: {
  readonly directTheme?: boolean;
  readonly fallbackTheme?: boolean;
} = {}): OpcPackage {
  const directTheme = options.directTheme ?? true;
  const fallbackTheme = options.fallbackTheme ?? true;
  const pkg = OpcPackage.create();
  pkg.setPart(
    PRESENTATION_URI,
    `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
      + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
      + '<p:sldIdLst/><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="5143500" cy="9144000"/>'
      + '</p:presentation>',
    PRESENTATION_CONTENT_TYPE,
  );
  pkg.setPart(
    SLIDE_ONE_URI,
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    SLIDE_CONTENT_TYPE,
  );
  pkg.setPart(
    '/ppt/slideMasters/slideMaster1.xml',
    `<p:sldMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    SLIDE_MASTER_CONTENT_TYPE,
  );
  pkg.setPart(
    '/ppt/theme/theme1.xml',
    `<a:theme xmlns:a="${DRAWING_NAMESPACE}" name="Direct"/>`,
    THEME_CONTENT_TYPE,
  );
  pkg.setPart(
    '/ppt/theme/theme2.xml',
    `<a:theme xmlns:a="${DRAWING_NAMESPACE}" name="Fallback"/>`,
    THEME_CONTENT_TYPE,
  );
  pkg.addRelationship(PRESENTATION_URI, {
    id: 'rId1',
    type: SLIDE_MASTER_RELATIONSHIP,
    target: 'slideMasters/slideMaster1.xml',
  });
  if (directTheme) {
    pkg.addRelationship(PRESENTATION_URI, {
      id: 'rId2',
      type: THEME_RELATIONSHIP,
      target: 'theme/theme1.xml',
    });
  }
  if (fallbackTheme) {
    pkg.addRelationship('/ppt/slideMasters/slideMaster1.xml', {
      type: THEME_RELATIONSHIP,
      target: '../theme/theme2.xml',
    });
  }
  return pkg;
}

function addNotesPart(
  pkg: OpcPackage,
  value: string,
  options: {
    readonly notesUri?: string;
    readonly slideUri?: string;
    readonly contentType?: string;
    readonly xml?: string;
  } = {},
): string {
  const notesUri = options.notesUri ?? PART_URI;
  const slideUri = options.slideUri ?? SLIDE_ONE_URI;
  pkg.setPart(
    notesUri,
    options.xml ?? createNotesSlideXml(value),
    options.contentType ?? NOTES_SLIDE_CONTENT_TYPE,
  );
  pkg.addRelationship(notesUri, {
    type: NOTES_MASTER_RELATIONSHIP,
    target: '../notesMasters/notesMaster1.xml',
  });
  pkg.addRelationship(notesUri, {
    type: SLIDE_RELATIONSHIP,
    target: `../slides/${slideUri.slice(slideUri.lastIndexOf('/') + 1)}`,
  });
  pkg.addRelationship(slideUri, {
    type: NOTES_SLIDE_RELATIONSHIP,
    target: `../notesSlides/${notesUri.slice(notesUri.lastIndexOf('/') + 1)}`,
  });
  return notesUri;
}

function expectPackageReadNoOp(pkg: OpcPackage, expected: string | undefined): void {
  const beforeParts = pkg.parts.map(({ uri, contentType, bytes }) => ({
    uri,
    contentType,
    bytes: bytes.slice(),
  }));
  const beforeRelationships = pkg.graph.map(({ uri, outgoing, incoming }) => ({
    uri,
    outgoing: outgoing.map((relationship) => ({ ...relationship })),
    incoming: incoming.map(({ sourceUri, relationship }) => ({
      sourceUri,
      relationship: { ...relationship },
    })),
  }));
  const beforeJournal = [...pkg.mutations];
  expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBe(expected);
  expect(pkg.parts.map(({ uri, contentType, bytes }) => ({
    uri,
    contentType,
    bytes,
  }))).toEqual(beforeParts);
  expect(pkg.graph.map(({ uri, outgoing, incoming }) => ({
    uri,
    outgoing,
    incoming,
  }))).toEqual(beforeRelationships);
  expect(pkg.mutations).toEqual(beforeJournal);
}

describe('slide speaker notes text codec', () => {
  it('normalizes line endings and rejects non-string or XML-unsafe input', () => {
    expect(normalizeSlideNotes('A\r\nB\rC\nD\tE', 'Slide notes')).toBe('A\nB\nC\nD\tE');
    for (const value of [undefined, null, 7, true, {}, [], Symbol('notes')]) {
      expect(() => normalizeSlideNotes(value, 'Slide notes')).toThrow(
        new TypeError('Slide notes must be a string'),
      );
    }
    for (const value of ['A\u0000B', 'A\u0008B', 'A\u000bB', 'A\u000cB', 'A\u001fB']) {
      expect(() => normalizeSlideNotes(value, 'Slide notes')).toThrow(
        new TypeError('Slide notes contains invalid XML characters'),
      );
    }
  });

  it('reads empty, plain, rich, field, break, and multi-paragraph text without patches', () => {
    const cases: readonly (readonly [string, string])[] = [
      [notesXml(), ''],
      [notesXml({ paragraphs: paragraph('<a:r><a:t>Speaker &amp; notes</a:t></a:r>') }), 'Speaker & notes'],
      [notesXml({ paragraphs: paragraph('<a:r><a:t xml:space="preserve">  leading and trailing  </a:t></a:r>') }), '  leading and trailing  '],
      [notesXml({ paragraphs: paragraph('<a:r><a:t/></a:r>') }), ''],
      [
        notesXml({
          paragraphs:
            paragraph('<a:r><a:t>First</a:t></a:r>')
            + paragraph('<a:fld id="field"><a:rPr/><a:t>soft</a:t></a:fld><a:br/><a:r><a:t>break</a:t></a:r>')
            + paragraph('<a:r><a:rPr/><a:t>Second</a:t></a:r>'),
        }),
        'First\nsoft\nbreak\nSecond',
      ],
    ];

    for (const [source, expected] of cases) {
      const xml = LosslessXmlDocument.parse(source);
      expect(readNotesBody(xml), source).toBe(expected);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('accepts alternate prefixes and ignores unrelated placeholder text', () => {
    const source = `<q:notes xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}" xmlns:x="urn:foreign">`
      + `<q:cSld><q:spTree>${groupShapeProperties().replaceAll('p:', 'q:').replaceAll('a:', 'd:')}`
      + `<q:sp><q:nvSpPr><q:cNvPr id="2" name="lookalike"/><q:cNvSpPr/><q:nvPr>`
      + `<q:ph x:type="body"/></q:nvPr></q:nvSpPr><q:txBody><d:p><d:r><d:t>ignore</d:t></d:r></d:p></q:txBody></q:sp>`
      + `${placeholderShape(3, 'sldImg', 'ignore image').replaceAll('p:', 'q:').replaceAll('a:', 'd:')}`
      + `${bodyPlaceholderShape(4, paragraph('<d:fld><d:t>Body</d:t></d:fld><d:br/><d:r><d:t>notes</d:t></d:r>'))
        .replaceAll('p:', 'q:').replaceAll('a:', 'd:')}`
      + `${placeholderShape(5, 'sldNum', 'ignore number').replaceAll('p:', 'q:').replaceAll('a:', 'd:')}`
      + `</q:spTree></q:cSld></q:notes>`;
    const xml = LosslessXmlDocument.parse(source);
    expect(readNotesBody(xml)).toBe('Body\nnotes');
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('accepts a default PresentationML namespace with a distinct DrawingML prefix', () => {
    const source = `<notes xmlns="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">`
      + '<cSld><spTree><nvGrpSpPr><cNvPr id="1" name=""/><cNvGrpSpPr/><nvPr/></nvGrpSpPr>'
      + '<grpSpPr/><sp><nvSpPr><cNvPr id="2" name="Notes"/><cNvSpPr/><nvPr>'
      + '<ph type="body"/></nvPr></nvSpPr><spPr/><txBody><d:bodyPr/><d:lstStyle/>'
      + '<d:p><d:r><d:t>Default prefix</d:t></d:r></d:p></txBody></sp></spTree></cSld></notes>';
    const xml = LosslessXmlDocument.parse(source);
    expect(readNotesBody(xml)).toBe('Default prefix');
    expect(replaceNotesBody(xml, 'Edited', PART_URI)).toBe(true);
    expect(readNotesBody(LosslessXmlDocument.parse(xml.serialize()))).toBe('Edited');
  });

  it('returns undefined for missing body and unsafe root, tree, placeholder, or text ownership', () => {
    const noBody = notesXml({ bodyShape: '' });
    const descendantBody = notesXml({
      bodyShape: `<p:sp><p:nvSpPr><p:cNvPr id="3" name="nested"/><p:cNvSpPr/><p:nvPr>`
        + `<p:extLst><p:ph type="body"/></p:extLst></p:nvPr></p:nvSpPr></p:sp>`,
    });
    const duplicateBodies = notesXml({
      extraShapes: bodyPlaceholderShape(5, paragraph('<a:r><a:t>Second body</a:t></a:r>')),
    });
    const repeatedChain = notesXml({
      bodyShape: `<p:sp><p:nvSpPr><p:cNvPr id="3" name="first"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>`
        + `<p:nvSpPr><p:cNvPr id="6" name="second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>`,
    });
    const duplicateTextBodies = notesXml({
      bodyShape: bodyPlaceholderShape(3, '', {
        textBodies: '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody><p:txBody/>',
      }),
    });
    const cases = [
      noBody,
      descendantBody,
      '<p:notes xmlns:p="urn:wrong"/>',
      `<p:notNotes xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      `${notesXml()}${notesXml()}`,
      `<p:notes xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      `<p:notes xmlns:p="${PRESENTATION_NAMESPACE}"><p:cSld/></p:notes>`,
      notesXml().replace('</p:cSld>', '</p:cSld><p:cSld/>'),
      notesXml({ bodyShape: '', cSldBefore: '<p:spTree/>' }),
      notesXml({
        bodyShape: `<p:sp><p:nvSpPr><p:cNvPr id="3" name="qualified"/><p:cNvSpPr/>`
          + '<p:nvPr><p:ph x:type="body"/></p:nvPr></p:nvSpPr></p:sp>',
      }),
      duplicateBodies,
      repeatedChain,
      duplicateTextBodies,
    ];
    for (const source of cases) {
      const xml = LosslessXmlDocument.parse(source);
      expect(readNotesBody(xml), source).toBeUndefined();
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
    expect(() => LosslessXmlDocument.parse('<p:notes>')).toThrow(LosslessXmlError);
  });

  it('replaces only body paragraphs and preserves all unowned XML', () => {
    const source = notesXml({
      paragraphs:
        paragraph('<a:pPr data="old"/><a:r><a:rPr/><a:t>Before</a:t></a:r>')
        + '<x:keep where="between-paragraphs"/>'
        + paragraph('<a:r><a:t>second</a:t></a:r>'),
    });
    const xml = LosslessXmlDocument.parse(source);
    expect(replaceNotesBody(xml, 'After\r\nLine 2 & <xml>', PART_URI)).toBe(true);
    const output = xml.serialize();
    expect(readNotesBody(LosslessXmlDocument.parse(output))).toBe('After\nLine 2 & <xml>');
    expect(output).toContain('<a:bodyPr data="KEEP"/>');
    expect(output).toContain('<a:lstStyle><a:lvl1pPr/></a:lstStyle>');
    expect(output).toContain('<p:spPr data="KEEP"/>');
    expect(output).toContain('<p:ph type="body" idx="1"/>');
    expect(output).toContain('<x:keep where="between-paragraphs"/>');
    expect(output).toContain('<x:keep where="shape"/>');
    expect(output).toContain('<x:keep where="cSld"/>');
    expect(output).toContain('<x:keep where="root"/>');
    expect(output).toContain('<a:t xml:space="preserve">After\nLine 2 &amp; &lt;xml&gt;</a:t>');
    const bodyTextXml = /<p:txBody data="KEEP">([\s\S]*?)<\/p:txBody>/.exec(output)?.[1];
    expect(bodyTextXml?.match(/<a:p(?:\s|>)/g)).toHaveLength(1);
  });

  it('keeps semantically equal normalized text as an exact no-op', () => {
    const source = notesXml({
      paragraphs: paragraph('<a:r><a:t>Same</a:t></a:r><a:br/><a:fld><a:t>value</a:t></a:fld>'),
    });
    const xml = LosslessXmlDocument.parse(source);
    expect(replaceNotesBody(xml, 'Same\r\nvalue', PART_URI)).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('inserts a missing canonical body with the next safe shape id', () => {
    const source = notesXml({
      bodyShape: '',
      extraShapes: placeholderShape(7, 'sldNum', 'ignore'),
    });
    const xml = LosslessXmlDocument.parse(source);
    expect(readNotesBody(xml)).toBeUndefined();
    expect(replaceNotesBody(xml, 'Inserted & <safe>', PART_URI)).toBe(true);
    const output = xml.serialize();
    expect(readNotesBody(LosslessXmlDocument.parse(output))).toBe('Inserted & <safe>');
    expect(output).toContain('<p:cNvPr id="8" name="Notes Placeholder 8"/>');
    expect(output).toContain('<p:ph type="body" idx="1"/>');
    expect(output).toContain('<a:t xml:space="preserve">Inserted &amp; &lt;safe&gt;</a:t>');
    expect(output).toContain('<x:keep where="cSld"/>');
  });

  it('rejects unsafe shape ids and ownership before making a patch', () => {
    const unsafe = [
      '<p:notes xmlns:p="urn:wrong"/>',
      notesXml().replace('</p:cSld>', '</p:cSld><p:cSld/>'),
      notesXml({ bodyShape: '', cSldBefore: '<p:spTree/>' }),
      notesXml({ bodyShape: '', extraShapes: placeholderShape(1, 'sldNum', 'duplicate') }),
      notesXml({ bodyShape: '', extraShapes: placeholderShape(Number.NaN, 'sldNum', 'invalid') }),
      notesXml({ bodyShape: '', extraShapes: placeholderShape(-1, 'sldNum', 'negative') }),
      notesXml({ bodyShape: '', extraShapes: placeholderShape(4_294_967_295, 'sldNum', 'exhausted') }),
      notesXml({
        bodyShape: '',
        extraShapes: '<p:sp><p:nvSpPr><p:cNvPr name="missing"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>',
      }),
      notesXml({
        bodyShape: '',
        extraShapes: '<p:sp><p:nvSpPr><p:cNvPr id="8" id="9" name="repeated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>',
      }),
      notesXml({
        bodyShape: bodyPlaceholderShape(3, '', {
          textBodies: '<p:txBody/><p:txBody/>',
        }),
      }),
      notesXml({ extraShapes: bodyPlaceholderShape(5, paragraph('<a:r><a:t>duplicate</a:t></a:r>')) }),
    ];
    for (const source of unsafe) expectRejectedWithoutPatch(source);
  });

  it('creates deterministic canonical notes-slide XML with escaped plain text', () => {
    const source = createNotesSlideXml('A & <B>\r\nLine 2');
    const xml = LosslessXmlDocument.parse(source);
    expect(readNotesBody(xml)).toBe('A & <B>\nLine 2');
    expect(source).toContain(`xmlns:p="${PRESENTATION_NAMESPACE}"`);
    expect(source).toContain(`xmlns:a="${DRAWING_NAMESPACE}"`);
    expect(source).toContain('<p:cNvPr id="2" name="Notes Placeholder 2"/>');
    expect(source).toContain('<a:t xml:space="preserve">A &amp; &lt;B&gt;\nLine 2</a:t>');
    expect(createNotesSlideXml('A & <B>\r\nLine 2')).toBe(source);
  });
});

describe('slide speaker notes package state', () => {
  it('reads valid state and returns undefined for absent or unsafe relationships without mutation', () => {
    const valid = packageWithNotesMaster();
    addNotesPart(valid, 'Speaker note');
    expectPackageReadNoOp(valid, 'Speaker note');

    const absent = packageWithNotesMaster();
    expectPackageReadNoOp(absent, undefined);

    const duplicate = packageWithNotesMaster();
    addNotesPart(duplicate, 'First');
    addNotesPart(duplicate, 'Second', {
      notesUri: '/ppt/notesSlides/notesSlide2.xml',
    });
    expectPackageReadNoOp(duplicate, undefined);

    const external = packageWithNotesMaster();
    external.addRelationship(SLIDE_ONE_URI, {
      type: NOTES_SLIDE_RELATIONSHIP,
      target: 'https://example.com/notes.xml',
      targetMode: 'External',
    });
    expectPackageReadNoOp(external, undefined);

    const unresolved = packageWithNotesMaster();
    unresolved.setPart(
      '/ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
        + `<Relationship Id="rId1" Type="${NOTES_SLIDE_RELATIONSHIP}" Target="../notesSlides/missing.xml"/>`
        + '</Relationships>',
      RELATIONSHIP_CONTENT_TYPE,
    );
    expectPackageReadNoOp(unresolved, undefined);

    const wrongContentType = packageWithNotesMaster();
    addNotesPart(wrongContentType, 'Wrong type', {
      contentType: 'application/octet-stream',
    });
    expectPackageReadNoOp(wrongContentType, undefined);

    const wrongRoot = packageWithNotesMaster();
    addNotesPart(wrongRoot, 'Wrong root', {
      xml: `<p:notNotes xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    });
    expectPackageReadNoOp(wrongRoot, undefined);

    const shared = packageWithNotesMaster();
    addNotesPart(shared, 'Shared');
    shared.addRelationship(SLIDE_TWO_URI, {
      type: NOTES_SLIDE_RELATIONSHIP,
      target: '../notesSlides/notesSlide1.xml',
    });
    expectPackageReadNoOp(shared, undefined);
  });

  it('edits, creates, clears, and rejects unsafe package state atomically', () => {
    const pkg = packageWithNotesMaster();
    addNotesPart(pkg, 'Before');
    const untouched = new Map(
      pkg.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'After\r\nLine 2')).toBe(true);
    expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBe('After\nLine 2');
    for (const { uri, bytes } of pkg.parts) {
      if (uri !== PART_URI) expect(bytes).toEqual(untouched.get(uri));
    }

    const sameParts = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const sameJournal = [...pkg.mutations];
    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'After\nLine 2')).toBe(false);
    expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(sameParts);
    expect(pkg.mutations).toEqual(sameJournal);

    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, undefined)).toBe(true);
    expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBeUndefined();
    expect(pkg.hasPart(PART_URI)).toBe(false);
    expect(pkg.hasPart(SLIDE_ONE_URI)).toBe(true);
    expect(pkg.hasPart(NOTES_MASTER_URI)).toBe(true);

    const created = packageWithNotesMaster();
    expect(replaceSlideNotes(created, PRESENTATION_URI, SLIDE_ONE_URI, '')).toBe(true);
    expect(readSlideNotes(created, PRESENTATION_URI, SLIDE_ONE_URI)).toBe('');
    expect(created.relationships(SLIDE_ONE_URI).filter(
      ({ type }) => type === NOTES_SLIDE_RELATIONSHIP,
    )).toHaveLength(1);

    const unsafe = packageWithNotesMaster();
    addNotesPart(unsafe, 'Unsafe');
    unsafe.addRelationship(SLIDE_TWO_URI, {
      type: NOTES_SLIDE_RELATIONSHIP,
      target: '../notesSlides/notesSlide1.xml',
    });
    const unsafeParts = unsafe.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const unsafeJournal = [...unsafe.mutations];
    expect(() => replaceSlideNotes(
      unsafe,
      PRESENTATION_URI,
      SLIDE_ONE_URI,
      'Rejected',
    )).toThrow(ModelParseError);
    expect(unsafe.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(unsafeParts);
    expect(unsafe.mutations).toEqual(unsafeJournal);
  });

  it('creates a missing notes master from the direct presentation theme', () => {
    const pkg = packageWithoutNotesMaster();
    const preserved = new Map(
      pkg.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'Created')).toBe(true);
    expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBe('Created');

    const masterRelationship = pkg.relationships(PRESENTATION_URI).find(
      ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
    );
    const masterUri = masterRelationship?.resolvedTarget;
    expect(masterUri).toBeDefined();
    expect(pkg.requirePart(masterUri!).contentType).toBe(NOTES_MASTER_CONTENT_TYPE);
    expect(pkg.relationships(masterUri!).find(
      ({ type }) => type === THEME_RELATIONSHIP,
    )?.resolvedTarget).toBe('/ppt/theme/theme1.xml');
    const presentationXml = new TextDecoder().decode(pkg.requirePart(PRESENTATION_URI).bytes);
    expect(presentationXml).toContain(
      `<p:notesMasterIdLst><p:notesMasterId r:id="${masterRelationship?.id}"/></p:notesMasterIdLst>`,
    );
    expect(presentationXml.indexOf('</p:sldIdLst>')).toBeLessThan(
      presentationXml.indexOf('<p:notesMasterIdLst>'),
    );
    expect(presentationXml.indexOf('</p:notesMasterIdLst>')).toBeLessThan(
      presentationXml.indexOf('<p:sldSz'),
    );
    const notesUri = pkg.relationships(SLIDE_ONE_URI).find(
      ({ type }) => type === NOTES_SLIDE_RELATIONSHIP,
    )!.resolvedTarget!;
    expect(pkg.relationships(notesUri).find(
      ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
    )?.resolvedTarget).toBe(masterUri);
    const masterXml = new TextDecoder().decode(pkg.requirePart(masterUri!).bytes);
    expect(masterXml).toContain('<p:notesMaster');
    expect(masterXml).toContain('<p:clrMap ');
    expect(masterXml).toContain('<p:hf hdr="1" ftr="1" dt="1" sldNum="1"/>');
    expect(masterXml).toContain('<p:notesStyle/>');
    for (const [uri, bytes] of preserved) {
      if (
        uri !== PRESENTATION_URI
        && uri !== '/ppt/_rels/presentation.xml.rels'
        && uri !== '/[Content_Types].xml'
      ) {
        expect(pkg.requirePart(uri).bytes).toEqual(bytes);
      }
    }
  });

  it('falls back to the first slide master theme when no direct presentation theme exists', () => {
    const pkg = packageWithoutNotesMaster({ directTheme: false });
    pkg.setPart(
      '/ppt/slideMasters/slideMaster2.xml',
      `<p:sldMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      SLIDE_MASTER_CONTENT_TYPE,
    );
    pkg.addRelationship(PRESENTATION_URI, {
      id: 'rId3',
      type: SLIDE_MASTER_RELATIONSHIP,
      target: 'slideMasters/slideMaster2.xml',
    });
    pkg.addRelationship('/ppt/slideMasters/slideMaster2.xml', {
      type: THEME_RELATIONSHIP,
      target: '../theme/theme1.xml',
    });
    const presentationPart = pkg.requirePart(PRESENTATION_URI);
    pkg.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(presentationPart.bytes).replace(
        '</p:sldMasterIdLst>',
        '<p:sldMasterId id="2147483649" r:id="rId3"/></p:sldMasterIdLst>',
      ),
      presentationPart.contentType,
    );
    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'Fallback')).toBe(true);
    const masterUri = pkg.relationships(PRESENTATION_URI).find(
      ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
    )!.resolvedTarget!;
    expect(pkg.relationships(masterUri).find(
      ({ type }) => type === THEME_RELATIONSHIP,
    )?.resolvedTarget).toBe('/ppt/theme/theme2.xml');
    expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBe('Fallback');
  });

  it('inserts a missing master list before slide size with alternate namespace prefixes', () => {
    const pkg = packageWithoutNotesMaster();
    const presentationPart = pkg.requirePart(PRESENTATION_URI);
    const source = new TextDecoder().decode(presentationPart.bytes)
      .replace('<p:sldIdLst/>', '')
      .replaceAll('<p:', '<q:')
      .replaceAll('</p:', '</q:')
      .replace('xmlns:p=', 'xmlns:q=')
      .replaceAll('r:id', 'rel:id')
      .replace('xmlns:r=', 'xmlns:rel=');
    pkg.setPart(PRESENTATION_URI, source, presentationPart.contentType);
    expect(replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'Prefixes')).toBe(true);
    const output = new TextDecoder().decode(pkg.requirePart(PRESENTATION_URI).bytes);
    const relationship = pkg.relationships(PRESENTATION_URI).find(
      ({ type }) => type === NOTES_MASTER_RELATIONSHIP,
    )!;
    expect(output).toContain(
      `<q:notesMasterIdLst><q:notesMasterId rel:id="${relationship.id}"/></q:notesMasterIdLst>`,
    );
    expect(output.indexOf('<q:notesMasterIdLst>')).toBeLessThan(output.indexOf('<q:sldSz'));
  });

  it('rejects partial notes-master state and ambiguous or absent themes without mutation', () => {
    const cases: OpcPackage[] = [];

    const orphanMaster = packageWithoutNotesMaster();
    orphanMaster.setPart(
      NOTES_MASTER_URI,
      `<p:notesMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      NOTES_MASTER_CONTENT_TYPE,
    );
    cases.push(orphanMaster);

    const listOnly = packageWithoutNotesMaster();
    const listPart = listOnly.requirePart(PRESENTATION_URI);
    listOnly.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(listPart.bytes).replace(
        '<p:sldSz',
        '<p:notesMasterIdLst><p:notesMasterId r:id="rId9"/></p:notesMasterIdLst><p:sldSz',
      ),
      listPart.contentType,
    );
    cases.push(listOnly);

    const relationshipOnly = packageWithoutNotesMaster();
    relationshipOnly.setPart(
      NOTES_MASTER_URI,
      `<p:notesMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      NOTES_MASTER_CONTENT_TYPE,
    );
    relationshipOnly.addRelationship(PRESENTATION_URI, {
      type: NOTES_MASTER_RELATIONSHIP,
      target: 'notesMasters/notesMaster1.xml',
    });
    cases.push(relationshipOnly);

    const duplicateRelationships = packageWithNotesMaster();
    duplicateRelationships.setPart(
      '/ppt/notesMasters/notesMaster2.xml',
      `<p:notesMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      NOTES_MASTER_CONTENT_TYPE,
    );
    duplicateRelationships.addRelationship(PRESENTATION_URI, {
      type: NOTES_MASTER_RELATIONSHIP,
      target: 'notesMasters/notesMaster2.xml',
    });
    cases.push(duplicateRelationships);

    const duplicateLists = packageWithNotesMaster();
    const duplicateListsPart = duplicateLists.requirePart(PRESENTATION_URI);
    duplicateLists.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(duplicateListsPart.bytes).replace(
        '</p:presentation>',
        '<p:notesMasterIdLst><p:notesMasterId r:id="rId1"/></p:notesMasterIdLst></p:presentation>',
      ),
      duplicateListsPart.contentType,
    );
    cases.push(duplicateLists);

    const duplicateIdentifiers = packageWithNotesMaster();
    const duplicateIdentifiersPart = duplicateIdentifiers.requirePart(PRESENTATION_URI);
    duplicateIdentifiers.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(duplicateIdentifiersPart.bytes).replace(
        '</p:notesMasterIdLst>',
        '<p:notesMasterId r:id="rId1"/></p:notesMasterIdLst>',
      ),
      duplicateIdentifiersPart.contentType,
    );
    cases.push(duplicateIdentifiers);

    const duplicateIdAttributes = packageWithNotesMaster();
    const duplicateIdAttributesPart = duplicateIdAttributes.requirePart(PRESENTATION_URI);
    duplicateIdAttributes.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(duplicateIdAttributesPart.bytes).replace(
        'r:id="rId1"',
        'r:id="rId1" r:id="rId1"',
      ),
      duplicateIdAttributesPart.contentType,
    );
    cases.push(duplicateIdAttributes);

    const externalMaster = packageWithNotesMaster();
    externalMaster.removeRelationship(PRESENTATION_URI, 'rId1');
    externalMaster.addRelationship(PRESENTATION_URI, {
      id: 'rId1',
      type: NOTES_MASTER_RELATIONSHIP,
      target: 'https://example.com/notesMaster.xml',
      targetMode: 'External',
    });
    cases.push(externalMaster);

    const missingMaster = packageWithoutNotesMaster();
    const missingPresentationPart = missingMaster.requirePart(PRESENTATION_URI);
    missingMaster.setPart(
      PRESENTATION_URI,
      new TextDecoder().decode(missingPresentationPart.bytes).replace(
        '<p:sldSz',
        '<p:notesMasterIdLst><p:notesMasterId r:id="rId9"/></p:notesMasterIdLst><p:sldSz',
      ),
      missingPresentationPart.contentType,
    );
    const missingRelationshipsPart = missingMaster.requirePart('/ppt/_rels/presentation.xml.rels');
    missingMaster.setPart(
      missingRelationshipsPart.uri,
      new TextDecoder().decode(missingRelationshipsPart.bytes).replace(
        '</Relationships>',
        `<Relationship Id="rId9" Type="${NOTES_MASTER_RELATIONSHIP}" Target="notesMasters/missing.xml"/></Relationships>`,
      ),
      missingRelationshipsPart.contentType,
    );
    cases.push(missingMaster);

    const wrongContentMaster = packageWithNotesMaster();
    const wrongContentPart = wrongContentMaster.requirePart(NOTES_MASTER_URI);
    wrongContentMaster.setPart(
      NOTES_MASTER_URI,
      wrongContentPart.bytes,
      'application/octet-stream',
    );
    cases.push(wrongContentMaster);

    const wrongRootMaster = packageWithNotesMaster();
    wrongRootMaster.setPart(
      NOTES_MASTER_URI,
      `<p:notNotesMaster xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      NOTES_MASTER_CONTENT_TYPE,
    );
    cases.push(wrongRootMaster);

    const duplicateDirectThemes = packageWithoutNotesMaster();
    duplicateDirectThemes.addRelationship(PRESENTATION_URI, {
      type: THEME_RELATIONSHIP,
      target: 'theme/theme2.xml',
    });
    cases.push(duplicateDirectThemes);

    const duplicateFallbackThemes = packageWithoutNotesMaster({ directTheme: false });
    duplicateFallbackThemes.addRelationship('/ppt/slideMasters/slideMaster1.xml', {
      type: THEME_RELATIONSHIP,
      target: '../theme/theme1.xml',
    });
    cases.push(duplicateFallbackThemes);

    cases.push(packageWithoutNotesMaster({ directTheme: false, fallbackTheme: false }));

    for (const pkg of cases) {
      const before = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
      const journal = [...pkg.mutations];
      expect(() => replaceSlideNotes(
        pkg,
        PRESENTATION_URI,
        SLIDE_ONE_URI,
        'Rejected',
      )).toThrow(ModelParseError);
      expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(before);
      expect(pkg.mutations).toEqual(journal);
    }
  });

  it('rolls missing-master and notes creation back as one transaction', () => {
    const pkg = packageWithoutNotesMaster();
    const before = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const journal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      replaceSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI, 'Temporary');
      throw new Error('rollback missing notes master');
    })).toThrow('rollback missing notes master');
    expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(before);
    expect(pkg.mutations).toEqual(journal);
    expect(readSlideNotes(pkg, PRESENTATION_URI, SLIDE_ONE_URI)).toBeUndefined();
  });
});
