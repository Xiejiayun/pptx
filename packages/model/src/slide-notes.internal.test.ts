import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  LosslessXmlError,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  createNotesSlideXml,
  normalizeSlideNotes,
  readNotesBody,
  replaceNotesBody,
} from './slide-notes.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART_URI = '/ppt/notesSlides/notesSlide1.xml';

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
