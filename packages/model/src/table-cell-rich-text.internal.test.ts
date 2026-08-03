import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { Relationship } from '@pptx/opc';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import type { ShapeHyperlinkReadContext } from './shape-hyperlink.internal.js';
import {
  readTableCellRichText,
  readTableCellText,
  requireEditablePlainTableCellText,
  requireEditableTableCellRichTextState,
} from './table-cell-rich-text.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PART_URI = '/ppt/slides/slide1.xml';

function parseCell(contents: string): {
  readonly xml: LosslessXmlDocument;
  readonly cell: XmlElement;
} {
  return parseSource(
    `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" ` +
      'xmlns:x="urn:test">' + contents + '<a:tcPr/></a:tc>',
  );
}

function parseSource(source: string): {
  readonly xml: LosslessXmlDocument;
  readonly cell: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  const cell = xml.roots[0];
  if (!cell) throw new Error('Fixture has no root cell');
  return { xml, cell };
}

function context(relationships: readonly Relationship[] = []): ShapeHyperlinkReadContext {
  return {
    relationships,
    slidePartUris: ['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml'],
  };
}

describe('table-cell rich text', () => {
  it('reads detached direct paragraphs, runs, fields, soft breaks, styles, and links', () => {
    const fixture = parseCell(
      '<a:txBody><a:bodyPr/><a:lstStyle/>' +
        '<a:p><a:r><a:rPr/><a:t>A</a:t></a:r><a:br/>' +
          '<a:fld id="field-1" type="slidenum"><a:rPr/><a:t>B</a:t></a:fld>' +
          '<a:endParaRPr/></a:p>' +
        '<a:p><a:endParaRPr/></a:p>' +
        '<a:p><a:r><a:rPr b="1"><a:hlinkClick r:id="rId7"/></a:rPr>' +
          '<a:t>C</a:t></a:r><a:endParaRPr/></a:p>' +
      '</a:txBody>',
    );
    const relationships: readonly Relationship[] = [{
      id: 'rId7',
      type: `${RELATIONSHIP_NAMESPACE}/hyperlink`,
      target: 'https://example.com',
      targetMode: 'External',
    }];

    const first = readTableCellRichText(
      fixture.xml,
      fixture.cell,
      context(relationships),
    );
    expect(first).toEqual([
      { runs: [{ text: 'A' }, { text: 'B', softBreakBefore: true }] },
      { runs: [] },
      {
        runs: [{
          text: 'C',
          style: { bold: true, hyperlink: { url: 'https://example.com' } },
        }],
      },
    ]);
    expect(readTableCellText(
      fixture.xml,
      fixture.cell,
      context(relationships),
    )).toBe('A\nB\n\nC');
    expect(fixture.xml.changed).toBe(false);

    const mutable = first as Array<{ runs: Array<{ text: string }> }>;
    mutable[0]!.runs[0]!.text = 'caller mutation';
    expect(readTableCellText(
      fixture.xml,
      fixture.cell,
      context(relationships),
    )).toBe('A\nB\n\nC');
    expect(readTableCellRichText(
      fixture.xml,
      fixture.cell,
      context(relationships),
    )).not.toBe(first);
  });

  it('returns empty snapshots for missing, repeated, foreign, or descendant text bodies', () => {
    const sources = [
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}"/>`,
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}"><a:txBody/><a:txBody/></a:tc>`,
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:x="urn:test">` +
        '<x:txBody><x:p><x:r><x:t>Foreign</x:t></x:r></x:p></x:txBody></a:tc>',
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}"><a:tcPr><a:txBody>` +
        '<a:p><a:r><a:t>Descendant</a:t></a:r></a:p>' +
        '</a:txBody></a:tcPr></a:tc>',
      `<x:tc xmlns:x="urn:test" xmlns:a="${DRAWING_NAMESPACE}">` +
        '<a:txBody><a:p><a:r><a:t>Wrong cell</a:t></a:r></a:p></a:txBody></x:tc>',
    ];

    for (const source of sources) {
      const fixture = parseSource(source);
      expect(readTableCellRichText(fixture.xml, fixture.cell, context([])), source)
        .toEqual([]);
      expect(readTableCellText(fixture.xml, fixture.cell, context([])), source).toBe('');
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('returns rich edit state only for one owned body with direct paragraphs', () => {
    const valid = parseCell(
      '<a:txBody><a:bodyPr/><a:p><a:r><a:t>Editable</a:t></a:r></a:p>' +
        '<a:p><a:endParaRPr/></a:p></a:txBody>',
    );
    expect(requireEditableTableCellRichTextState(
      valid.xml,
      valid.cell,
      context([]),
      PART_URI,
    ).paragraphs).toEqual([
      { runs: [{ text: 'Editable' }] },
      { runs: [] },
    ]);

    const invalid = [
      parseCell('<a:txBody><a:bodyPr/></a:txBody>'),
      parseCell('<a:txBody/><a:txBody/>'),
      parseCell('<x:txBody><x:p/></x:txBody>'),
      parseCell('<a:tcPr><a:txBody><a:p/></a:txBody></a:tcPr>'),
    ];
    for (const fixture of invalid) {
      expect(() => requireEditableTableCellRichTextState(
        fixture.xml,
        fixture.cell,
        context([]),
        PART_URI,
      )).toThrow(ModelParseError);
      expect(() => requireEditableTableCellRichTextState(
        fixture.xml,
        fixture.cell,
        context([]),
        PART_URI,
      )).toThrow(PART_URI);
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('resolves only an exact single-paragraph single-run plain text target', () => {
    const valid = [
      parseCell('<a:txBody><a:p><a:r><a:t>Plain</a:t></a:r></a:p></a:txBody>'),
      parseCell(
        '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>' +
          '<a:r><a:rPr b="1"/><a:t>Styled</a:t></a:r>' +
          '<a:endParaRPr lang="en-US"/></a:p></a:txBody>',
      ),
    ];
    expect(valid.map(({ xml, cell }) =>
      xml.text(requireEditablePlainTableCellText(xml, cell, PART_URI))))
      .toEqual(['Plain', 'Styled']);

    const invalidContents = [
      '<a:txBody><a:p/><a:p/></a:txBody>',
      '<a:txBody><a:p><a:r><a:t>A</a:t></a:r><a:r><a:t>B</a:t></a:r></a:p></a:txBody>',
      '<a:txBody><a:p><a:r><a:t>A</a:t></a:r><a:br/></a:p></a:txBody>',
      '<a:txBody><a:p><a:fld><a:t>A</a:t></a:fld></a:p></a:txBody>',
      '<a:txBody><a:p><a:r><a:t>A</a:t><a:t>B</a:t></a:r></a:p></a:txBody>',
      '<a:txBody><a:p><x:r><x:t>A</x:t></x:r></a:p></a:txBody>',
      '<a:txBody><a:p><a:r><x:t>A</x:t></a:r></a:p></a:txBody>',
      '<a:tcPr><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tcPr>',
    ];
    for (const contents of invalidContents) {
      const fixture = parseCell(contents);
      expect(() => requireEditablePlainTableCellText(
        fixture.xml,
        fixture.cell,
        PART_URI,
      ), contents).toThrow(ModelParseError);
      expect(fixture.xml.changed).toBe(false);
    }
  });
});
