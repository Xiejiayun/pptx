import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { Relationship } from '@pptx/opc';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import {
  readTableCellHyperlink,
  requireEditableTableCellHyperlinkState,
} from './table-cell-hyperlink.internal.js';
import type { ShapeHyperlinkReadContext } from './shape-hyperlink.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HYPERLINK_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/hyperlink`;
const SLIDE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/slide`;

function parseCell(contents: string): {
  readonly xml: LosslessXmlDocument;
  readonly cell: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(
    `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" ` +
    'xmlns:x="urn:test">' +
    `<a:txBody>${contents}</a:txBody><a:tcPr/></a:tc>`,
  );
  return { xml, cell: xml.roots[0]! };
}

function parseSource(source: string): {
  readonly xml: LosslessXmlDocument;
  readonly cell: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  return { xml, cell: xml.roots[0]! };
}

function paragraph(click = ''): string {
  return '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">' +
    `${click}</a:rPr><a:t>Linked</a:t></a:r><a:endParaRPr/></a:p>`;
}

function external(
  target = 'https://example.com?a=1&b=2',
  id = 'rId7',
): Relationship {
  return {
    id,
    type: HYPERLINK_RELATIONSHIP,
    target,
    targetMode: 'External',
  };
}

function internal(
  resolvedTarget = '/ppt/slides/slide2.xml',
  id = 'rId7',
): Relationship {
  return {
    id,
    type: SLIDE_RELATIONSHIP,
    target: resolvedTarget.slice(resolvedTarget.lastIndexOf('/') + 1),
    targetMode: 'Internal',
    resolvedTarget,
  };
}

function context(
  relationships: readonly Relationship[],
  slidePartUris = ['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml'],
): ShapeHyperlinkReadContext {
  return { relationships, slidePartUris };
}

describe('table-cell hyperlinks', () => {
  it('reads detached frozen external and internal single-run hyperlinks', () => {
    const absentTooltip = parseCell(paragraph('<a:hlinkClick r:id="rId7"/>'));
    const first = readTableCellHyperlink(
      absentTooltip.xml,
      absentTooltip.cell,
      context([external()]),
    );
    const second = readTableCellHyperlink(
      absentTooltip.xml,
      absentTooltip.cell,
      context([external()]),
    );
    expect(first).toEqual({ url: 'https://example.com?a=1&b=2' });
    expect(Object.isFrozen(first)).toBe(true);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);

    const emptyTooltip = parseCell(paragraph(
      '<a:hlinkClick r:id="rId7" tooltip=""/>',
    ));
    expect(readTableCellHyperlink(
      emptyTooltip.xml,
      emptyTooltip.cell,
      context([external()]),
    )).toEqual({ url: 'https://example.com?a=1&b=2', tooltip: '' });

    const slide = parseCell(paragraph(
      '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
    ));
    expect(readTableCellHyperlink(
      slide.xml,
      slide.cell,
      context([internal()]),
    )).toEqual({ slide: 2 });
  });

  it('imports supported PptxGenJS run hyperlink attributes', () => {
    const parsed = parseCell(paragraph(
      '<a:hlinkClick r:id="rId7" invalidUrl="" action="" tgtFrame="" ' +
      'tooltip="Visit &amp; learn" history="1" highlightClick="0" endSnd="0"/>',
    ));
    expect(readTableCellHyperlink(
      parsed.xml,
      parsed.cell,
      context([external()]),
    )).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    });
  });

  it('requires a safe editable run and distinguishes absent from linked state', () => {
    const unlinked = parseCell(paragraph());
    const absent = requireEditableTableCellHyperlinkState(
      unlinked.cell,
      context([]),
      '/ppt/slides/slide1.xml',
    );
    expect(absent.properties.localName).toBe('rPr');
    expect(absent.hyperlink).toBeUndefined();
    expect(absent.relationshipId).toBeUndefined();

    const linked = parseCell(paragraph(
      '<a:hlinkClick r:id="rId7" tooltip=""/>',
    ));
    const state = requireEditableTableCellHyperlinkState(
      linked.cell,
      context([external('https://example.com', 'rId7')]),
      '/ppt/slides/slide1.xml',
    );
    expect(state.properties.localName).toBe('rPr');
    expect(state.hyperlink).toEqual({ url: 'https://example.com', tooltip: '' });
    expect(state.relationshipId).toBe('rId7');
  });

  it('rejects structurally unsafe or undecodable editable state', () => {
    const cases = [
      {
        parsed: parseCell('<a:bodyPr/><a:lstStyle/><a:p/><a:p/>'),
        relationships: [],
      },
      {
        parsed: parseCell(
          '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>A</a:t></a:r>' +
          '<a:r><a:rPr/><a:t>B</a:t></a:r></a:p>',
        ),
        relationships: [],
      },
      {
        parsed: parseCell(
          '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t></a:r></a:p>',
        ),
        relationships: [],
      },
      {
        parsed: parseCell(paragraph('<x:hlinkClick r:id="rId7"/>')),
        relationships: [external()],
      },
      {
        parsed: parseCell(paragraph(
          '<a:hlinkClick r:id="rId7"/><a:hlinkClick r:id="rId8"/>',
        )),
        relationships: [external(), external('https://other.example', 'rId8')],
      },
      {
        parsed: parseCell(paragraph('<a:hlinkClick r:id="missing"/>')),
        relationships: [external()],
      },
      {
        parsed: parseCell(paragraph(
          '<a:hlinkClick r:id="rId7" action="ppaction://unsupported"/>',
        )),
        relationships: [external()],
      },
    ];
    for (const { parsed, relationships } of cases) {
      expect(() => requireEditableTableCellHyperlinkState(
        parsed.cell,
        context(relationships),
        '/ppt/slides/slide1.xml',
      )).toThrow(ModelParseError);
      expect(parsed.xml.changed).toBe(false);
    }
  });

  it('returns undefined for unsupported cell text ownership', () => {
    const unsupported = [
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Plain</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p/><a:p/>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>A</a:t></a:r>' +
        '<a:r><a:rPr/><a:t>B</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:fld><a:rPr/><a:t>Field</a:t></a:fld></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:br/><a:r><a:rPr/><a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:rPr/>' +
        '<a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:pPr/><a:r><a:rPr>' +
        '<a:hlinkClick r:id="rId7"/></a:rPr><a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId7"/>' +
        '</a:rPr><a:t>A</a:t></a:r><a:pPr/></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/><a:r><a:rPr>' +
        '<a:hlinkClick r:id="rId7"/></a:rPr><a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId7"/>' +
        '</a:rPr><a:t>A</a:t></a:r><a:endParaRPr/><a:endParaRPr/></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t><a:rPr>' +
        '<a:hlinkClick r:id="rId7"/></a:rPr></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId7"/>' +
        '<a:hlinkClick r:id="rId8"/></a:rPr><a:t>A</a:t></a:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><x:r><a:rPr><a:hlinkClick r:id="rId7"/>' +
        '</a:rPr><a:t>A</a:t></x:r></a:p>',
      '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><x:keep>' +
        '<a:hlinkClick r:id="rId7"/></x:keep></a:rPr><a:t>A</a:t></a:r></a:p>',
    ];

    for (const contents of unsupported) {
      const parsed = parseCell(contents);
      expect(readTableCellHyperlink(
        parsed.xml,
        parsed.cell,
        context([external(), external('https://two.example', 'rId8')]),
      ), contents).toBeUndefined();
    }

    const repeatedBody = parseSource(
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">` +
      `<a:txBody>${paragraph()}</a:txBody><a:txBody>${paragraph()}</a:txBody>` +
      '<a:tcPr/></a:tc>',
    );
    expect(readTableCellHyperlink(
      repeatedBody.xml,
      repeatedBody.cell,
      context([external()]),
    )).toBeUndefined();

    const wrongNamespaceBody = parseSource(
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" ` +
      'xmlns:x="urn:test"><x:txBody/><a:txBody>' +
      `${paragraph('<a:hlinkClick r:id="rId7"/>')}</a:txBody>` +
      '<a:tcPr/></a:tc>',
    );
    expect(readTableCellHyperlink(
      wrongNamespaceBody.xml,
      wrongNamespaceBody.cell,
      context([external()]),
    )).toBeUndefined();

    const wrongNamespaceParagraph = parseCell(
      '<a:bodyPr/><a:lstStyle/><x:p/><a:p><a:r><a:rPr>' +
      '<a:hlinkClick r:id="rId7"/></a:rPr><a:t>A</a:t></a:r></a:p>',
    );
    expect(readTableCellHyperlink(
      wrongNamespaceParagraph.xml,
      wrongNamespaceParagraph.cell,
      context([external()]),
    )).toBeUndefined();

    const wrongCell = parseSource(
      `<x:tc xmlns:x="urn:test" xmlns:a="${DRAWING_NAMESPACE}" ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}"><a:txBody>${paragraph(
        '<a:hlinkClick r:id="rId7"/>',
      )}</a:txBody></x:tc>`,
    );
    expect(readTableCellHyperlink(
      wrongCell.xml,
      wrongCell.cell,
      context([external()]),
    )).toBeUndefined();
  });

  it('returns undefined for unsafe relationship and click state', () => {
    const sources = [
      { click: '<a:hlinkClick r:id=""/>', relationships: [external()] },
      { click: '<a:hlinkClick r:id="missing"/>', relationships: [external()] },
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        relationships: [external(), external('https://duplicate.example')],
      },
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        relationships: [internal()],
      },
      {
        click: '<a:hlinkClick r:id="rId7" action="ppaction://unsupported"/>',
        relationships: [external()],
      },
      {
        click: '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        relationships: [internal('/ppt/slides/missing.xml')],
      },
      {
        click: '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        relationships: [internal()],
        slides: ['/ppt/slides/slide2.xml', '/ppt/slides/slide2.xml'],
      },
      {
        click: '<a:hlinkClick r:id="rId7"><a:unknown/></a:hlinkClick>',
        relationships: [external()],
      },
    ];

    for (const source of sources) {
      const parsed = parseCell(paragraph(source.click));
      expect(readTableCellHyperlink(
        parsed.xml,
        parsed.cell,
        context(source.relationships, source.slides),
      ), source.click).toBeUndefined();
    }
  });
});
