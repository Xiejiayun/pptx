import { describe, expect, it } from 'vitest';
import { normalizeRichText, renderRichTextParagraphs } from './rich-text.internal.js';

describe('uniform rich text hyperlink rendering', () => {
  it('renders one canonical hyperlink into every non-empty run', () => {
    const paragraphs = normalizeRichText([{
      runs: [
        { text: 'Default underline' },
        {
          text: 'Explicit underline',
          style: {
            underline: false,
            color: { kind: 'scheme', value: 'accent2' },
            fontFamily: 'Aptos',
          },
        },
        { text: '' },
      ],
    }]);
    const baseline = renderRichTextParagraphs(paragraphs);
    const rendered = renderRichTextParagraphs(paragraphs, {
      defaultHyperlink: { url: 'https://example.com?a=1&b=2', tooltip: '' },
      hyperlinkRelationshipId: 'rId7',
    });

    expect(baseline).not.toContain('hlinkClick');
    expect(baseline).not.toContain('u="sng"');
    expect(rendered.match(/<a:hlinkClick /g)).toHaveLength(2);
    expect(rendered.match(/r:id="rId7"/g)).toHaveLength(2);
    expect(rendered.match(/tooltip=""/g)).toHaveLength(2);
    expect(rendered).toContain('<a:rPr lang="en-US" u="sng" dirty="0">');
    expect(rendered).toContain('<a:rPr lang="en-US" u="none" dirty="0">');
    expect(rendered).toContain(
      '<a:schemeClr val="accent2"/></a:solidFill>' +
      '<a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>' +
      '<a:hlinkClick r:id="rId7" tooltip=""/>',
    );
  });

  it('requires uniform hyperlink values and relationship IDs together', () => {
    const paragraphs = normalizeRichText([{ runs: [{ text: 'Pairing' }] }]);
    expect(() => renderRichTextParagraphs(paragraphs, {
      defaultHyperlink: { slide: 2 },
    })).toThrow(/relationship ID/i);
    expect(() => renderRichTextParagraphs(paragraphs, {
      hyperlinkRelationshipId: 'rId7',
    })).toThrow(/hyperlink/i);
  });
});
