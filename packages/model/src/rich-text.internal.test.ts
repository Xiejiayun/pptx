import { describe, expect, it } from 'vitest';
import { normalizeRichText, renderRichTextParagraphs } from './rich-text.internal.js';
import { normalizeHyperlink } from './shape-hyperlink.internal.js';

describe('rich text run hyperlink normalization', () => {
  it('normalizes direct links, suppression, and detached nested values', () => {
    const input = { url: 'https://local.example', tooltip: '' };
    const paragraphs = normalizeRichText([{
      runs: [
        { text: 'URL', style: { hyperlink: input } },
        { text: 'Slide', style: { hyperlink: { slide: 2, tooltip: 'Next' } } },
        { text: 'Suppressed', style: { hyperlink: false } },
      ],
    }]);
    input.url = 'https://changed.example';

    expect(paragraphs[0]!.runs[0]!.style?.hyperlink).toEqual({
      url: 'https://local.example',
      tooltip: '',
    });
    expect(Object.isFrozen(paragraphs[0]!.runs[0]!.style?.hyperlink)).toBe(true);
    expect(paragraphs[0]!.runs[1]!.style?.hyperlink).toEqual({
      slide: 2,
      tooltip: 'Next',
    });
    expect(paragraphs[0]!.runs[2]!.style?.hyperlink).toBe(false);
  });

  it('rejects invalid and empty-run direct links', () => {
    expect(() => normalizeRichText([{
      runs: [{ text: '', style: { hyperlink: { url: 'https://empty.example' } } }],
    }])).toThrow(/non-empty text/i);

    for (const hyperlink of [
      null,
      true,
      1,
      'https://coerced.example',
      {},
      { url: 'https://both.example', slide: 1 },
      { url: 'https://example.com', tooltip: 1 },
    ]) {
      expect(() => normalizeRichText([{
        runs: [{ text: 'Invalid', style: { hyperlink } }],
      }]), String(hyperlink)).toThrow();
    }
  });

  it('rejects unsafe style containers without invoking hyperlink accessors', () => {
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'hyperlink', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return { url: 'https://unsafe.example' };
      },
    });
    const inherited = Object.create({ hyperlink: { url: 'https://inherited.example' } });
    const symbol = { [Symbol('unsafe')]: true };
    class Style {}

    for (const style of [accessor, inherited, symbol, new Style()]) {
      expect(() => normalizeRichText([{
        runs: [{ text: 'Unsafe', style }],
      }])).toThrow();
    }
    expect(accessorCalls).toBe(0);
  });
});

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

describe('rich text run hyperlink rendering', () => {
  it('resolves local values over outer defaults with suppression and independent IDs', () => {
    const paragraphs = normalizeRichText([{
      runs: [
        { text: 'Inherited' },
        { text: 'Local', style: { hyperlink: { url: 'https://local.example' } } },
        { text: 'Suppressed', style: { hyperlink: false } },
        {
          text: 'No underline',
          style: {
            hyperlink: { slide: 2, tooltip: '' },
            underline: false,
          },
        },
        { text: 'Same target', style: { hyperlink: { url: 'https://local.example' } } },
      ],
    }]);
    const rendered = renderRichTextParagraphs(paragraphs, {
      defaultHyperlink: normalizeHyperlink({ url: 'https://outer.example' }, 'outer'),
      hyperlinkRelationshipId: 'rIdOuter',
      runHyperlinkRelationshipIds: [[
        undefined,
        'rIdLocal',
        undefined,
        'rIdInternal',
        'rIdSameTarget',
      ]],
    });
    const runs = [...rendered.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map(([run]) => run);

    expect(runs).toHaveLength(5);
    expect(runs[0]).toContain('r:id="rIdOuter"');
    expect(runs[1]).toContain('r:id="rIdLocal"');
    expect(runs[2]).not.toContain('hlinkClick');
    expect(runs[3]).toContain('u="none"');
    expect(runs[3]).toContain('r:id="rIdInternal"');
    expect(runs[3]).toContain('tooltip=""');
    expect(runs[3]).toContain('action="ppaction://hlinksldjump"');
    expect(runs[4]).toContain('r:id="rIdSameTarget"');
    expect(rendered.match(/<a:hlinkClick /g)).toHaveLength(4);
  });

  it('requires every direct hyperlink and relationship ID together', () => {
    const linked = normalizeRichText([{
      runs: [{ text: 'Local', style: { hyperlink: { url: 'https://local.example' } } }],
    }]);
    expect(() => renderRichTextParagraphs(linked)).toThrow(/relationship ID/i);

    const plain = normalizeRichText([{ runs: [{ text: 'Plain' }] }]);
    expect(() => renderRichTextParagraphs(plain, {
      runHyperlinkRelationshipIds: [['rIdUnexpected']],
    })).toThrow(/hyperlink/i);
  });
});
