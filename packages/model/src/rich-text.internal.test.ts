import { describe, expect, it } from 'vitest';
import { normalizeRichText, renderRichTextParagraphs } from './rich-text.internal.js';
import { normalizeHyperlink } from './shape-hyperlink.internal.js';

describe('rich text breakLine normalization', () => {
  it('splits non-final runs into canonical paragraphs and removes transient markers', () => {
    const input = [{
      align: 'center',
      rtl: true,
      marginLeft: 12,
      marginRight: 18,
      indent: 4,
      bullet: false,
      level: 2,
      spacing: { before: 3, after: 5, line: { kind: 'multiple', factor: 1.25 } },
      tabStops: [{ position: 1.5, alignment: 'right' }],
      runs: [
        { text: 'First', style: { bold: true }, breakLine: true },
        { text: 'Second', softBreakBefore: true, breakLine: false },
        { text: '', breakLine: true },
        { text: 'Final', breakLine: true },
      ],
    }, {
      runs: [],
    }];
    const before = structuredClone(input);

    const paragraphs = normalizeRichText(input);

    const properties = {
      align: 'center',
      bullet: false,
      indent: 4,
      level: 2,
      marginLeft: 12,
      marginRight: 18,
      rtl: true,
      spacing: { before: 3, after: 5, line: { kind: 'multiple', factor: 1.25 } },
      tabStops: [{ positionEmu: 1_371_600, alignment: 'right' }],
    };
    expect(paragraphs).toEqual([
      { ...properties, runs: [{ text: 'First', style: { bold: true } }] },
      {
        ...properties,
        runs: [
          { text: 'Second', softBreakBefore: true },
          { text: '' },
        ],
      },
      { ...properties, runs: [{ text: 'Final' }] },
      { runs: [] },
    ]);
    expect(paragraphs[0]).not.toBe(paragraphs[1]);
    expect(paragraphs[0]!.runs).not.toBe(paragraphs[1]!.runs);
    expect(Object.hasOwn(paragraphs[0]!.runs[0]!, 'breakLine')).toBe(false);
    expect(Object.hasOwn(paragraphs[1]!.runs[0]!, 'breakLine')).toBe(false);
    expect(input).toEqual(before);
  });

  it('preserves empty and consecutive splits without adding a trailing paragraph', () => {
    expect(normalizeRichText([{
      runs: [
        { text: '', breakLine: true },
        { text: '', breakLine: true },
        { text: 'End' },
      ],
    }])).toEqual([
      { runs: [{ text: '' }] },
      { runs: [{ text: '' }] },
      { runs: [{ text: 'End' }] },
    ]);

    expect(normalizeRichText([{
      runs: [
        { text: 'Together', breakLine: false },
        { text: 'Still together', breakLine: undefined },
        { text: 'Trailing only', breakLine: true },
      ],
    }])).toEqual([{
      runs: [
        { text: 'Together' },
        { text: 'Still together' },
        { text: 'Trailing only' },
      ],
    }]);
  });

  it('preserves an explicit soft break on the first run after a paragraph split', () => {
    const paragraphs = normalizeRichText([{
      runs: [
        { text: 'Before', breakLine: true },
        { text: 'After', softBreakBefore: true },
      ],
    }]);

    expect(paragraphs).toEqual([
      { runs: [{ text: 'Before' }] },
      { runs: [{ text: 'After', softBreakBefore: true }] },
    ]);
    expect(renderRichTextParagraphs(paragraphs)).toMatch(
      /<a:p>[\s\S]*?<a:t xml:space="preserve">Before<\/a:t>[\s\S]*?<\/a:p>\s*<a:p>[\s\S]*?<a:br\/>[\s\S]*?<a:t xml:space="preserve">After<\/a:t>/,
    );
  });

  it('strictly rejects invalid values and unsafe run containers without accessors', () => {
    for (const breakLine of [
      null,
      'true',
      1,
      {},
      Symbol('break'),
      new Boolean(true),
      () => true,
    ]) {
      expect(() => normalizeRichText([{
        runs: [{ text: 'Invalid', breakLine }],
      }]), String(breakLine)).toThrow(/breakLine must be a boolean/i);
    }

    let accessorCalls = 0;
    const accessor = Object.defineProperty({ text: 'Accessor' }, 'breakLine', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return true;
      },
    });
    const inherited = Object.create({ breakLine: true }) as { text?: string };
    inherited.text = 'Inherited';
    class Run {
      text = 'Class';
      breakLine = true;
    }

    for (const run of [accessor, inherited, new Run(), { text: 'Symbol', [Symbol('break')]: true }]) {
      expect(() => normalizeRichText([{ runs: [run] }])).toThrow();
    }
    expect(accessorCalls).toBe(0);
  });
});

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

describe('rich text run style defaults', () => {
  it('resolves run defaults with explicit overrides and hyperlink color suppression', () => {
    const paragraphs = normalizeRichText([{
      runs: [
        { text: 'Inherited' },
        {
          text: 'Local false',
          style: {
            fontFamily: 'Arial',
            fontSize: 10,
            bold: false,
            color: { kind: 'srgb', value: 'FF0000' },
          },
        },
        {
          text: 'Linked',
          style: { hyperlink: { url: 'https://example.com' } },
        },
      ],
    }, { runs: [] }]);

    const rendered = renderRichTextParagraphs(paragraphs, {
      defaultFontFamily: 'Aptos',
      defaultFontSize: 18.25,
      defaultBold: true,
      defaultColor: { kind: 'scheme', value: 'accent1' },
      suppressDefaultColorForHyperlinks: true,
      runHyperlinkRelationshipIds: [[undefined, undefined, 'rId7'], []],
    });
    const runs = [...rendered.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)]
      .map(([run]) => run);

    expect(runs).toHaveLength(3);
    expect(runs[0]).toContain('<a:rPr lang="en-US" sz="1825" b="1" dirty="0">');
    expect(runs[0]).toContain('<a:schemeClr val="accent1"/>');
    expect(runs[0]).toContain(
      '<a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>',
    );
    expect(runs[1]).toContain('<a:rPr lang="en-US" sz="1000" b="0" dirty="0">');
    expect(runs[1]).toContain('<a:srgbClr val="FF0000"/>');
    expect(runs[1]).toContain(
      '<a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/>',
    );
    expect(runs[2]).toContain('<a:rPr lang="en-US" sz="1825" b="1" u="sng" dirty="0">');
    expect(runs[2]).not.toContain('<a:solidFill>');
    expect(runs[2]).toContain(
      '<a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>',
    );
    expect(runs[2]).toContain('<a:hlinkClick r:id="rId7"/>');
    expect(rendered.match(
      /<a:endParaRPr lang="en-US" sz="1825" dirty="0"><a:latin typeface="Aptos"\/><a:ea typeface="Aptos"\/><a:cs typeface="Aptos"\/><\/a:endParaRPr>/g,
    )).toHaveLength(2);
  });

  it('keeps legacy rendering byte-identical when defaults are omitted', () => {
    const paragraphs = normalizeRichText([{ runs: [{ text: 'Legacy' }] }]);

    expect(renderRichTextParagraphs(paragraphs)).toBe(
      '<a:p><a:pPr indent="0" marL="0"><a:buNone/></a:pPr>' +
      '<a:r><a:rPr lang="en-US" dirty="0">' +
      '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
      '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>' +
      '<a:cs typeface="+mn-cs"/></a:rPr>' +
      '<a:t xml:space="preserve">Legacy</a:t></a:r>' +
      '<a:endParaRPr lang="en-US" dirty="0"/></a:p>',
    );
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
