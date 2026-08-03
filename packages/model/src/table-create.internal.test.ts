import { describe, expect, it } from 'vitest';
import {
  distributeTableDimension,
  normalizeTableDefinition,
  renderTableGraphicFrame,
} from './table-create.internal.js';

describe('table auto-page option normalization', () => {
  it('normalizes detached strict controls, defaults, and scalar margins', () => {
    const margin = [100, 200, 300, 400] as [number, number, number, number];
    const definition = normalizeTableDefinition(
      [['Header'], ['Body']],
      {
        autoPage: true,
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageSlideStartY: 500,
        slideMargin: margin,
        rowHeights: [40, 60],
      },
    );
    expect(definition.autoPage).toEqual({
      repeatHeader: true,
      headerRows: 1,
      slideStartY: 500,
      slideMargin: [100, 200, 300, 400],
      measureContent: false,
    });
    expect(Object.isFrozen(definition.autoPage)).toBe(true);
    expect(Object.isFrozen(definition.autoPage!.slideMargin)).toBe(true);
    margin[0] = 999;
    expect(definition.autoPage!.slideMargin).toEqual([100, 200, 300, 400]);

    expect(normalizeTableDefinition(
      [['Header'], ['Body']],
      {
        autoPage: true,
        autoPageRepeatHeader: true,
        slideMargin: 250,
        rowHeights: 50,
      },
    ).autoPage).toEqual({
      repeatHeader: true,
      headerRows: 1,
      slideMargin: [250, 250, 250, 250],
      measureContent: false,
    });
    expect(normalizeTableDefinition(
      [['A'], ['B']],
      { autoPage: true, rowHeights: [10, 20] },
    ).autoPage).toEqual({
      repeatHeader: false,
      headerRows: 0,
      measureContent: false,
    });
    expect(normalizeTableDefinition(
      [['A']],
      { autoPage: false, rowHeights: [10] },
    ).autoPage).toBeUndefined();
  });

  it('normalizes automatic rows and strict detached table/cell measurement weights', () => {
    const definition = normalizeTableDefinition(
      [
        [{
          text: 'A',
          options: {
            autoPageCharWeight: -0.25,
            autoPageLineWeight: 0,
          },
        }],
        ['B'],
      ],
      {
        autoPage: true,
        autoPageCharWeight: 0,
        autoPageLineWeight: 0.5,
        rowHeights: [0, 400_000],
      },
    );

    expect(definition.autoPage).toEqual({
      repeatHeader: false,
      headerRows: 0,
      charWeight: 0,
      lineWeight: 0.5,
      measureContent: true,
    });
    expect(definition.rows[0]![0]).toMatchObject({
      autoPageCharWeight: -0.25,
      autoPageLineWeight: 0,
    });
    expect(definition.rowHeights).toEqual([0, 400_000]);
    expect(definition.autoRowHeight).toBe(true);
    expect(definition.height).toBe(400_000);
    expect(Object.isFrozen(definition.autoPage)).toBe(true);

    expect(normalizeTableDefinition(
      [['A'], ['B']],
      { autoPage: true },
    )).toMatchObject({
      autoRowHeight: true,
      rowHeights: [0, 0],
      autoPage: {
        repeatHeader: false,
        headerRows: 0,
        measureContent: true,
      },
    });
    expect(normalizeTableDefinition(
      [['A'], ['B']],
      { autoPage: true, rowHeights: 0 },
    )).toMatchObject({
      autoRowHeight: true,
      rowHeights: [0, 0],
      height: 0,
      autoPage: { measureContent: true },
    });
    expect(normalizeTableDefinition(
      [['A'], ['B']],
      {
        autoPage: true,
        autoPageLineWeight: 0,
        rowHeights: [10, 20],
      },
    )).toMatchObject({
      autoRowHeight: false,
      rowHeights: [10, 20],
      autoPage: { lineWeight: 0, measureContent: true },
    });
    expect(normalizeTableDefinition(
      [['A']],
      {
        autoPage: true,
        autoPageCharWeight: -1,
        autoPageLineWeight: 1,
        rowHeights: [10],
      },
    ).autoPage).toEqual({
      repeatHeader: false,
      headerRows: 0,
      charWeight: -1,
      lineWeight: 1,
      measureContent: true,
    });
    expect(normalizeTableDefinition(
      [['A']],
      {
        autoPage: true,
        autoPageCharWeight: undefined,
        autoPageLineWeight: undefined,
        rowHeights: [10],
      },
    ).autoPage).toEqual({
      repeatHeader: false,
      headerRows: 0,
      measureContent: false,
    });
  });

  it('selects measured mode for a cell weight without materializing the table default', () => {
    const definition = normalizeTableDefinition([[
      { text: 'A', options: { autoPageCharWeight: 1 } },
      { text: 'B', options: { autoPageLineWeight: -1 } },
    ]], {
      autoPage: true,
      rowHeights: [100],
    });
    expect(definition.autoPage).toEqual({
      repeatHeader: false,
      headerRows: 0,
      measureContent: true,
    });
    expect(definition.rows[0]!.map((cell) => ({
      char: cell.autoPageCharWeight,
      line: cell.autoPageLineWeight,
    }))).toEqual([
      { char: 1, line: undefined },
      { char: undefined, line: -1 },
    ]);
  });

  it('detaches measurement weights and canonicalizes negative zero', () => {
    const cellOptions = { autoPageCharWeight: -0 };
    const tableOptions = {
      autoPage: true,
      autoPageLineWeight: -0,
      rowHeights: [100],
    };
    const definition = normalizeTableDefinition([[
      { text: 'A', options: cellOptions },
    ]], tableOptions);

    cellOptions.autoPageCharWeight = 1;
    tableOptions.autoPageLineWeight = 1;
    expect(definition.autoPage).toMatchObject({ lineWeight: 0 });
    expect(definition.rows[0]![0]!.autoPageCharWeight).toBe(0);
    expect(Object.is(definition.autoPage!.lineWeight, -0)).toBe(false);
    expect(Object.is(definition.rows[0]![0]!.autoPageCharWeight, -0)).toBe(false);
  });

  it.each([
    ['non-boolean enable', { autoPage: 'yes', rowHeights: [10, 10] }, TypeError],
    ['controls while disabled', {
      autoPage: false,
      autoPageRepeatHeader: true,
      rowHeights: [10, 10],
    }, TypeError],
    ['header count without repeat', {
      autoPage: true,
      autoPageHeaderRows: 1,
      rowHeights: [10, 10],
    }, TypeError],
    ['zero header count', {
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 0,
      rowHeights: [10, 10],
    }, RangeError],
    ['too many header rows', {
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 3,
      rowHeights: [10, 10],
    }, RangeError],
    ['zero direct row with explicit height', {
      autoPage: true,
      rowHeights: [0, 10],
      height: 10,
    }, RangeError],
    ['placeholder owner', {
      autoPage: true,
      placeholder: 'Body',
      rowHeights: [10, 10],
    }, TypeError],
    ['negative start Y', {
      autoPage: true,
      autoPageSlideStartY: -1,
      rowHeights: [10, 10],
    }, RangeError],
    ['fractional start Y', {
      autoPage: true,
      autoPageSlideStartY: 1.5,
      rowHeights: [10, 10],
    }, TypeError],
    ['short margin tuple', {
      autoPage: true,
      slideMargin: [1, 2, 3],
      rowHeights: [10, 10],
    }, RangeError],
    ['negative margin', {
      autoPage: true,
      slideMargin: [1, 2, -1, 4],
      rowHeights: [10, 10],
    }, RangeError],
    ['table char weight below range', {
      autoPage: true,
      autoPageCharWeight: -1.000_001,
      rowHeights: [10, 10],
    }, RangeError],
    ['table line weight above range', {
      autoPage: true,
      autoPageLineWeight: 1.000_001,
      rowHeights: [10, 10],
    }, RangeError],
    ['non-numeric table weight', {
      autoPage: true,
      autoPageCharWeight: '0',
      rowHeights: [10, 10],
    }, TypeError],
    ['non-finite table weight', {
      autoPage: true,
      autoPageLineWeight: Number.NaN,
      rowHeights: [10, 10],
    }, TypeError],
    ['weight while disabled', {
      autoPage: false,
      autoPageCharWeight: 0,
      rowHeights: [10, 10],
    }, TypeError],
    ['weight without auto-page', {
      autoPageLineWeight: 0,
      rowHeights: [10, 10],
    }, TypeError],
  ])('rejects %s before package access', (_name, options, error) => {
    expect(() => normalizeTableDefinition([['A'], ['B']], options)).toThrow(error);
  });

  it.each([
    ['char below range', { autoPageCharWeight: -1.1 }, RangeError],
    ['char above range', { autoPageCharWeight: 1.1 }, RangeError],
    ['line string', { autoPageLineWeight: '0' }, TypeError],
    ['line infinity', { autoPageLineWeight: Number.POSITIVE_INFINITY }, TypeError],
  ])('rejects invalid cell measurement %s', (_name, cellOptions, error) => {
    expect(() => normalizeTableDefinition([[
      { text: 'A', options: cellOptions },
    ]], { autoPage: true, rowHeights: [10] })).toThrow(error);
  });

  it('rejects cell measurement controls unless the enclosing table enables auto-page', () => {
    for (const options of [undefined, { autoPage: false }]) {
      expect(() => normalizeTableDefinition([[
        { text: 'A', options: { autoPageLineWeight: 0 } },
      ]], options)).toThrow(TypeError);
    }
  });

  it('rejects unsafe control descriptors, tuples, and object shapes', () => {
    const accessor = { autoPage: true, rowHeights: [10, 10] };
    Object.defineProperty(accessor, 'slideMargin', { get: () => 1, enumerable: true });
    expect(() => normalizeTableDefinition([['A'], ['B']], accessor)).toThrow(TypeError);

    const sparse = [1, 2, 3, 4];
    delete sparse[2];
    expect(() => normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      slideMargin: sparse,
      rowHeights: [10, 10],
    })).toThrow(TypeError);

    const symbol = [1, 2, 3, 4];
    Object.defineProperty(symbol, Symbol('unsafe'), { value: true });
    expect(() => normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      slideMargin: symbol,
      rowHeights: [10, 10],
    })).toThrow(TypeError);

    class Margin extends Array<number> {}
    expect(() => normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      slideMargin: new Margin(1, 2, 3, 4),
      rowHeights: [10, 10],
    })).toThrow(TypeError);

    let accessorCalls = 0;
    const tableAccessor = { autoPage: true, rowHeights: [10, 10] };
    Object.defineProperty(tableAccessor, 'autoPageCharWeight', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 0;
      },
    });
    expect(() => normalizeTableDefinition([['A'], ['B']], tableAccessor))
      .toThrow(TypeError);

    const cellAccessor = {};
    Object.defineProperty(cellAccessor, 'autoPageLineWeight', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 0;
      },
    });
    expect(() => normalizeTableDefinition([[
      { text: 'A', options: cellAccessor },
    ]], { autoPage: true, rowHeights: [10] })).toThrow(TypeError);
    expect(accessorCalls).toBe(0);

    class CellWeights {
      autoPageCharWeight = 0;
    }
    expect(() => normalizeTableDefinition([[
      { text: 'A', options: new CellWeights() },
    ]], { autoPage: true, rowHeights: [10] })).toThrow(TypeError);

    const inheritedTable = Object.assign(
      Object.create({ autoPageCharWeight: 0 }),
      { autoPage: true, rowHeights: [10, 10] },
    );
    expect(() => normalizeTableDefinition([['A'], ['B']], inheritedTable))
      .toThrow(TypeError);

    const tableSymbol = { autoPage: true, rowHeights: [10, 10] };
    Object.defineProperty(tableSymbol, Symbol('weight'), { value: 0 });
    expect(() => normalizeTableDefinition([['A'], ['B']], tableSymbol))
      .toThrow(TypeError);

    const cellSymbol = { autoPageCharWeight: 0 };
    Object.defineProperty(cellSymbol, Symbol('weight'), { value: 0 });
    expect(() => normalizeTableDefinition([[
      { text: 'A', options: cellSymbol },
    ]], { autoPage: true, rowHeights: [10] })).toThrow(TypeError);
  });
});

describe('table creation internals', () => {
  it('normalizes and renders detached table-cell hyperlinks', () => {
    const url = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const slide = Object.assign(Object.create(null), { slide: 2, tooltip: '' });
    const definition = normalizeTableDefinition([[
      { text: 'URL', options: { hyperlink: url } },
      {
        text: 'Slide',
        options: Object.assign(Object.create(null), { hyperlink: slide }),
      },
      { text: 'Plain', options: { hyperlink: undefined } },
      { text: '', options: { hyperlink: { url: 'https://empty.example' } } },
    ]], {});

    expect(definition.rows).toEqual([[
      {
        text: 'URL',
        hyperlink: {
          url: 'https://example.com?a=1&b=2',
          tooltip: 'Visit & learn',
        },
      },
      { text: 'Slide', hyperlink: { slide: 2, tooltip: '' } },
      { text: 'Plain' },
      { text: '', hyperlink: { url: 'https://empty.example' } },
    ]]);
    expect(Object.isFrozen(definition.rows[0]![0]!.hyperlink)).toBe(true);
    url.url = 'https://changed.example';
    url.tooltip = 'Changed';
    slide.slide = 1;
    expect(definition.rows[0]![0]!.hyperlink).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    });
    expect(definition.rows[0]![1]!.hyperlink).toEqual({ slide: 2, tooltip: '' });

    const rendered = renderTableGraphicFrame(
      7,
      definition,
      undefined,
      undefined,
      [['rId7', 'rId8', undefined, 'rId9']],
    );
    expect(rendered).toContain(
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
    );
    expect(rendered.match(/<a:hlinkClick\b/g)).toHaveLength(3);
    expect(rendered.match(/\bu="sng"/g)).toHaveLength(3);
    expect(rendered).toContain(
      '<a:hlinkClick r:id="rId7" tooltip="Visit &amp; learn"/>',
    );
    expect(rendered).toContain(
      '<a:hlinkClick r:id="rId8" tooltip="" action="ppaction://hlinksldjump"/>',
    );
    expect(rendered).toContain(
      '<a:hlinkClick r:id="rId9"/>',
    );
    expect(rendered.slice(
      rendered.indexOf('<p:nvGraphicFramePr>'),
      rendered.indexOf('</p:nvGraphicFramePr>'),
    )).not.toContain('<a:hlinkClick');

    const unlinked = renderTableGraphicFrame(
      8,
      normalizeTableDefinition([['Same']], undefined),
    );
    const undefinedLink = renderTableGraphicFrame(
      8,
      normalizeTableDefinition([[
        { text: 'Same', options: { hyperlink: undefined } },
      ]], undefined),
    );
    expect(undefinedLink).toBe(unlinked);
    expect(unlinked).not.toContain('xmlns:r=');
  });

  it('normalizes and renders rich and multi-paragraph table-cell text', () => {
    const source = [{
      align: 'right' as const,
      runs: [
        { text: 'Inherited', style: { bold: true } },
        { text: ' shared', style: { italic: false } },
        {
          text: ' explicit',
          style: { hyperlink: { url: 'https://run.example', tooltip: 'Run' } },
        },
        {
          text: ' suppressed',
          softBreakBefore: true,
          style: { hyperlink: false as const, italic: true },
        },
      ],
    }, {
      bullet: { kind: 'bullet' as const, character: '→', indent: 18 },
      runs: [{
        text: 'Slide',
        style: { hyperlink: { slide: 2, tooltip: '' }, underline: false },
      }],
    }];
    const definition = normalizeTableDefinition([[
      {
        text: source,
        options: {
          align: 'center',
          hyperlink: { url: 'https://cell.example' },
        },
      },
      'one\r\ntwo\rempty\n',
    ]], undefined);

    expect(definition.rows[0]![0]!.text).toBe(
      'Inherited shared explicit\n suppressed\nSlide',
    );
    expect(definition.rows[0]![0]!.richText).toEqual(source);
    expect(definition.rows[0]![1]!.text).toBe('one\ntwo\nempty\n');
    expect(definition.rows[0]![1]!.richText).toEqual([
      { runs: [{ text: 'one', style: {} }] },
      { runs: [{ text: 'two', style: {} }] },
      { runs: [{ text: 'empty', style: {} }] },
      { runs: [{ text: '', style: {} }] },
    ]);

    source[0]!.runs[0]!.text = 'MUTATED';
    expect(definition.rows[0]![0]!.richText?.[0]?.runs[0]).toEqual({
      text: 'Inherited',
      style: { bold: true },
    });

    const rendered = renderTableGraphicFrame(
      9,
      definition,
      undefined,
      undefined,
      [['rIdCell', undefined]],
      [[[
        [undefined, undefined, 'rIdRun', undefined],
        ['rIdSlide'],
      ], [
        [undefined],
        [undefined],
        [undefined],
        [undefined],
      ]]],
    );
    expect(rendered.match(/<a:p>/g)).toHaveLength(6);
    expect(rendered.match(/<a:hlinkClick\b/g)).toHaveLength(4);
    expect(rendered.match(/<a:hlinkClick r:id="rIdCell"\/>/g)).toHaveLength(2);
    expect(rendered).toContain(
      '<a:hlinkClick r:id="rIdRun" tooltip="Run"/>',
    );
    expect(rendered).toContain(
      '<a:hlinkClick r:id="rIdSlide" tooltip="" action="ppaction://hlinksldjump"/>',
    );
    expect(rendered).toContain('<a:pPr algn="r"');
    expect(rendered).toContain('<a:br/><a:r><a:rPr lang="en-US" i="1"');
    expect(rendered).toContain('<a:buChar char="→"/>');
    expect(rendered).toContain('xmlns:r=');

    const plainString = renderTableGraphicFrame(
      10,
      normalizeTableDefinition([['Same']], undefined),
    );
    const plainObject = renderTableGraphicFrame(
      10,
      normalizeTableDefinition([[{ text: 'Same' }]], undefined),
    );
    expect(plainObject).toBe(plainString);
  });

  it('expands logical colspan and rowspan cells into a complete physical matrix', () => {
    const rich = [{ runs: [{ text: 'Merged', style: { bold: true } }] }];
    const options = {
      colspan: 2,
      rowspan: 2,
      hyperlink: { url: 'https://merged.example' },
    };
    const definition = normalizeTableDefinition([
      [{ text: rich, options }, 'Right'],
      ['Bottom right'],
    ], {
      fontFamily: 'Aptos',
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '4472C4' },
      },
    });

    expect(definition.rows).toHaveLength(2);
    expect(definition.rows[0]).toHaveLength(3);
    expect(definition.rows[1]).toHaveLength(3);
    expect(definition.rows[0]![0]).toMatchObject({
      text: 'Merged',
      rowspan: 2,
      colspan: 2,
      fontFamily: 'Aptos',
      hyperlink: { url: 'https://merged.example' },
    });
    expect(definition.rows[0]![1]).toEqual({
      text: '',
      continuation: {
        rowSpan: 2,
        horizontal: true,
      },
    });
    expect(definition.rows[1]![0]).toEqual({
      text: '',
      continuation: {
        gridSpan: 2,
        vertical: true,
      },
    });
    expect(definition.rows[1]![1]).toEqual({
      text: '',
      continuation: {
        vertical: true,
        horizontal: true,
      },
    });
    expect(definition.columnWidths).toEqual([914_400, 914_400, 914_400]);
    expect(definition.rowHeights).toEqual([0, 0]);

    const xml = renderTableGraphicFrame(
      14,
      definition,
      undefined,
      undefined,
      [
        ['rId7', undefined, undefined],
        [undefined, undefined, undefined],
      ],
    );
    expect(xml).toContain('<a:tc rowSpan="2" gridSpan="2"><a:txBody>');
    expect(xml).toContain('<a:tc rowSpan="2" hMerge="1"><a:tcPr/></a:tc>');
    expect(xml).toContain('<a:tc gridSpan="2" vMerge="1"><a:tcPr/></a:tc>');
    expect(xml).toContain('<a:tc vMerge="1" hMerge="1"><a:tcPr/></a:tc>');
    expect(xml.match(/<a:hlinkClick\b/g)).toHaveLength(1);
    expect(xml).toContain('<a:hlinkClick r:id="rId7"/>');
    expect(xml.match(/typeface="Aptos"/g)).toHaveLength(18);
    expect(xml.match(/<a:solidFill><a:srgbClr val="4472C4"\/><\/a:solidFill>/g))
      .toHaveLength(3);
    expect(() => renderTableGraphicFrame(
      14,
      definition,
      undefined,
      undefined,
      [
        ['rId7', 'rIdContinuation', undefined],
        [undefined, undefined, undefined],
      ],
    )).toThrow(/continuation cells cannot contain hyperlink relationships/);

    rich[0]!.runs[0]!.text = 'MUTATED';
    options.colspan = 1;
    options.rowspan = 1;
    options.hyperlink.url = 'https://changed.example';
    expect(definition.rows[0]![0]).toMatchObject({
      text: 'Merged',
      rowspan: 2,
      colspan: 2,
      hyperlink: { url: 'https://merged.example' },
    });
  });

  it('accepts fully covered empty logical rows and rejects invalid span layouts', () => {
    const covered = normalizeTableDefinition([
      [
        { text: 'A', options: { rowspan: 2 } },
        { text: 'B', options: { rowspan: 2 } },
      ],
      [],
    ], undefined);
    expect(covered.rows).toEqual([
      [
        { text: 'A', rowspan: 2 },
        { text: 'B', rowspan: 2 },
      ],
      [
        { text: '', continuation: { vertical: true } },
        { text: '', continuation: { vertical: true } },
      ],
    ]);

    const invalidLayouts = [
      [[]],
      [['A', 'B'], []],
      [['A', 'B'], ['C']],
      [['A'], ['B', 'C']],
      [[{ text: 'A', options: { rowspan: 2 } }]],
      [[
        { text: 'A', options: { rowspan: 2 } },
        'B',
        { text: 'C', options: { rowspan: 2 } },
      ], [
        { text: 'D', options: { colspan: 2 } },
      ]],
    ];
    for (const rows of invalidLayouts) {
      expect(() => normalizeTableDefinition(rows, undefined), JSON.stringify(rows)).toThrow();
    }

    const invalidSpans = [
      0,
      -1,
      1.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.MAX_SAFE_INTEGER,
      Number.MAX_SAFE_INTEGER + 1,
      '2',
      true,
    ];
    for (const span of invalidSpans) {
      expect(() => normalizeTableDefinition([[
        { text: 'A', options: { colspan: span } },
      ]], undefined), String(span)).toThrow();
      expect(() => normalizeTableDefinition([[
        { text: 'A', options: { rowspan: span } },
      ]], undefined), String(span)).toThrow();
    }

    expect(normalizeTableDefinition([[
      { text: 'A', options: { colspan: 1, rowspan: 1 } },
    ]], undefined).rows).toEqual([[{ text: 'A' }]]);
  });

  it('normalizes and renders strict table-cell text style defaults', () => {
    const tableColor = { kind: 'scheme' as const, value: 'accent1' };
    const tableSpacing = {
      before: 6,
      after: 8,
      line: { kind: 'multiple' as const, factor: 1.5 },
    };
    const cellSpacing = { before: 3 };
    const cellColor = { kind: 'srgb' as const, value: '00AA00' };
    const definition = normalizeTableDefinition([[
      'Plain',
      { text: 'False', options: { bold: false, spacing: cellSpacing } },
      {
        text: [{
          spacing: { after: 9, line: false },
          runs: [
            { text: 'Inherited' },
            { text: 'Local', style: { fontSize: 10, bold: false } },
          ],
        }, { spacing: false, runs: [] }],
        options: {
          fontFamily: 'Courier New',
          color: cellColor,
        },
      },
    ]], {
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: tableColor,
      spacing: tableSpacing,
    });

    expect(definition.rows[0]!.map((cell) => ({
      fontFamily: cell.fontFamily,
      fontSize: cell.fontSize,
      bold: cell.bold,
      color: cell.color,
      spacing: cell.spacing,
    }))).toEqual([
      {
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'scheme', value: 'accent1' },
        spacing: tableSpacing,
      },
      {
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: false,
        color: { kind: 'scheme', value: 'accent1' },
        spacing: { ...tableSpacing, before: 3 },
      },
      {
        fontFamily: 'Courier New',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'srgb', value: '00AA00' },
        spacing: tableSpacing,
      },
    ]);

    tableColor.value = 'accent2';
    tableSpacing.before = 60;
    tableSpacing.line.factor = 2;
    cellSpacing.before = 30;
    cellColor.value = 'FF0000';
    expect(definition.rows[0]![0]!.color).toEqual({ kind: 'scheme', value: 'accent1' });
    expect(definition.rows[0]![0]!.spacing).toEqual({
      before: 6,
      after: 8,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect(definition.rows[0]![1]!.spacing?.before).toBe(3);
    expect(definition.rows[0]![2]!.color).toEqual({ kind: 'srgb', value: '00AA00' });

    const xml = renderTableGraphicFrame(11, definition);
    const cells = [...xml.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map(([cell]) => cell);
    expect(cells).toHaveLength(3);
    expect(cells[0]).toContain('<a:rPr lang="en-US" sz="1825" b="1" dirty="0">');
    expect(cells[0]).toContain('<a:schemeClr val="accent1"/>');
    expect(cells[0]).toContain(
      '<a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>',
    );
    expect(cells[0]).toContain('<a:spcPct val="150000"/>');
    expect(cells[0]).toContain('<a:spcPts val="600"/>');
    expect(cells[0]).toContain('<a:spcPts val="800"/>');

    expect(cells[1]).toContain('<a:rPr lang="en-US" sz="1825" b="0" dirty="0">');
    expect(cells[1]).toContain('<a:spcPts val="300"/>');
    expect(cells[1]).toContain('<a:spcPts val="800"/>');
    expect(cells[1]).toContain('<a:spcPct val="150000"/>');

    const richRuns = [...cells[2]!.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map(([run]) => run);
    expect(richRuns).toHaveLength(2);
    expect(richRuns[0]).toContain('<a:rPr lang="en-US" sz="1825" b="1" dirty="0">');
    expect(richRuns[0]).toContain('<a:srgbClr val="00AA00"/>');
    expect(richRuns[0]).toContain('typeface="Courier New"');
    expect(richRuns[1]).toContain('<a:rPr lang="en-US" sz="1000" b="0" dirty="0">');
    expect(richRuns[1]).toContain('<a:srgbClr val="00AA00"/>');
    expect(richRuns[1]).toContain('typeface="Courier New"');
    const paragraphs = [...cells[2]!.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
      .map(([paragraph]) => paragraph);
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]).not.toContain('<a:lnSpc>');
    expect(paragraphs[0]).toContain('<a:spcPts val="600"/>');
    expect(paragraphs[0]).toContain('<a:spcPts val="900"/>');
    expect(paragraphs[1]).not.toMatch(/<a:(?:lnSpc|spcBef|spcAft)>/);
    expect(cells[2]!.match(
      /<a:endParaRPr lang="en-US" sz="1825" dirty="0"><a:latin typeface="Courier New"\/><a:ea typeface="Courier New"\/><a:cs typeface="Courier New"\/><\/a:endParaRPr>/g,
    )).toHaveLength(2);
  });

  it('suppresses local hyperlink colors while preserving default hyperlink colors and omitted bytes', () => {
    const definition = normalizeTableDefinition([[
      {
        text: [{
          runs: [
            { text: 'Inherited link', style: { hyperlink: { url: 'https://one.example' } } },
            {
              text: 'Explicit link',
              style: {
                color: { kind: 'srgb', value: 'FF0000' },
                hyperlink: { url: 'https://two.example' },
              },
            },
          ],
        }],
      },
      {
        text: 'Default link',
        options: { hyperlink: { url: 'https://default.example' } },
      },
    ]], {
      fontFamily: 'Aptos',
      fontSize: 18,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
    });
    const xml = renderTableGraphicFrame(
      12,
      definition,
      undefined,
      undefined,
      [[undefined, 'rIdDefault']],
      [[[['rId7', 'rId8']], [[undefined]]]],
    );
    const runs = [...xml.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map(([run]) => run);

    expect(runs).toHaveLength(3);
    expect(runs[0]).toContain('sz="1800" b="1" u="sng"');
    expect(runs[0]).toContain('typeface="Aptos"');
    expect(runs[0]).not.toContain('<a:solidFill>');
    expect(runs[0]).toContain('r:id="rId7"');
    expect(runs[1]).toContain('<a:srgbClr val="FF0000"/>');
    expect(runs[1]).toContain('r:id="rId8"');
    expect(runs[2]).toContain('sz="1800" b="1" u="sng"');
    expect(runs[2]).toContain('typeface="Aptos"');
    expect(runs[2]).toContain('<a:schemeClr val="accent1"/>');
    expect(runs[2]).toContain('r:id="rIdDefault"');

    const omitted = renderTableGraphicFrame(
      13,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      13,
      normalizeTableDefinition([[{
        text: 'Same',
        options: {
          fontFamily: undefined,
          fontSize: undefined,
          bold: undefined,
          color: undefined,
          spacing: undefined,
        },
      }]], {
        fontFamily: undefined,
        fontSize: undefined,
        bold: undefined,
        color: undefined,
        spacing: undefined,
      }),
    );
    expect(runtimeUndefined).toBe(omitted);
  });

  it('strictly rejects malformed table-cell text style defaults', () => {
    const invalidDefaults = [
      { fontFamily: '' },
      { fontFamily: 'bad\u0000font' },
      { fontSize: Number.NaN },
      { fontSize: 0.99 },
      { fontSize: 4000.01 },
      { bold: 1 },
      { color: { kind: 'srgb', value: 'XYZ' } },
      { spacing: {} },
      { spacing: { before: -1 } },
    ];
    for (const options of invalidDefaults) {
      expect(() => normalizeTableDefinition([['Invalid']], options)).toThrow();
      expect(() => normalizeTableDefinition([[
        { text: 'Invalid', options },
      ]], {})).toThrow();
    }

    expect(() => normalizeTableDefinition([['Alias']], { fontFace: 'Aptos' })).toThrow(
      /unsupported property fontFace/,
    );
    expect(() => normalizeTableDefinition([[
      { text: 'Alias', options: { paraSpaceAfter: 6 } },
    ]], {})).toThrow(/unsupported property paraSpaceAfter/);

    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'fontFamily', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 'Aptos';
      },
    });
    class Defaults {
      fontFamily = 'Aptos';
    }
    for (const options of [
      accessor,
      Object.assign(Object.create({ fontFamily: 'Aptos' }), { bold: true }),
      { fontFamily: 'Aptos', [Symbol('unsafe')]: true },
      new Defaults(),
    ]) {
      expect(() => normalizeTableDefinition([['Invalid']], options)).toThrow();
      expect(() => normalizeTableDefinition([[
        { text: 'Invalid', options },
      ]], {})).toThrow();
    }
    expect(accessorCalls).toBe(0);
  });

  it('rejects invalid table-cell hyperlinks and relationship ID matrices', () => {
    const invalid = [
      {},
      { url: 'https://example.com', slide: 2 },
      { url: '' },
      { url: 42 },
      { slide: 0 },
      { slide: 1.5 },
      { slide: Number.MAX_SAFE_INTEGER + 1 },
      { url: 'https://example.com', tooltip: 7 },
      { url: 'https://example.com', _rId: 'rId7' },
    ];
    for (const hyperlink of invalid) {
      expect(() => normalizeTableDefinition([[
        { text: 'Invalid', options: { hyperlink } },
      ]], undefined), JSON.stringify(hyperlink)).toThrow();
    }

    const accessor = {};
    Object.defineProperty(accessor, 'url', { get: () => 'https://example.com' });
    class HyperlinkOptions {
      url = 'https://example.com';
    }
    for (const hyperlink of [accessor, new HyperlinkOptions()]) {
      expect(() => normalizeTableDefinition([[
        { text: 'Invalid', options: { hyperlink } },
      ]], undefined)).toThrow(TypeError);
    }

    const inherited = Object.create({ url: 'https://example.com' });
    expect(() => normalizeTableDefinition([[
      { text: 'Invalid', options: { hyperlink: inherited } },
    ]], undefined)).toThrow(TypeError);

    const definition = normalizeTableDefinition([[
      { text: 'URL', options: { hyperlink: { url: 'https://example.com' } } },
      'Plain',
    ]], undefined);
    for (const ids of [
      [],
      [['rId7']],
      [['rId7', 'rId8']],
      [[undefined, undefined]],
    ]) {
      expect(() => renderTableGraphicFrame(
        7,
        definition,
        undefined,
        undefined,
        ids,
      ), JSON.stringify(ids)).toThrow(TypeError);
    }
  });

  it('normalizes detached table input and renders deterministic exact geometry', () => {
    const sourceRows = [
      ['A & <1>', '', 'C1'],
      ['A2', 'B2', 'C2'],
    ];
    const sourceOptions = {
      name: 'Table "A"',
      x: 457_200,
      y: 685_800,
      width: 2_743_201,
      height: 1_371_601,
    };
    const definition = normalizeTableDefinition(sourceRows, sourceOptions);
    const xml = renderTableGraphicFrame(7, definition);

    expect(definition).toEqual({
      rows: [
        [{ text: 'A & <1>' }, { text: '' }, { text: 'C1' }],
        [{ text: 'A2' }, { text: 'B2' }, { text: 'C2' }],
      ],
      name: 'Table "A"',
      x: 457_200,
      y: 685_800,
      width: 2_743_201,
      height: 1_371_601,
      autoRowHeight: false,
      columnWidths: [914_401, 914_400, 914_400],
      rowHeights: [685_801, 685_800],
    });
    expect(distributeTableDimension(6, 3)).toEqual([2, 2, 2]);
    expect(distributeTableDimension(2_743_201, 3)).toEqual([914_401, 914_400, 914_400]);
    expect(xml).toContain('id="7" name="Table &quot;A&quot;"');
    expect(xml).toContain(
      '<a:gridCol w="914401"/><a:gridCol w="914400"/><a:gridCol w="914400"/>',
    );
    expect(xml).toContain('<a:tr h="685801">');
    expect(xml).toContain('<a:tr h="685800">');
    expect(xml).toContain('<a:t xml:space="preserve">A &amp; &lt;1&gt;</a:t>');
    expect(xml).toMatch(/<a:t xml:space="preserve"><\/a:t>/);
    expect(xml).toContain('marL="91440" marR="91440" marT="45720" marB="45720"');
    expect(xml).toContain(
      '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>',
    );
    expect(xml.match(/<a:tc>/g)).toHaveLength(6);
    expect(xml).toContain(
      'uri="http://schemas.openxmlformats.org/drawingml/2006/table"',
    );
    expect(xml).not.toContain('p14:modId');

    sourceRows[0]![0] = 'MUTATED';
    sourceOptions.width = 1;
    expect(definition.rows[0]![0]!.text).toBe('A & <1>');
    expect(definition.width).toBe(2_743_201);

    const defaults = normalizeTableDefinition([['A', 'B']], undefined);
    const defaultXml = renderTableGraphicFrame(2, defaults);
    expect(defaults).toEqual({
      rows: [[{ text: 'A' }, { text: 'B' }]],
      x: 457_200,
      y: 457_200,
      width: 1_828_800,
      height: 914_400,
      autoRowHeight: true,
      columnWidths: [914_400, 914_400],
      rowHeights: [0],
    });
    expect(defaultXml).toContain('<a:off x="457200" y="457200"/>');
    expect(defaultXml).toContain('<a:ext cx="1828800" cy="914400"/>');
    expect(defaultXml).toContain('<a:gridCol w="914400"/><a:gridCol w="914400"/>');
    expect(defaultXml).toContain('<a:tr h="0">');
  });

  it('normalizes exact scalar and per-column widths without retaining input', () => {
    const unequalSource = [914_400.2, 1_828_799.7, 2_743_200];
    const unequal = normalizeTableDefinition(
      [['A', 'B', 'C']],
      { columnWidths: unequalSource },
    );
    expect(unequal).toMatchObject({
      width: 5_486_400,
      columnWidths: [914_400, 1_828_800, 2_743_200],
    });
    expect(renderTableGraphicFrame(4, unequal)).toContain(
      '<a:gridCol w="914400"/><a:gridCol w="1828800"/><a:gridCol w="2743200"/>',
    );
    unequalSource[0] = 1;
    expect(unequal.columnWidths[0]).toBe(914_400);

    const scalar = normalizeTableDefinition(
      [['A', 'B', 'C']],
      { columnWidths: 1_143_000 },
    );
    expect(scalar.width).toBe(3_429_000);
    expect(scalar.columnWidths).toEqual([1_143_000, 1_143_000, 1_143_000]);

    const matching = normalizeTableDefinition(
      [['A', 'B', 'C']],
      {
        width: 5_486_400,
        columnWidths: [914_400, 1_828_800, 2_743_200],
      },
    );
    expect(matching.width).toBe(5_486_400);
    expect(matching.columnWidths).toEqual([914_400, 1_828_800, 2_743_200]);

    expect(normalizeTableDefinition(
      [['A', 'B']],
      { columnWidths: undefined },
    ).columnWidths).toEqual([914_400, 914_400]);
  });

  it('normalizes exact scalar and per-row heights without retaining input', () => {
    const unequalSource = [457_200.2, 914_399.7, 1_371_600];
    const unequal = normalizeTableDefinition(
      [['A'], ['B'], ['C']],
      { rowHeights: unequalSource },
    );
    expect(unequal).toMatchObject({
      height: 2_743_200,
      autoRowHeight: false,
      rowHeights: [457_200, 914_400, 1_371_600],
    });
    expect(renderTableGraphicFrame(5, unequal)).toMatch(
      /<a:tr h="457200">[\s\S]*<a:tr h="914400">[\s\S]*<a:tr h="1371600">/,
    );
    unequalSource[0] = 1;
    expect(unequal.rowHeights[0]).toBe(457_200);

    const scalar = normalizeTableDefinition(
      [['A'], ['B'], ['C']],
      { rowHeights: 685_800 },
    );
    expect(scalar.height).toBe(2_057_400);
    expect(scalar.rowHeights).toEqual([685_800, 685_800, 685_800]);

    const matching = normalizeTableDefinition(
      [['A'], ['B'], ['C']],
      { height: 2_743_200, rowHeights: [457_200, 914_400, 1_371_600] },
    );
    expect(matching.height).toBe(2_743_200);
    expect(matching.rowHeights).toEqual([457_200, 914_400, 1_371_600]);

    expect(normalizeTableDefinition(
      [['A'], ['B']],
      { rowHeights: undefined },
    ).rowHeights).toEqual([0, 0]);
  });

  it('normalizes detached cell fills while keeping empty options byte-identical', () => {
    const objectCell = { text: 'Plain' };
    const sourceColor = { kind: 'srgb' as const, value: '#ff0000' };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceColor,
      transparency: 33.3334,
    };
    const nullColor = Object.assign(Object.create(null), {
      kind: 'scheme',
      value: 'accent2',
    });
    const nullFill = Object.assign(Object.create(null), {
      kind: 'solid',
      color: nullColor,
      transparency: 25,
    });
    const nullOptions = Object.assign(Object.create(null), { fill: nullFill });
    const nullCell = Object.assign(Object.create(null), {
      text: 'Null prototype',
      options: nullOptions,
    });
    const objectRows = [[
      'String',
      objectCell,
      { text: 'Empty options', options: {} },
      { text: 'Undefined fill', options: { fill: undefined } },
      { text: 'None', options: { fill: { kind: 'none' as const } } },
      { text: 'Solid', options: { fill: sourceFill } },
      nullCell,
      { text: 'Opaque', options: { fill: {
        kind: 'solid' as const,
        color: { kind: 'srgb' as const, value: '00FF00' },
        transparency: 0,
      } } },
      { text: 'Invisible', options: { fill: {
        kind: 'solid' as const,
        color: { kind: 'srgb' as const, value: '445566' },
        transparency: 100,
      } } },
    ]];
    const definition = normalizeTableDefinition(objectRows, undefined);

    expect(definition.rows).toEqual([[
      { text: 'String' },
      { text: 'Plain' },
      { text: 'Empty options' },
      { text: 'Undefined fill' },
      { text: 'None', fill: { kind: 'none' } },
      {
        text: 'Solid',
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
          transparency: 33.333,
        },
      },
      {
        text: 'Null prototype',
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
      },
      {
        text: 'Opaque',
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '00FF00' },
          transparency: 0,
        },
      },
      {
        text: 'Invisible',
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
      },
    ]]);

    const equivalentRows = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { fill: undefined } }]],
    ];
    const equivalentXml = equivalentRows.map((rows) =>
      renderTableGraphicFrame(8, normalizeTableDefinition(rows, undefined)));
    expect(new Set(equivalentXml).size).toBe(1);

    const xml = renderTableGraphicFrame(8, definition);
    expect(xml).toContain(
      '</a:lnB><a:noFill/></a:tcPr>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="66667"/></a:srgbClr></a:solidFill>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr></a:solidFill>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:solidFill>',
    );

    objectCell.text = 'MUTATED';
    sourceColor.value = '000000';
    sourceFill.transparency = 1;
    nullColor.value = 'accent6';
    nullFill.transparency = 50;
    expect(definition.rows[0]![5]).toEqual({
      text: 'Solid',
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 33.333,
      },
    });
    expect(definition.rows[0]![6]).toEqual({
      text: 'Null prototype',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      },
    });
  });

  it('materializes detached strict table fills under cell fills', () => {
    const sourceColor = { kind: 'scheme', value: 'accent1' };
    const sourceFill = {
      kind: 'solid',
      color: sourceColor,
      transparency: 33.3334,
    };
    const definition = normalizeTableDefinition([[
      'String',
      { text: 'Object' },
      { text: 'Empty options', options: {} },
      { text: 'Undefined', options: { fill: undefined } },
      { text: 'Cell none', options: { fill: { kind: 'none' } } },
      { text: 'Cell solid', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 25,
      } } },
    ]], {
      fill: sourceFill,
      margin: { top: 9 },
      valign: 'middle',
    });
    const tableFill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 33.333,
    };

    expect(definition.rows[0]!.map(({ fill }) => fill)).toEqual([
      tableFill,
      tableFill,
      tableFill,
      tableFill,
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 25,
      },
    ]);
    const xml = renderTableGraphicFrame(9, definition);
    expect(xml.match(/<a:schemeClr val="accent1">/g)).toHaveLength(4);
    expect(xml).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="114300" marB="45720" anchor="ctr">'
      + '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>'
      + '<a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>'
      + '<a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>'
      + '<a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>'
      + '<a:solidFill><a:schemeClr val="accent1">'
      + '<a:alpha val="66667"/></a:schemeClr></a:solidFill></a:tcPr>',
    );
    expect(xml).toContain('</a:lnB><a:noFill/></a:tcPr>');

    sourceColor.value = 'accent6';
    sourceFill.transparency = 1;
    expect(definition.rows[0]![0]!.fill).toEqual(tableFill);

    const tableNone = normalizeTableDefinition(
      [[
        'Inherited none',
        { text: 'Solid override', options: { fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '00FF00' },
          transparency: 0,
        } } },
      ]],
      { fill: { kind: 'none' } },
    );
    expect(tableNone.rows[0]!.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    ]);
    const noneXml = renderTableGraphicFrame(10, tableNone);
    expect(noneXml).toContain('</a:lnB><a:noFill/></a:tcPr>');
    expect(noneXml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00">'
      + '<a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );

    const explicitZero = normalizeTableDefinition([['Opaque']], { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    } });
    expect(explicitZero.rows[0]![0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    });
    expect(renderTableGraphicFrame(11, explicitZero)).toContain(
      '<a:solidFill><a:srgbClr val="00FF00">'
      + '<a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );

    const transparent = normalizeTableDefinition([['Transparent']], { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: '445566' },
      transparency: 100,
    } });
    expect(transparent.rows[0]![0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '445566' },
      transparency: 100,
    });
    expect(renderTableGraphicFrame(12, transparent)).toContain(
      '<a:solidFill><a:srgbClr val="445566">'
      + '<a:alpha val="0"/></a:srgbClr></a:solidFill>',
    );

    const omitted = renderTableGraphicFrame(
      13,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      13,
      normalizeTableDefinition([['Same']], { fill: undefined }),
    );
    expect(runtimeUndefined).toBe(omitted);
  });

  it('materializes detached strict table borders under whole cell border values', () => {
    const sourceColor = { kind: 'scheme', value: 'accent1' };
    const sourceBorder = {
      kind: 'line',
      color: sourceColor,
      width: 1.500004,
      style: 'dash',
    };
    const definition = normalizeTableDefinition([[
      'String',
      { text: 'Object' },
      { text: 'Empty options', options: {} },
      { text: 'Undefined', options: { border: undefined } },
      { text: 'Empty border', options: { border: {} } },
      { text: 'All undefined', options: {
        border: [undefined, undefined, undefined, undefined],
      } },
      { text: 'Cell partial', options: {
        border: { left: { kind: 'none' } },
      } },
      { text: 'Cell none', options: { border: { kind: 'none' } } },
      { text: 'Cell tuple', options: { border: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 0,
        },
        undefined,
        { kind: 'none' },
        undefined,
      ] } },
    ]], {
      border: sourceBorder,
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
      margin: { top: 9 },
      valign: 'middle',
    });
    const tableBorders = {
      top: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      right: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      left: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
    };
    expect(definition.rows[0]!.slice(0, 6).map(({ borders }) => borders)).toEqual(
      Array(6).fill(tableBorders),
    );
    expect(definition.rows[0]![6]!.borders).toEqual({ left: { kind: 'none' } });
    expect(definition.rows[0]![7]!.borders).toEqual({
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(definition.rows[0]![8]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 0,
      },
      bottom: { kind: 'none' },
    });

    const xml = renderTableGraphicFrame(41, definition);
    expect(xml.match(/<a:schemeClr val="accent1"\/>/g)).toHaveLength(24);
    expect(xml).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="114300" marB="45720" anchor="ctr">'
      + '<a:lnL w="19050" cap="flat" cmpd="sng" algn="ctr">'
      + '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
      + '<a:prstDash val="sysDash"/><a:round/>'
      + '<a:headEnd type="none" w="med" len="med"/>'
      + '<a:tailEnd type="none" w="med" len="med"/></a:lnL>'
      + '<a:lnR w="19050" cap="flat" cmpd="sng" algn="ctr">',
    );
    expect(xml).toMatch(
      /<a:lnL[\s\S]*<\/a:lnL><a:lnR[\s\S]*<\/a:lnR><a:lnT[\s\S]*<\/a:lnT><a:lnB[\s\S]*<\/a:lnB><a:solidFill>/,
    );

    sourceColor.value = 'accent6';
    sourceBorder.width = 9;
    expect(definition.rows[0]![0]!.borders).toEqual(tableBorders);

    const tableNone = normalizeTableDefinition([[
      'Inherited none',
      { text: 'Line override', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 2,
        style: 'solid',
      } } },
    ]], { border: { kind: 'none' } });
    expect(tableNone.rows[0]!.map(({ borders }) => borders)).toEqual([
      {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      },
      {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 2,
          style: 'solid',
        },
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 2,
          style: 'solid',
        },
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 2,
          style: 'solid',
        },
        left: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 2,
          style: 'solid',
        },
      },
    ]);

    const tuple = normalizeTableDefinition([['Tuple']], { border: [
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 0 },
      undefined,
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 0.5,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 1584,
        style: 'solid',
      },
    ] });
    expect(tuple.rows[0]![0]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 0,
      },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 0.5,
      },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 1584,
        style: 'solid',
      },
    });
    expect(renderTableGraphicFrame(42, tuple)).toContain(
      '<a:lnT w="0" cap="flat" cmpd="sng" algn="ctr">'
      + '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:round/>',
    );

    const named = normalizeTableDefinition([['Named']], { border: {
      right: { kind: 'none' },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 2,
        style: 'dash',
      },
    } });
    expect(named.rows[0]![0]!.borders).toEqual({
      right: { kind: 'none' },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 2,
        style: 'dash',
      },
    });

    const omitted = renderTableGraphicFrame(
      43,
      normalizeTableDefinition([['Same']], {}),
    );
    expect(renderTableGraphicFrame(
      43,
      normalizeTableDefinition([['Same']], { border: undefined }),
    )).toBe(omitted);
    expect(renderTableGraphicFrame(
      43,
      normalizeTableDefinition([['Same']], { border: {} }),
    )).toBe(omitted);
    expect(renderTableGraphicFrame(
      43,
      normalizeTableDefinition([['Same']], {
        border: [undefined, undefined, undefined, undefined],
      }),
    )).toBe(omitted);
  });

  it('normalizes detached cell borders and renders canonical LRTB XML before fill', () => {
    const sourceColor = { kind: 'srgb' as const, value: '#ff0000' };
    const sourceLine = {
      kind: 'line' as const,
      color: sourceColor,
      width: 1.500004,
      style: 'solid' as const,
    };
    const rows = [[
      'String',
      { text: 'Empty options', options: {} },
      { text: 'Undefined border', options: { border: undefined } },
      { text: 'Empty border', options: { border: {} } },
      { text: 'None', options: { border: { kind: 'none' as const } } },
      { text: 'Scalar', options: { border: sourceLine } },
      { text: 'Tuple', options: { border: [
        {
          kind: 'line' as const,
          color: { kind: 'scheme' as const, value: 'accent1' },
          width: 2,
          style: 'dash' as const,
        },
        undefined,
        {
          kind: 'line' as const,
          color: { kind: 'srgb' as const, value: '00FF00' },
          width: 0,
        },
        { kind: 'none' as const },
      ] } },
      { text: 'Named + fill', options: {
        border: { top: sourceLine, left: { kind: 'none' as const } },
        fill: {
          kind: 'solid' as const,
          color: { kind: 'scheme' as const, value: 'accent2' },
          transparency: 25,
        },
      } },
      { text: 'Undefined tuple', options: {
        border: [undefined, undefined, undefined, undefined],
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);

    expect(definition.rows).toEqual([[
      { text: 'String' },
      { text: 'Empty options' },
      { text: 'Undefined border' },
      { text: 'Empty border' },
      {
        text: 'None',
        borders: {
          top: { kind: 'none' },
          right: { kind: 'none' },
          bottom: { kind: 'none' },
          left: { kind: 'none' },
        },
      },
      {
        text: 'Scalar',
        borders: {
          top: {
            kind: 'line',
            color: { kind: 'srgb', value: 'FF0000' },
            width: 1.5,
            style: 'solid',
          },
          right: {
            kind: 'line',
            color: { kind: 'srgb', value: 'FF0000' },
            width: 1.5,
            style: 'solid',
          },
          bottom: {
            kind: 'line',
            color: { kind: 'srgb', value: 'FF0000' },
            width: 1.5,
            style: 'solid',
          },
          left: {
            kind: 'line',
            color: { kind: 'srgb', value: 'FF0000' },
            width: 1.5,
            style: 'solid',
          },
        },
      },
      {
        text: 'Tuple',
        borders: {
          top: {
            kind: 'line',
            color: { kind: 'scheme', value: 'accent1' },
            width: 2,
            style: 'dash',
          },
          bottom: {
            kind: 'line',
            color: { kind: 'srgb', value: '00FF00' },
            width: 0,
          },
          left: { kind: 'none' },
        },
      },
      {
        text: 'Named + fill',
        borders: {
          top: {
            kind: 'line',
            color: { kind: 'srgb', value: 'FF0000' },
            width: 1.5,
            style: 'solid',
          },
          left: { kind: 'none' },
        },
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
      },
      { text: 'Undefined tuple' },
    ]]);

    const equivalentRows = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { border: undefined } }]],
      [[{ text: 'Same', options: { border: {} } }]],
      [[{ text: 'Same', options: {
        border: [undefined, undefined, undefined, undefined],
      } }]],
      [[{ text: 'Same', options: { border: { kind: 'none' } } }]],
    ];
    const equivalentXml = equivalentRows.map((inputRows) =>
      renderTableGraphicFrame(9, normalizeTableDefinition(inputRows, undefined)));
    expect(new Set(equivalentXml).size).toBe(1);

    const none = (tag: string): string =>
      `<a:${tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${tag}>`;
    const line = (
      tag: string,
      width: number,
      color: string,
      dash?: 'solid' | 'sysDash',
    ): string => {
      const dashXml = dash === undefined ? '' : `<a:prstDash val="${dash}"/>`;
      return `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${color}</a:solidFill>${dashXml}<a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
    };
    const srgb = (value: string): string => `<a:srgbClr val="${value}"/>`;
    const scheme = (value: string): string => `<a:schemeClr val="${value}"/>`;
    const tupleXml = renderTableGraphicFrame(
      10,
      normalizeTableDefinition([[rows[0]![6]]], undefined),
    );
    expect(tupleXml).toContain([
      none('lnL'),
      none('lnR'),
      line('lnT', 25_400, scheme('accent1'), 'sysDash'),
      line('lnB', 0, srgb('00FF00')),
    ].join(''));
    const combinedXml = renderTableGraphicFrame(
      11,
      normalizeTableDefinition([[rows[0]![7]]], undefined),
    );
    expect(combinedXml).toContain([
      none('lnL'),
      none('lnR'),
      line('lnT', 19_050, srgb('FF0000'), 'solid'),
      none('lnB'),
      '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr></a:solidFill>',
    ].join(''));

    sourceColor.value = '000000';
    sourceLine.width = 4;
    expect(definition.rows[0]![5]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1.5,
        style: 'solid',
      },
      right: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1.5,
        style: 'solid',
      },
      bottom: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1.5,
        style: 'solid',
      },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1.5,
        style: 'solid',
      },
    });
  });

  it('normalizes detached cell margins and renders canonical defaults before border and fill', () => {
    const sourceMargins = { top: 1.500004, left: -2 };
    const nullMargins = Object.assign(Object.create(null), {
      right: 5,
      bottom: 6,
    });
    const rows = [[
      'String',
      { text: 'Empty options', options: {} },
      { text: 'Undefined margin', options: { margin: undefined } },
      { text: 'Empty margin', options: { margin: {} } },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'Tuple', options: { margin: [1, 2, 3, 4] } },
      { text: 'Named', options: { margin: sourceMargins } },
      { text: 'Null prototype', options: { margin: nullMargins } },
      { text: 'Combined', options: {
        margin: { top: 4, left: 8 },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'C00000' },
          width: 2,
        },
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
        },
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);

    expect(definition.rows[0]!.map(({ text, margins }) => ({ text, margins }))).toEqual([
      { text: 'String', margins: undefined },
      { text: 'Empty options', margins: undefined },
      { text: 'Undefined margin', margins: undefined },
      { text: 'Empty margin', margins: {} },
      { text: 'Zero', margins: { top: 0, right: 0, bottom: 0, left: 0 } },
      { text: 'Tuple', margins: { top: 1, right: 2, bottom: 3, left: 4 } },
      { text: 'Named', margins: { top: 1.5, left: -2 } },
      { text: 'Null prototype', margins: { right: 5, bottom: 6 } },
      { text: 'Combined', margins: { top: 4, left: 8 } },
    ]);
    expect(definition.rows[0]![8]).toMatchObject({
      borders: {
        top: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
        right: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
        bottom: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
        left: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
      },
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
    });

    const equivalentRows = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { margin: undefined } }]],
      [[{ text: 'Same', options: { margin: {} } }]],
      [[{ text: 'Same', options: {
        margin: { top: undefined, right: undefined },
      } }]],
    ];
    const equivalentXml = equivalentRows.map((inputRows) =>
      renderTableGraphicFrame(12, normalizeTableDefinition(inputRows, undefined)));
    expect(new Set(equivalentXml).size).toBe(1);
    expect(equivalentXml[0]).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720">',
    );

    const renderCell = (cell: unknown): string => renderTableGraphicFrame(
      13,
      normalizeTableDefinition([[cell]], undefined),
    );
    expect(renderCell(rows[0]![4])).toContain(
      '<a:tcPr marL="0" marR="0" marT="0" marB="0">',
    );
    expect(renderCell(rows[0]![5])).toContain(
      '<a:tcPr marL="50800" marR="25400" marT="12700" marB="38100">',
    );
    expect(renderCell(rows[0]![6])).toContain(
      '<a:tcPr marL="-25400" marR="91440" marT="19050" marB="45720">',
    );
    expect(renderCell(rows[0]![8])).toMatch(
      /<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
    );

    sourceMargins.top = 99;
    sourceMargins.left = 99;
    nullMargins.right = 99;
    nullMargins.bottom = 99;
    expect(definition.rows[0]![6]!.margins).toEqual({ top: 1.5, left: -2 });
    expect(definition.rows[0]![7]!.margins).toEqual({ right: 5, bottom: 6 });
  });

  it('normalizes and renders strict cell vertical alignment after margins', () => {
    const nullOptions = Object.assign(Object.create(null), { valign: 'bottom' });
    const rows = [[
      'String',
      { text: 'Empty options', options: {} },
      { text: 'Undefined', options: { valign: undefined } },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Middle', options: { valign: 'middle' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
      { text: 'Null prototype', options: nullOptions },
      { text: 'Combined', options: {
        valign: 'middle',
        margin: { top: 4, left: 8 },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'C00000' },
          width: 2,
        },
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
        },
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);

    expect(definition.rows[0]!.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      undefined,
      undefined,
      undefined,
      'top',
      'middle',
      'bottom',
      'bottom',
      'middle',
    ]);

    const equivalent = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { valign: undefined } }]],
    ].map((input) => renderTableGraphicFrame(
      20,
      normalizeTableDefinition(input, undefined),
    ));
    expect(new Set(equivalent).size).toBe(1);
    expect(equivalent[0]).not.toContain(' anchor=');

    const renderCell = (cell: unknown): string => renderTableGraphicFrame(
      21,
      normalizeTableDefinition([[cell]], undefined),
    );
    expect(renderCell(rows[0]![3])).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="t">',
    );
    expect(renderCell(rows[0]![4])).toContain(' marB="45720" anchor="ctr">');
    expect(renderCell(rows[0]![5])).toContain(' marB="45720" anchor="b">');
    expect(renderCell(rows[0]![7])).toMatch(
      /<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720" anchor="ctr"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
    );
  });

  it('normalizes and renders strict table cell text direction', () => {
    const nullOptions = Object.assign(Object.create(null), {
      textDirection: 'wordArtVert',
    });
    const rows = [[
      'String',
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { textDirection: undefined } },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
      { text: 'Null prototype', options: nullOptions },
      { text: 'Combined', options: {
        align: 'center',
        textDirection: 'vert270',
        valign: 'middle',
        margin: { top: 4, left: 8 },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'C00000' },
          width: 2,
        },
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
        },
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);

    expect(definition.rows[0]!.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
      undefined,
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVert',
      'vert270',
    ]);

    const equivalent = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { textDirection: undefined } }]],
      [[{ text: 'Same', options: { textDirection: 'horz' } }]],
    ].map((input) => renderTableGraphicFrame(
      50,
      normalizeTableDefinition(input, undefined),
    ));
    expect(new Set(equivalent).size).toBe(1);
    expect(equivalent[0]).not.toMatch(/<a:tcPr[^>]*\svert=/);

    const xml = renderTableGraphicFrame(51, definition);
    expect([...xml.matchAll(/<a:tcPr([^>]*)>/g)].map((match) =>
      match[1]!.match(/\svert="([^"]+)"/)?.[1])).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVert',
      'vert270',
    ]);
    expect(xml).toMatch(
      /marB="45720" anchor="ctr" vert="vert270"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
    );
    expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);
  });

  it('normalizes and renders strict table cell text fit', () => {
    const nullOptions = Object.assign(Object.create(null), { fit: 'resize' });
    const rows = [[
      'String',
      { text: 'Object' },
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { fit: undefined } },
      { text: 'None', options: { fit: 'none' } },
      { text: 'Shrink', options: { fit: 'shrink' } },
      { text: 'Resize', options: { fit: 'resize' } },
      { text: 'Null prototype', options: nullOptions },
      { text: 'Combined', options: {
        align: 'center',
        fit: 'shrink',
        margin: { top: 4, left: 8 },
        textDirection: 'vert270',
        valign: 'middle',
        border: { kind: 'none' },
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);

    expect(definition.rows[0]!.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      'none',
      'shrink',
      'resize',
      'resize',
      'shrink',
    ]);

    const equivalent = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { fit: undefined } }]],
      [[{ text: 'Same', options: { fit: 'none' } }]],
    ].map((input) => renderTableGraphicFrame(
      70,
      normalizeTableDefinition(input, undefined),
    ));
    expect(new Set(equivalent).size).toBe(1);
    expect(equivalent[0]).toContain('<a:bodyPr/><a:lstStyle/>');

    const xml = renderTableGraphicFrame(71, definition);
    const cells = [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
      .map((match) => match[0]);
    expect(cells.map((cell) =>
      cell.match(/<a:(normAutofit|spAutoFit)\/>/)?.[1])).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      'normAutofit',
      'spAutoFit',
      'spAutoFit',
      'normAutofit',
    ]);
    expect(cells[5]).toContain(
      '<a:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>',
    );
    expect(cells[6]).toContain(
      '<a:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/>',
    );
    expect(cells[8]).toMatch(
      /<a:bodyPr><a:normAutofit\/><\/a:bodyPr><a:lstStyle\/>[\s\S]*<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720" anchor="ctr" vert="vert270">/,
    );
    expect(xml).not.toMatch(/<a:tcPr[^>]*\sfit=/);
    expect(xml).not.toContain('<a:noAutofit/>');

    nullOptions.fit = 'none';
    expect(definition.rows[0]![7]!.textFit).toBe('resize');
  });

  it('materializes strict table text direction onto uncovered cells', () => {
    const rows = [[
      'String',
      { text: 'Object' },
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { textDirection: undefined } },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    ]];
    const definition = normalizeTableDefinition(rows, {
      textDirection: 'vert270',
    });

    expect(definition.rows[0]!.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
    ]);

    const xml = renderTableGraphicFrame(60, definition);
    expect([...xml.matchAll(/<a:tcPr([^>]*)>/g)].map((match) =>
      match[1]!.match(/\svert="([^"]+)"/)?.[1])).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
    ]);
    expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);

    const omitted = renderTableGraphicFrame(
      61,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      61,
      normalizeTableDefinition([['Same']], { textDirection: undefined }),
    );
    const horizontal = renderTableGraphicFrame(
      61,
      normalizeTableDefinition([['Same']], { textDirection: 'horz' }),
    );
    expect(runtimeUndefined).toBe(omitted);
    expect(horizontal).toBe(omitted);
    expect(omitted).not.toMatch(/<a:tcPr[^>]*\svert=/);

    for (const textDirection of [
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
    ] as const) {
      const nullPrototypeOptions = Object.assign(Object.create(null), {
        textDirection,
      });
      const normalized = normalizeTableDefinition(
        [['Detached']],
        nullPrototypeOptions,
      );
      nullPrototypeOptions.textDirection = 'horz';
      expect(normalized.rows[0]![0]!.textDirection).toBe(textDirection);
      const rendered = renderTableGraphicFrame(62, normalized);
      const direct = rendered.match(/<a:tcPr[^>]*\svert="([^"]+)"/)?.[1];
      expect(direct).toBe(textDirection === 'horz' ? undefined : textDirection);
    }
  });

  it('normalizes and renders strict table cell horizontal alignment', () => {
    const nullOptions = Object.assign(Object.create(null), { align: 'right' });
    const rows = [[
      'String',
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { align: undefined } },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Center', options: { align: 'center' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify this sentence', options: { align: 'justify' } },
      { text: 'Null prototype', options: nullOptions },
      { text: 'Combined', options: {
        align: 'center',
        valign: 'bottom',
        margin: { top: 4, left: 8 },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'C00000' },
          width: 2,
        },
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
        },
      } },
    ]];
    const definition = normalizeTableDefinition(rows, undefined);
    expect(definition.rows[0]!.map(({ alignment }) => alignment)).toEqual([
      undefined,
      undefined,
      undefined,
      'left',
      'center',
      'right',
      'justify',
      'right',
      'center',
    ]);

    const equivalent = [
      [['Same']],
      [[{ text: 'Same' }]],
      [[{ text: 'Same', options: {} }]],
      [[{ text: 'Same', options: { align: undefined } }]],
    ].map((input) => renderTableGraphicFrame(
      20,
      normalizeTableDefinition(input, undefined),
    ));
    expect(new Set(equivalent).size).toBe(1);
    expect(equivalent[0]).not.toMatch(/<a:pPr[^>]*\salgn=/);

    const renderCell = (cell: unknown): string => renderTableGraphicFrame(
      21,
      normalizeTableDefinition([[cell]], undefined),
    );
    expect(renderCell(rows[0]![3])).toContain('<a:pPr algn="l"');
    expect(renderCell(rows[0]![4])).toContain('<a:pPr algn="ctr"');
    expect(renderCell(rows[0]![5])).toContain('<a:pPr algn="r"');
    expect(renderCell(rows[0]![6])).toContain('<a:pPr algn="just"');
    const combined = renderCell(rows[0]![8]);
    expect(combined).toContain(
      '<a:p><a:pPr algn="ctr" indent="0" marL="0"><a:buNone/></a:pPr><a:r>',
    );
    expect(combined).toContain('marB="45720" anchor="b">');
    expect(combined).not.toMatch(/<a:tcPr[^>]*\salgn=/);
    expect(combined).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
  });

  it('materializes strict table horizontal alignment onto uncovered cells', () => {
    const rows = [[
      'String',
      { text: 'Object' },
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { align: undefined } },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify', options: { align: 'justify' } },
    ]];
    const definition = normalizeTableDefinition(rows, { align: 'center' });
    expect(definition.rows[0]!.map(({ alignment }) => alignment)).toEqual([
      'center',
      'center',
      'center',
      'center',
      'left',
      'right',
      'justify',
    ]);

    const xml = renderTableGraphicFrame(33, definition);
    expect([...xml.matchAll(/<a:pPr[^>]*\salgn="([^"]+)"/g)]
      .map((match) => match[1])).toEqual([
      'ctr',
      'ctr',
      'ctr',
      'ctr',
      'l',
      'r',
      'just',
    ]);
    expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
    expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);

    const tableValues = [
      ['left', 'l'],
      ['center', 'ctr'],
      ['right', 'r'],
      ['justify', 'just'],
    ] as const;
    for (const [align, token] of tableValues) {
      const value = normalizeTableDefinition([['A']], { align });
      expect(value.rows[0]![0]!.alignment).toBe(align);
      expect(renderTableGraphicFrame(34, value)).toContain(`algn="${token}"`);
    }

    const omitted = renderTableGraphicFrame(
      35,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      35,
      normalizeTableDefinition([['Same']], { align: undefined }),
    );
    expect(runtimeUndefined).toBe(omitted);
    expect(omitted).not.toMatch(/<a:pPr[^>]*\salgn=/);
  });

  it('materializes strict table vertical alignment onto uncovered cells', () => {
    const rows = [[
      'String',
      { text: 'Object' },
      { text: 'Empty', options: {} },
      { text: 'Undefined', options: { valign: undefined } },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ]];
    const definition = normalizeTableDefinition(rows, { valign: 'middle' });
    expect(definition.rows[0]!.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'middle',
      'middle',
      'middle',
      'top',
      'bottom',
    ]);

    const xml = renderTableGraphicFrame(31, definition);
    const anchors = [...xml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)]
      .map((match) => match[1]);
    expect(anchors).toEqual(['ctr', 'ctr', 'ctr', 'ctr', 't', 'b']);
    expect(xml).not.toMatch(/<a:bodyPr[^>]* anchor=/);

    const tableValues = ['top', 'middle', 'bottom'] as const;
    expect(tableValues.map((valign) => normalizeTableDefinition(
      [['Inherited']],
      { valign },
    ).rows[0]![0]!.verticalAlignment)).toEqual(tableValues);
    expect(tableValues.map((valign) => renderTableGraphicFrame(
      33,
      normalizeTableDefinition([['Inherited']], { valign }),
    ).match(/<a:tcPr[^>]* anchor="([^"]+)"/)?.[1])).toEqual(['t', 'ctr', 'b']);

    const omitted = renderTableGraphicFrame(
      32,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      32,
      normalizeTableDefinition([['Same']], { valign: undefined }),
    );
    expect(runtimeUndefined).toBe(omitted);
    expect(omitted).not.toContain(' anchor=');
  });

  it('materializes strict table margins under cell margin sides', () => {
    const tableMargin = { top: 9, left: 18 };
    const definition = normalizeTableDefinition([[
      'String',
      { text: 'Object' },
      { text: 'Empty options', options: {} },
      { text: 'Undefined', options: { margin: undefined } },
      { text: 'Empty margin', options: { margin: {} } },
      { text: 'Undefined sides', options: { margin: {
        top: undefined,
        right: undefined,
        bottom: undefined,
        left: undefined,
      } } },
      { text: 'Partial', options: { margin: { bottom: 12 } } },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'Tuple', options: { margin: [1, 2, 3, 4] } },
    ]], { margin: tableMargin, valign: 'middle' });
    tableMargin.top = 99;
    tableMargin.left = 99;

    expect(definition.rows[0]!.map(({ margins }) => margins)).toEqual([
      { top: 9, left: 18 },
      { top: 9, left: 18 },
      { top: 9, left: 18 },
      { top: 9, left: 18 },
      { top: 9, left: 18 },
      { top: 9, left: 18 },
      { top: 9, bottom: 12, left: 18 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ]);

    const xml = renderTableGraphicFrame(41, definition);
    const margins = [...xml.matchAll(
      /<a:tcPr marL="(-?\d+)" marR="(-?\d+)" marT="(-?\d+)" marB="(-?\d+)"/g,
    )].map((match) => match.slice(1).map(Number));
    expect(margins).toEqual([
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 45720],
      [228600, 91440, 114300, 152400],
      [0, 0, 0, 0],
      [50800, 25400, 12700, 38100],
    ]);
    expect(xml).toMatch(/marB="45720" anchor="ctr"><a:lnL/);
    expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns)=/);

    expect(normalizeTableDefinition([['Scalar']], { margin: 0 })
      .rows[0]![0]!.margins).toEqual({
      top: 0,
      right: 0,
      bottom: 0,
      left: 0,
    });
    const tableTuple = [1, 2, 3, 4] as [number, number, number, number];
    const tupleDefinition = normalizeTableDefinition([['Tuple']], { margin: tableTuple });
    tableTuple.fill(99);
    expect(tupleDefinition.rows[0]![0]!.margins).toEqual({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    });

    const omitted = renderTableGraphicFrame(
      42,
      normalizeTableDefinition([['Same']], {}),
    );
    const runtimeUndefined = renderTableGraphicFrame(
      42,
      normalizeTableDefinition([['Same']], { margin: undefined }),
    );
    const empty = renderTableGraphicFrame(
      42,
      normalizeTableDefinition([['Same']], { margin: {} }),
    );
    expect(runtimeUndefined).toBe(omitted);
    expect(empty).toBe(omitted);
  });

  it('strictly rejects malformed matrices and options without invoking accessors', () => {
    const sparseOuter = Array(1);
    const sparseRow = [Array(2)];
    sparseRow[0]![0] = 'A';
    const extraOuter = Object.assign([['A']], { extra: true });
    const extraRow = [Object.assign(['A'], { extra: true })];
    const symbolOuter = Object.assign([['A']], { [Symbol('extra')]: true });
    const accessorRow = [['A']];
    let rowAccessorCalls = 0;
    Object.defineProperty(accessorRow[0]!, '0', {
      get() {
        rowAccessorCalls += 1;
        return 'A';
      },
      enumerable: true,
      configurable: true,
    });
    const invalidRows = [
      null,
      false,
      '',
      [],
      ['A'],
      [[], []],
      [['A'], ['B', 'C']],
      sparseOuter,
      sparseRow,
      extraOuter,
      extraRow,
      symbolOuter,
      accessorRow,
      [[1]],
      [[null]],
      [['bad\u0000xml']],
      Symbol('rows'),
    ];
    for (const rows of invalidRows) {
      expect(() => normalizeTableDefinition(rows, undefined)).toThrow();
    }
    expect(rowAccessorCalls).toBe(0);

    const inheritedCell = Object.create({ text: 'Inherited' });
    const accessorCell = {};
    let cellAccessorCalls = 0;
    Object.defineProperty(accessorCell, 'text', {
      get() {
        cellAccessorCalls += 1;
        return 'Accessor';
      },
      enumerable: true,
      configurable: true,
    });
    const optionsAccessorCell = { text: 'Accessor options' };
    Object.defineProperty(optionsAccessorCell, 'options', {
      get() {
        cellAccessorCalls += 1;
        return {};
      },
      enumerable: true,
      configurable: true,
    });
    const accessorCellOptions = {};
    Object.defineProperty(accessorCellOptions, 'fill', {
      get() {
        cellAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    const accessorBorderOptions = {};
    Object.defineProperty(accessorBorderOptions, 'border', {
      get() {
        cellAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    const accessorMarginOptions = {};
    Object.defineProperty(accessorMarginOptions, 'margin', {
      get() {
        cellAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    const accessorValignOptions = {};
    Object.defineProperty(accessorValignOptions, 'valign', {
      get() {
        cellAccessorCalls += 1;
        return 'top';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorAlignOptions = {};
    Object.defineProperty(accessorAlignOptions, 'align', {
      get() {
        cellAccessorCalls += 1;
        return 'center';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTextDirectionOptions = {};
    Object.defineProperty(accessorTextDirectionOptions, 'textDirection', {
      get() {
        cellAccessorCalls += 1;
        return 'vert';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorMargin = { left: 1 };
    Object.defineProperty(accessorMargin, 'top', {
      get() {
        cellAccessorCalls += 1;
        return 2;
      },
      enumerable: true,
      configurable: true,
    });
    const accessorMarginTuple = [1, 2, 3, 4];
    Object.defineProperty(accessorMarginTuple, '2', {
      get() {
        cellAccessorCalls += 1;
        return 3;
      },
      enumerable: true,
      configurable: true,
    });
    const accessorBorder: Record<string, unknown> = {};
    Object.defineProperty(accessorBorder, 'kind', {
      get() {
        cellAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorNamedBorder: Record<string, unknown> = {};
    Object.defineProperty(accessorNamedBorder, 'top', {
      get() {
        cellAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    const accessorBorderTuple: unknown[] = [undefined, undefined, undefined, undefined];
    Object.defineProperty(accessorBorderTuple, '0', {
      get() {
        cellAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    const accessorBorderColor = { kind: 'srgb' };
    Object.defineProperty(accessorBorderColor, 'value', {
      get() {
        cellAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorFill = {};
    Object.defineProperty(accessorFill, 'kind', {
      get() {
        cellAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorColor = { kind: 'srgb' };
    Object.defineProperty(accessorColor, 'value', {
      get() {
        cellAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
      configurable: true,
    });
    class ExoticCell {
      text = 'Class';
    }
    class ExoticCellOptions {
      fill = undefined;
    }
    class ExoticFill {
      kind = 'none';
    }
    class ExoticBorder {
      kind = 'none';
    }
    class ExoticBorderColor {
      kind = 'srgb';
      value = 'FF0000';
    }
    class ExoticBorderTuple extends Array<unknown> {}
    const exoticBorderTuple = new ExoticBorderTuple();
    exoticBorderTuple.push({ kind: 'none' }, undefined, undefined, undefined);
    class ExoticMargin {
      top = 1;
    }
    class ExoticMarginTuple extends Array<number> {}
    const exoticMarginTuple = new ExoticMarginTuple();
    exoticMarginTuple.push(1, 2, 3, 4);
    const sparseBorderTuple = Array(4);
    sparseBorderTuple[0] = { kind: 'none' };
    const extraBorderTuple = Object.assign(
      [{ kind: 'none' }, undefined, undefined, undefined],
      { extra: true },
    );
    const borderSymbol = Symbol('border extra');
    const inheritedMargin = Object.create({ top: 1 });
    const symbolMargin = { top: 1, [Symbol('margin extra')]: 2 };
    const sparseMarginTuple = [1, 2, 3, 4];
    delete sparseMarginTuple[2];
    const extraMarginTuple = Object.assign([1, 2, 3, 4], { extra: true });
    const invalidValigns = [
      null,
      false,
      true,
      0,
      '',
      'Top',
      ' top ',
      'mid',
      'center',
      'just',
      'dist',
      'distributed',
      [],
      {},
      Symbol('top'),
    ];
    const invalidAlignments = [
      null,
      false,
      true,
      0,
      '',
      'Left',
      ' center ',
      'l',
      'ctr',
      'r',
      'just',
      'dist',
      'thaiDist',
      'justLow',
      [],
      {},
      Symbol('center'),
    ];
    const invalidTextDirections = [
      null,
      false,
      true,
      0,
      '',
      'Vert',
      ' vert ',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
      [],
      {},
      Symbol('vert'),
    ];
    const invalidFits = [
      null,
      false,
      true,
      0,
      '',
      'None',
      ' shrink',
      'resize ',
      'auto',
      [],
      {},
      Symbol('fit'),
    ];
    let cellFitAccessorCalls = 0;
    const accessorFitOptions: Record<string, unknown> = {};
    Object.defineProperty(accessorFitOptions, 'fit', {
      get() {
        cellFitAccessorCalls += 1;
        return 'shrink';
      },
      enumerable: true,
    });
    const invalidCells = [
      null,
      1,
      true,
      [],
      new Date(0),
      new ExoticCell(),
      {},
      inheritedCell,
      accessorCell,
      optionsAccessorCell,
      { text: 'A', options: null },
      { text: 'A', options: [] },
      { text: 'A', options: new Date(0) },
      { text: 'A', options: new ExoticCellOptions() },
      { text: 'A', options: Object.create({ fill: undefined }) },
      { text: 'A', options: accessorCellOptions },
      { text: 'A', options: accessorBorderOptions },
      { text: 'A', options: accessorMarginOptions },
      { text: 'A', options: accessorValignOptions },
      { text: 'A', options: accessorAlignOptions },
      { text: 'A', options: accessorTextDirectionOptions },
      { text: 'A', options: accessorFitOptions },
      { text: 'A', options: { unknown: true } },
      { text: 'A', options: Object.assign({}, { [Symbol('extra')]: true }) },
      { text: 'A', options: { fill: accessorFill } },
      { text: 'A', options: { fill: new ExoticFill() } },
      { text: 'A', options: { fill: { kind: 'solid' } } },
      { text: 'A', options: { fill: {
        kind: 'solid',
        color: accessorColor,
      } } },
      { text: 'A', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFF' },
      } } },
      { text: 'A', options: { border: accessorBorder } },
      { text: 'A', options: { border: accessorNamedBorder } },
      { text: 'A', options: { border: accessorBorderTuple } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: accessorBorderColor,
        width: 1,
      } } },
      { text: 'A', options: { border: Object.create({ kind: 'none' }) } },
      { text: 'A', options: { border: Object.create({ top: { kind: 'none' } }) } },
      { text: 'A', options: { border: new ExoticBorder() } },
      { text: 'A', options: { border: exoticBorderTuple } },
      { text: 'A', options: { border: [] } },
      { text: 'A', options: { border: [{ kind: 'none' }] } },
      { text: 'A', options: { border: [
        { kind: 'none' },
        undefined,
        undefined,
      ] } },
      { text: 'A', options: { border: sparseBorderTuple } },
      { text: 'A', options: { border: extraBorderTuple } },
      { text: 'A', options: { border: { top: [{ kind: 'none' }] } } },
      { text: 'A', options: { border: Object.assign(
        { kind: 'none' },
        { [borderSymbol]: true },
      ) } },
      { text: 'A', options: { border: { top: Object.assign(
        { kind: 'none' },
        { [borderSymbol]: true },
      ) } } },
      { text: 'A', options: { border: { kind: 'line', color: Object.create({
        kind: 'srgb',
        value: 'FF0000',
      }), width: 1 } } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: new ExoticBorderColor(),
        width: 1,
      } } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: Object.assign(
          { kind: 'srgb', value: 'FF0000' },
          { [borderSymbol]: true },
        ),
        width: 1,
      } } },
      { text: 'A', options: { border: { kind: 'unknown' } } },
      { text: 'A', options: { border: { kind: 'line' } } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
      } } },
      { text: 'A', options: { border: { kind: 'line', width: 1 } } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFF' },
        width: 1,
      } } },
      { text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'scheme', value: 'unknown' },
        width: 1,
      } } },
      ...[-0.001, 1584.001, Number.NaN, Number.NEGATIVE_INFINITY, Number.POSITIVE_INFINITY]
        .map((width) => ({ text: 'A', options: { border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width,
        } } })),
      { text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1,
        style: 'dot',
      } } },
      { text: 'A', options: { margin: accessorMargin } },
      { text: 'A', options: { margin: accessorMarginTuple } },
      { text: 'A', options: { margin: new ExoticMargin() } },
      { text: 'A', options: { margin: inheritedMargin } },
      { text: 'A', options: { margin: symbolMargin } },
      { text: 'A', options: { margin: sparseMarginTuple } },
      { text: 'A', options: { margin: [1, 2, 3] } },
      { text: 'A', options: { margin: [1, 2, 3, 4, 5] } },
      { text: 'A', options: { margin: extraMarginTuple } },
      { text: 'A', options: { margin: exoticMarginTuple } },
      { text: 'A', options: { margin: { top: null } } },
      { text: 'A', options: { margin: { right: false } } },
      { text: 'A', options: { margin: { bottom: '1' } } },
      { text: 'A', options: { margin: { left: {} } } },
      ...[
        Number.NaN,
        Number.NEGATIVE_INFINITY,
        Number.POSITIVE_INFINITY,
        2_147_483_648 / 12_700,
        -2_147_483_649 / 12_700,
      ].map((margin) => ({ text: 'A', options: { margin } })),
      ...invalidValigns.map((valign) => ({ text: 'A', options: { valign } })),
      ...invalidAlignments.map((align) => ({ text: 'A', options: { align } })),
      ...invalidTextDirections.map((textDirection) => ({
        text: 'A',
        options: { textDirection },
      })),
      ...invalidFits.map((fit) => ({ text: 'A', options: { fit } })),
      { text: 'A', extra: true },
      Object.assign({ text: 'A' }, { [Symbol('extra')]: true }),
      { text: 1 },
      { text: ['rich'] },
      { text: 'bad\u0000xml' },
    ];
    for (const cell of invalidCells) {
      expect(() => normalizeTableDefinition([[cell]], undefined)).toThrow();
    }
    expect(cellAccessorCalls).toBe(0);
    expect(cellFitAccessorCalls).toBe(0);

    const nullPrototype = Object.assign(Object.create(null), {
      name: 'Null prototype',
      width: 914_400,
    });
    expect(normalizeTableDefinition([['A']], nullPrototype)).toMatchObject({
      name: 'Null prototype',
      width: 914_400,
    });

    class ExoticOptions {}
    const symbolOptions = Object.assign({}, { [Symbol('extra')]: true });
    const accessorOptions: Record<string, unknown> = {};
    let optionAccessorCalls = 0;
    Object.defineProperty(accessorOptions, 'name', {
      get() {
        optionAccessorCalls += 1;
        return 'Accessor';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableValignOptions: Record<string, unknown> = {};
    let tableValignAccessorCalls = 0;
    Object.defineProperty(accessorTableValignOptions, 'valign', {
      get() {
        tableValignAccessorCalls += 1;
        return 'top';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableTextDirectionOptions: Record<string, unknown> = {};
    let tableTextDirectionAccessorCalls = 0;
    Object.defineProperty(accessorTableTextDirectionOptions, 'textDirection', {
      get() {
        tableTextDirectionAccessorCalls += 1;
        return 'vert';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableAlignOptions: Record<string, unknown> = {};
    let tableAlignAccessorCalls = 0;
    Object.defineProperty(accessorTableAlignOptions, 'align', {
      get() {
        tableAlignAccessorCalls += 1;
        return 'center';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableMarginOptions: Record<string, unknown> = {};
    let tableMarginAccessorCalls = 0;
    Object.defineProperty(accessorTableMarginOptions, 'margin', {
      get() {
        tableMarginAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableFillOptions: Record<string, unknown> = {};
    const accessorTableFill: Record<string, unknown> = {};
    const accessorTableFillColor: Record<string, unknown> = { kind: 'srgb' };
    let tableFillAccessorCalls = 0;
    Object.defineProperty(accessorTableFillOptions, 'fill', {
      get() {
        tableFillAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableFill, 'kind', {
      get() {
        tableFillAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableFillColor, 'value', {
      get() {
        tableFillAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorTableBorderOptions: Record<string, unknown> = {};
    const accessorTableBorder: Record<string, unknown> = {};
    const accessorTableNamedBorder: Record<string, unknown> = {};
    const accessorTableBorderTuple: unknown[] = [
      undefined,
      undefined,
      undefined,
      undefined,
    ];
    const accessorTableBorderColor: Record<string, unknown> = { kind: 'srgb' };
    let tableBorderAccessorCalls = 0;
    Object.defineProperty(accessorTableBorderOptions, 'border', {
      get() {
        tableBorderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableBorder, 'kind', {
      get() {
        tableBorderAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableNamedBorder, 'top', {
      get() {
        tableBorderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableBorderTuple, '0', {
      get() {
        tableBorderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorTableBorderColor, 'value', {
      get() {
        tableBorderAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
      configurable: true,
    });
    const inheritedTableFill = Object.create({ kind: 'none' });
    const symbolTableFill = { kind: 'none', [Symbol('fill extra')]: true };
    const invalidTableFills = [
      null,
      false,
      'FF0000',
      [],
      {},
      accessorTableFill,
      new ExoticFill(),
      inheritedTableFill,
      symbolTableFill,
      { kind: 'none', transparency: 0 },
      { kind: 'none', extra: true },
      { kind: 'solid' },
      { kind: 'solid', color: accessorTableFillColor },
      { kind: 'solid', color: Object.create({ kind: 'srgb', value: 'FF0000' }) },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'Accent1' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: Number.NaN,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: Number.NEGATIVE_INFINITY,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: Number.POSITIVE_INFINITY,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: -1,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 101,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        extra: true,
      },
    ];
    const invalidTableBorders = [
      null,
      false,
      'FF0000',
      [],
      [{ kind: 'none' }],
      [{ kind: 'none' }, undefined, undefined],
      [{ kind: 'none' }, undefined, undefined, undefined, undefined],
      sparseBorderTuple,
      extraBorderTuple,
      exoticBorderTuple,
      accessorTableBorder,
      accessorTableNamedBorder,
      accessorTableBorderTuple,
      new ExoticBorder(),
      Object.create({ kind: 'none' }),
      Object.create({ top: { kind: 'none' } }),
      Object.assign({ kind: 'none' }, { [borderSymbol]: true }),
      { top: Object.assign({ kind: 'none' }, { [borderSymbol]: true }) },
      { kind: 'none', width: 0 },
      { kind: 'none', extra: true },
      { top: [{ kind: 'none' }] },
      { kind: 'unknown' },
      { kind: 'line' },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'line', width: 1 },
      {
        kind: 'line',
        color: accessorTableBorderColor,
        width: 1,
      },
      {
        kind: 'line',
        color: Object.create({ kind: 'srgb', value: 'FF0000' }),
        width: 1,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFF' },
        width: 1,
      },
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'Accent1' },
        width: 1,
      },
      ...[-0.001, 1584.001, Number.NaN, Number.NEGATIVE_INFINITY, Number.POSITIVE_INFINITY]
        .map((width) => ({
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width,
        })),
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1,
        style: 'dot',
      },
    ];
    const invalidTableMargins = [
      null,
      false,
      '1',
      Number.NaN,
      Number.NEGATIVE_INFINITY,
      Number.POSITIVE_INFINITY,
      accessorMargin,
      accessorMarginTuple,
      new ExoticMargin(),
      inheritedMargin,
      symbolMargin,
      sparseMarginTuple,
      [1, 2, 3],
      [1, 2, 3, 4, 5],
      extraMarginTuple,
      exoticMarginTuple,
      { top: null },
      { right: false },
      { bottom: '1' },
      { left: {} },
      2_147_483_648 / 12_700,
      -2_147_483_649 / 12_700,
    ];
    const invalidOptions = [
      null,
      false,
      '',
      [],
      new ExoticOptions(),
      symbolOptions,
      accessorOptions,
      accessorTableValignOptions,
      accessorTableTextDirectionOptions,
      accessorTableAlignOptions,
      accessorTableMarginOptions,
      accessorTableFillOptions,
      accessorTableBorderOptions,
      { extra: true },
      { name: 1 },
      { name: 'bad\u0000name' },
      { x: '1' },
      { y: Number.NaN },
      { width: Number.POSITIVE_INFINITY },
      { height: Number.NEGATIVE_INFINITY },
      { x: Number.MAX_SAFE_INTEGER + 2 },
      { width: 0 },
      { width: -1 },
      { height: 0 },
      { height: -1 },
      { fit: 'shrink' },
      ...invalidTableFills.map((fill) => ({ fill })),
      ...invalidTableBorders.map((border) => ({ border })),
      ...invalidTableMargins.map((margin) => ({ margin })),
      ...invalidValigns.map((valign) => ({ valign })),
      ...invalidTextDirections.map((textDirection) => ({ textDirection })),
      ...invalidAlignments.map((align) => ({ align })),
    ];
    for (const options of invalidOptions) {
      expect(() => normalizeTableDefinition([['A']], options)).toThrow();
    }
    expect(() => normalizeTableDefinition([['A', 'B']], { width: 1 })).toThrow(RangeError);
    expect(() => normalizeTableDefinition([['A'], ['B']], { height: 1 })).toThrow(RangeError);
    expect(optionAccessorCalls).toBe(0);
    expect(tableValignAccessorCalls).toBe(0);
    expect(tableTextDirectionAccessorCalls).toBe(0);
    expect(tableAlignAccessorCalls).toBe(0);
    expect(tableMarginAccessorCalls).toBe(0);
    expect(tableFillAccessorCalls).toBe(0);
    expect(tableBorderAccessorCalls).toBe(0);
    expect(cellAccessorCalls).toBe(0);

    const sparseWidths = Array(3);
    sparseWidths[0] = 1;
    sparseWidths[2] = 1;
    const extraWidths = Object.assign([1, 1, 1], { extra: true });
    const symbolWidths = Object.assign([1, 1, 1], { [Symbol('extra')]: true });
    const accessorWidths = [1, 1, 1];
    let widthAccessorCalls = 0;
    Object.defineProperty(accessorWidths, '1', {
      get() {
        widthAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    const invalidColumnWidths = [
      null,
      false,
      '',
      {},
      new Uint32Array([1, 1, 1]),
      [],
      [1],
      [1, 1],
      [1, 1, 1, 1],
      sparseWidths,
      extraWidths,
      symbolWidths,
      accessorWidths,
      [1, '2', 3],
      [1, null, 3],
      [1, [2], 3],
      [1, Number.NaN, 3],
      [1, Number.POSITIVE_INFINITY, 3],
      [1, Number.MAX_SAFE_INTEGER + 2, 3],
      [1, 0, 3],
      [1, -1, 3],
    ];
    for (const columnWidths of invalidColumnWidths) {
      expect(() => normalizeTableDefinition(
        [['A', 'B', 'C']],
        { columnWidths },
      )).toThrow();
    }
    expect(widthAccessorCalls).toBe(0);
    expect(() => normalizeTableDefinition(
      [['A', 'B']],
      { columnWidths: [Number.MAX_SAFE_INTEGER, 1] },
    )).toThrow(RangeError);
    expect(() => normalizeTableDefinition(
      [['A', 'B']],
      { width: 4, columnWidths: [1, 2] },
    )).toThrow(RangeError);

    const sparseHeights = Array(3);
    sparseHeights[0] = 1;
    sparseHeights[2] = 1;
    const extraHeights = Object.assign([1, 1, 1], { extra: true });
    const symbolHeights = Object.assign([1, 1, 1], { [Symbol('extra')]: true });
    const accessorHeights = [1, 1, 1];
    let heightAccessorCalls = 0;
    Object.defineProperty(accessorHeights, '1', {
      get() {
        heightAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    const invalidRowHeights = [
      null,
      false,
      '',
      {},
      new Uint32Array([1, 1, 1]),
      [],
      [1],
      [1, 1],
      [1, 1, 1, 1],
      sparseHeights,
      extraHeights,
      symbolHeights,
      accessorHeights,
      [1, '2', 3],
      [1, null, 3],
      [1, [2], 3],
      [1, Number.NaN, 3],
      [1, Number.POSITIVE_INFINITY, 3],
      [1, Number.MAX_SAFE_INTEGER + 2, 3],
      [1, 0, 3],
      [1, -1, 3],
    ];
    for (const rowHeights of invalidRowHeights) {
      expect(() => normalizeTableDefinition(
        [['A'], ['B'], ['C']],
        { rowHeights },
      )).toThrow();
    }
    expect(heightAccessorCalls).toBe(0);
    expect(() => normalizeTableDefinition(
      [['A'], ['B']],
      { rowHeights: [Number.MAX_SAFE_INTEGER, 1] },
    )).toThrow(RangeError);
    expect(() => normalizeTableDefinition(
      [['A'], ['B']],
      { height: 4, rowHeights: [1, 2] },
    )).toThrow(RangeError);
  });
});
