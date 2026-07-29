import { describe, expect, it } from 'vitest';
import {
  distributeTableDimension,
  normalizeTableDefinition,
  renderTableGraphicFrame,
} from './table-create.internal.js';

describe('table creation internals', () => {
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
      [['line\nbreak']],
      [['carriage\rreturn']],
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
      { text: 'A', extra: true },
      Object.assign({ text: 'A' }, { [Symbol('extra')]: true }),
      { text: 1 },
      { text: ['rich'] },
      { text: 'line\nbreak' },
      { text: 'carriage\rreturn' },
      { text: 'bad\u0000xml' },
    ];
    for (const cell of invalidCells) {
      expect(() => normalizeTableDefinition([[cell]], undefined)).toThrow();
    }
    expect(cellAccessorCalls).toBe(0);

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
    const invalidOptions = [
      null,
      false,
      '',
      [],
      new ExoticOptions(),
      symbolOptions,
      accessorOptions,
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
    ];
    for (const options of invalidOptions) {
      expect(() => normalizeTableDefinition([['A']], options)).toThrow();
    }
    expect(() => normalizeTableDefinition([['A', 'B']], { width: 1 })).toThrow(RangeError);
    expect(() => normalizeTableDefinition([['A'], ['B']], { height: 1 })).toThrow(RangeError);
    expect(optionAccessorCalls).toBe(0);

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
