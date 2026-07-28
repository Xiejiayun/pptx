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
        ['A & <1>', '', 'C1'],
        ['A2', 'B2', 'C2'],
      ],
      name: 'Table "A"',
      x: 457_200,
      y: 685_800,
      width: 2_743_201,
      height: 1_371_601,
      autoRowHeight: false,
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
    expect(definition.rows[0]![0]).toBe('A & <1>');
    expect(definition.width).toBe(2_743_201);

    const defaults = normalizeTableDefinition([['A', 'B']], undefined);
    const defaultXml = renderTableGraphicFrame(2, defaults);
    expect(defaults).toEqual({
      rows: [['A', 'B']],
      x: 457_200,
      y: 457_200,
      width: 1_828_800,
      height: 914_400,
      autoRowHeight: true,
    });
    expect(defaultXml).toContain('<a:off x="457200" y="457200"/>');
    expect(defaultXml).toContain('<a:ext cx="1828800" cy="914400"/>');
    expect(defaultXml).toContain('<a:gridCol w="914400"/><a:gridCol w="914400"/>');
    expect(defaultXml).toContain('<a:tr h="0">');
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
  });
});
