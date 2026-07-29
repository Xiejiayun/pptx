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
      accessorTableMarginOptions,
      accessorTableFillOptions,
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
      ...invalidTableFills.map((fill) => ({ fill })),
      ...invalidTableMargins.map((margin) => ({ margin })),
      ...invalidValigns.map((valign) => ({ valign })),
    ];
    for (const options of invalidOptions) {
      expect(() => normalizeTableDefinition([['A']], options)).toThrow();
    }
    expect(() => normalizeTableDefinition([['A', 'B']], { width: 1 })).toThrow(RangeError);
    expect(() => normalizeTableDefinition([['A'], ['B']], { height: 1 })).toThrow(RangeError);
    expect(optionAccessorCalls).toBe(0);
    expect(tableValignAccessorCalls).toBe(0);
    expect(tableMarginAccessorCalls).toBe(0);
    expect(tableFillAccessorCalls).toBe(0);
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
