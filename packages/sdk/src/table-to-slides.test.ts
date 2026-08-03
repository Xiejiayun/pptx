import { describe, expect, it } from 'vitest';
import { inches } from '@pptx/model';
import {
  normalizeTableToSlidesRequest,
  resolveHtmlTableColumns,
  snapshotHtmlTable,
  snapshotHtmlTableById,
} from './table-to-slides.js';

interface CellFixtureOptions {
  readonly localName?: 'td' | 'th';
  readonly width?: number;
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly style?: Readonly<Record<string, string>>;
}

interface CellFixture {
  localName: 'td' | 'th';
  innerText: string;
  offsetWidth: number;
  colSpan: number;
  rowSpan: number;
  style: Readonly<Record<string, string>>;
  getAttribute(name: string): string | null;
}

interface TableFixtureInput {
  readonly head?: readonly (readonly CellFixture[])[];
  readonly bodies?: readonly (readonly (readonly CellFixture[])[])[];
  readonly foot?: readonly (readonly CellFixture[])[];
  readonly id?: string;
}

function cell(text: string, options: CellFixtureOptions = {}): CellFixture {
  const attributes = { ...(options.attributes ?? {}) };
  return {
    localName: options.localName ?? 'td',
    innerText: text,
    offsetWidth: options.width ?? 100,
    colSpan: options.colSpan ?? 1,
    rowSpan: options.rowSpan ?? 1,
    style: options.style ?? {},
    getAttribute(name: string): string | null {
      return Object.hasOwn(attributes, name) ? attributes[name]! : null;
    },
  };
}

function tableFixture(input: TableFixtureInput) {
  const rows = (value: readonly (readonly CellFixture[])[] | undefined) =>
    value === undefined ? undefined : { rows: value.map((cells) => ({ cells })) };
  const table = {
    localName: 'table',
    tHead: rows(input.head) ?? null,
    tBodies: (input.bodies ?? []).map((body) => rows(body)!),
    tFoot: rows(input.foot) ?? null,
    ownerDocument: undefined as unknown,
  };
  const calls = {
    ids: [] as string[],
    styles: 0,
  };
  const defaultView = {
    getComputedStyle(element: unknown) {
      calls.styles += 1;
      const style = (element as { style?: Readonly<Record<string, string>> }).style ?? {};
      return Object.freeze({
        getPropertyValue(name: string): string {
          return style[name] ?? '';
        },
      });
    },
  };
  const document = {
    defaultView,
    getElementById(id: string): unknown {
      calls.ids.push(id);
      return id === (input.id ?? 'table') ? table : null;
    },
  };
  table.ownerDocument = document;
  const cells = [
    ...(input.head ?? []).flat(),
    ...(input.bodies ?? []).flat(2),
    ...(input.foot ?? []).flat(),
  ];
  return {
    table,
    document,
    cells,
    calls,
    getComputedStyle: defaultView.getComputedStyle.bind(defaultView),
  };
}

describe('HTML table row snapshots', () => {
  it('snapshots thead, multiple tbody sections, and tfoot without retaining DOM', () => {
    const dom = tableFixture({
      head: [[cell('Head', { colSpan: 2, width: 200, localName: 'th' })]],
      bodies: [
        [[cell('A\r\nB', { width: 80 }), cell('C', { width: 120 })]],
        [[cell('D', { rowSpan: 2, width: 80 }), cell('E', { width: 120 })]],
      ],
      foot: [[cell('Foot A', { width: 80 }), cell('Foot B', { width: 120 })]],
    });
    const snapshot = snapshotHtmlTable(dom.table, dom.getComputedStyle);
    expect(snapshot.headRowCount).toBe(1);
    expect(snapshot.widthSourceRowIndex).toBe(0);
    expect(snapshot.rows.map((row) => row.map(({ text }) => text))).toEqual([
      ['Head'], ['A\nB', 'C'], ['D', 'E'], ['Foot A', 'Foot B'],
    ]);
    expect(snapshot.rows[0]![0]).toMatchObject({
      colspan: 2,
      header: true,
      offsetWidth: 200,
      options: {},
    });
    expect(snapshot.rows[2]![0]).toMatchObject({ rowspan: 2 });
    expect(snapshot.rows[1]![0]).not.toHaveProperty('colspan');
    expect(snapshot.rows[1]![0]).not.toHaveProperty('rowspan');
    dom.cells[0]!.innerText = 'mutated';
    expect(snapshot.rows[0]![0]!.text).toBe('Head');
    expect(dom.calls.styles).toBe(dom.cells.length);
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.rows)).toBe(true);
    expect(Object.isFrozen(snapshot.rows[0])).toBe(true);
    expect(Object.isFrozen(snapshot.rows[0]![0])).toBe(true);
    expect(Object.isFrozen(snapshot.rows[0]![0]!.options)).toBe(true);
  });

  it('uses getElementById without selector interpolation and preserves section order', () => {
    const id = 'sales:2026 [data="q"] table';
    const dom = tableFixture({
      id,
      head: [[cell('H', { localName: 'th' })]],
      bodies: [[[cell('B1')]], [[cell('B2')]]],
      foot: [[cell('F')]],
    });
    const snapshot = snapshotHtmlTableById(id, dom.document);
    expect(dom.calls.ids).toEqual([id]);
    expect(snapshot.rows.map((row) => row[0]!.text)).toEqual(['H', 'B1', 'B2', 'F']);
  });

  it('snapshots width attributes and reads each platform cell value once', () => {
    let textReads = 0;
    let widthReads = 0;
    let colSpanReads = 0;
    let rowSpanReads = 0;
    const attributes: string[] = [];
    const tracked = {
      localName: 'th',
      get innerText() {
        textReads += 1;
        return 'Tracked\rvalue';
      },
      get offsetWidth() {
        widthReads += 1;
        return 250;
      },
      get colSpan() {
        colSpanReads += 1;
        return 2;
      },
      get rowSpan() {
        rowSpanReads += 1;
        return 1;
      },
      getAttribute(name: string) {
        attributes.push(name);
        return name === 'data-pptx-width' ? '2.5' : '1.25';
      },
    };
    const snapshot = snapshotHtmlTable(
      { localName: 'table', tHead: null, tBodies: [{ rows: [{ cells: [tracked] }] }], tFoot: null },
      () => ({}),
    );
    expect(snapshot.rows[0]![0]).toMatchObject({
      text: 'Tracked\nvalue',
      offsetWidth: 250,
      colspan: 2,
      pptxWidth: '2.5',
      pptxMinWidth: '1.25',
    });
    expect({ textReads, widthReads, colSpanReads, rowSpanReads }).toEqual({
      textReads: 1,
      widthReads: 1,
      colSpanReads: 1,
      rowSpanReads: 1,
    });
    expect(attributes).toEqual(['data-pptx-width', 'data-pptx-min-width']);
  });

  it.each([
    [undefined, /browser document/i],
    [{}, /getElementById/i],
    [{ getElementById: () => null }, /was not found/i],
    [{ getElementById: () => 'table' }, /must be a table/i],
    [{ getElementById: () => ({ localName: 'div' }) }, /must be a table/i],
  ])('rejects invalid DOM lookup state %#', (documentValue, expected) => {
    expect(() => snapshotHtmlTableById('table', documentValue)).toThrow(expected);
  });

  it('rejects missing view/style support and malformed table structures', () => {
    const dom = tableFixture({ bodies: [[[cell('A')]]] });
    (dom.table as { ownerDocument: unknown }).ownerDocument = {};
    expect(() => snapshotHtmlTableById('table', dom.document)).toThrow(/defaultView/i);

    (dom.table as { ownerDocument: unknown }).ownerDocument = { defaultView: {} };
    expect(() => snapshotHtmlTableById('table', dom.document)).toThrow(/getComputedStyle/i);

    expect(() => snapshotHtmlTable(
      { localName: 'table', tHead: null, tBodies: [], tFoot: null },
      () => ({}),
    )).toThrow(/at least one row/i);
    expect(() => snapshotHtmlTable(
      { localName: 'table', tHead: null, tBodies: [{ rows: [{ cells: [] }] }], tFoot: null },
      () => ({}),
    )).toThrow(/at least one cell/i);
  });

  it('wraps platform failures with deterministic lookup and cell context', () => {
    expect(() => snapshotHtmlTableById('table', {
      getElementById() {
        throw new Error('lookup failed');
      },
    })).toThrow(/document getElementById failed/i);
    expect(() => snapshotHtmlTable(
      {
        localName: 'table',
        tHead: null,
        tBodies: [{
          rows: [{
            cells: [{
              localName: 'td',
              innerText: 'A',
              offsetWidth: 1,
              colSpan: 1,
              rowSpan: 1,
              getAttribute() {
                throw new Error('attribute failed');
              },
            }],
          }],
        }],
        tFoot: null,
      },
      () => ({}),
    )).toThrow(/cell 0:0 getAttribute\(data-pptx-width\) failed/i);
  });

  it.each([
    [{ localName: 'td', innerText: 'A', offsetWidth: -1, colSpan: 1, rowSpan: 1 }, /offsetWidth/i],
    [{ localName: 'td', innerText: 1, offsetWidth: 1, colSpan: 1, rowSpan: 1 }, /innerText/i],
    [{ localName: 'td', innerText: 'A', offsetWidth: 1, colSpan: 0, rowSpan: 1 }, /colSpan/i],
    [{ innerText: 'A', offsetWidth: 1, colSpan: 1, rowSpan: 1 }, /localName/i],
  ])('rejects malformed cells %#', (invalidCell, expected) => {
    expect(() => snapshotHtmlTable(
      {
        localName: 'table',
        tHead: null,
        tBodies: [{ rows: [{ cells: [{ ...invalidCell, getAttribute: () => null }] }] }],
        tFoot: null,
      },
      () => ({}),
    )).toThrow(expected);
  });
});

describe('HTML table computed CSS snapshots', () => {
  it('maps exact computed cell CSS to editable native table options', () => {
    const styled = cell('Styled', {
      style: {
        color: 'rgb(1, 2, 3)',
        'background-color': 'rgb(240, 241, 242)',
        'font-family': '"Noto Sans", Arial, sans-serif',
        'font-size': '18.5px',
        'font-weight': '600',
        'text-align': 'right',
        'vertical-align': 'bottom',
        direction: 'ltr',
        'padding-top': '7.5px',
        'padding-right': '11px',
        'padding-bottom': '3.25px',
        'padding-left': '5px',
        'border-top-style': 'solid',
        'border-top-width': '2px',
        'border-top-color': 'rgba(10, 20, 30, 0.5)',
        'border-right-style': 'dashed',
        'border-right-width': '1.5px',
        'border-right-color': 'rgb(40, 50, 60)',
        'border-bottom-style': 'none',
        'border-bottom-width': '0px',
        'border-bottom-color': 'rgb(0, 0, 0)',
        'border-left-style': 'dotted',
        'border-left-width': '3px',
        'border-left-color': 'rgb(70, 80, 90)',
      },
    });
    const dom = tableFixture({ bodies: [[[styled]]] });
    const options = snapshotHtmlTable(dom.table, dom.getComputedStyle).rows[0]![0]!.options;
    expect(options).toEqual({
      align: 'right',
      bold: true,
      color: { kind: 'srgb', value: '010203' },
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'F0F1F2' } },
      fontFamily: 'Noto Sans',
      fontSize: 18.5,
      margin: [7.5, 11, 3.25, 5],
      valign: 'bottom',
      border: {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: '0A141E' },
          width: 2,
          style: 'solid',
        },
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: '28323C' },
          width: 1.5,
          style: 'dash',
        },
        bottom: { kind: 'none' },
        left: {
          kind: 'line',
          color: { kind: 'srgb', value: '46505A' },
          width: 3,
          style: 'dash',
        },
      },
    });
    expect(Object.isFrozen(options)).toBe(true);
    expect(Object.isFrozen(options.margin)).toBe(true);
    expect(Object.isFrozen(options.border)).toBe(true);
    (styled.style as Record<string, string>).color = 'rgb(255, 255, 255)';
    expect(options.color).toEqual({ kind: 'srgb', value: '010203' });
  });

  it('maps transparent fill, decimal channels, font weights, and directional alignment', () => {
    const dom = tableFixture({
      bodies: [[
        [cell('transparent', { style: {
          color: 'rgba(1.4, 2.5, 3.6, 0.25)',
          'background-color': 'rgba(0, 0, 0, 0)',
          'font-weight': '400',
          'text-align': 'start',
          direction: 'rtl',
          'vertical-align': 'middle',
        } })],
        [cell('bold', { style: {
          'background-color': 'transparent',
          'font-weight': 'bolder',
          'text-align': 'end',
          direction: 'rtl',
          'vertical-align': 'top',
        } })],
      ]],
    });
    const rows = snapshotHtmlTable(dom.table, dom.getComputedStyle).rows;
    expect(rows[0]![0]!.options).toMatchObject({
      align: 'right',
      bold: false,
      color: { kind: 'srgb', value: '010304' },
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } },
      valign: 'middle',
    });
    expect(rows[1]![0]!.options).toMatchObject({
      align: 'left',
      bold: true,
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } },
      valign: 'top',
    });
  });

  it.each([
    ['normal', false],
    ['400', false],
    ['500', true],
    ['bold', true],
    ['bolder', true],
  ])('maps font-weight %s to bold=%s', (weight, bold) => {
    const dom = tableFixture({ bodies: [[[cell('A', { style: { 'font-weight': weight } })]]] });
    expect(snapshotHtmlTable(dom.table, dom.getComputedStyle).rows[0]![0]!.options.bold)
      .toBe(bold);
  });

  it.each([
    ['left', 'ltr', 'left'],
    ['center', 'ltr', 'center'],
    ['right', 'ltr', 'right'],
    ['justify', 'rtl', 'justify'],
    ['start', 'ltr', 'left'],
    ['end', 'ltr', 'right'],
  ])('maps text alignment %s in %s to %s', (alignment, direction, expected) => {
    const dom = tableFixture({ bodies: [[[cell('A', { style: {
      'text-align': alignment,
      direction,
    } })]]] });
    expect(snapshotHtmlTable(dom.table, dom.getComputedStyle).rows[0]![0]!.options.align)
      .toBe(expected);
  });

  it('accepts modern space-separated RGB and percentage alpha', () => {
    const dom = tableFixture({ bodies: [[[cell('A', { style: {
      color: 'rgb(10.4 20.5 30.6 / 25%)',
    } })]]] });
    expect(snapshotHtmlTable(dom.table, dom.getComputedStyle).rows[0]![0]!.options.color)
      .toEqual({ kind: 'srgb', value: '0A151F' });
  });

  it('omits intentionally empty and generic computed values', () => {
    const dom = tableFixture({ bodies: [[[cell('A', { style: {
      'font-family': 'sans-serif',
      'vertical-align': 'baseline',
    } })]]] });
    expect(snapshotHtmlTable(dom.table, dom.getComputedStyle).rows[0]![0]!.options)
      .toEqual({});
  });

  it('reads every required computed property exactly once per cell', () => {
    const reads = new Map<string, number>();
    const tracked = {
      localName: 'td',
      innerText: 'A',
      offsetWidth: 1,
      colSpan: 1,
      rowSpan: 1,
      getAttribute: () => null,
    };
    snapshotHtmlTable(
      { localName: 'table', tHead: null, tBodies: [{ rows: [{ cells: [tracked] }] }], tFoot: null },
      () => ({
        getPropertyValue(name: string) {
          reads.set(name, (reads.get(name) ?? 0) + 1);
          return '';
        },
      }),
    );
    expect(reads.size).toBe(24);
    expect([...reads.values()].every((count) => count === 1)).toBe(true);
  });

  it.each([
    [{ color: 'lab(10% 0 0)' }, /color.*rgb/i],
    [{ color: 'rgb(256, 0, 0)' }, /channel/i],
    [{ color: 'rgba(0, 0, 0, )' }, /alpha/i],
    [{ 'font-size': 'NaNpx' }, /font-size/i],
    [{ 'padding-left': '-1px' }, /padding-left/i],
    [{ 'border-top-style': 'solid', 'border-top-width': 'wide', 'border-top-color': 'rgb(0,0,0)' }, /border-top-width/i],
    [{ 'border-left-style': 'sparkle', 'border-left-width': '1px', 'border-left-color': 'rgb(0,0,0)' }, /border-left-style/i],
    [{ 'text-align': 'match-parent' }, /text-align/i],
    [{ 'text-align': 'start', direction: 'sideways' }, /direction/i],
  ])('rejects malformed non-empty computed CSS %#', (style, expected) => {
    const dom = tableFixture({ bodies: [[[cell('A', { style })]]] });
    expect(() => snapshotHtmlTable(dom.table, dom.getComputedStyle)).toThrow(expected);
  });
});

describe('HTML table column width resolution', () => {
  it('maps a 900px 25/75 table to exact EMU proportions', () => {
    const dom = tableFixture({ bodies: [[[
      cell('A', { width: 225 }),
      cell('B', { width: 675 }),
    ]]] });
    const snapshot = snapshotHtmlTable(dom.table, dom.getComputedStyle);
    const resolved = resolveHtmlTableColumns(snapshot, inches(10.8));
    expect(resolved).toEqual({
      widths: [inches(2.7), inches(8.1)],
      width: inches(10.8),
    });
    expect(Object.isFrozen(resolved)).toBe(true);
    expect(Object.isFrozen(resolved.widths)).toBe(true);
  });

  it('uses stable largest remainders and expands colspan pixel weights', () => {
    const equal = tableFixture({ bodies: [[[
      cell('A', { width: 1 }),
      cell('B', { width: 1 }),
      cell('C', { width: 1 }),
    ]]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(equal.table, equal.getComputedStyle),
      10,
    ).widths).toEqual([4, 3, 3]);

    const spanned = tableFixture({ bodies: [[[
      cell('AB', { width: 300, colSpan: 2 }),
      cell('C', { width: 100 }),
    ]]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(spanned.table, spanned.getComputedStyle),
      400,
    ).widths).toEqual([150, 150, 100]);
  });

  it('applies fixed and minimum header constraints with true water filling', () => {
    const dom = tableFixture({ head: [[
      cell('A', {
        localName: 'th',
        width: 100,
        attributes: { 'data-pptx-min-width': '0.8' },
      }),
      cell('B', {
        localName: 'th',
        width: 200,
        attributes: { 'data-pptx-width': '4.5' },
      }),
      cell('C', { localName: 'th', width: 100 }),
    ]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(dom.table, dom.getComputedStyle),
      inches(8),
    )).toEqual({
      widths: [inches(1.75), inches(4.5), inches(1.75)],
      width: inches(8),
    });
  });

  it('expands the table when fixed plus minimum constraints overflow', () => {
    const dom = tableFixture({ head: [[
      cell('A', {
        localName: 'th',
        width: 1,
        attributes: { 'data-pptx-width': '1.5' },
      }),
      cell('B', {
        localName: 'th',
        width: 1,
        attributes: { 'data-pptx-min-width': '1' },
      }),
    ]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(dom.table, dom.getComputedStyle),
      inches(2),
    )).toEqual({
      widths: [inches(1.5), inches(1)],
      width: inches(2.5),
    });
  });

  it('lets fixed width win over minimum and distributes colspan constraints', () => {
    const dom = tableFixture({ head: [[
      cell('AB', {
        localName: 'th',
        colSpan: 2,
        width: 200,
        attributes: {
          'data-pptx-width': '3',
          'data-pptx-min-width': 'not-a-number',
        },
      }),
      cell('C', { localName: 'th', width: 100 }),
    ]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(dom.table, dom.getComputedStyle),
      inches(6),
    )).toEqual({
      widths: [inches(1.5), inches(1.5), inches(3)],
      width: inches(6),
    });
  });

  it('uses explicit scalar and vector widths before pixels or attributes', () => {
    const dom = tableFixture({ head: [[
      cell('A', {
        localName: 'th',
        width: 0,
        attributes: { 'data-pptx-width': 'bad' },
      }),
      cell('B', { localName: 'th', width: 0 }),
    ]] });
    const snapshot = snapshotHtmlTable(dom.table, dom.getComputedStyle);
    expect(resolveHtmlTableColumns(snapshot, 200, 100)).toEqual({
      widths: [100, 100],
      width: 200,
    });
    expect(resolveHtmlTableColumns(snapshot, 200, [80, 120])).toEqual({
      widths: [80, 120],
      width: 200,
    });
    expect(() => resolveHtmlTableColumns(snapshot, 201, [80, 120]))
      .toThrow(/equal.*target width/i);
  });

  it('rejects fully hidden flexible tables but permits all-fixed hidden tables', () => {
    const hidden = tableFixture({ bodies: [[[
      cell('A', { width: 0 }),
      cell('B', { width: 0 }),
    ]]] });
    expect(() => resolveHtmlTableColumns(
      snapshotHtmlTable(hidden.table, hidden.getComputedStyle),
      100,
    )).toThrow(/hidden|columnWidths/i);

    const fixed = tableFixture({ head: [[
      cell('A', {
        localName: 'th',
        width: 0,
        attributes: { 'data-pptx-width': '1' },
      }),
      cell('B', {
        localName: 'th',
        width: 0,
        attributes: { 'data-pptx-width': '2' },
      }),
    ]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(fixed.table, fixed.getComputedStyle),
      inches(10),
    )).toEqual({
      widths: [inches(1), inches(2)],
      width: inches(3),
    });
  });

  it.each(['', '-1', 'NaN', 'Infinity', '0x10'])('rejects invalid width attribute %j', (width) => {
    const dom = tableFixture({ head: [[cell('A', {
      localName: 'th',
      attributes: { 'data-pptx-width': width },
    })]] });
    expect(() => resolveHtmlTableColumns(
      snapshotHtmlTable(dom.table, dom.getComputedStyle),
      inches(2),
    )).toThrow(/data-pptx-width/i);
  });

  it.each(['', '-1', 'NaN', 'Infinity', '0x10'])(
    'rejects invalid minimum width attribute %j',
    (minimum) => {
      const dom = tableFixture({ head: [[cell('A', {
        localName: 'th',
        attributes: { 'data-pptx-min-width': minimum },
      })]] });
      expect(() => resolveHtmlTableColumns(
        snapshotHtmlTable(dom.table, dom.getComputedStyle),
        inches(2),
      )).toThrow(/data-pptx-min-width/i);
    },
  );

  it('accepts a zero minimum width while preserving the one-EMU column floor', () => {
    const dom = tableFixture({ head: [[
      cell('A', {
        localName: 'th',
        width: 1,
        attributes: { 'data-pptx-min-width': '0' },
      }),
      cell('B', { localName: 'th', width: 1 }),
    ]] });
    expect(resolveHtmlTableColumns(
      snapshotHtmlTable(dom.table, dom.getComputedStyle),
      3,
    ).widths).toEqual([2, 1]);
  });

  it('rejects impossible physical column and safe-integer allocations', () => {
    const excessive = tableFixture({ bodies: [[[
      cell('A', { colSpan: 1_000_001 }),
    ]]] });
    expect(() => resolveHtmlTableColumns(
      snapshotHtmlTable(excessive.table, excessive.getComputedStyle),
      2_000_002,
    )).toThrow(/physical columns/i);

    const two = tableFixture({ bodies: [[[
      cell('A'),
      cell('B'),
    ]]] });
    const snapshot = snapshotHtmlTable(two.table, two.getComputedStyle);
    expect(() => resolveHtmlTableColumns(snapshot, 1)).toThrow(/one EMU per column/i);
    expect(() => resolveHtmlTableColumns(
      snapshot,
      Number.MAX_SAFE_INTEGER,
      [Number.MAX_SAFE_INTEGER, 1],
    )).toThrow(/safe integer|overflow/i);

    const sparse = [1, 2];
    delete sparse[0];
    expect(() => resolveHtmlTableColumns(snapshot, 3, sparse)).toThrow(/dense/i);
  });
});

describe('tableToSlides request normalization', () => {
  it('defaults autoPage and freezes a detached request', () => {
    const request = normalizeTableToSlidesRequest('table', undefined);
    expect(request).toEqual({ id: 'table', autoPage: true });
    expect(Object.isFrozen(request)).toBe(true);
  });

  it('normalizes every scalar and tuple request field', () => {
    const options = {
      name: 'Sales',
      masterSlideName: 'REPORT',
      autoPage: true,
      autoPageCharWeight: -0,
      autoPageLineWeight: 0.5,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 2,
      autoPageSlideStartY: 100,
      slideMargin: [1, 2, 3, 4] as const,
      x: -10.4,
      y: 20.6,
      width: 1_000.4,
      height: 2_000.4,
      columnWidths: [400.4, 600.4],
    };
    const request = normalizeTableToSlidesRequest('table', options);
    expect(request).toMatchObject({
      id: 'table',
      name: 'Sales',
      masterSlideName: 'REPORT',
      autoPage: true,
      autoPageCharWeight: 0,
      autoPageLineWeight: 0.5,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 2,
      autoPageSlideStartY: 100,
      slideMargin: [1, 2, 3, 4],
      x: -10,
      y: 21,
      width: 1_000,
      height: 2_000,
      columnWidths: [400, 600],
    });
    options.columnWidths[0] = 999;
    expect(request.columnWidths).toEqual([400, 600]);
    expect(Object.isFrozen(request.slideMargin)).toBe(true);
    expect(Object.isFrozen(request.columnWidths)).toBe(true);
  });

  it('normalizes strict outer addition records without invoking accessors', () => {
    const request = normalizeTableToSlidesRequest('table', {
      addImage: { source: Uint8Array.of(1), options: { width: 10, height: 20 } },
      addShape: { type: 'rect', options: { x: 1 } },
      addTable: { rows: [['A']], options: { x: 2 } },
      addText: { text: 'T', options: { x: 3 } },
    });
    expect(request.addImage).toMatchObject({ source: Uint8Array.of(1) });
    expect(request.addShape).toMatchObject({ type: 'rect' });
    expect(request.addTable).toMatchObject({ rows: [['A']] });
    expect(request.addText).toMatchObject({ text: 'T' });
    expect([
      request.addImage,
      request.addShape,
      request.addTable,
      request.addText,
    ].every(Object.isFrozen)).toBe(true);

    let reads = 0;
    const accessor = Object.defineProperty({}, 'source', {
      enumerable: true,
      get() {
        reads += 1;
        return Uint8Array.of(1);
      },
    });
    expect(() => normalizeTableToSlidesRequest('table', { addImage: accessor }))
      .toThrow(/data property/i);
    expect(reads).toBe(0);
  });

  it.each([
    ['', undefined, /non-empty string/i],
    [1, undefined, /non-empty string/i],
    ['table', null, /options must be an object/i],
    ['table', [], /options must be an object/i],
    ['table', Object.create({ autoPage: true }), /ordinary object/i],
    ['table', new (class Options {})(), /ordinary object/i],
    ['table', { unknown: true }, /unsupported property unknown/i],
    ['table', { autoPage: 1 }, /autoPage must be a boolean/i],
    ['table', { autoPageRepeatHeader: 'yes' }, /autoPageRepeatHeader must be a boolean/i],
    ['table', { autoPageCharWeight: -1.1 }, /between -1 and 1/i],
    ['table', { autoPageLineWeight: Number.NaN }, /finite/i],
    ['table', { autoPageHeaderRows: 0 }, /positive safe integer/i],
    ['table', { autoPageSlideStartY: -1 }, /non-negative/i],
    ['table', { x: Number.POSITIVE_INFINITY }, /finite/i],
    ['table', { width: 0 }, /positive/i],
    ['table', { height: -1 }, /positive/i],
    ['table', { columnWidths: [] }, /non-empty/i],
    ['table', { columnWidths: [1, 0] }, /positive/i],
    ['table', { masterSlideName: '' }, /non-empty string/i],
    ['table', { autoPage: false, slideMargin: 1 }, /require autoPage to be true/i],
    ['table', { autoPage: true, autoPageHeaderRows: 1 }, /requires autoPageRepeatHeader/i],
  ])('rejects invalid requests %#', (elementId, options, expected) => {
    expect(() => normalizeTableToSlidesRequest(elementId, options)).toThrow(expected);
  });

  it('rejects accessors, symbols, sparse tuples, and malformed addition records', () => {
    let reads = 0;
    const accessor = Object.defineProperty({}, 'autoPage', {
      enumerable: true,
      get() {
        reads += 1;
        return true;
      },
    });
    expect(() => normalizeTableToSlidesRequest('table', accessor)).toThrow(/data property/i);
    expect(reads).toBe(0);
    expect(() => normalizeTableToSlidesRequest('table', { [Symbol('x')]: true }))
      .toThrow(/unsupported property Symbol\(x\)/i);

    const margin = [1, 2, 3, 4];
    delete margin[2];
    expect(() => normalizeTableToSlidesRequest('table', { slideMargin: margin }))
      .toThrow(/dense/i);
    const widths = [1, 2];
    delete widths[0];
    expect(() => normalizeTableToSlidesRequest('table', { columnWidths: widths }))
      .toThrow(/dense/i);

    expect(() => normalizeTableToSlidesRequest('table', { addShape: { options: {} } }))
      .toThrow(/type/i);
    expect(() => normalizeTableToSlidesRequest('table', { addText: { text: 1 } }))
      .toThrow(/text/i);
    expect(() => normalizeTableToSlidesRequest('table', {
      addTable: { rows: [['A']], extra: true },
    })).toThrow(/unsupported property extra/i);
  });
});
