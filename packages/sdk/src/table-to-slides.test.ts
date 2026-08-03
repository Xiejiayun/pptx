import { describe, expect, it } from 'vitest';
import {
  normalizeTableToSlidesRequest,
  snapshotHtmlTable,
  snapshotHtmlTableById,
} from './table-to-slides.js';

interface CellFixtureOptions {
  readonly localName?: 'td' | 'th';
  readonly width?: number;
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly attributes?: Readonly<Record<string, string>>;
}

interface CellFixture {
  localName: 'td' | 'th';
  innerText: string;
  offsetWidth: number;
  colSpan: number;
  rowSpan: number;
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
    getComputedStyle(_element: unknown): Readonly<Record<string, never>> {
      calls.styles += 1;
      return Object.freeze({});
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
