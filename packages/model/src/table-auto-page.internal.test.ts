import { describe, expect, it } from 'vitest';
import { normalizeTableDefinition } from './table-create.internal.js';
import {
  planTableAutoPages as partitionTableAutoPages,
  resolveTableAutoPageLayout,
  type TableAutoPageMargins,
} from './table-auto-page.internal.js';
import { EMU_PER_INCH, type Emu, type SlideSize } from './units.js';

const emu = (value: number): Emu => value as Emu;

const SLIDE: SlideSize = { width: emu(1_000), height: emu(100) };

function definition(
  rows: readonly (readonly unknown[])[],
  options: Readonly<Record<string, unknown>>,
) {
  return normalizeTableDefinition(rows, {
    autoPage: true,
    ...options,
  });
}

function planTableAutoPages(
  source: ReturnType<typeof definition>,
  slideSize: Readonly<SlideSize>,
  layoutMargins?: Readonly<TableAutoPageMargins>,
  bottomEdgeOverride?: number,
) {
  const region = resolveTableAutoPageLayout(
    source,
    slideSize,
    layoutMargins,
    bottomEdgeOverride,
  );
  return partitionTableAutoPages(source, region);
}

function firstColumnText(
  pages: ReturnType<typeof planTableAutoPages>,
): readonly (readonly string[])[] {
  return pages.map((page) => page.rows.map((row) => row[0]!.text));
}

describe('table auto-page layout region', () => {
  it('resolves one frozen layout region for measurement and partition', () => {
    const source = definition([['A']], {
      y: 20,
      rowHeights: [10],
      autoPageSlideStartY: 5,
      slideMargin: [3, 4, 7, 6],
    });

    const region = resolveTableAutoPageLayout(source, SLIDE);
    expect(region).toEqual({
      firstY: 20,
      continuationY: 5,
      bottomEdge: 93,
      firstCapacity: 73,
      continuationCapacity: 88,
    });
    expect(Object.isFrozen(region)).toBe(true);
    expect(partitionTableAutoPages({ ...source, y: 99 }, region)[0]!.y).toBe(20);
  });

  it('uses explicit margins before layout margins and canonical defaults', () => {
    const layoutMargins: TableAutoPageMargins = {
      top: 10,
      right: 20,
      bottom: 30,
      left: 40,
    };
    const layout = resolveTableAutoPageLayout(definition([['A']], {
      y: 20,
      rowHeights: [10],
    }), SLIDE, layoutMargins);
    expect(layout).toEqual({
      firstY: 20,
      continuationY: 10,
      bottomEdge: 70,
      firstCapacity: 50,
      continuationCapacity: 60,
    });

    const explicit = resolveTableAutoPageLayout(definition([['A']], {
      y: 20,
      rowHeights: [10],
      slideMargin: [5, 6, 10, 8],
    }), SLIDE, layoutMargins);
    expect(explicit).toEqual({
      firstY: 20,
      continuationY: 5,
      bottomEdge: 90,
      firstCapacity: 70,
      continuationCapacity: 85,
    });

    const canonicalSlide: SlideSize = {
      width: emu(2_000_000),
      height: emu(1_000_000),
    };
    const canonical = resolveTableAutoPageLayout(definition([['A']], {
      y: 500_000,
      rowHeights: [10],
    }), canonicalSlide);
    const defaultMargin = EMU_PER_INCH / 2;
    expect(canonical).toEqual({
      firstY: 500_000,
      continuationY: defaultMargin,
      bottomEdge: 1_000_000 - defaultMargin,
      firstCapacity: 1_000_000 - defaultMargin - 500_000,
      continuationCapacity: 1_000_000 - (2 * defaultMargin),
    });
  });

  it('applies the exact stricter bottom edge and clips a looser override', () => {
    const source = definition([['A']], {
      y: 20,
      rowHeights: [10],
      autoPageSlideStartY: 5,
      slideMargin: [0, 0, 10, 0],
    });
    expect(resolveTableAutoPageLayout(source, SLIDE, undefined, 80)).toEqual({
      firstY: 20,
      continuationY: 5,
      bottomEdge: 80,
      firstCapacity: 60,
      continuationCapacity: 75,
    });
    expect(resolveTableAutoPageLayout(source, SLIDE, undefined, 99)).toEqual({
      firstY: 20,
      continuationY: 5,
      bottomEdge: 90,
      firstCapacity: 70,
      continuationCapacity: 85,
    });
  });

  it.each([
    ['zero bottom override', () => resolveTableAutoPageLayout(
      definition([['A']], { y: 0, rowHeights: [10], slideMargin: 0 }),
      SLIDE,
      undefined,
      0,
    ), /bottom edge/i],
    ['fractional bottom override', () => resolveTableAutoPageLayout(
      definition([['A']], { y: 0, rowHeights: [10], slideMargin: 0 }),
      SLIDE,
      undefined,
      1.5,
    ), /bottom edge/i],
    ['override at source Y', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 20,
        rowHeights: [10],
        autoPageSlideStartY: 5,
        slideMargin: 0,
      }),
      SLIDE,
      undefined,
      20,
    ), /source height/i],
    ['override before source Y', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 20,
        rowHeights: [10],
        autoPageSlideStartY: 5,
        slideMargin: 0,
      }),
      SLIDE,
      undefined,
      19,
    ), /source height/i],
    ['override at continuation Y', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 5,
        rowHeights: [10],
        autoPageSlideStartY: 20,
        slideMargin: 0,
      }),
      SLIDE,
      undefined,
      20,
    ), /continuation height/i],
    ['horizontal margin exhaustion', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 0,
        rowHeights: [10],
        slideMargin: [0, 600, 0, 400],
      }),
      SLIDE,
    ), /horizontal/i],
    ['vertical margin exhaustion', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 0,
        rowHeights: [10],
        slideMargin: [50, 0, 50, 0],
      }),
      SLIDE,
    ), /vertical/i],
    ['unsafe slide width', () => resolveTableAutoPageLayout(
      definition([['A']], { y: 0, rowHeights: [10], slideMargin: 0 }),
      { width: emu(Number.MAX_SAFE_INTEGER + 1), height: emu(100) },
    ), /slide width/i],
    ['unsafe slide height', () => resolveTableAutoPageLayout(
      definition([['A']], { y: 0, rowHeights: [10], slideMargin: 0 }),
      { width: emu(100), height: emu(Number.MAX_SAFE_INTEGER + 1) },
    ), /slide height/i],
    ['unsafe bottom override', () => resolveTableAutoPageLayout(
      definition([['A']], { y: 0, rowHeights: [10], slideMargin: 0 }),
      SLIDE,
      undefined,
      Number.MAX_SAFE_INTEGER + 1,
    ), /bottom edge/i],
    ['overflowing margin sum', () => resolveTableAutoPageLayout(
      definition([['A']], {
        y: 0,
        rowHeights: [10],
        slideMargin: [0, Number.MAX_SAFE_INTEGER, 0, 1],
      }),
      { width: emu(Number.MAX_SAFE_INTEGER), height: emu(100) },
    ), /safe integer/i],
  ])('rejects %s', (_name, resolve, message) => {
    expect(resolve).toThrow(message);
  });
});

describe('table auto-page partition planner', () => {
  it('partitions exact EMU rows and repeats headers', () => {
    const pages = planTableAutoPages(definition(
      [['H'], ['A'], ['B'], ['C']],
      {
        y: 20,
        rowHeights: [20, 30, 30, 30],
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageSlideStartY: 10,
        slideMargin: [0, 0, 30, 0],
      },
    ), SLIDE);

    expect(firstColumnText(pages)).toEqual([
      ['H', 'A'],
      ['H', 'B'],
      ['H', 'C'],
    ]);
    expect(pages.map(({ y, height }) => [y, height])).toEqual([
      [20, 50],
      [10, 50],
      [10, 50],
    ]);
    expect(pages.map(({ rowHeights }) => rowHeights)).toEqual([
      [20, 30],
      [20, 30],
      [20, 30],
    ]);
    expect(pages.every(({ autoPage }) => autoPage === undefined)).toBe(true);
    expect(Object.isFrozen(pages)).toBe(true);
    expect(pages.every(Object.isFrozen)).toBe(true);
    expect(pages.every(({ rows, rowHeights }) =>
      Object.isFrozen(rows) && Object.isFrozen(rowHeights))).toBe(true);
  });

  it('uses exact boundaries, keeps no-overflow input on one ordinary page, and omits headers', () => {
    const exact = planTableAutoPages(definition(
      [['A'], ['B'], ['C']],
      {
        y: 20,
        rowHeights: [30, 30, 30],
        autoPageSlideStartY: 0,
        slideMargin: [0, 0, 20, 0],
      },
    ), SLIDE);
    expect(firstColumnText(exact)).toEqual([['A', 'B'], ['C']]);
    expect(exact.map(({ y, height }) => [y, height])).toEqual([[20, 60], [0, 30]]);

    const one = planTableAutoPages(definition(
      [['A'], ['B']],
      { y: 10, rowHeights: [20, 30], slideMargin: 0 },
    ), SLIDE);
    expect(firstColumnText(one)).toEqual([['A', 'B']]);
    expect(one[0]!.height).toBe(50);
  });

  it('allows a header-only source page and never creates a header-only continuation', () => {
    const pages = planTableAutoPages(definition(
      [['H'], ['A'], ['B']],
      {
        y: 65,
        rowHeights: [20, 30, 30],
        autoPageRepeatHeader: true,
        autoPageSlideStartY: 0,
        slideMargin: 0,
      },
    ), SLIDE);
    expect(firstColumnText(pages)).toEqual([['H'], ['H', 'A', 'B']]);
    expect(pages.map(({ height }) => height)).toEqual([20, 80]);
  });

  it('uses explicit margins before layout margins and layout margins before canonical defaults', () => {
    const layoutMargins: TableAutoPageMargins = {
      top: 10,
      right: 20,
      bottom: 30,
      left: 40,
    };
    const layoutPages = planTableAutoPages(definition(
      [['A'], ['B'], ['C']],
      { y: 40, rowHeights: [30, 30, 30] },
    ), SLIDE, layoutMargins);
    expect(layoutPages.map(({ y }) => y)).toEqual([40, 10]);
    expect(firstColumnText(layoutPages)).toEqual([['A'], ['B', 'C']]);

    const explicitPages = planTableAutoPages(definition(
      [['A'], ['B'], ['C']],
      {
        y: 40,
        rowHeights: [30, 30, 30],
        slideMargin: [5, 6, 10, 8],
      },
    ), SLIDE, layoutMargins);
    expect(explicitPages.map(({ y }) => y)).toEqual([40, 5]);
    expect(firstColumnText(explicitPages)).toEqual([['A'], ['B', 'C']]);

    const canonical = planTableAutoPages(definition(
      [['A']],
      { y: 500_000, rowHeights: [100_000] },
    ), { width: emu(12_192_000), height: emu(6_858_000) });
    expect(canonical).toHaveLength(1);
  });

  it('keeps complete vertical and rectangular merge blocks on one page', () => {
    const pages = planTableAutoPages(definition(
      [
        ['H', 'H2'],
        [{ text: 'Merged', options: { rowspan: 2, colspan: 2 } }],
        [],
        ['Tail', 'Tail2'],
      ],
      {
        y: 40,
        rowHeights: [10, 25, 25, 25],
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageSlideStartY: 0,
        slideMargin: 0,
      },
    ), SLIDE);
    expect(firstColumnText(pages)).toEqual([
      ['H', 'Merged', ''],
      ['H', 'Tail'],
    ]);
    expect(pages[0]!.rows[1]![0]!.rowspan).toBe(2);
    expect(pages[0]!.rows[2]![0]!.continuation).toEqual({
      gridSpan: 2,
      vertical: true,
    });
  });

  it('groups adjacent active rowspans into minimal contiguous blocks', () => {
    const pages = planTableAutoPages(definition(
      [
        [{ text: 'A', options: { rowspan: 2 } }, 'A2'],
        ['B2'],
        [{ text: 'C', options: { rowspan: 2 } }, 'C2'],
        ['D2'],
        ['Tail', 'Tail2'],
      ],
      {
        y: 40,
        rowHeights: [20, 20, 20, 20, 20],
        autoPageSlideStartY: 0,
        slideMargin: 0,
      },
    ), SLIDE);
    expect(firstColumnText(pages)).toEqual([
      ['A', ''],
      ['C', '', 'Tail'],
    ]);
  });

  it.each([
    ['header crosses body', () => definition(
      [
        [{ text: 'H', options: { rowspan: 2 } }, 'H2'],
        ['B2'],
      ],
      {
        y: 0,
        rowHeights: [20, 20],
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        slideMargin: 0,
      },
    ), /header/i],
    ['header exceeds first page', () => definition(
      [['H'], ['B']],
      {
        y: 90,
        rowHeights: [20, 20],
        autoPageRepeatHeader: true,
        slideMargin: 0,
      },
    ), /header|source/i],
    ['first block cannot fit', () => definition(
      [['A'], ['B']],
      { y: 80, rowHeights: [30, 20], slideMargin: 0 },
    ), /first|source/i],
    ['block cannot fit continuation', () => definition(
      [['H'], ['A']],
      {
        y: 0,
        rowHeights: [20, 90],
        autoPageRepeatHeader: true,
        autoPageSlideStartY: 0,
        slideMargin: 0,
      },
    ), /continuation|block/i],
    ['automatic content is not materialized', () => definition(
      [['A']],
      { y: 0, slideMargin: 0 },
    ), /materialized/i],
    ['zero continuation area', () => definition(
      [['A']],
      {
        y: 0,
        rowHeights: [10],
        autoPageSlideStartY: 100,
        slideMargin: 0,
      },
    ), /continuation|height/i],
    ['horizontal margins exhaust slide', () => definition(
      [['A']],
      { y: 0, rowHeights: [10], slideMargin: [0, 600, 0, 400] },
    ), /horizontal/i],
    ['vertical margins exhaust slide', () => definition(
      [['A']],
      { y: 0, rowHeights: [10], slideMargin: [50, 0, 50, 0] },
    ), /vertical/i],
  ])('rejects %s without producing a partial plan', (_name, make, message) => {
    expect(() => planTableAutoPages(make(), SLIDE)).toThrow(message);
  });

  it('rejects unsafe layout margins and slide dimensions', () => {
    const source = definition([['A']], { y: 0, rowHeights: [10] });
    expect(() => planTableAutoPages(source, { width: emu(0), height: emu(100) }))
      .toThrow(/slide width/i);
    expect(() => planTableAutoPages(source, SLIDE, {
      top: 0,
      right: 0,
      bottom: -1,
      left: 0,
    })).toThrow(/margin/i);
  });
});
