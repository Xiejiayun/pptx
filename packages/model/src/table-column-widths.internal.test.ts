import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  normalizeTableColumnWidthInput,
  readTableColumnWidths,
  replaceTableColumnWidths,
} from './table-column-widths.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';
const DEFAULT_GRID =
  '<a:tblGrid><a:gridCol w="914400"/><a:gridCol w="1828800"/></a:tblGrid>';
const DEFAULT_TRANSFORM =
  '<p:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></p:xfrm>';

function parseFrame(
  grid = DEFAULT_GRID,
  transform = DEFAULT_TRANSFORM,
): { xml: LosslessXmlDocument; frame: XmlElement } {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
      transform +
      '<a:graphic><a:graphicData><a:tbl>' +
      grid +
      '<a:tr h="914400">' +
      '<a:tc gridSpan="2"><a:tcPr/></a:tc>' +
      '<a:tc hMerge="1" vMerge="1"><a:tcPr/></a:tc>' +
      '</a:tr></a:tbl></a:graphicData></a:graphic>' +
      '</p:graphicFrame>',
  );
}

function parseSource(
  source: string,
): { xml: LosslessXmlDocument; frame: XmlElement } {
  const xml = LosslessXmlDocument.parse(source);
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no root frame');
  return { xml, frame };
}

describe('table column widths', () => {
  it('reads only a complete unique direct grid into detached snapshots', () => {
    const valid = parseFrame(
      '<a:tblGrid keep="GRID">' +
        '<a:gridCol w="0914400" keep="A"/>' +
        '<x:opaque xmlns:x="urn:test"><a:gridCol w="1"/></x:opaque>' +
        '<a:gridCol w="1828800" keep="B"/>' +
        '</a:tblGrid>',
    );
    expect(readTableColumnWidths(valid.xml, valid.frame)).toEqual([
      914_400,
      1_828_800,
    ]);
    const snapshot = readTableColumnWidths(valid.xml, valid.frame) as number[];
    snapshot[0] = 1;
    expect(readTableColumnWidths(valid.xml, valid.frame)).toEqual([
      914_400,
      1_828_800,
    ]);
    expect(valid.xml.changed).toBe(false);

    const alternate = parseSource(
      '<p:graphicFrame xmlns:p="p" xmlns:q="a">' +
        DEFAULT_TRANSFORM +
        '<q:graphic><q:graphicData><q:tbl>' +
        '<q:tblGrid><q:gridCol w="11"/><q:gridCol w="22"/></q:tblGrid>' +
        '</q:tbl></q:graphicData></q:graphic></p:graphicFrame>',
    );
    expect(readTableColumnWidths(alternate.xml, alternate.frame)).toEqual([11, 22]);
  });

  it('returns undefined for malformed or ambiguous direct grid paths', () => {
    const malformedGrids = [
      '',
      '<a:tblGrid/>',
      '<a:tblGrid><a:gridCol/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w=""/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="0"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="-1"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="+1"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="1.5"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="1e3"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="9007199254740992"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol x:w="914400" xmlns:x="x"/></a:tblGrid>',
      '<a:tblGrid><a:gridCol w="914400" w="1828800"/></a:tblGrid>',
      '<a:tblGrid><x:keep xmlns:x="urn:test"><a:gridCol w="914400"/></x:keep></a:tblGrid>',
      DEFAULT_GRID + DEFAULT_GRID,
    ];
    for (const grid of malformedGrids) {
      const { xml, frame } = parseFrame(grid);
      expect(readTableColumnWidths(xml, frame), grid).toBeUndefined();
      expect(xml.changed).toBe(false);
    }

    const ambiguousPaths = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic/><a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_GRID +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic><a:graphicData/><a:graphicData><a:tbl>' +
        DEFAULT_GRID +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_GRID +
        '</a:tbl><a:tbl>' +
        DEFAULT_GRID +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<x:keep xmlns:x="urn:test"><a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_GRID +
        '</a:tbl></a:graphicData></a:graphic></x:keep></p:graphicFrame>',
    ];
    for (const source of ambiguousPaths) {
      const { xml, frame } = parseSource(source);
      expect(readTableColumnWidths(xml, frame), source).toBeUndefined();
    }
  });

  it('normalizes scalar and array inputs without retaining or invoking them', () => {
    expect(normalizeTableColumnWidthInput(914_400.4)).toEqual({
      kind: 'scalar',
      value: 914_400,
    });
    const input = [914_400.4, 1_828_800.6];
    expect(normalizeTableColumnWidthInput(input)).toEqual({
      kind: 'array',
      values: [914_400, 1_828_801],
      sum: 2_743_201,
    });
    input[0] = 1;
    expect(normalizeTableColumnWidthInput(input)).toEqual({
      kind: 'array',
      values: [1, 1_828_801],
      sum: 1_828_802,
    });

    const accessor = [1];
    let calls = 0;
    Object.defineProperty(accessor, '0', {
      get() {
        calls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    expect(() => normalizeTableColumnWidthInput(accessor)).toThrow(TypeError);
    expect(calls).toBe(0);
  });

  it('rejects invalid scalar and descriptor-unsafe array inputs', () => {
    const hole = new Array(2);
    hole[1] = 1;
    const extra = [1];
    Object.defineProperty(extra, 'extra', { value: true, enumerable: true });
    const symbol = [1];
    Object.defineProperty(symbol, Symbol('width'), { value: true });
    const invalid = [
      undefined,
      null,
      '1',
      true,
      {},
      new Uint32Array([1]),
      [],
      hole,
      extra,
      symbol,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      0,
      -1,
      Number.MAX_SAFE_INTEGER + 1,
      [1, 0],
      [1, Number.NaN],
      [Number.MAX_SAFE_INTEGER, 1],
    ];
    for (const value of invalid) {
      expect(() => normalizeTableColumnWidthInput(value), String(value)).toThrow();
    }
  });

  it('patches only changed width and transform tokens while preserving opaque XML', () => {
    const { xml, frame } = parseFrame(
      '<a:tblGrid keep="GRID">' +
        '<a:gridCol w="0914400" keep="A"/>' +
        '<x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
        '<a:gridCol w="1828800" keep="B"/>' +
        '</a:tblGrid>',
      '<p:xfrm keep="XFRM"><a:off x="0" y="0"/>' +
        '<a:ext cx="2743200" cy="914400" keep="EXT"/></p:xfrm>',
    );
    expect(replaceTableColumnWidths(
      xml,
      frame,
      normalizeTableColumnWidthInput([914_400, 2_743_200]),
      PART_URI,
    )).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain(
      '<a:gridCol w="0914400" keep="A"/>' +
        '<x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
        '<a:gridCol w="2743200" keep="B"/>',
    );
    expect(updated).toContain(
      '<p:xfrm keep="XFRM"><a:off x="0" y="0"/>' +
        '<a:ext cx="3657600" cy="914400" keep="EXT"/></p:xfrm>',
    );
    expect(updated).toContain('gridSpan="2"');
    expect(updated).toContain('hMerge="1" vMerge="1"');
  });

  it('broadcasts scalars, repairs mismatches, and preserves semantic no-ops', () => {
    const scalar = parseFrame();
    expect(replaceTableColumnWidths(
      scalar.xml,
      scalar.frame,
      normalizeTableColumnWidthInput(1_143_000),
      PART_URI,
    )).toBe(true);
    const reparsedScalar = parseSource(scalar.xml.serialize());
    expect(readTableColumnWidths(
      reparsedScalar.xml,
      reparsedScalar.frame,
    )).toEqual([
      1_143_000,
      1_143_000,
    ]);
    expect(scalar.xml.serialize()).toContain('cx="2286000"');

    const mismatch = parseFrame(DEFAULT_GRID, DEFAULT_TRANSFORM.replace('2743200', '0'));
    expect(replaceTableColumnWidths(
      mismatch.xml,
      mismatch.frame,
      normalizeTableColumnWidthInput([914_400, 1_828_800]),
      PART_URI,
    )).toBe(true);
    expect(mismatch.xml.serialize()).toContain('cx="2743200"');

    const noOp = parseFrame(
      '<a:tblGrid><a:gridCol w="0914400"/><a:gridCol w="01828800"/></a:tblGrid>',
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="02743200" cy="914400"/></p:xfrm>',
    );
    const original = noOp.xml.source;
    expect(replaceTableColumnWidths(
      noOp.xml,
      noOp.frame,
      normalizeTableColumnWidthInput([914_400, 1_828_800]),
      PART_URI,
    )).toBe(false);
    expect(noOp.xml.changed).toBe(false);
    expect(noOp.xml.serialize()).toBe(original);
  });

  it('rejects wrong lengths, overflow, and malformed grid or transform before patching', () => {
    const valid = parseFrame();
    for (const input of [[1], [1, 2, 3]]) {
      expect(() => replaceTableColumnWidths(
        valid.xml,
        valid.frame,
        normalizeTableColumnWidthInput(input),
        PART_URI,
      )).toThrow(TypeError);
    }
    expect(() => replaceTableColumnWidths(
      valid.xml,
      valid.frame,
      normalizeTableColumnWidthInput(Number.MAX_SAFE_INTEGER),
      PART_URI,
    )).toThrow(RangeError);
    expect(valid.xml.changed).toBe(false);

    const malformedGrids = [
      '',
      '<a:tblGrid/>',
      '<a:tblGrid><a:gridCol/></a:tblGrid>',
      DEFAULT_GRID + DEFAULT_GRID,
    ];
    for (const grid of malformedGrids) {
      const candidate = parseFrame(grid);
      expect(() => replaceTableColumnWidths(
        candidate.xml,
        candidate.frame,
        normalizeTableColumnWidthInput(1),
        PART_URI,
      )).toThrow(ModelParseError);
      expect(candidate.xml.changed).toBe(false);
    }

    const malformedTransforms = [
      '',
      DEFAULT_TRANSFORM + DEFAULT_TRANSFORM,
      '<x:keep xmlns:x="x">' + DEFAULT_TRANSFORM + '</x:keep>',
      '<p:xfrm/>',
      '<p:xfrm><x:keep xmlns:x="x"><a:ext cx="1"/></x:keep></p:xfrm>',
      '<p:xfrm><a:ext cx="1"/><a:ext cx="1"/></p:xfrm>',
      '<p:xfrm><a:ext/></p:xfrm>',
      '<p:xfrm><a:ext cx="-1"/></p:xfrm>',
      '<p:xfrm><a:ext cx="1.5"/></p:xfrm>',
      '<p:xfrm><a:ext cx="1e3"/></p:xfrm>',
      '<p:xfrm><a:ext cx="9007199254740992"/></p:xfrm>',
      '<p:xfrm><a:ext cx="1" cx="2"/></p:xfrm>',
      '<p:xfrm xmlns:x="x"><a:ext x:cx="1"/></p:xfrm>',
    ];
    for (const transform of malformedTransforms) {
      const candidate = parseFrame(DEFAULT_GRID, transform);
      let thrown: unknown;
      try {
        replaceTableColumnWidths(
          candidate.xml,
          candidate.frame,
          normalizeTableColumnWidthInput(1),
          PART_URI,
        );
      } catch (error) {
        thrown = error;
      }
      expect(thrown, transform).toBeInstanceOf(ModelParseError);
      expect(thrown).toMatchObject({ partUri: PART_URI });
      expect(candidate.xml.changed).toBe(false);
    }
  });
});
