import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  normalizeTableRowHeightInput,
  readTableRowHeights,
  replaceTableRowHeights,
} from './table-row-heights.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';
const DEFAULT_ROWS =
  '<a:tr h="0"><a:tc rowSpan="2"><a:tcPr/></a:tc></a:tr>' +
  '<a:tr h="914400"><a:tc hMerge="1"><a:tcPr/></a:tc></a:tr>' +
  '<a:tr h="1828800"><a:tc vMerge="1"><a:tcPr/></a:tc></a:tr>';
const DEFAULT_TRANSFORM =
  '<p:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="2743200"/></p:xfrm>';

function parseFrame(
  rows = DEFAULT_ROWS,
  transform = DEFAULT_TRANSFORM,
): { xml: LosslessXmlDocument; frame: XmlElement } {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
      transform +
      '<a:graphic><a:graphicData><a:tbl>' +
      '<a:tblGrid><a:gridCol w="2743200"/></a:tblGrid>' +
      rows +
      '</a:tbl></a:graphicData></a:graphic>' +
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

describe('table row heights', () => {
  it('reads only complete unique direct rows into detached snapshots', () => {
    const valid = parseFrame(
      '<a:tr h="000" keep="R1"><a:tc rowSpan="2"><a:tcPr/></a:tc></a:tr>' +
        '<x:opaque xmlns:x="urn:test"><a:tr h="1"/></x:opaque>' +
        '<a:tr h="0914400" keep="R2"><a:tc hMerge="1" vMerge="1"><a:tcPr/></a:tc></a:tr>',
      '',
    );
    expect(readTableRowHeights(valid.xml, valid.frame)).toEqual([0, 914_400]);
    const snapshot = readTableRowHeights(valid.xml, valid.frame) as number[];
    snapshot[0] = 1;
    expect(readTableRowHeights(valid.xml, valid.frame)).toEqual([0, 914_400]);
    expect(valid.xml.changed).toBe(false);

    const alternate = parseSource(
      '<p:graphicFrame xmlns:p="p" xmlns:q="a">' +
        '<q:graphic><q:graphicData><q:tbl>' +
        '<q:tr h="0"/><q:tr h="22"/>' +
        '</q:tbl></q:graphicData></q:graphic></p:graphicFrame>',
    );
    expect(readTableRowHeights(alternate.xml, alternate.frame)).toEqual([0, 22]);
  });

  it('returns undefined for malformed rows or ambiguous direct table paths', () => {
    const malformedRows = [
      '',
      '<a:tr/>',
      '<a:tr h=""/>',
      '<a:tr h="-1"/>',
      '<a:tr h="+1"/>',
      '<a:tr h="1.5"/>',
      '<a:tr h="1e3"/>',
      '<a:tr h="9007199254740992"/>',
      '<a:tr x:h="914400" xmlns:x="x"/>',
      '<a:tr h="914400" h="1828800"/>',
      '<x:keep xmlns:x="urn:test"><a:tr h="914400"/></x:keep>',
    ];
    for (const rows of malformedRows) {
      const { xml, frame } = parseFrame(rows);
      expect(readTableRowHeights(xml, frame), rows).toBeUndefined();
      expect(xml.changed).toBe(false);
    }

    const ambiguousPaths = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic/><a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_ROWS +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic><a:graphicData/><a:graphicData><a:tbl>' +
        DEFAULT_ROWS +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_ROWS +
        '</a:tbl><a:tbl>' +
        DEFAULT_ROWS +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
        '<x:keep xmlns:x="urn:test"><a:graphic><a:graphicData><a:tbl>' +
        DEFAULT_ROWS +
        '</a:tbl></a:graphicData></a:graphic></x:keep></p:graphicFrame>',
    ];
    for (const source of ambiguousPaths) {
      const { xml, frame } = parseSource(source);
      expect(readTableRowHeights(xml, frame), source).toBeUndefined();
    }
  });

  it('normalizes scalar and array inputs without retaining or invoking them', () => {
    expect(normalizeTableRowHeightInput(0)).toEqual({
      kind: 'scalar',
      value: 0,
    });
    expect(normalizeTableRowHeightInput(914_400.4)).toEqual({
      kind: 'scalar',
      value: 914_400,
    });
    const input = [0.4, 914_400.6];
    const normalized = normalizeTableRowHeightInput(input);
    expect(normalized).toEqual({
      kind: 'array',
      values: [0, 914_401],
    });
    input[0] = 1;
    expect(normalized).toEqual({
      kind: 'array',
      values: [0, 914_401],
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
    expect(() => normalizeTableRowHeightInput(accessor)).toThrow(TypeError);
    expect(calls).toBe(0);
  });

  it('rejects invalid scalar and descriptor-unsafe array inputs', () => {
    const hole = new Array(2);
    hole[1] = 1;
    const extra = [1];
    Object.defineProperty(extra, 'extra', { value: true, enumerable: true });
    const symbol = [1];
    Object.defineProperty(symbol, Symbol('height'), { value: true });
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
      -1,
      -0.4,
      Number.MAX_SAFE_INTEGER + 1,
      [1, -1],
      [1, Number.NaN],
    ];
    for (const value of invalid) {
      expect(() => normalizeTableRowHeightInput(value), String(value)).toThrow();
    }
  });

  it('patches only changed row and transform tokens while preserving opaque XML', () => {
    const { xml, frame } = parseFrame(
      '<a:tr h="0457200" keep="R1"><a:tc rowSpan="2"><a:tcPr/></a:tc></a:tr>' +
        '<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>' +
        '<a:tr h="914400" keep="R2"><a:tc hMerge="1"><a:tcPr/></a:tc></a:tr>' +
        '<a:tr h="1828800" keep="R3"><a:tc vMerge="1"><a:tcPr/></a:tc></a:tr>',
      '<p:xfrm keep="XFRM"><a:off x="0" y="0"/>' +
        '<a:ext cx="2743200" cy="2743200" keep="EXT"/></p:xfrm>',
    );
    expect(replaceTableRowHeights(
      xml,
      frame,
      normalizeTableRowHeightInput([457_200, 1_143_000, 594_360]),
      PART_URI,
    )).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain('<a:tr h="0457200" keep="R1">');
    expect(updated).toContain('<a:tr h="1143000" keep="R2">');
    expect(updated).toContain('<a:tr h="594360" keep="R3">');
    expect(updated).toContain(
      '<a:ext cx="2743200" cy="2194560" keep="EXT"/>',
    );
    expect(updated).toContain(
      '<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>',
    );
    expect(updated).toContain('rowSpan="2"');
    expect(updated).toContain('hMerge="1"');
    expect(updated).toContain('vMerge="1"');
  });

  it('broadcasts scalars and separates explicit from automatic transform behavior', () => {
    const scalar = parseFrame();
    expect(replaceTableRowHeights(
      scalar.xml,
      scalar.frame,
      normalizeTableRowHeightInput(1_143_000),
      PART_URI,
    )).toBe(true);
    const reparsedScalar = parseSource(scalar.xml.serialize());
    expect(readTableRowHeights(
      reparsedScalar.xml,
      reparsedScalar.frame,
    )).toEqual([1_143_000, 1_143_000, 1_143_000]);
    expect(scalar.xml.serialize()).toContain('cy="3429000"');

    const automatic = parseFrame();
    expect(replaceTableRowHeights(
      automatic.xml,
      automatic.frame,
      normalizeTableRowHeightInput(0),
      PART_URI,
    )).toBe(true);
    expect(automatic.xml.serialize().match(/<a:tr h="0"/g)).toHaveLength(3);
    expect(automatic.xml.serialize()).toContain('cy="2743200"');

    const mismatch = parseFrame(
      '<a:tr h="914400"/><a:tr h="1828800"/>',
      DEFAULT_TRANSFORM.replace('cy="2743200"', 'cy="0"'),
    );
    expect(replaceTableRowHeights(
      mismatch.xml,
      mismatch.frame,
      normalizeTableRowHeightInput([914_400, 1_828_800]),
      PART_URI,
    )).toBe(true);
    expect(mismatch.xml.serialize()).toContain('cy="2743200"');

    const mixed = parseFrame(
      '<a:tr h="914400"/><a:tr h="1828800"/>',
      DEFAULT_TRANSFORM.replace('cy="2743200"', 'cy="02743200"'),
    );
    expect(replaceTableRowHeights(
      mixed.xml,
      mixed.frame,
      normalizeTableRowHeightInput([0, Number.MAX_SAFE_INTEGER]),
      PART_URI,
    )).toBe(true);
    expect(mixed.xml.serialize()).toContain(
      '<a:ext cx="2743200" cy="02743200"/>',
    );
    expect(mixed.xml.serialize()).toContain(
      '<a:tr h="0"/><a:tr h="9007199254740991"/>',
    );

    const noOp = parseFrame(
      '<a:tr h="000"/><a:tr h="0914400"/>',
      DEFAULT_TRANSFORM.replace('cy="2743200"', 'cy="02743200"'),
    );
    const original = noOp.xml.source;
    expect(replaceTableRowHeights(
      noOp.xml,
      noOp.frame,
      normalizeTableRowHeightInput([0, 914_400]),
      PART_URI,
    )).toBe(false);
    expect(noOp.xml.changed).toBe(false);
    expect(noOp.xml.serialize()).toBe(original);
  });

  it('rejects wrong lengths, explicit overflow, and malformed rows or transform before patching', () => {
    const valid = parseFrame();
    for (const input of [[1], [1, 2], [1, 2, 3, 4]]) {
      expect(() => replaceTableRowHeights(
        valid.xml,
        valid.frame,
        normalizeTableRowHeightInput(input),
        PART_URI,
      )).toThrow(TypeError);
    }
    expect(() => replaceTableRowHeights(
      valid.xml,
      valid.frame,
      normalizeTableRowHeightInput(Number.MAX_SAFE_INTEGER),
      PART_URI,
    )).toThrow(RangeError);
    expect(() => replaceTableRowHeights(
      valid.xml,
      valid.frame,
      normalizeTableRowHeightInput([Number.MAX_SAFE_INTEGER, 1, 1]),
      PART_URI,
    )).toThrow(RangeError);
    expect(valid.xml.changed).toBe(false);

    const malformedRows = [
      '',
      '<a:tr/>',
      '<a:tr h="1" h="2"/>',
    ];
    for (const rows of malformedRows) {
      const candidate = parseFrame(rows);
      expect(() => replaceTableRowHeights(
        candidate.xml,
        candidate.frame,
        normalizeTableRowHeightInput(1),
        PART_URI,
      )).toThrow(ModelParseError);
      expect(candidate.xml.changed).toBe(false);
    }

    const malformedTransforms = [
      '',
      DEFAULT_TRANSFORM + DEFAULT_TRANSFORM,
      '<x:keep xmlns:x="x">' + DEFAULT_TRANSFORM + '</x:keep>',
      '<p:xfrm/>',
      '<p:xfrm><x:keep xmlns:x="x"><a:ext cy="1"/></x:keep></p:xfrm>',
      '<p:xfrm><a:ext cy="1"/><a:ext cy="1"/></p:xfrm>',
      '<p:xfrm><a:ext/></p:xfrm>',
      '<p:xfrm><a:ext cy="-1"/></p:xfrm>',
      '<p:xfrm><a:ext cy="+1"/></p:xfrm>',
      '<p:xfrm><a:ext cy="1.5"/></p:xfrm>',
      '<p:xfrm><a:ext cy="1e3"/></p:xfrm>',
      '<p:xfrm><a:ext cy="9007199254740992"/></p:xfrm>',
      '<p:xfrm><a:ext cy="1" cy="2"/></p:xfrm>',
      '<p:xfrm xmlns:x="x"><a:ext x:cy="1"/></p:xfrm>',
    ];
    for (const transform of malformedTransforms) {
      const candidate = parseFrame(DEFAULT_ROWS, transform);
      let thrown: unknown;
      try {
        replaceTableRowHeights(
          candidate.xml,
          candidate.frame,
          normalizeTableRowHeightInput(1),
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
