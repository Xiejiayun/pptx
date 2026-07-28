import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { inches, PptxDocument, ShapeModel, TableModel } from '@pptx/sdk';
import { importPptxGenJS } from './index.js';

interface BorderProps {
  readonly type?: 'none' | 'dash' | 'solid';
  readonly color?: string;
  readonly pt?: number;
}

interface PptxGenJSInstance {
  readonly version: string;
  readonly SchemeColor: {
    readonly accent1: 'accent1';
    readonly accent2: 'accent2';
  };
  layout: string;
  rtlMode: unknown;
  addSlide(): {
    addText(
      text: string | readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
      options: Record<string, unknown>,
    ): void;
    addTable(
      rows: readonly (readonly {
        readonly text?: string;
        readonly options?: Record<string, unknown>;
      }[])[],
      options: Record<string, unknown>,
    ): void;
  };
  write(options: { outputType: 'uint8array'; compression: boolean }): Promise<Uint8Array>;
}

const require = createRequire(import.meta.url);
const PptxGenJS = require('pptxgenjs') as new () => PptxGenJSInstance;

describe('importPptxGenJS', () => {
  it('matches native basic table creation to public PptxGenJS plain-table output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    const rows = [
      ['A1', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
    ] as const;
    generatedSlide.addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1.5, w: 6, h: 2 },
    );

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(rows, {
      x: inches(1),
      y: inches(1.5),
      width: inches(6),
      height: inches(2),
    });

    expect(importedTable).toBeInstanceOf(TableModel);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
    );

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).toContain(
        'uri="http://schemas.openxmlformats.org/drawingml/2006/table"',
      );
      const columnWidths = [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
        .map((match) => Number(match[1]));
      const rowHeights = [...xml.matchAll(/<a:tr h="(\d+)">/g)]
        .map((match) => Number(match[1]));
      expect(columnWidths).toHaveLength(3);
      expect(columnWidths.reduce((sum, width) => sum + width, 0)).toBe(5_486_400);
      expect(rowHeights).toEqual([914_400, 914_400]);
      expect(xml.match(/<a:tc>/g)).toHaveLength(6);
      expect(xml.match(/marL="91440" marR="91440" marT="45720" marB="45720"/g))
        .toHaveLength(6);
      const properties = xml.match(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/)?.[1];
      expect(properties).toBeDefined();
      const left = properties!.indexOf('<a:lnL ');
      const right = properties!.indexOf('<a:lnR ');
      const top = properties!.indexOf('<a:lnT ');
      const bottom = properties!.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }
    expect(nativeXml).not.toContain('p14:modId');
    expect(nativeXml).toContain('<a:ext cx="5486400" cy="1828800"/>');

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNativeTable = reopenedNative.slides[0]!.shapes[0] as TableModel;
    const reopenedImportedTable = reopenedImported.slides[0]!.shapes[0] as TableModel;
    expect(reopenedNativeTable.rows).toEqual(nativeTable.rows);
    expect(reopenedImportedTable.rows).toEqual(importedTable.rows);
    expect(reopenedNativeTable.transform).toEqual(nativeTable.transform);
    expect(reopenedImportedTable.transform).toEqual(importedTable.transform);
  });

  it('imports PptxGenJS table-cell text directions with exact four-value semantics', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        { text: 'Horizontal', options: { textDirection: 'horz' } },
        { text: 'Vertical', options: { textDirection: 'vert' } },
        { text: 'Rotate 270', options: { textDirection: 'vert270' } },
        { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
        { text: 'Invalid', options: { textDirection: 'eaVert' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, textDirection: 'vert270' },
    );
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Explicit horizontal', options: { textDirection: 'horz' } },
      ]],
      { x: 0.5, y: 2, w: 12, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(2);
    expect(tables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
    expect(tables[0]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Inherited',
      'Horizontal',
      'Vertical',
      'Rotate 270',
      'Stacked',
      'Invalid',
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Explicit horizontal',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* vert="vert270"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* vert="vert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="wordArtVert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="eaVert"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:tcPr[^>]* vert="horz"/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(reopenedTables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
  });

  it('imports PptxGenJS table fit-like runtime options as fit-less cells', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Fit none', options: { fit: 'none' } },
        { text: 'Fit shrink', options: { fit: 'shrink' } },
        { text: 'Fit resize', options: { fit: 'resize' } },
        { text: 'Auto fit', options: { autoFit: true } },
        { text: 'Shrink text', options: { shrinkText: true } },
        {
          text: 'Conflicting',
          options: { fit: 'resize', autoFit: true, shrinkText: true, textDirection: 'vert' },
        },
      ]],
      {
        x: 0.5,
        y: 0.5,
        w: 12,
        h: 1,
        fit: 'resize',
        autoFit: true,
        shrinkText: true,
        textDirection: 'vert270',
      },
    );

    const document = await importPptxGenJS(generated);
    const table = document.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(table?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(table?.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Auto fit',
      'Shrink text',
      'Conflicting',
    ]);
    expect(table?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).not.toMatch(/<a:(?:noAutofit|normAutofit|spAutoFit)\b/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);
  });

  it('imports PptxGenJS table-cell vertical alignments from direct cell anchors', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited bottom', options: {} },
        { text: 'Top', options: { valign: 'top' } },
        { text: 'Middle', options: { valign: 'middle' } },
        { text: 'Bottom', options: { valign: 'bottom' } },
        { text: 'Invalid mid', options: { valign: 'mid' } },
        { text: 'Invalid distributed', options: { valign: 'distributed' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, valign: 'bottom' },
    );
    slide.addTable(
      [[{ text: 'Inherited top', options: {} }]],
      { x: 0.5, y: 2, w: 3, h: 1, valign: 'top' },
    );
    slide.addTable(
      [[{ text: 'Inherited middle', options: {} }]],
      { x: 4, y: 2, w: 3, h: 1, valign: 'middle' },
    );
    slide.addTable(
      [[{ text: 'Omitted direct alignment', options: {} }]],
      { x: 7.5, y: 2, w: 3, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(4);
    expect(tables[0]!.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(tables.slice(1).map((table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Inherited bottom',
      'Top',
      'Middle',
      'Bottom',
      'Invalid mid',
      'Invalid distributed',
      'Inherited top',
      'Inherited middle',
      'Omitted direct alignment',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* anchor="t"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="ctr"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="b"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="mid"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* anchor="distributed"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:bodyPr[^>]* anchor=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(reopenedTables.slice(1).map(
      (table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
  });

  it('imports PptxGenJS table-cell margins from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted defaults', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table zero', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, margin: 0 },
    );
    slide.addTable(
      [[{ text: 'Table 0.1 inch', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, margin: 0.1 },
    );
    slide.addTable(
      [[{ text: 'Table inch tuple', options: {} }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, margin: [0.05, 0.1, 0.15, 0.2] },
    );
    slide.addTable(
      [[
        { text: 'Inherited 0.1', options: {} },
        { text: 'Cell zero', options: { margin: 0 } },
        { text: 'Cell quarter inch', options: { margin: 0.25 } },
        { text: 'Cell inch tuple', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
        { text: 'Cell scalar one point', options: { margin: 1 } },
        { text: 'Cell point tuple', options: { margin: [1, 2, 3, 4] } },
        { text: 'Cell negative inch', options: { margin: -0.1 } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, margin: 0.1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins));
    expect(snapshots).toEqual([
      [{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }],
      [{ top: 0, right: 0, bottom: 0, left: 0 }],
      [{ top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 }],
      [{ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 }],
      [
        { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
        { top: 0, right: 0, bottom: 0, left: 0 },
        { top: 18, right: 18, bottom: 18, left: 18 },
        { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
        { top: 1, right: 1, bottom: 1, left: 1 },
        { top: 1, right: 2, bottom: 3, left: 4 },
        { top: -7.2, right: -7.2, bottom: -7.2, left: -7.2 },
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted defaults',
      'Table zero',
      'Table 0.1 inch',
      'Table inch tuple',
      'Inherited 0.1',
      'Cell zero',
      'Cell quarter inch',
      'Cell inch tuple',
      'Cell scalar one point',
      'Cell point tuple',
      'Cell negative inch',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* marL=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marR=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marT=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marB=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr marL="91440" marR="91440" marT="91440" marB="91440">/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr marL="182880" marR="91440" marT="45720" marB="137160">/g)).toHaveLength(2);
    expect(xml).toContain('<a:tcPr marL="12700" marR="12700" marT="12700" marB="12700">');
    expect(xml).toContain('<a:tcPr marL="50800" marR="25400" marT="12700" marB="38100">');
    expect(xml).toContain('<a:tcPr marL="-91440" marR="-91440" marT="-91440" marB="-91440">');
    expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns)=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins))).toEqual(snapshots);
  });

  it('imports PptxGenJS table-cell fills from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted fill', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Type none', options: { fill: { type: 'none' } } }]],
      { x: 0.2, y: 1, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table red', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, fill: { color: 'FF0000' } },
    );
    slide.addTable(
      [[{ text: 'Table theme alpha', options: {} }]],
      {
        x: 0.2,
        y: 2.6,
        w: 2,
        h: 0.5,
        fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      },
    );
    slide.addTable(
      [[
        { text: 'Inherited blue', options: {} },
        { text: 'Cell yellow alpha', options: { fill: { color: 'FFFF00', transparency: 50 } } },
        { text: 'Explicit zero', options: { fill: { color: '00FF00', transparency: 0 } } },
        { text: 'Fractional', options: { fill: { color: '112233', transparency: 33.333 } } },
        { text: 'Full transparency', options: { fill: { color: '445566', transparency: 100 } } },
        { text: 'Deprecated alpha', options: { fill: { color: generated.SchemeColor.accent2, alpha: 25 } } },
        { text: 'Runtime negative', options: { fill: { color: '778899', transparency: -1 } } },
        { text: 'Runtime overflow', options: { fill: { color: 'AABBCC', transparency: 101 } } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, fill: { color: '0000FF' } },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) => table.rows[0]!.cells.map(({ fill }) => fill));
    expect(snapshots).toEqual([
      [undefined],
      [undefined],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } }],
      [{
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      }],
      [
        { kind: 'solid', color: { kind: 'srgb', value: '0000FF' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 33.333,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        undefined,
        undefined,
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted fill',
      'Type none',
      'Table red',
      'Table theme alpha',
      'Inherited blue',
      'Cell yellow alpha',
      'Explicit zero',
      'Fractional',
      'Full transparency',
      'Deprecated alpha',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const directFillXml = [...xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)]
      .map((match) => match[1]!.replace(/<a:ln[LRBT][\s\S]*?<\/a:ln[LRBT]>/g, ''))
      .map((properties) => properties.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0]);
    expect(directFillXml.filter(Boolean)).toHaveLength(10);
    expect(directFillXml[0]).toBeUndefined();
    expect(directFillXml[1]).toBeUndefined();
    expect(xml).toContain('<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="112233"><a:alpha val="66667"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="778899"><a:alpha val="101000"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="AABBCC"><a:alpha val="-1000"/></a:srgbClr></a:solidFill>');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(snapshots);
  });

  it('imports PptxGenJS table-cell borders from materialized direct lines', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    const tableTuple = [
      { type: 'none' },
      { type: 'dash', color: '00FF00', pt: 1.5 },
      { type: 'solid', color: '0000FF', pt: 0 },
      { type: 'solid' },
    ] satisfies [BorderProps, BorderProps, BorderProps, BorderProps];
    const partialCellTuple = [
      undefined,
      { type: 'dash' },
      undefined,
      { type: 'solid' },
    ] as unknown as [BorderProps, BorderProps, BorderProps, BorderProps];

    slide.addTable(
      [[{ text: 'Omitted border', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table scalar', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, border: { type: 'solid', color: 'FF0000', pt: 2 } },
    );
    slide.addTable(
      [[{ text: 'Table tuple', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, border: tableTuple },
    );
    slide.addTable(
      [[{
        text: 'Cell scalar zero',
        options: { border: { type: 'solid', color: 'FFFF00', pt: 0 } },
      }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, border: { type: 'solid', color: 'AAAAAA', pt: 3 } },
    );
    slide.addTable(
      [[{ text: 'Cell partial tuple', options: { border: partialCellTuple } }]],
      { x: 0.2, y: 3.4, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Default border values', options: {} }]],
      { x: 0.2, y: 4.2, w: 2, h: 0.5, border: {} },
    );
    slide.addTable(
      [[
        { text: 'Runtime negative', options: { border: { type: 'solid', color: '778899', pt: -1 } } },
        { text: 'Runtime overflow', options: { border: { type: 'dash', color: 'AABBCC', pt: 2000 } } },
      ]],
      { x: 0.2, y: 5, w: 4, h: 0.5 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(7);
    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({ top: value, right: value, bottom: value, left: value });
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders));
    expect(snapshots).toEqual([
      [four(none)],
      [four(line('FF0000', 2))],
      [{
        top: none,
        right: line('00FF00', 1.5, 'dash'),
        bottom: line('0000FF', 1),
        left: line('666666', 1),
      }],
      [four(line('FFFF00', 0))],
      [{
        top: none,
        right: line('666666', 1, 'dash'),
        bottom: none,
        left: line('666666', 1),
      }],
      [four(line('666666', 1))],
      [undefined, undefined],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted border',
      'Table scalar',
      'Table tuple',
      'Cell scalar zero',
      'Cell partial tuple',
      'Default border values',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).toContain(
      '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>',
    );
    expect(xml).toContain('<a:prstDash val="solid"/>');
    expect(xml).toContain('<a:prstDash val="sysDash"/>');
    expect(xml).toContain('<a:lnB w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>');
    expect(xml).toContain('<a:lnL w="-12700" cap="flat" cmpd="sng" algn="ctr">');
    expect(xml).toContain('<a:lnL w="25400000" cap="flat" cmpd="sng" algn="ctr">');
    for (const properties of xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)) {
      const direct = properties[1]!;
      const left = direct.indexOf('<a:lnL ');
      const right = direct.indexOf('<a:lnR ');
      const top = direct.indexOf('<a:lnT ');
      const bottom = direct.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders))).toEqual(snapshots);
  });

  it('imports public PptxGenJS output and continues editing in the OOXML kernel', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addText('Created by PptxGenJS', { x: 1, y: 1, w: 7, h: 1, align: 'center' });
    generatedSlide.addText(
      [
        {
          text: 'Bold red',
          options: { bold: true, fontFace: 'Aptos', fontSize: 24, color: 'ff0000' },
        },
        {
          text: 'italic',
          options: { italic: true, fontSize: 14, color: '4472C4', softBreakBefore: true },
        },
      ],
      { x: 1, y: 2, w: 7, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Left', options: { align: 'left' } },
        { text: 'Center', options: { align: 'center' } },
        { text: 'Right', options: { align: 'right' } },
        { text: 'Justify', options: { align: 'justify' } },
      ],
      { x: 1, y: 3, w: 7, h: 2, align: 'left' },
    );
    generatedSlide.addText('Standard\nSecond', { x: 1, y: 5, w: 3, h: 1, bullet: true });
    generatedSlide.addText('Custom', {
      x: 4,
      y: 5,
      w: 3,
      h: 1,
      bullet: { characterCode: '25BA', indent: 18 },
    });
    generatedSlide.addText('Public numberType', {
      x: 7,
      y: 5,
      w: 3,
      h: 1,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3, indent: 22 },
    });
    generatedSlide.addText('Deprecated style', {
      x: 10,
      y: 5,
      w: 2,
      h: 1,
      bullet: { type: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
    });
    generatedSlide.addText('Exact first\nExact second', {
      x: 1,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 28,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 6.25,
      paraSpaceAfter: 8.5,
    });
    generatedSlide.addText('Multiple', {
      x: 4,
      y: 6,
      w: 3,
      h: 1,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 4.25,
      paraSpaceAfter: 7.75,
    });
    generatedSlide.addText('Zero spacing', {
      x: 7,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 0,
      lineSpacingMultiple: 0,
      paraSpaceBefore: 0,
      paraSpaceAfter: 0,
    });
    generatedSlide.addText('Level one', {
      x: 10,
      y: 6,
      w: 2,
      h: 0.5,
      bullet: true,
      indentLevel: 1,
    });
    generatedSlide.addText('Custom level two', {
      x: 10,
      y: 6.5,
      w: 2,
      h: 0.5,
      bullet: { characterCode: '25BA', indent: 18 },
      indentLevel: 2,
    });
    generatedSlide.addText('Number level three', {
      x: 10,
      y: 7,
      w: 2,
      h: 0.5,
      bullet: { type: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      indentLevel: 3,
    });
    generatedSlide.addText('No bullet level two', {
      x: 10,
      y: 7.5,
      w: 2,
      h: 0.5,
      indentLevel: 2,
    });
    generatedSlide.addText('Left\tCenter\tRight\tDecimal', {
      x: 1,
      y: 7.25,
      w: 8,
      h: 0.5,
      tabStops: [
        { position: 1 },
        { position: 2.25, alignment: 'ctr' },
        { position: 3.5, alignment: 'r' },
        { position: 4.75, alignment: 'dec' },
      ],
    });
    generatedSlide.addText('Empty tabs', {
      x: 1,
      y: 7.75,
      w: 3,
      h: 0.5,
      tabStops: [],
    });
    generatedSlide.addText(
      [
        { text: 'First\tA', options: { breakLine: true, tabStops: [{ position: 1.5, alignment: 'r' }] } },
        { text: 'Second\tB', options: { tabStops: [{ position: 2.5, alignment: 'ctr' }] } },
      ],
      { x: 5, y: 7.5, w: 4, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Single', options: { underline: true } },
        { text: ' Double', options: { underline: { style: 'dbl', color: 'ff0000' } } },
        { text: ' Wavy', options: { underline: { style: 'wavyDbl' } } },
        { text: ' None', options: { underline: { style: 'none' } } },
        { text: ' Dot dash', options: { underline: { style: 'dotDashHeavy' } } },
      ],
      { x: 9, y: 3, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'True', options: { strike: true } },
        { text: ' False', options: { strike: false } },
        { text: ' Single', options: { strike: 'sngStrike' } },
        { text: ' Double', options: { strike: 'dblStrike' } },
        { text: ' None', options: { strike: 'noStrike' } },
      ],
      { x: 9, y: 4, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Yellow', options: { highlight: 'ffff00' } },
        { text: ' Theme', options: { highlight: 'accent2' } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 5, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { outline: { color: 'ff0000', size: 1.5 } } },
        { text: ' Theme', options: { outline: { color: 'accent1', size: 2 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 6, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { glow: { color: 'ff0000', size: 8, opacity: 0.5 } } },
        { text: ' Theme', options: { glow: { color: 'accent1', size: 2.5, opacity: 1 } } },
        { text: ' Default', options: { glow: { size: 0, opacity: 0 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 7, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Super', options: { superscript: true } },
        { text: ' Sub', options: { subscript: true } },
        { text: ' Custom+', options: { baseline: 600 } },
        { text: ' Custom-', options: { baseline: -800 } },
        { text: ' Fraction', options: { baseline: 1.5 } },
        { text: ' Zero', options: { baseline: 0 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 8, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Positive', options: { charSpacing: 2.5 } },
        { text: ' Negative', options: { charSpacing: -1.25 } },
        { text: ' Fraction', options: { charSpacing: 0.004 } },
        { text: ' Zero', options: { charSpacing: 0 } },
        { text: ' Combined', options: { charSpacing: 3, baseline: 600 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 9, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Inherited', options: {} },
        { text: ' German', options: { lang: 'de-DE' } },
        { text: ' Explicit default', options: { lang: 'en-US' } },
        { text: ' Empty inherits', options: { lang: '' } },
      ],
      { x: 9, y: 10, w: 3, h: 1, lang: 'fr-CA', objectName: 'Language outer' },
    );
    generatedSlide.addText('Omitted margin', {
      x: 0,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Margin omitted',
    });
    generatedSlide.addText('Zero margin', {
      x: 0,
      y: 0.5,
      w: 2,
      h: 0.5,
      margin: 0,
      objectName: 'Margin zero',
    });
    generatedSlide.addText('Scalar margin', {
      x: 0,
      y: 1,
      w: 2,
      h: 0.5,
      margin: 10,
      objectName: 'Margin scalar',
    });
    generatedSlide.addText('Tuple margin', {
      x: 0,
      y: 1.5,
      w: 2,
      h: 0.5,
      margin: [4, 8, 8, 4],
      objectName: 'Margin tuple',
    });
    generatedSlide.addText('Fractional margin', {
      x: 0,
      y: 2,
      w: 2,
      h: 0.5,
      margin: 0.125,
      objectName: 'Margin fractional',
    });
    generatedSlide.addText('Negative margin', {
      x: 0,
      y: 2.5,
      w: 2,
      h: 0.5,
      margin: -0.5,
      objectName: 'Margin negative',
    });
    generatedSlide.addText('Asymmetric probe', {
      x: 0,
      y: 3,
      w: 2,
      h: 0.5,
      margin: [1, 2, 3, 4],
      objectName: 'Margin asymmetric probe',
    });
    generatedSlide.addText('Omitted vertical alignment', {
      x: 2,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Vertical omitted',
    });
    generatedSlide.addText('Top vertical alignment', {
      x: 2,
      y: 0.5,
      w: 2,
      h: 0.5,
      valign: 'top',
      objectName: 'Vertical top',
    });
    generatedSlide.addText('Middle vertical alignment', {
      x: 2,
      y: 1,
      w: 2,
      h: 0.5,
      valign: 'middle',
      objectName: 'Vertical middle',
    });
    generatedSlide.addText('Bottom vertical alignment', {
      x: 2,
      y: 1.5,
      w: 2,
      h: 0.5,
      valign: 'bottom',
      objectName: 'Vertical bottom',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run vertical alignment', options: { valign: 'bottom' } }],
      {
        x: 2,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Vertical run ignored',
      },
    );
    generatedSlide.addText('Omitted text wrapping', {
      x: 4,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Wrap omitted',
    });
    generatedSlide.addText('Enabled text wrapping', {
      x: 4,
      y: 0.5,
      w: 2,
      h: 0.5,
      wrap: true,
      objectName: 'Wrap true',
    });
    generatedSlide.addText('Disabled text wrapping', {
      x: 4,
      y: 1,
      w: 2,
      h: 0.5,
      wrap: false,
      objectName: 'Wrap false',
    });
    generatedSlide.addText('Invalid text wrapping', {
      x: 4,
      y: 1.5,
      w: 2,
      h: 0.5,
      wrap: 'false',
      objectName: 'Wrap invalid fallback',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text wrapping', options: { wrap: false } }],
      {
        x: 4,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Wrap run ignored',
      },
    );
    const textDirections = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    generatedSlide.addText('Omitted text direction', {
      x: 6,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Direction omitted',
    });
    for (const [index, direction] of textDirections.entries()) {
      generatedSlide.addText(direction, {
        x: 6,
        y: (index + 1) * 0.5,
        w: 2,
        h: 0.5,
        vert: direction,
        objectName: `Direction ${direction}`,
      });
    }
    generatedSlide.addText('Invalid text direction', {
      x: 6,
      y: 4,
      w: 2,
      h: 0.5,
      vert: 'vertical',
      objectName: 'Direction invalid passthrough',
    });
    generatedSlide.addText('Ignored textDirection alias', {
      x: 6,
      y: 4.5,
      w: 2,
      h: 0.5,
      textDirection: 'vert',
      objectName: 'Direction alias ignored',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text direction', options: { vert: 'vert270', textDirection: 'vert' } }],
      {
        x: 6,
        y: 5,
        w: 2,
        h: 0.5,
        objectName: 'Direction run ignored',
      },
    );
    generatedSlide.addText('Omitted text fit', {
      x: 8,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Fit omitted',
    });
    generatedSlide.addText('No text fit', {
      x: 8,
      y: 0.5,
      w: 2,
      h: 0.5,
      fit: 'none',
      objectName: 'Fit none',
    });
    generatedSlide.addText('Shrink text fit', {
      x: 8,
      y: 1,
      w: 2,
      h: 0.5,
      fit: 'shrink',
      objectName: 'Fit shrink',
    });
    generatedSlide.addText('Resize text fit', {
      x: 8,
      y: 1.5,
      w: 2,
      h: 0.5,
      fit: 'resize',
      objectName: 'Fit resize',
    });
    generatedSlide.addText('Invalid text fit', {
      x: 8,
      y: 2,
      w: 2,
      h: 0.5,
      fit: 'SHRINK',
      objectName: 'Fit invalid ignored',
    });
    generatedSlide.addText('Legacy shrink text fit', {
      x: 8,
      y: 2.5,
      w: 2,
      h: 0.5,
      shrinkText: true,
      objectName: 'Fit legacy shrink',
    });
    generatedSlide.addText('Legacy resize text fit', {
      x: 8,
      y: 3,
      w: 2,
      h: 0.5,
      autoFit: true,
      objectName: 'Fit legacy resize',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text fit', options: { fit: 'shrink', shrinkText: true, autoFit: true } }],
      {
        x: 8,
        y: 3.5,
        w: 2,
        h: 0.5,
        objectName: 'Fit run ignored',
      },
    );
    generatedSlide.addText('مرحبا\nالعالم', {
      x: 0,
      y: 4,
      w: 2,
      h: 1,
      objectName: 'RTL true',
      rtlMode: true,
    });
    generatedSlide.addText('Explicit false', {
      x: 0,
      y: 5,
      w: 2,
      h: 0.5,
      objectName: 'RTL false',
      rtlMode: false,
    });
    generatedSlide.addText('Omitted', {
      x: 0,
      y: 5.5,
      w: 2,
      h: 0.5,
      objectName: 'RTL omitted',
    });
    generatedSlide.addText(
      [
        { text: 'Run one', options: { rtlMode: true } },
        { text: ' Run two', options: { rtlMode: true } },
      ],
      {
        x: 0,
        y: 6,
        w: 2,
        h: 0.5,
        objectName: 'RTL run probe',
      },
    );
    const document = await importPptxGenJS(generated);
    expect(document.slides[0]?.title.text).toBe('Created by PptxGenJS');
    expect((document.slides[0]!.shapes[0] as ShapeModel).richText[0]!.align).toBe('center');
    const rich = document.slides[0]!.shapes[1] as ShapeModel;
    expect(rich.text).toBe('Bold red\nitalic');
    expect(rich.richText[0]!.runs).toEqual([
      {
        text: 'Bold red',
        style: {
          fontFamily: 'Aptos',
          fontSize: 24,
          lang: 'en-US',
          bold: true,
          color: { kind: 'srgb', value: 'FF0000' },
        },
      },
      {
        text: 'italic',
        softBreakBefore: true,
        style: {
          fontSize: 14,
          lang: 'en-US',
          italic: true,
          color: { kind: 'srgb', value: '4472C4' },
        },
      },
    ]);
    const aligned = document.slides[0]!.shapes[2] as ShapeModel;
    expect(aligned.richText.map(({ align }) => align)).toEqual(['left', 'center', 'right', 'justify']);
    aligned.richText = aligned.richText.map((paragraph, index) => ({
      runs: paragraph.runs,
      ...(index === 3
        ? { align: 'center' as const }
        : paragraph.align
          ? { align: paragraph.align }
          : {}),
    }));
    expect((document.slides[0]!.shapes[3] as ShapeModel).richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect((document.slides[0]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((document.slides[0]!.shapes[5] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 3,
      indent: 22,
    });
    expect((document.slides[0]!.shapes[6] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
      indent: 24,
    });
    expect((document.slides[0]!.shapes[7] as ShapeModel).richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    ]);
    expect((document.slides[0]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((document.slides[0]!.shapes[9] as ShapeModel).richText[0]!.spacing).toBeUndefined();
    expect((document.slides[0]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '•', indent: 27 },
      level: 1,
    });
    expect((document.slides[0]!.shapes[11] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '►', indent: 18 },
      level: 2,
    });
    expect((document.slides[0]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      level: 3,
    });
    const noBulletLevel = (document.slides[0]!.shapes[13] as ShapeModel).richText[0]!;
    expect(noBulletLevel.level).toBe(2);
    expect(noBulletLevel.bullet).toBeUndefined();
    expect((document.slides[0]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((document.slides[0]!.shapes[15] as ShapeModel).richText[0]!.tabStops).toEqual([]);
    expect((document.slides[0]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((document.slides[0]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((document.slides[0]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((document.slides[0]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const spaced = (document.slides[0]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(spaced.map(({ style }) => style?.characterSpacing)).toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(spaced[4]!.style!.baseline).toBe('superscript');
    const languages = (document.slides[0]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(languages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const shapeByName = (name: string): ShapeModel => {
      const shape = document.slides[0]!.shapes.find((candidate) => candidate.name === name);
      expect(shape).toBeInstanceOf(ShapeModel);
      return shape as ShapeModel;
    };
    expect(shapeByName('RTL true').richText.map(({ rtl }) => rtl)).toEqual([true, true]);
    expect(shapeByName('RTL false').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    expect(shapeByName('RTL omitted').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    const rtlRunProbe = shapeByName('RTL run probe');
    expect(rtlRunProbe.text).toBe('Run one Run two');
    expect(rtlRunProbe.richText[0]!.runs.map(({ style }) =>
      (style as Record<string, unknown> | undefined)?.rtlMode)).toEqual([undefined, undefined]);
    expect(shapeByName('Margin omitted').textMargins).toBeUndefined();
    expect(shapeByName('Margin zero').textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(shapeByName('Margin scalar').textMargins).toEqual({
      top: 10,
      right: 10,
      bottom: 10,
      left: 10,
    });
    expect(shapeByName('Margin tuple').textMargins).toEqual({ top: 4, right: 8, bottom: 8, left: 4 });
    expect(shapeByName('Margin fractional').textMargins).toEqual({
      top: 1_588 / 12_700,
      right: 1_588 / 12_700,
      bottom: 1_588 / 12_700,
      left: 1_588 / 12_700,
    });
    expect(shapeByName('Margin negative').textMargins).toEqual({
      top: -0.5,
      right: -0.5,
      bottom: -0.5,
      left: -0.5,
    });
    expect(shapeByName('Margin asymmetric probe').textMargins).toEqual({
      top: 4,
      right: 2,
      bottom: 3,
      left: 1,
    });
    expect([
      'Vertical omitted',
      'Vertical top',
      'Vertical middle',
      'Vertical bottom',
      'Vertical run ignored',
    ].map((name) => shapeByName(name).verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
      'middle',
    ]);
    expect([
      'Wrap omitted',
      'Wrap true',
      'Wrap false',
      'Wrap invalid fallback',
      'Wrap run ignored',
    ].map((name) => shapeByName(name).textWrap)).toEqual([
      true,
      true,
      false,
      true,
      true,
    ]);
    expect([
      'Direction omitted',
      ...textDirections.map((direction) => `Direction ${direction}`),
      'Direction invalid passthrough',
      'Direction alias ignored',
      'Direction run ignored',
    ].map((name) => shapeByName(name).textDirection)).toEqual([
      undefined,
      ...textDirections,
      undefined,
      undefined,
      undefined,
    ]);
    expect([
      'Fit omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Fit invalid ignored',
      'Fit legacy shrink',
      'Fit legacy resize',
      'Fit run ignored',
    ].map((name) => shapeByName(name).textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
      undefined,
      'shrink',
      'resize',
      undefined,
    ]);
    const importedXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(importedXml).toContain('lIns="12700" tIns="50800" rIns="25400" bIns="38100"');
    expect(importedXml).toMatch(/name="Vertical omitted"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical top"[\s\S]*?<a:bodyPr[^>]*anchor="t"/);
    expect(importedXml).toMatch(/name="Vertical middle"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical bottom"[\s\S]*?<a:bodyPr[^>]*anchor="b"/);
    expect(importedXml).toMatch(/name="Vertical run ignored"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Wrap omitted"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap true"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap false"[\s\S]*?<a:bodyPr[^>]*wrap="none"/);
    expect(importedXml).toMatch(/name="Wrap invalid fallback"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap run ignored"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(
      /name="Direction omitted"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    for (const direction of textDirections) {
      expect(importedXml).toMatch(
        new RegExp(`name="Direction ${direction}"[\\s\\S]*?<a:bodyPr[^>]* vert="${direction}"`),
      );
    }
    expect(importedXml).toMatch(
      /name="Direction invalid passthrough"[\s\S]*?<a:bodyPr[^>]* vert="vertical"/,
    );
    expect(importedXml).toMatch(
      /name="Direction alias ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Direction run ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Fit omitted"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit none"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit invalid ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit run ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="fr-CA" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="de-DE" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:endParaRPr lang="fr-CA" dirty="0"\/>/,
    );
    expect(importedXml).toMatch(
      /name="RTL true"[\s\S]*?<a:p><a:pPr rtl="1"[^>]*>[\s\S]*?<a:p><a:pPr rtl="1"/,
    );
    expect(importedXml).toMatch(
      /name="RTL false"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="RTL omitted"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    const rtlRunStart = importedXml.indexOf('name="RTL run probe"');
    const rtlRunEnd = importedXml.indexOf('</p:sp>', rtlRunStart);
    expect(importedXml.slice(rtlRunStart, rtlRunEnd).match(/<a:pPr rtl="1"/g)).toHaveLength(2);
    document.slides[0]!.title.text = 'Edited by the OOXML kernel';
    document.duplicateSlide(0);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ title }) => title.text)).toEqual([
      'Edited by the OOXML kernel',
      'Edited by the OOXML kernel',
    ]);
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).richText[0]!.runs[1]).toMatchObject({
      text: 'italic',
      softBreakBefore: true,
      style: { italic: true, color: { kind: 'srgb', value: '4472C4' } },
    });
    expect((reopened.slides[1]!.shapes[2] as ShapeModel).richText.map(({ align }) => align)).toEqual([
      'left',
      'center',
      'right',
      'center',
    ]);
    expect((reopened.slides[1]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((reopened.slides[1]!.shapes[6] as ShapeModel).richText[0]!.bullet).toMatchObject({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
    });
    expect((reopened.slides[1]!.shapes[7] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 6.25,
      after: 8.5,
      line: { kind: 'exact', points: 28 },
    });
    expect((reopened.slides[1]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((reopened.slides[1]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', indent: 27 },
      level: 1,
    });
    expect((reopened.slides[1]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', indent: 22 },
      level: 3,
    });
    expect((reopened.slides[1]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((reopened.slides[1]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((reopened.slides[1]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((reopened.slides[1]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((reopened.slides[1]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const reopenedSpaced = (reopened.slides[1]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(reopenedSpaced.map(({ style }) => style?.characterSpacing))
      .toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(reopenedSpaced[4]!.style!.baseline).toBe('superscript');
    const reopenedLanguages = (reopened.slides[1]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(reopenedLanguages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const reopenedMargins = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Margin '))
      .map(({ name, textMargins }) => [name, textMargins]);
    expect(reopenedMargins).toEqual([
      ['Margin omitted', undefined],
      ['Margin zero', { top: 0, right: 0, bottom: 0, left: 0 }],
      ['Margin scalar', { top: 10, right: 10, bottom: 10, left: 10 }],
      ['Margin tuple', { top: 4, right: 8, bottom: 8, left: 4 }],
      ['Margin fractional', {
        top: 1_588 / 12_700,
        right: 1_588 / 12_700,
        bottom: 1_588 / 12_700,
        left: 1_588 / 12_700,
      }],
      ['Margin negative', { top: -0.5, right: -0.5, bottom: -0.5, left: -0.5 }],
      ['Margin asymmetric probe', { top: 4, right: 2, bottom: 3, left: 1 }],
    ]);
    const reopenedVerticalAlignment = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Vertical '))
      .map(({ name, verticalAlignment }) => [name, verticalAlignment]);
    expect(reopenedVerticalAlignment).toEqual([
      ['Vertical omitted', 'middle'],
      ['Vertical top', 'top'],
      ['Vertical middle', 'middle'],
      ['Vertical bottom', 'bottom'],
      ['Vertical run ignored', 'middle'],
    ]);
    const reopenedWrapping = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Wrap '))
      .map(({ name, textWrap }) => [name, textWrap]);
    expect(reopenedWrapping).toEqual([
      ['Wrap omitted', true],
      ['Wrap true', true],
      ['Wrap false', false],
      ['Wrap invalid fallback', true],
      ['Wrap run ignored', true],
    ]);
    const reopenedDirections = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Direction '))
      .map(({ name, textDirection }) => [name, textDirection]);
    expect(reopenedDirections).toEqual([
      ['Direction omitted', undefined],
      ...textDirections.map((direction) => [`Direction ${direction}`, direction]),
      ['Direction invalid passthrough', undefined],
      ['Direction alias ignored', undefined],
      ['Direction run ignored', undefined],
    ]);
    const reopenedFit = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Fit '))
      .map(({ name, textFit }) => [name, textFit]);
    expect(reopenedFit).toEqual([
      ['Fit omitted', undefined],
      ['Fit none', undefined],
      ['Fit shrink', 'shrink'],
      ['Fit resize', 'resize'],
      ['Fit invalid ignored', undefined],
      ['Fit legacy shrink', 'shrink'],
      ['Fit legacy resize', 'resize'],
      ['Fit run ignored', undefined],
    ]);
    const reopenedRtl = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('RTL '))
      .map(({ name, richText }) => [name, richText.map(({ rtl }) => rtl)]);
    expect(reopenedRtl).toEqual([
      ['RTL true', [true, true]],
      ['RTL false', [undefined]],
      ['RTL omitted', [undefined]],
      ['RTL run probe', [true]],
    ]);
  }, 30_000);

  it('imports and reopens PptxGenJS rich text transparency from real output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const slide = generated.addSlide();
    slide.addText(
      [
        { text: 'Omitted', options: { color: 'FF0000' } },
        { text: ' Zero', options: { color: '00FF00', transparency: 0 } },
        { text: ' Quarter', options: { color: '0000FF', transparency: 25 } },
        { text: ' Fractional', options: { color: '112233', transparency: 50.5555 } },
        { text: ' Invisible', options: { color: '445566', transparency: 100 } },
        { text: ' Theme', options: { color: 'accent1', transparency: 40 } },
        { text: ' Default', options: { transparency: 60 } },
      ],
      { x: 1, y: 1, w: 10, h: 1, objectName: 'Transparency probe' },
    );

    const document = await importPptxGenJS(generated);
    const shape = document.slides[0]!.shapes.find(({ name }) => name === 'Transparency probe');
    expect(shape).toBeInstanceOf(ShapeModel);
    const runs = (shape as ShapeModel).richText[0]!.runs;
    expect(runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
    expect(runs.map(({ style }) => style?.color)).toEqual([
      { kind: 'srgb', value: 'FF0000' },
      { kind: 'srgb', value: '00FF00' },
      { kind: 'srgb', value: '0000FF' },
      { kind: 'srgb', value: '112233' },
      { kind: 'srgb', value: '445566' },
      { kind: 'scheme', value: 'accent1' },
      { kind: 'srgb', value: '000000' },
    ]);
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/<a:alpha val="\d+"\/>/g)).toEqual([
      '<a:alpha val="75000"/>',
      '<a:alpha val="49445"/>',
      '<a:alpha val="0"/>',
      '<a:alpha val="60000"/>',
      '<a:alpha val="40000"/>',
    ]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Transparency probe',
    ) as ShapeModel;
    expect(reopenedShape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
  });

  it('imports only direct PptxGenJS presentation RTL and reopens it', async () => {
    const cases: readonly [string, boolean, unknown, boolean | undefined][] = [
      ['omitted', false, undefined, undefined],
      ['true', true, true, true],
      ['false', true, false, undefined],
      ['truthy', true, 'yes', true],
    ];
    for (const [name, assign, value, expected] of cases) {
      const generated = new PptxGenJS();
      if (assign) generated.rtlMode = value;
      generated.addSlide();
      const document = await importPptxGenJS(generated);
      const journal = [...document.opcPackage.mutations];

      expect(document.rtlMode, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);
      const presentationXml = new TextDecoder().decode(
        document.opcPackage.requirePart(document.presentationPartUri).bytes,
      );
      if (expected === true) {
        expect(presentationXml, name).toMatch(/<p:presentation\b[^>]*\srtl="1"/);
      } else {
        expect(presentationXml, name).not.toMatch(/<p:presentation\b[^>]*\srtl=/);
      }
      expect(presentationXml, name).toMatch(/<a:lvl1pPr\b[^>]*\srtl="0"/);

      if (name === 'true') {
        const reopened = await PptxDocument.open(await document.write());
        expect(reopened.rtlMode).toBe(true);
      }
    }
  }, 20_000);

  it('imports PptxGenJS non-list zero margins and indents without aliasing bullet indentation', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    slide.addText('Plain', { name: 'Margin plain', x: 1, y: 1, w: 3, h: 0.5 });
    slide.addText([{ text: 'Rich' }], { name: 'Margin rich', x: 1, y: 2, w: 3, h: 0.5 });
    slide.addText('Bullet', {
      name: 'Margin bullet',
      x: 1,
      y: 3,
      w: 3,
      h: 0.5,
      bullet: true,
    });
    slide.addText('Number', {
      name: 'Margin number',
      x: 1,
      y: 4,
      w: 3,
      h: 0.5,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 1, indent: 22 },
    });
    const document = await importPptxGenJS(generated);
    const shapes = document.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);

    expect(shapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes[2]!.richText[0]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(shapes[3]!.richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 1,
      indent: 22,
    });
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/indent="0" marL="0"/g)).toHaveLength(2);
    expect(slideXml).toMatch(/marL="342900" indent="-342900"/);
    expect(slideXml).toMatch(/marL="279400" indent="-279400"/);
    expect(slideXml).not.toContain('marR=');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShapes = reopened.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
  }, 20_000);

  it('keeps pptxgenjs out of every non-adapter package dependency list', async () => {
    const packagesDirectory = fileURLToPath(new URL('../..', import.meta.url));
    const packageNames = ['lossless-xml', 'model', 'opc', 'sdk', 'validator'];
    for (const packageName of packageNames) {
      const manifest = JSON.parse(await readFile(`${packagesDirectory}/${packageName}/package.json`, 'utf8')) as {
        dependencies?: Record<string, string>;
      };
      expect(manifest.dependencies?.pptxgenjs, packageName).toBeUndefined();
    }
  });
});
