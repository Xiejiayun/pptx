import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const tarball = resolve(process.argv[2] ?? '');
if (!tarball.endsWith('.tgz')) throw new Error('Usage: node scripts/smoke-npm-package.mjs <package.tgz>');
const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const directory = await mkdtemp(join(tmpdir(), 'jiayunxie-pptx-smoke-'));
try {
  await writeFile(join(directory, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  run('npm', [
    'install',
    '--ignore-scripts',
    '--no-audit',
    '--no-fund',
    '--package-lock=false',
    '--cache',
    join(directory, '.npm-cache'),
    tarball,
  ], directory);

  const installed = join(directory, 'node_modules', '@jiayunxie', 'pptx');
  const manifest = JSON.parse(await readFile(join(installed, 'package.json'), 'utf8'));
  if (manifest.name !== '@jiayunxie/pptx' || manifest.version !== '0.1.0') {
    throw new Error(`Unexpected package identity: ${manifest.name}@${manifest.version}`);
  }
  if (JSON.stringify(manifest).includes('workspace:')) throw new Error('Packed manifest contains workspace protocol');
  if (manifest.exports?.['.']?.browser !== './dist/browser.js') {
    throw new Error('Packed manifest is missing the browser conditional export');
  }
  for (const section of ['dependencies', 'optionalDependencies', 'peerDependencies']) {
    for (const name of Object.keys(manifest[section] ?? {})) {
      if (name.startsWith('@pptx/')) throw new Error(`Packed manifest contains internal runtime dependency: ${name}`);
    }
  }

  const distFiles = (await listFiles(join(installed, 'dist'))).filter((name) => /\.(?:js|d\.ts)$/.test(name));
  for (const file of distFiles) {
    const source = await readFile(file, 'utf8');
    if (source.includes('@pptx/')) throw new Error(`Bundled output contains internal import: ${file}`);
  }
  const browserSource = await readFile(join(installed, 'dist/browser.js'), 'utf8');
  if (/\b(?:from|import)\s*['"]node:/.test(browserSource)) {
    throw new Error('Browser bundle contains a static node: import');
  }

  await writeFile(
    join(directory, 'smoke.mjs'),
    `import { inches, PRESET_SHAPE_TYPES, PptxDocument, ShapeModel, TableModel, GradientCodec, importPptxGenJS, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const created = PptxDocument.create({ rtlMode: true });
const shapeDeck = PptxDocument.create();
const shapeSlide = shapeDeck.addSlide();
const defaultShape = shapeSlide.addShape('rect');
const customShape = shapeSlide.addShape('foldedCorner', {
  name: 'Packed folded corner',
  x: inches(2),
  y: inches(3),
  width: inches(4),
  height: inches(2),
  rotation: 2_700_000,
  flipHorizontal: true,
  flipVertical: true,
});
const initialDefaultPreset = defaultShape.presetType;
const initialCustomPreset = customShape.presetType;
const shapeIdentity = shapeSlide.shapes[0] === defaultShape && shapeSlide.shapes[1] === customShape;
const shapePartCountBeforeEdit = shapeDeck.opcPackage.parts.length;
const shapeRelationshipCountBeforeEdit = shapeSlide.relationships.length;
defaultShape.presetType = 'ellipse';
customShape.presetType = 'roundRect';
const shapeEditIsolation = shapeDeck.opcPackage.parts.length === shapePartCountBeforeEdit &&
  shapeSlide.relationships.length === shapeRelationshipCountBeforeEdit;
const duplicateShapeSlide = shapeDeck.duplicateSlide(0);
const duplicateDefaultShape = duplicateShapeSlide.shapes[0];
if (!(duplicateDefaultShape instanceof ShapeModel)) throw new Error('Packed duplicate shape failed');
duplicateDefaultShape.presetType = 'star5';
const reopenedShapeDeck = await PptxDocument.open(await shapeDeck.write());
const reopenedShapeTypes = reopenedShapeDeck.slides.map((slide) =>
  slide.shapes.map((shape) => shape instanceof ShapeModel ? shape.presetType : undefined));
const reopenedShapeNames = reopenedShapeDeck.slides.map((slide) =>
  slide.shapes.map(({ name }) => name));
const presetShapes = PRESET_SHAPE_TYPES.length === 178 &&
  Object.isFrozen(PRESET_SHAPE_TYPES) &&
  PRESET_SHAPE_TYPES.includes('foldedCorner') &&
  !PRESET_SHAPE_TYPES.includes('folderCorner') &&
  defaultShape instanceof ShapeModel &&
  customShape instanceof ShapeModel &&
  initialDefaultPreset === 'rect' &&
  initialCustomPreset === 'foldedCorner' &&
  shapeIdentity &&
  shapeEditIsolation &&
  defaultShape.transform.x === inches(1) &&
  defaultShape.transform.y === inches(1) &&
  defaultShape.transform.width === inches(1) &&
  defaultShape.transform.height === inches(1) &&
  defaultShape.presetType === 'ellipse' &&
  customShape.name === 'Packed folded corner' &&
  customShape.transform.x === inches(2) &&
  customShape.transform.y === inches(3) &&
  customShape.transform.width === inches(4) &&
  customShape.transform.height === inches(2) &&
  customShape.transform.rotation === 2_700_000 &&
  customShape.transform.flipHorizontal === true &&
  customShape.transform.flipVertical === true &&
  customShape.presetType === 'roundRect' &&
  duplicateDefaultShape.presetType === 'star5' &&
  JSON.stringify(reopenedShapeTypes) === JSON.stringify([
    ['ellipse', 'roundRect'],
    ['star5', 'roundRect'],
  ]) &&
  JSON.stringify(reopenedShapeNames) === JSON.stringify([
    ['Shape 2', 'Packed folded corner'],
    ['Shape 2', 'Packed folded corner'],
  ]);
const createdText = created.addSlide().addText('Smoke\\n\\nParagraph', { align: 'center', fit: 'shrink', valign: 'top', vert: 'vert270', wrap: false, bullet: true, level: 2, margin: 10, rtlMode: true, spacing: { before: 4, after: 6, line: { kind: 'exact', points: 20 } }, tabStops: [{ position: 1.25 }, { position: 2.5, alignment: 'right' }] });
const initialTextWrap = createdText.textWrap;
const initialTextDirection = createdText.textDirection;
const initialTextFit = createdText.textFit;
createdText.text = 'Updated\\r\\nParagraph';
createdText.textMargins = { top: 4, left: 8 };
createdText.verticalAlignment = 'bottom';
createdText.textWrap = true;
const updatedTextWrap = createdText.textWrap;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
const updatedTextDirection = createdText.textDirection;
createdText.textDirection = undefined;
createdText.textFit = 'resize';
const updatedTextFit = createdText.textFit;
createdText.textFit = 'none';
createdText.textFit = undefined;
const richText = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: 1, spacing: { line: { kind: 'multiple', factor: 1.5 } }, tabStops: [{ position: 1.5, alignment: 'center' }], runs: [{ text: 'Bold', style: { bold: true, fontSize: 18, color: { kind: 'srgb', value: 'ff0000' }, glow: { color: { kind: 'srgb', value: '00ff00' }, opacity: 0.5, size: 8 }, highlight: { kind: 'srgb', value: 'ffff00' }, outline: { color: { kind: 'srgb', value: '0000ff' }, size: 1.5 }, underline: true, strike: true } }, { text: 'Blue', softBreakBefore: true, style: { lang: 'de-DE', color: { kind: 'scheme', value: 'accent1' }, glow: { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2.5 }, highlight: { kind: 'scheme', value: 'accent2' }, outline: { color: { kind: 'scheme', value: 'accent3' }, size: 2 }, underline: { style: 'dbl', color: { kind: 'srgb', value: '00ff00' } }, strike: 'dblStrike' } }] }, { rtl: false, runs: [{ text: 'LTR override' }] }], { lang: 'fr-CA', rtlMode: true });
const marginText = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, marginLeft: false, runs: [{ text: 'Bullet' }] }, { marginLeft: false, runs: [{ text: 'Absent' }] }], { paragraphMarginLeft: 24 });
const initialParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const bulletMarginIsolation = marginText.richText[1].marginLeft === undefined && marginText.richText[1].bullet.indent === 27;
marginText.richText = [{ marginLeft: 6, runs: [{ text: 'Six' }] }, { marginLeft: 0, runs: [{ text: 'Zero' }] }, { marginLeft: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }];
const updatedParagraphMargins = marginText.richText.map(({ marginLeft }) => marginLeft);
const rightMarginText = created.slides[0].addRichText([{ marginRight: 12, runs: [{ text: 'Twelve' }] }, { bullet: true, runs: [{ text: 'Bullet' }] }, { marginRight: false, runs: [{ text: 'Absent' }] }], { paragraphMarginRight: 24 });
const initialParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const bulletRightMarginCoexistence = rightMarginText.richText[1].marginRight === 24 && rightMarginText.richText[1].bullet.indent === 27;
rightMarginText.richText = [{ marginRight: 6, runs: [{ text: 'Six' }] }, { marginRight: 0, runs: [{ text: 'Zero' }] }, { marginRight: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }, { bullet: true, marginRight: 9, runs: [{ text: 'Bullet' }] }];
const updatedParagraphRightMargins = rightMarginText.richText.map(({ marginRight }) => marginRight);
const indentText = created.slides[0].addRichText([{ runs: [{ text: 'Default' }] }, { indent: -18, runs: [{ text: 'Hanging' }] }, { indent: false, runs: [{ text: 'Absent' }] }, { bullet: true, indent: false, runs: [{ text: 'Bullet' }] }], { paragraphIndent: 24 });
const initialParagraphIndents = indentText.richText.map(({ indent }) => indent);
const bulletIndentIsolation = indentText.richText[3].indent === undefined && indentText.richText[3].bullet.indent === 27;
indentText.richText = [{ indent: 6, runs: [{ text: 'Positive' }] }, { indent: -6, runs: [{ text: 'Negative' }] }, { indent: 0, runs: [{ text: 'Zero' }] }, { indent: false, runs: [{ text: 'Cleared' }] }, { runs: [{ text: 'Omitted' }] }];
const updatedParagraphIndents = indentText.richText.map(({ indent }) => indent);
const transparencyText = created.slides[0].addRichText([{ runs: [{ text: 'Quarter', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 25 } }, { text: 'Fractional', style: { transparency: 50.5555 } }, { text: 'Invisible', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 } }, { text: 'Default', style: { transparency: 60 } }] }]);
const initialTransparencies = transparencyText.richText[0].runs.map(({ style }) => style?.transparency);
transparencyText.richText = [{ runs: [{ text: 'Opaque', style: { transparency: 0 } }, { text: 'Mostly', style: { transparency: 75 } }, { text: 'Cleared' }] }];
const updatedTransparencies = transparencyText.richText[0].runs.map(({ style }) => style?.transparency);
const tableSlide = created.slides[0];
const creationColor = { kind: 'srgb', value: '#D9EAF7' };
const creationFill = { kind: 'solid', color: creationColor, transparency: 33.3334 };
const tableCreationFillColor = { kind: 'scheme', value: 'accent4' };
const tableCreationFill = { kind: 'solid', color: tableCreationFillColor, transparency: 40 };
const tableCreationBorderColor = { kind: 'scheme', value: 'accent4' };
const tableCreationBorder = { kind: 'line', color: tableCreationBorderColor, width: 1.5, style: 'dash' };
const creationBorderColor = { kind: 'srgb', value: '#C00000' };
const creationBorder = { kind: 'line', color: creationBorderColor, width: 2, style: 'solid' };
const creationMargin = { top: 4, left: 8 };
const createdTable = tableSlide.addTable([
  [
    { text: 'Region', options: { align: 'left', border: creationBorder, fill: creationFill, fit: 'shrink', margin: creationMargin, textDirection: 'vert', valign: 'top' } },
    { text: 'Revenue', options: {
      border: [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, undefined, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
      fit: 'resize',
      margin: [1, 2, 3, 4],
      textDirection: 'vert270',
      valign: 'middle',
    } },
  ],
  [
    { text: 'East', options: { align: 'right', border: { top: { kind: 'line', color: { kind: 'scheme', value: 'accent3' }, width: 1, style: 'dash' }, left: { kind: 'none' } }, fit: 'none', margin: 0, textDirection: 'wordArtVert', valign: 'bottom' } },
    { text: '', options: { align: 'justify', border: { kind: 'none' }, fill: { kind: 'none' }, margin: {}, textDirection: 'horz' } },
  ],
], { name: 'Created smoke table', align: 'center', columnWidths: [inches(1), inches(3)], rowHeights: [inches(0.5), inches(1.5)], fill: tableCreationFill, margin: { top: 9, left: 18 }, valign: 'middle' });
const tableBorderSlide = created.addSlide();
const createdTableBorderDefault = tableBorderSlide.addTable([[
  'Inherited border',
  { text: 'None override', options: { border: { kind: 'none' } } },
]], { name: 'Created table border default', columnWidths: inches(1), rowHeights: inches(0.5), border: tableCreationBorder });
const createdTableDirectionDefault = tableBorderSlide.addTable([[
  'Inherited direction',
  { text: 'Horizontal override', options: { textDirection: 'horz' } },
]], { name: 'Created table direction default', columnWidths: inches(1), rowHeights: inches(0.5), textDirection: 'vert270' });
const tableCellObjectCreation = JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ text }) => text))) === JSON.stringify([['Region', 'Revenue'], ['East', '']]);
const initialCreatedFill = createdTable.rows[0].cells[0].fill;
const initialTableDefaultFill = createdTable.rows[1].cells[0].fill;
const initialTableNoneOverride = createdTable.rows[1].cells[1].fill;
const initialCreatedBorders = createdTable.rows.map(({ cells }) => cells.map(({ borders }) => borders));
const initialTableDefaultBorders = createdTableBorderDefault.rows[0].cells[0].borders;
const initialTableBorderNoneOverride = createdTableBorderDefault.rows[0].cells[1].borders;
const initialCreatedMargins = createdTable.rows.map(({ cells }) => cells.map(({ margins }) => margins));
const initialCreatedAlignments = createdTable.rows.map(({ cells }) => cells.map(({ verticalAlignment }) => verticalAlignment));
const initialCreatedDirections = createdTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
const initialCreatedFits = createdTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
const initialTableDirectionDefault = createdTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection);
creationColor.value = '000000';
creationFill.transparency = 1;
tableCreationFillColor.value = 'accent6';
tableCreationFill.transparency = 1;
tableCreationBorderColor.value = 'accent6';
tableCreationBorder.width = 9;
creationBorderColor.value = '000000';
creationBorder.width = 9;
creationMargin.top = 99;
creationMargin.left = 99;
const detachedCreatedFill = createdTable.rows[0].cells[0].fill;
const detachedTableDefaultFill = createdTable.rows[1].cells[0].fill;
const detachedCreatedBorders = createdTable.rows[0].cells[0].borders;
const detachedTableDefaultBorders = createdTableBorderDefault.rows[0].cells[0].borders;
const detachedCreatedMargins = createdTable.rows[0].cells[0].margins;
const createdTableDefaults = createdTable instanceof TableModel && createdTable.transform.x === inches(0.5) && createdTable.transform.y === inches(0.5) && createdTable.rows[1].cells[1].margins?.top === 9 && createdTable.rows[1].cells[1].margins?.left === 18;
const createdTableXml = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
const createdTableCells = [...createdTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0]);
const createdTableHorizontalAlignments = createdTableCells.map((cellXml) => cellXml.match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const createdTableDirections = createdTableCells.map((cellXml) => cellXml.match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const createdTableFits = createdTableCells.map((cellXml) => cellXml.match(/<a:(normAutofit|spAutoFit)\\/>/)?.[1]);
const createdTableAnchors = createdTableCells.flatMap((cellXml) => [...cellXml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)].map((match) => match[1]));
const createdTableGrid = [...createdTableXml.matchAll(/<a:gridCol w="(\\d+)"\\/>/g)].map((match) => Number(match[1]));
const createdTableRows = [...createdTableXml.matchAll(/<a:tr h="(\\d+)">/g)].map((match) => Number(match[1]));
const initialTableColumnWidths = createdTable.columnWidths;
const initialTableRowHeights = createdTable.rowHeights;
createdTable.setColumnWidths([inches(1.5), inches(2.5)]);
createdTable.setRowHeights([inches(0.75), inches(1.25)]);
const explicitTableHeight = createdTable.transform.height;
createdTable.setRowHeights([0, inches(1.25)]);
const automaticTableHeightPreserved = createdTable.transform.height === explicitTableHeight;
createdTable.setRowHeights([inches(0.75), inches(1.25)]);
createdTable.setCellText(1, 0, 'Edited East');
createdTable.setCellFill(1, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 50 });
createdTable.setCellBorders(1, 0, { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, width: 1, style: 'solid' });
createdTable.setCellVerticalAlignment(0, 0, 'bottom');
createdTable.setCellVerticalAlignment(1, 1, undefined);
createdTable.setCellMargins(0, 0, { bottom: 9 });
createdTable.setCellMargins(1, 1, undefined);
createdTable.setCellFill(1, 1, undefined);
createdTableBorderDefault.setCellBorders(0, 0, undefined);
const reopenedCreated = await PptxDocument.open(await created.write());
const reopenedCreatedTable = reopenedCreated.slides[0].shapes.find((shape) => shape.name === 'Created smoke table');
const reopenedCreatedTableXml = new TextDecoder().decode(reopenedCreated.opcPackage.requirePart(reopenedCreated.slides[0].partUri).bytes);
const reopenedCreatedTableHorizontalAlignments = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const reopenedCreatedTableDirections = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const reopenedCreatedFits = reopenedCreatedTable instanceof TableModel
  ? reopenedCreatedTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))
  : undefined;
const reopenedTableBorderDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created table border default');
const reopenedTableDirectionDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created table direction default');
const tableCreation = createdTableDefaults && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rows[1].cells[0].text === 'Edited East' && reopenedCreatedTable.rows[1].cells[1].text === '' && reopenedCreatedTable.rows[1].cells[1].verticalAlignment === undefined;
const tableColumnWidths = createdTable.transform.width === inches(4) && createdTableGrid.length === 2 && createdTableGrid[0] === inches(1) && createdTableGrid[1] === inches(3) && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.transform.width === inches(4);
const tableColumnWidthEditing = initialTableColumnWidths?.join(',') === [inches(1), inches(3)].join(',') && createdTable.columnWidths?.join(',') === [inches(1.5), inches(2.5)].join(',') && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.columnWidths?.join(',') === [inches(1.5), inches(2.5)].join(',');
const tableRowHeights = createdTable.transform.height === inches(2) && createdTableRows.length === 2 && createdTableRows[0] === inches(0.5) && createdTableRows[1] === inches(1.5) && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.transform.height === inches(2);
const tableRowHeightEditing = initialTableRowHeights?.join(',') === [inches(0.5), inches(1.5)].join(',') && automaticTableHeightPreserved && createdTable.rowHeights?.join(',') === [inches(0.75), inches(1.25)].join(',') && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rowHeights?.join(',') === [inches(0.75), inches(1.25)].join(',');
const tableCellFillCreation = initialCreatedFill?.kind === 'solid' && initialCreatedFill.color.kind === 'srgb' && initialCreatedFill.color.value === 'D9EAF7' && initialCreatedFill.transparency === 33.333 && detachedCreatedFill?.kind === 'solid' && detachedCreatedFill.color.kind === 'srgb' && detachedCreatedFill.color.value === 'D9EAF7' && detachedCreatedFill.transparency === 33.333 && createdTable.rows[0].cells[1].fill?.kind === 'solid' && createdTable.rows[0].cells[1].fill.color.kind === 'scheme' && createdTable.rows[0].cells[1].fill.color.value === 'accent2' && createdTable.rows[0].cells[1].fill.transparency === 25 && initialTableNoneOverride?.kind === 'none' && reopenedCreatedTable instanceof TableModel && reopenedCreatedTable.rows[0].cells[0].fill?.kind === 'solid' && reopenedCreatedTable.rows[0].cells[0].fill.color.value === 'D9EAF7' && reopenedCreatedTable.rows[0].cells[0].fill.transparency === 33.333 && reopenedCreatedTable.rows[0].cells[1].fill?.kind === 'solid' && reopenedCreatedTable.rows[0].cells[1].fill.color.kind === 'scheme' && reopenedCreatedTable.rows[0].cells[1].fill.color.value === 'accent2' && reopenedCreatedTable.rows[0].cells[1].fill.transparency === 25;
const tableFillCreation = tableCellFillCreation &&
  initialTableDefaultFill?.kind === 'solid' &&
  initialTableDefaultFill.color.kind === 'scheme' &&
  initialTableDefaultFill.color.value === 'accent4' &&
  initialTableDefaultFill.transparency === 40 &&
  detachedTableDefaultFill?.kind === 'solid' &&
  detachedTableDefaultFill.color.kind === 'scheme' &&
  detachedTableDefaultFill.color.value === 'accent4' &&
  detachedTableDefaultFill.transparency === 40 &&
  createdTableCells[2]?.includes('</a:lnB><a:solidFill><a:schemeClr val="accent4"><a:alpha val="60000"/></a:schemeClr></a:solidFill>') === true &&
  createdTable.rows[1].cells[0].fill?.kind === 'solid' &&
  createdTable.rows[1].cells[0].fill.color.kind === 'scheme' &&
  createdTable.rows[1].cells[0].fill.color.value === 'accent1' &&
  createdTable.rows[1].cells[0].fill.transparency === 50 &&
  createdTable.rows[1].cells[1].fill === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[1].cells[0].fill?.kind === 'solid' &&
  reopenedCreatedTable.rows[1].cells[0].fill.color.kind === 'scheme' &&
  reopenedCreatedTable.rows[1].cells[0].fill.color.value === 'accent1' &&
  reopenedCreatedTable.rows[1].cells[0].fill.transparency === 50 &&
  reopenedCreatedTable.rows[1].cells[1].fill === undefined;
const creationSides = ['top', 'right', 'bottom', 'left'];
const isCreationLine = (border, colorKind, colorValue, width, style) => border?.kind === 'line' && border.color.kind === colorKind && border.color.value === colorValue && border.width === width && border.style === style;
const allCreationLines = (borders, colorKind, colorValue, width, style) => borders !== undefined && creationSides.every((side) => isCreationLine(borders[side], colorKind, colorValue, width, style));
const allCreationNone = (borders) => borders !== undefined && creationSides.every((side) => borders[side]?.kind === 'none');
const initialScalarBorders = initialCreatedBorders[0][0];
const initialTupleBorders = initialCreatedBorders[0][1];
const initialNamedBorders = initialCreatedBorders[1][0];
const initialNoneBorders = initialCreatedBorders[1][1];
const tableCellBorderCreation = tableCellFillCreation &&
  allCreationLines(initialScalarBorders, 'srgb', 'C00000', 2, 'solid') &&
  allCreationLines(detachedCreatedBorders, 'srgb', 'C00000', 2, 'solid') &&
  isCreationLine(initialTupleBorders?.top, 'scheme', 'accent1', 1.5, 'dash') &&
  initialTupleBorders?.right?.kind === 'none' &&
  isCreationLine(initialTupleBorders?.bottom, 'srgb', '00FF00', 0, undefined) &&
  initialTupleBorders?.left?.kind === 'none' &&
  isCreationLine(initialNamedBorders?.top, 'scheme', 'accent3', 1, 'dash') &&
  initialNamedBorders?.right?.kind === 'none' && initialNamedBorders?.bottom?.kind === 'none' && initialNamedBorders?.left?.kind === 'none' &&
  allCreationNone(initialNoneBorders) &&
  allCreationLines(createdTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') &&
  reopenedCreatedTable instanceof TableModel &&
  allCreationLines(reopenedCreatedTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid') &&
  isCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.top, 'scheme', 'accent1', 1.5, 'dash') &&
  reopenedCreatedTable.rows[0].cells[1].borders?.right?.kind === 'none' &&
  isCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.bottom, 'srgb', '00FF00', 0, undefined) &&
  reopenedCreatedTable.rows[0].cells[1].borders?.left?.kind === 'none' &&
  allCreationLines(reopenedCreatedTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') &&
  allCreationNone(reopenedCreatedTable.rows[1].cells[1].borders);
const tableBorderCreation = tableCellBorderCreation &&
  allCreationLines(initialTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  allCreationNone(initialTableBorderNoneOverride) &&
  allCreationLines(detachedTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  createdTableBorderDefault.rows[0].cells[0].borders === undefined &&
  reopenedTableBorderDefault instanceof TableModel &&
  reopenedTableBorderDefault.rows[0].cells[0].borders === undefined &&
  allCreationNone(reopenedTableBorderDefault.rows[0].cells[1].borders);
const marginVector = (margins) => [margins?.top, margins?.right, margins?.bottom, margins?.left];
const tableCellMarginCreation = tableCellBorderCreation &&
  JSON.stringify(initialCreatedMargins.map((row) => row.map(marginVector))) === JSON.stringify([
    [[4, 7.2, 3.6, 8], [1, 2, 3, 4]],
    [[0, 0, 0, 0], [9, 7.2, 3.6, 18]],
  ]) &&
  JSON.stringify(marginVector(detachedCreatedMargins)) === JSON.stringify([4, 7.2, 3.6, 8]) &&
  createdTable.rows[0].cells[0].margins?.top === undefined &&
  createdTable.rows[0].cells[0].margins?.right === undefined &&
  createdTable.rows[0].cells[0].margins?.bottom === 9 &&
  createdTable.rows[0].cells[0].margins?.left === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[0].cells[0].margins?.top === undefined &&
  reopenedCreatedTable.rows[0].cells[0].margins?.right === undefined &&
  reopenedCreatedTable.rows[0].cells[0].margins?.bottom === 9 &&
  reopenedCreatedTable.rows[0].cells[0].margins?.left === undefined;
const tableMarginCreation = tableCellMarginCreation &&
  createdTableXml.includes('<a:tcPr marL="228600" marR="91440" marT="114300" marB="45720" anchor="ctr">') &&
  createdTable.rows[1].cells[1].margins === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[1].cells[1].margins === undefined;
const tableCellVerticalAlignmentCreation = tableCellMarginCreation &&
  JSON.stringify(initialCreatedAlignments) === JSON.stringify([
    ['top', 'middle'],
    ['bottom', 'middle'],
  ]) &&
  createdTable.rows[0].cells[0].verticalAlignment === 'bottom' &&
  createdTable.rows[0].cells[1].verticalAlignment === 'middle' &&
  createdTable.rows[1].cells[0].verticalAlignment === 'bottom' &&
  createdTable.rows[1].cells[1].verticalAlignment === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[0].cells[0].verticalAlignment === 'bottom' &&
  reopenedCreatedTable.rows[0].cells[1].verticalAlignment === 'middle' &&
  reopenedCreatedTable.rows[1].cells[0].verticalAlignment === 'bottom' &&
  reopenedCreatedTable.rows[1].cells[1].verticalAlignment === undefined;
const tableVerticalAlignmentCreation = tableCellVerticalAlignmentCreation &&
  JSON.stringify(createdTableAnchors) === JSON.stringify(['t', 'ctr', 'b', 'ctr']) &&
  createdTableCells.every((cellXml) => !/<a:bodyPr[^>]* anchor=/.test(cellXml));
const tableCellHorizontalAlignmentCreation = tableVerticalAlignmentCreation &&
  JSON.stringify(createdTableHorizontalAlignments) === JSON.stringify(['l', 'ctr', 'r', 'just']) &&
  JSON.stringify(reopenedCreatedTableHorizontalAlignments) === JSON.stringify(['l', 'ctr', 'r', 'just']);
const tableHorizontalAlignmentCreation = tableCellHorizontalAlignmentCreation &&
  createdTableHorizontalAlignments[1] === 'ctr' &&
  reopenedCreatedTableHorizontalAlignments[1] === 'ctr';
const tableCellTextDirectionCreation = tableHorizontalAlignmentCreation &&
  JSON.stringify(initialCreatedDirections) === JSON.stringify([
    ['vert', 'vert270'],
    ['wordArtVert', undefined],
  ]) &&
  JSON.stringify(createdTableDirections) === JSON.stringify([
    'vert',
    'vert270',
    'wordArtVert',
    undefined,
  ]) &&
  reopenedCreatedTable instanceof TableModel &&
  JSON.stringify(reopenedCreatedTable.rows.map(({ cells }) =>
    cells.map(({ textDirection }) => textDirection))) === JSON.stringify([
    ['vert', 'vert270'],
    ['wordArtVert', undefined],
  ]) &&
  JSON.stringify(reopenedCreatedTableDirections) === JSON.stringify([
    'vert',
    'vert270',
    'wordArtVert',
    undefined,
  ]);
const tableCellTextFitCreation = tableCellTextDirectionCreation &&
  JSON.stringify(initialCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]) &&
  JSON.stringify(createdTableFits) === JSON.stringify([
    'normAutofit',
    'spAutoFit',
    undefined,
    undefined,
  ]) &&
  JSON.stringify(reopenedCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]);
const tableTextDirectionCreation = tableCellTextDirectionCreation &&
  JSON.stringify(initialTableDirectionDefault) === JSON.stringify([
    'vert270',
    undefined,
  ]) &&
  reopenedTableDirectionDefault instanceof TableModel &&
  JSON.stringify(reopenedTableDirectionDefault.rows[0].cells.map(
    ({ textDirection }) => textDirection)) === JSON.stringify([
    'vert270',
    undefined,
  ]);
const tablePart = created.opcPackage.requirePart(tableSlide.partUri);
const tableXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Smoke table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:bodyPr custom="TARGET"><a:noAutofit/></a:bodyPr><a:p><a:pPr algn="ctr"/><a:r><a:t>Target</a:t></a:r></a:p></a:txBody><a:tcPr vert="horz" anchor="ctr" marL="12700" marR="25400" marT="38100" marB="50800"><a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="sysDash"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr custom="NEIGHBOR"><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>Neighbor</a:t></a:r></a:p></a:txBody><a:tcPr vert="vert" anchor="b" marL="63500" marR="76200" marT="88900" marB="101600" keep="ADJACENT"><a:lnL w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
created.opcPackage.setPart(tableSlide.partUri, new TextDecoder().decode(tablePart.bytes).replace('</p:spTree>', tableXml + '</p:spTree>'), tablePart.contentType);
const table = tableSlide.shapes.find((shape) => shape.name === 'Smoke table');
const initialCellDirection = table?.rows[0]?.cells[0]?.textDirection;
const initialCellFit = table?.rows[0]?.cells[0]?.textFit;
const initialCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
const initialCellMargins = table?.rows[0]?.cells[0]?.margins;
const initialCellFill = table?.rows[0]?.cells[0]?.fill;
const initialCellBorders = table?.rows[0]?.cells[0]?.borders;
const initialHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
const initialNeighborHorizontalAlignment = table?.rows[0]?.cells[1]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'left');
const leftHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'center');
const centerHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'right');
const rightHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'justify');
const justifyHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, undefined);
const clearedHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellBorders(0, 0, { kind: 'line', color: { kind: 'srgb', value: '#0000FF' }, width: 2, style: 'solid' });
const scalarCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }, undefined]);
const tupleCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, { left: { kind: 'none' } });
const partialCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, undefined);
const clearedCellBorders = table?.rows[0]?.cells[0]?.borders;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '#FF0000' } });
const opaqueCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 0 });
const explicitOpaqueCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '112233' }, transparency: 33.333 });
const fractionalCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, { kind: 'none' });
const noneCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, undefined);
const clearedCellFill = table?.rows[0]?.cells[0]?.fill;
table?.setCellMargins(0, 0, 4);
const uniformCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, [1, 2, 3, 4]);
const tupleCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, { top: 5, left: 7 });
const partialCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, undefined);
const clearedCellMargins = table?.rows[0]?.cells[0]?.margins;
table?.setCellTextFit(0, 0, 'shrink');
const shrinkCellFit = table?.rows[0]?.cells[0]?.textFit;
const beforeSameFit = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
table?.setCellTextFit(0, 0, 'shrink');
const sameFitPreserved = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes) === beforeSameFit;
table?.setCellTextFit(0, 0, 'resize');
const resizeCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'none');
const noneClearedCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'shrink');
table?.setCellTextFit(0, 0, undefined);
const undefinedClearedCellFit = table?.rows[0]?.cells[0]?.textFit;
table?.setCellVerticalAlignment(0, 0, 'top');
const topCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'middle');
const middleCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'bottom');
const bottomCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, undefined);
const clearedCellAlignment = table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellTextDirection(0, 0, 'vert270');
const rotatedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'wordArtVert');
const stackedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'horz');
const horizontalCellDirection = table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, undefined);
const clearedCellDirection = table?.rows[0]?.cells[0]?.textDirection;
const reopenedEdited = await PptxDocument.open(await created.write());
const reopenedEditedTable = reopenedEdited.slides[0].shapes.find(
  (shape) => shape.name === 'Smoke table',
);
const reopenedClearedHorizontalAlignment = reopenedEditedTable instanceof TableModel
  ? reopenedEditedTable.rows[0].cells[0].horizontalAlignment
  : null;
const reopenedNeighborHorizontalAlignment = reopenedEditedTable instanceof TableModel
  ? reopenedEditedTable.rows[0].cells[1].horizontalAlignment
  : null;
const inheritedLanguage = richText.richText[0].runs[0].style.lang;
const localLanguage = richText.richText[0].runs[1].style.lang;
const initialRtl = richText.richText.map(({ rtl }) => rtl);
const presentationRtlEnabled = created.rtlMode;
created.rtlMode = false;
const presentationRtlDisabled = created.rtlMode;
created.rtlMode = undefined;
const presentationRtlCleared = created.rtlMode;
const paragraphRtlAfterGlobalClear = richText.richText.map(({ rtl }) => rtl);
const metadata = PptxDocument.create({ title: 'Packed & <Title>' });
const createdPresentationTitle = metadata.title;
metadata.title = 'Edited title';
const editedPresentationTitle = metadata.title;
const reopenedMetadata = await PptxDocument.open(await metadata.write());
const reopenedPresentationTitle = reopenedMetadata.title;
metadata.title = '';
const emptyPresentationTitle = metadata.title;
metadata.title = undefined;
const clearedPresentationTitle = metadata.title;
const authorship = PptxDocument.create({ author: 'Packed & <Author>' });
const createdPresentationAuthor = authorship.author;
authorship.author = 'Edited author';
const editedPresentationAuthor = authorship.author;
const reopenedAuthorship = await PptxDocument.open(await authorship.write());
const reopenedPresentationAuthor = reopenedAuthorship.author;
authorship.author = '';
const emptyPresentationAuthor = authorship.author;
authorship.author = undefined;
const clearedPresentationAuthor = authorship.author;
const editorship = PptxDocument.create({ lastModifiedBy: 'Packed & <Editor>' });
const createdPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = 'Edited editor';
const editedPresentationLastModifiedBy = editorship.lastModifiedBy;
const reopenedEditorship = await PptxDocument.open(await editorship.write());
const reopenedPresentationLastModifiedBy = reopenedEditorship.lastModifiedBy;
editorship.lastModifiedBy = '';
const emptyPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = undefined;
const clearedPresentationLastModifiedBy = editorship.lastModifiedBy;
const chronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
});
const createdPresentationCreatedAt = chronology.createdAt;
const createdPresentationModifiedAt = chronology.modifiedAt;
chronology.createdAt = '2026-07-30T00:00:00Z';
const editedPresentationCreatedAt = chronology.createdAt;
chronology.modifiedAt = '2026-07-30T01:02:03Z';
const editedPresentationModifiedAt = chronology.modifiedAt;
const reopenedChronology = await PptxDocument.open(await chronology.write());
const reopenedPresentationCreatedAt = reopenedChronology.createdAt;
const reopenedPresentationModifiedAt = reopenedChronology.modifiedAt;
chronology.modifiedAt = undefined;
const clearedPresentationModifiedAt = chronology.modifiedAt;
const modifiedAtCreatedIsolation = chronology.createdAt;
chronology.createdAt = undefined;
const clearedPresentationCreatedAt = chronology.createdAt;
const subjectMatter = PptxDocument.create({ subject: 'Packed & <Subject>' });
const createdPresentationSubject = subjectMatter.subject;
subjectMatter.subject = 'Edited subject';
const editedPresentationSubject = subjectMatter.subject;
const reopenedSubjectMatter = await PptxDocument.open(await subjectMatter.write());
const reopenedPresentationSubject = reopenedSubjectMatter.subject;
subjectMatter.subject = '';
const emptyPresentationSubject = subjectMatter.subject;
subjectMatter.subject = undefined;
const clearedPresentationSubject = subjectMatter.subject;
const revisioned = PptxDocument.create({ revision: '007' });
const createdPresentationRevision = revisioned.revision;
revisioned.revision = '42';
const editedPresentationRevision = revisioned.revision;
const reopenedRevisioned = await PptxDocument.open(await revisioned.write());
const reopenedPresentationRevision = reopenedRevisioned.revision;
revisioned.revision = undefined;
const clearedPresentationRevision = revisioned.revision;
const organization = PptxDocument.create({ company: 'Packed & <Company>' });
const createdPresentationCompany = organization.company;
organization.company = 'Edited company';
const editedPresentationCompany = organization.company;
const reopenedOrganization = await PptxDocument.open(await organization.write());
const reopenedPresentationCompany = reopenedOrganization.company;
organization.company = '';
const emptyPresentationCompany = organization.company;
organization.company = undefined;
const clearedPresentationCompany = organization.company;
const themed = PptxDocument.create({
  theme: { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' },
});
const createdTheme = themed.theme;
themed.theme = { headFontFace: 'Noto Sans Display' };
const replacedTheme = themed.theme;
themed.masterLayoutTheme.presentationTheme.setFonts({ minorLatin: 'Noto Sans' });
const reopenedTheme = (await PptxDocument.open(await themed.write())).theme;
const sectioned = PptxDocument.create();
const firstSection = sectioned.addSection({ title: 'Packed & <Intro>' });
const packedSectionEscaped = new TextDecoder().decode(
  sectioned.opcPackage.requirePart(sectioned.presentationPartUri).bytes,
).includes('name="Packed &amp; &lt;Intro&gt;"');
const assignedSlide = sectioned.addSlide({ sectionTitle: 'Packed & <Intro>' });
const automaticSlide = sectioned.addSlide();
const secondAssignedSlide = sectioned.addSlide({ sectionTitle: 'Packed & <Intro>' });
const dataSection = sectioned.addSection({ title: 'Data', order: 1 });
sectioned.assignSlideToSection(1, dataSection.id);
const defaultSection = sectioned.sections.find(({ title }) => title === 'Default-1');
if (!defaultSection) throw new Error('Packed default section was not created');
sectioned.deleteSection(defaultSection.id);
sectioned.renameSection(firstSection.id, 'Edited intro');
sectioned.moveSection(dataSection.id, 0);
const detachedSections = sectioned.sections;
detachedSections[0].title = 'Detached caller title';
detachedSections[0].slideIds.push(999);
const currentSections = sectioned.sections;
const reopenedSections = (await PptxDocument.open(await sectioned.write())).sections;
const hiddenDeck = PptxDocument.create();
const packedVisibleSlide = hiddenDeck.addSlide();
const packedHiddenSlide = hiddenDeck.addSlide();
packedHiddenSlide.hidden = true;
const packedHiddenDuplicate = hiddenDeck.duplicateSlide(1);
packedVisibleSlide.hidden = true;
packedHiddenSlide.hidden = false;
const reopenedHiddenDeck = await PptxDocument.open(await hiddenDeck.write());
const reopenedHiddenStates = reopenedHiddenDeck.slides.map(({ hidden }) => hidden);
const reopenedHiddenRootStates = reopenedHiddenDeck.slides.map(({ partUri }) =>
  /<p:sld\\b[^>]*\\sshow="0"/.test(new TextDecoder().decode(
    reopenedHiddenDeck.opcPackage.requirePart(partUri).bytes,
  )));
const speakerNotesDeck = PptxDocument.create();
const packedLazyNotesSlide = speakerNotesDeck.addSlide();
const packedLazyNotesInitial = packedLazyNotesSlide.notes;
packedLazyNotesSlide.addNotes('Temporary notes');
packedLazyNotesSlide.notes = undefined;
const packedEmptyNotesSlide = speakerNotesDeck.addSlide();
packedEmptyNotesSlide.notes = '';
const packedOriginalNotesSlide = speakerNotesDeck.addSlide().addNotes('Original');
const packedNotesDuplicate = speakerNotesDeck.duplicateSlide(2);
packedOriginalNotesSlide.notes = 'Edited';
const reopenedNotesDeck = await PptxDocument.open(await speakerNotesDeck.write());
const reopenedNotesSnapshots = reopenedNotesDeck.slides.map(({ notes }) => notes);
const reopenedNotesUris = reopenedNotesDeck.slides.flatMap((slide) => {
  const uri = slide.relationships.find(({ type }) => type.endsWith('/notesSlide'))?.resolvedTarget;
  return uri === undefined ? [] : [uri];
});
const reopenedNotesParts = reopenedNotesDeck.opcPackage.parts.filter(
  ({ contentType }) => contentType === 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
);
const reopenedNotesRetargeted = reopenedNotesDeck.slides.every((slide) => {
  const uri = slide.relationships.find(({ type }) => type.endsWith('/notesSlide'))?.resolvedTarget;
  if (uri === undefined) return slide.notes === undefined;
  const backlinks = reopenedNotesDeck.opcPackage.relationships(uri).filter(
    ({ type }) => type.endsWith('/slide'),
  );
  return backlinks.length === 1 && backlinks[0].resolvedTarget === slide.partUri;
});
const reopenedNotesMasterUris = reopenedNotesUris.map((uri) =>
  reopenedNotesDeck.opcPackage.relationships(uri).find(
    ({ type }) => type.endsWith('/notesMaster'),
  )?.resolvedTarget);
richText.richText = [{ align: 'justify', bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 }, level: 3, spacing: { before: 5, after: 7, line: { kind: 'exact', points: 22 } }, tabStops: [{ position: 2.75, alignment: 'decimal' }], runs: [{ text: 'Updated rich', style: { lang: 'ja-JP', baseline: 'superscript', characterSpacing: 2.5, italic: true, glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 }, highlight: { kind: 'srgb', value: '00ff00' }, outline: { color: { kind: 'scheme', value: 'accent1' }, size: 0.75 }, underline: { style: 'wavyHeavy', color: { kind: 'scheme', value: 'accent2' } }, strike: false } }] }];
const custom = PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
custom.slideSize = { width: inches(10), height: inches(7.5) };
const customXml = new TextDecoder().decode(custom.opcPackage.requirePart('/ppt/presentation.xml').bytes);
const checks = {
  PptxDocument: typeof PptxDocument === 'function',
  presetShapes,
  presentationRtl: presentationRtlEnabled === true && presentationRtlDisabled === false && presentationRtlCleared === undefined && paragraphRtlAfterGlobalClear[0] === true && paragraphRtlAfterGlobalClear[1] === false,
  presentationTitle: createdPresentationTitle === 'Packed & <Title>' && editedPresentationTitle === 'Edited title' && reopenedPresentationTitle === 'Edited title' && emptyPresentationTitle === '' && clearedPresentationTitle === undefined,
  presentationAuthor: createdPresentationAuthor === 'Packed & <Author>' && editedPresentationAuthor === 'Edited author' && reopenedPresentationAuthor === 'Edited author' && emptyPresentationAuthor === '' && clearedPresentationAuthor === undefined,
  presentationLastModifiedBy: createdPresentationLastModifiedBy === 'Packed & <Editor>' && editedPresentationLastModifiedBy === 'Edited editor' && reopenedPresentationLastModifiedBy === 'Edited editor' && emptyPresentationLastModifiedBy === '' && clearedPresentationLastModifiedBy === undefined,
  presentationCreatedAt: createdPresentationCreatedAt === '2024-02-29T12:34:56.123+05:30' && editedPresentationCreatedAt === '2026-07-30T00:00:00Z' && reopenedPresentationCreatedAt === '2026-07-30T00:00:00Z' && clearedPresentationCreatedAt === undefined,
  presentationModifiedAt: createdPresentationModifiedAt === '2024-03-01T01:02:03.456+08:00' && editedPresentationModifiedAt === '2026-07-30T01:02:03Z' && reopenedPresentationModifiedAt === '2026-07-30T01:02:03Z' && clearedPresentationModifiedAt === undefined && modifiedAtCreatedIsolation === '2026-07-30T00:00:00Z',
  presentationSubject: createdPresentationSubject === 'Packed & <Subject>' && editedPresentationSubject === 'Edited subject' && reopenedPresentationSubject === 'Edited subject' && emptyPresentationSubject === '' && clearedPresentationSubject === undefined,
  presentationRevision: createdPresentationRevision === '007' && editedPresentationRevision === '42' && reopenedPresentationRevision === '42' && clearedPresentationRevision === undefined,
  presentationCompany: createdPresentationCompany === 'Packed & <Company>' && editedPresentationCompany === 'Edited company' && reopenedPresentationCompany === 'Edited company' && emptyPresentationCompany === '' && clearedPresentationCompany === undefined,
  presentationThemeFonts: createdTheme?.headFontFace === 'Aptos Display' && createdTheme.bodyFontFace === 'Aptos' && replacedTheme?.headFontFace === 'Noto Sans Display' && replacedTheme.bodyFontFace === 'Calibri' && reopenedTheme?.headFontFace === 'Noto Sans Display' && reopenedTheme.bodyFontFace === 'Noto Sans',
  presentationSections: packedSectionEscaped && assignedSlide.slideId === 256 && automaticSlide.slideId === 257 && secondAssignedSlide.slideId === 258 && /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/.test(firstSection.id) && /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/.test(dataSection.id) && currentSections.length === 2 && currentSections[0].id === dataSection.id && currentSections[0].title === 'Data' && currentSections[0].slideIds.join(',') === String(automaticSlide.slideId) && currentSections[1].id === firstSection.id && currentSections[1].title === 'Edited intro' && currentSections[1].slideIds.join(',') === [assignedSlide.slideId, secondAssignedSlide.slideId].join(',') && reopenedSections?.length === 2 && reopenedSections[0].id === dataSection.id && reopenedSections[0].slideIds.join(',') === String(automaticSlide.slideId) && reopenedSections[1].id === firstSection.id && reopenedSections[1].title === 'Edited intro' && reopenedSections[1].slideIds.join(',') === [assignedSlide.slideId, secondAssignedSlide.slideId].join(','),
  hiddenSlides: packedHiddenDuplicate.hidden === true && reopenedHiddenStates.join(',') === 'true,false,true' && reopenedHiddenRootStates.join(',') === 'true,false,true',
  speakerNotes: packedLazyNotesInitial === undefined && packedNotesDuplicate.notes === 'Original' && JSON.stringify(reopenedNotesSnapshots) === JSON.stringify([undefined, '', 'Edited', 'Original']) && reopenedNotesParts.length === 3 && reopenedNotesUris.length === 3 && new Set(reopenedNotesUris).size === 3 && reopenedNotesRetargeted && reopenedNotesMasterUris.every((uri) => uri !== undefined) && new Set(reopenedNotesMasterUris).size === 1,
  paragraphMarginLeft: initialParagraphMargins[0] === 12 && initialParagraphMargins[1] === undefined && initialParagraphMargins[2] === undefined && bulletMarginIsolation && updatedParagraphMargins[0] === 6 && updatedParagraphMargins[1] === 0 && updatedParagraphMargins[2] === undefined && updatedParagraphMargins[3] === undefined,
  paragraphMarginRight: initialParagraphRightMargins[0] === 12 && initialParagraphRightMargins[1] === 24 && initialParagraphRightMargins[2] === undefined && bulletRightMarginCoexistence && updatedParagraphRightMargins[0] === 6 && updatedParagraphRightMargins[1] === 0 && updatedParagraphRightMargins[2] === undefined && updatedParagraphRightMargins[3] === undefined && updatedParagraphRightMargins[4] === 9,
  paragraphIndent: initialParagraphIndents[0] === 24 && initialParagraphIndents[1] === -18 && initialParagraphIndents[2] === undefined && initialParagraphIndents[3] === undefined && bulletIndentIsolation && updatedParagraphIndents[0] === 6 && updatedParagraphIndents[1] === -6 && updatedParagraphIndents[2] === 0 && updatedParagraphIndents[3] === undefined && updatedParagraphIndents[4] === undefined,
  richTextTransparency: initialTransparencies[0] === 25 && initialTransparencies[1] === 50.555 && initialTransparencies[2] === 100 && initialTransparencies[3] === 60 && updatedTransparencies[0] === 0 && updatedTransparencies[1] === 75 && updatedTransparencies[2] === undefined,
  tableCreation,
  tableCellObjectCreation,
  tableCellFillCreation,
  tableFillCreation,
  tableCellBorderCreation,
  tableBorderCreation,
  tableCellMarginCreation,
  tableMarginCreation,
  tableCellVerticalAlignmentCreation,
  tableVerticalAlignmentCreation,
  tableCellHorizontalAlignmentCreation,
  tableHorizontalAlignmentCreation,
  tableCellTextDirectionCreation,
  tableTextDirectionCreation,
  tableColumnWidths,
  tableColumnWidthEditing,
  tableRowHeights,
  tableRowHeightEditing,
  tableCellTextDirection: table instanceof TableModel && initialCellDirection === 'horz' && rotatedCellDirection === 'vert270' && stackedCellDirection === 'wordArtVert' && horizontalCellDirection === 'horz' && clearedCellDirection === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].textDirection === 'vert' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellTextFitCreation,
  tableCellTextFit: table instanceof TableModel && initialCellFit === 'none' && shrinkCellFit === 'shrink' && sameFitPreserved && resizeCellFit === 'resize' && noneClearedCellFit === undefined && undefinedClearedCellFit === undefined && table.rows[0].cells[0].textFit === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].textFit === 'resize' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellVerticalAlignment: table instanceof TableModel && initialCellAlignment === 'middle' && topCellAlignment === 'top' && middleCellAlignment === 'middle' && bottomCellAlignment === 'bottom' && clearedCellAlignment === undefined && table.rows[0].cells[0].verticalAlignment === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].verticalAlignment === 'bottom' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellHorizontalAlignmentEditing: table instanceof TableModel && initialHorizontalAlignment === 'center' && initialNeighborHorizontalAlignment === undefined && leftHorizontalAlignment === 'left' && centerHorizontalAlignment === 'center' && rightHorizontalAlignment === 'right' && justifyHorizontalAlignment === 'justify' && clearedHorizontalAlignment === undefined && table.rows[0].cells[0].horizontalAlignment === undefined && table.rows[0].cells[1].horizontalAlignment === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].text === 'Neighbor' && reopenedClearedHorizontalAlignment === undefined && reopenedNeighborHorizontalAlignment === undefined,
  tableCellMargins: table instanceof TableModel && initialCellMargins?.top === 3 && initialCellMargins?.right === 2 && initialCellMargins?.bottom === 4 && initialCellMargins?.left === 1 && uniformCellMargins?.top === 4 && uniformCellMargins?.right === 4 && uniformCellMargins?.bottom === 4 && uniformCellMargins?.left === 4 && tupleCellMargins?.top === 1 && tupleCellMargins?.right === 2 && tupleCellMargins?.bottom === 3 && tupleCellMargins?.left === 4 && partialCellMargins?.top === 5 && partialCellMargins?.right === undefined && partialCellMargins?.bottom === undefined && partialCellMargins?.left === 7 && clearedCellMargins === undefined && table.rows[0].cells[0].margins === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].margins?.top === 7 && table.rows[0].cells[1].margins?.right === 6 && table.rows[0].cells[1].margins?.bottom === 8 && table.rows[0].cells[1].margins?.left === 5 && table.rows[0].cells[1].text === 'Neighbor',
  tableCellBorders: table instanceof TableModel && initialCellBorders?.top?.kind === 'line' && initialCellBorders.top.color.kind === 'scheme' && initialCellBorders.top.color.value === 'accent2' && initialCellBorders.top.width === 1.5 && initialCellBorders.top.style === 'dash' && initialCellBorders.right?.kind === 'none' && initialCellBorders.bottom?.kind === 'none' && initialCellBorders.left?.kind === 'line' && initialCellBorders.left.color.kind === 'srgb' && initialCellBorders.left.color.value === 'FF0000' && initialCellBorders.left.width === 1 && initialCellBorders.left.style === 'solid' && scalarCellBorders?.top?.kind === 'line' && scalarCellBorders.top.color.kind === 'srgb' && scalarCellBorders.top.color.value === '0000FF' && scalarCellBorders.top.width === 2 && scalarCellBorders.top.style === 'solid' && scalarCellBorders.right?.kind === 'line' && scalarCellBorders.bottom?.kind === 'line' && scalarCellBorders.left?.kind === 'line' && tupleCellBorders?.top?.kind === 'line' && tupleCellBorders.top.color.kind === 'scheme' && tupleCellBorders.top.color.value === 'accent1' && tupleCellBorders.top.width === 1.5 && tupleCellBorders.top.style === 'dash' && tupleCellBorders.right?.kind === 'line' && tupleCellBorders.right.color.kind === 'srgb' && tupleCellBorders.right.color.value === '00FF00' && tupleCellBorders.right.width === 0 && tupleCellBorders.right.style === undefined && tupleCellBorders.bottom?.kind === 'none' && tupleCellBorders.left === undefined && partialCellBorders?.left?.kind === 'none' && partialCellBorders.top === undefined && partialCellBorders.right === undefined && partialCellBorders.bottom === undefined && clearedCellBorders === undefined && table.rows[0].cells[0].borders === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].borders?.left?.kind === 'line' && table.rows[0].cells[1].borders.left.color.kind === 'srgb' && table.rows[0].cells[1].borders.left.color.value === '333333' && table.rows[0].cells[1].borders.left.width === 2 && table.rows[0].cells[1].borders.left.style === 'solid' && table.rows[0].cells[1].text === 'Neighbor',
  tableCellFill: table instanceof TableModel && initialCellFill?.kind === 'solid' && initialCellFill.color.kind === 'scheme' && initialCellFill.color.value === 'accent1' && initialCellFill.transparency === 25 && opaqueCellFill?.kind === 'solid' && opaqueCellFill.color.kind === 'srgb' && opaqueCellFill.color.value === 'FF0000' && opaqueCellFill.transparency === undefined && explicitOpaqueCellFill?.kind === 'solid' && explicitOpaqueCellFill.color.kind === 'scheme' && explicitOpaqueCellFill.color.value === 'accent2' && explicitOpaqueCellFill.transparency === 0 && fractionalCellFill?.kind === 'solid' && fractionalCellFill.color.kind === 'srgb' && fractionalCellFill.color.value === '112233' && fractionalCellFill.transparency === 33.333 && noneCellFill?.kind === 'none' && clearedCellFill === undefined && table.rows[0].cells[0].fill === undefined && table.rows[0].cells[0].text === 'Target' && table.rows[0].cells[1].fill?.kind === 'solid' && table.rows[0].cells[1].fill.color.kind === 'srgb' && table.rows[0].cells[1].fill.color.value === '70AD47' && table.rows[0].cells[1].fill.transparency === 50 && table.rows[0].cells[1].text === 'Neighbor',
  createText: createdText.text === 'Updated\\nParagraph' && initialTextWrap === false && updatedTextWrap === true && createdText.textWrap === undefined && initialTextDirection === 'vert270' && updatedTextDirection === 'wordArtVert' && createdText.textDirection === undefined && initialTextFit === 'shrink' && updatedTextFit === 'resize' && createdText.textFit === undefined && createdText.verticalAlignment === 'bottom' && createdText.textMargins.top === 4 && createdText.textMargins.left === 8 && createdText.textMargins.right === undefined && createdText.richText.every(({ align, bullet, level, rtl, spacing, tabStops }) => align === 'center' && bullet?.kind === 'bullet' && bullet.indent === 27 && level === 2 && rtl === true && spacing?.line?.kind === 'exact' && Array.isArray(tabStops) && tabStops[0]?.position === 1.25 && tabStops[1]?.alignment === 'right') && created.slides[0].shapes[0] === createdText,
  richText: inheritedLanguage === 'fr-CA' && localLanguage === 'de-DE' && initialRtl[0] === true && initialRtl[1] === false && richText.text === 'Updated rich' && richText.richText[0].rtl === undefined && richText.richText[0].align === 'justify' && richText.richText[0].bullet.style === 'romanUcPeriod' && richText.richText[0].level === 3 && richText.richText[0].spacing.line.kind === 'exact' && Array.isArray(richText.richText[0].tabStops) && richText.richText[0].tabStops[0].alignment === 'decimal' && richText.richText[0].runs[0].style.lang === 'ja-JP' && richText.richText[0].runs[0].style.baseline === 'superscript' && richText.richText[0].runs[0].style.characterSpacing === 2.5 && richText.richText[0].runs[0].style.italic === true && richText.richText[0].runs[0].style.glow.color.value === 'accent3' && richText.richText[0].runs[0].style.glow.opacity === 0.25 && richText.richText[0].runs[0].style.glow.size === 6 && richText.richText[0].runs[0].style.highlight.value === '00FF00' && richText.richText[0].runs[0].style.outline.color.value === 'accent1' && richText.richText[0].runs[0].style.outline.size === 0.75 && richText.richText[0].runs[0].style.underline.style === 'wavyHeavy' && richText.richText[0].runs[0].style.underline.color.value === 'accent2' && richText.richText[0].runs[0].style.strike === false,
  customSlideSize: custom.slideSize.width === inches(10) && customXml.includes('<p:sldSz cx="9144000" cy="6858000"/>'),
  GradientCodec: typeof GradientCodec === 'function',
  importPptxGenJS: typeof importPptxGenJS === 'function',
  transitions: typeof transitions.TransitionCodec === 'function',
  animations: typeof animations.AnimationTimingCodec === 'function',
  advancedCharts: typeof advancedCharts.AdvancedChartCodec === 'function',
  smartArt: typeof smartArt.SmartArtDiagramCodec === 'function',
};
if (Object.values(checks).some((value) => !value)) throw new Error(JSON.stringify(checks));
process.stdout.write(JSON.stringify(checks));
`,
  );
  const apiResult = run(process.execPath, ['smoke.mjs'], directory);
  const apiChecks = JSON.parse(apiResult.stdout);

  await writeFile(
    join(directory, 'browser-smoke.mjs'),
    `import { inches, PRESET_SHAPE_TYPES, PptxDocument, TableModel, transitions, animations, advancedCharts, smartArt } from '@jiayunxie/pptx';
const resolved = import.meta.resolve('@jiayunxie/pptx');
if (!resolved.endsWith('/dist/browser.js')) throw new Error('Browser condition resolved to ' + resolved);
const checks = [PptxDocument, transitions.TransitionCodec, animations.AnimationTimingCodec, advancedCharts.AdvancedChartCodec, smartArt.SmartArtDiagramCodec];
if (checks.some((value) => typeof value !== 'function')) throw new Error('Browser API surface is incomplete');
if (PRESET_SHAPE_TYPES.length !== 178 || !Object.isFrozen(PRESET_SHAPE_TYPES)) {
  throw new Error('Browser preset catalog failed');
}
const browserShapeDeck = PptxDocument.create();
const browserShape = browserShapeDeck.addSlide().addShape('foldedCorner');
browserShape.presetType = 'star5';
const reopenedBrowserShape = await PptxDocument.open(await browserShapeDeck.writeBlob());
if (reopenedBrowserShape.slides[0]?.shapes[0]?.presetType !== 'star5') {
  throw new Error('Browser preset shape failed');
}
const created = PptxDocument.create({ rtlMode: true, slideSize: '16:9' });
const browserText = created.addSlide().addText('Browser\\nText', { align: 'center', fit: 'resize', valign: 'bottom', vert: 'vert', wrap: false, bullet: true, level: 2, margin: [0, 0, 0, 0], rtlMode: true, spacing: { line: { kind: 'multiple', factor: 1.25 } }, tabStops: [{ position: 1.25 }] });
if (browserText.textWrap !== false || browserText.verticalAlignment !== 'bottom' || browserText.textDirection !== 'vert' || browserText.textFit !== 'resize' || browserText.richText.some(({ rtl }) => rtl !== true) || browserText.richText[0].tabStops[0].position !== 1.25 || browserText.textMargins.top !== 0 || browserText.textMargins.right !== 0 || browserText.textMargins.bottom !== 0 || browserText.textMargins.left !== 0) throw new Error('Browser create-text API failed');
const browserRich = created.slides[0].addRichText([{ align: 'right', bullet: { kind: 'number', style: 'alphaUcPeriod' }, level: 3, spacing: { before: 4, after: 6 }, tabStops: [{ position: 2.5, alignment: 'decimal' }], runs: [{ text: 'Rich', style: { lang: 'ja-JP', baseline: 'subscript', characterSpacing: 0, bold: true, glow: { opacity: 0.75, size: 4 }, highlight: { kind: 'scheme', value: 'accent1' }, outline: { color: { kind: 'srgb', value: 'ff0000' }, size: 1.25 }, underline: { style: 'wavyDbl' }, strike: 'dblStrike' } }] }], { rtlMode: true }).richText[0];
if (browserRich.rtl !== true || browserRich.tabStops[0].alignment !== 'decimal' || browserRich.runs[0].style.lang !== 'ja-JP' || browserRich.runs[0].style.baseline !== 'subscript' || browserRich.runs[0].style.characterSpacing !== 0 || browserRich.runs[0].style.glow.color.value !== 'FFFFFF' || browserRich.runs[0].style.glow.opacity !== 0.75 || browserRich.runs[0].style.glow.size !== 4 || browserRich.runs[0].style.highlight.value !== 'accent1' || browserRich.runs[0].style.outline.color.value !== 'FF0000' || browserRich.runs[0].style.outline.size !== 1.25 || browserRich.runs[0].style.underline.style !== 'wavyDbl' || browserRich.runs[0].style.strike !== 'dblStrike') throw new Error('Browser rich-text API failed');
const browserTransparency = created.slides[0].addRichText([{ runs: [{ text: 'Half', style: { transparency: 50 } }] }]);
if (browserTransparency.richText[0].runs[0].style.transparency !== 50) throw new Error('Browser transparency create failed');
browserTransparency.richText = [{ runs: [{ text: 'Cleared' }] }];
if (browserTransparency.richText[0].runs[0].style?.transparency !== undefined) throw new Error('Browser transparency clear failed');
const browserMargin = created.slides[0].addRichText([{ marginLeft: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginLeft: 24 });
if (browserMargin.richText[0].marginLeft !== 12) throw new Error('Browser paragraph margin create failed');
browserMargin.richText = [{ marginLeft: false, runs: [{ text: 'Cleared' }] }];
if (browserMargin.richText[0].marginLeft !== undefined) throw new Error('Browser paragraph margin clear failed');
const browserRightMargin = created.slides[0].addRichText([{ bullet: true, marginRight: 12, runs: [{ text: 'Margin' }] }], { paragraphMarginRight: 24 });
if (browserRightMargin.richText[0].marginRight !== 12 || browserRightMargin.richText[0].bullet.indent !== 27) throw new Error('Browser paragraph right margin create failed');
browserRightMargin.richText = [{ marginRight: false, runs: [{ text: 'Cleared' }] }];
if (browserRightMargin.richText[0].marginRight !== undefined) throw new Error('Browser paragraph right margin clear failed');
const browserIndent = created.slides[0].addRichText([{ indent: -12, runs: [{ text: 'Indent' }] }, { bullet: true, indent: false, runs: [{ text: 'Bullet' }] }], { paragraphIndent: 24 });
if (browserIndent.richText[0].indent !== -12 || browserIndent.richText[1].indent !== undefined || browserIndent.richText[1].bullet.indent !== 27) throw new Error('Browser paragraph indent create failed');
browserIndent.richText = [{ indent: false, runs: [{ text: 'Cleared' }] }];
if (browserIndent.richText[0].indent !== undefined) throw new Error('Browser paragraph indent clear failed');
const tableSlide = created.slides[0];
const browserCreationColor = { kind: 'srgb', value: '#D9EAF7' };
const browserCreationFill = { kind: 'solid', color: browserCreationColor, transparency: 33.3334 };
const browserTableFillColor = { kind: 'scheme', value: 'accent4' };
const browserTableFill = { kind: 'solid', color: browserTableFillColor, transparency: 40 };
const browserTableBorderColor = { kind: 'scheme', value: 'accent4' };
const browserTableBorder = { kind: 'line', color: browserTableBorderColor, width: 1.5, style: 'dash' };
const browserCreationBorderColor = { kind: 'srgb', value: '#C00000' };
const browserCreationBorder = { kind: 'line', color: browserCreationBorderColor, width: 2, style: 'solid' };
const browserCreationMargin = { top: 4, left: 8 };
const createdTable = tableSlide.addTable([
  [
    { text: 'Region', options: { align: 'left', border: browserCreationBorder, fill: browserCreationFill, fit: 'shrink', margin: browserCreationMargin, textDirection: 'vert', valign: 'top' } },
    { text: 'Revenue', options: {
      border: [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, undefined, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
      fit: 'resize',
      margin: [1, 2, 3, 4],
      textDirection: 'vert270',
      valign: 'middle',
    } },
  ],
  [
    { text: 'West', options: { align: 'right', border: { top: { kind: 'line', color: { kind: 'scheme', value: 'accent3' }, width: 1, style: 'dash' }, left: { kind: 'none' } }, fit: 'none', margin: 0, textDirection: 'wordArtVert', valign: 'bottom' } },
    { text: '', options: { align: 'justify', border: { kind: 'none' }, fill: { kind: 'none' }, margin: {}, textDirection: 'horz' } },
  ],
], { name: 'Created browser table', align: 'center', columnWidths: inches(1.25), rowHeights: inches(0.75), fill: browserTableFill, margin: { top: 9, left: 18 }, valign: 'middle' });
const browserTableDefaultsSlide = created.addSlide();
const browserTableBorderDefault = browserTableDefaultsSlide.addTable([[
  'Inherited border',
  { text: 'None override', options: { border: { kind: 'none' } } },
]], { name: 'Created browser table border default', columnWidths: inches(1), rowHeights: inches(0.5), border: browserTableBorder });
const browserTableDirectionDefault = browserTableDefaultsSlide.addTable([[
  'Inherited direction',
  { text: 'Horizontal override', options: { textDirection: 'horz' } },
]], { name: 'Created browser table direction default', columnWidths: inches(1), rowHeights: inches(0.5), textDirection: 'vert270' });
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ text }) => text))) !== JSON.stringify([['Region', 'Revenue'], ['West', '']])) throw new Error('Browser table cell object creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ verticalAlignment }) => verticalAlignment))) !== JSON.stringify([['top', 'middle'], ['bottom', 'middle']])) throw new Error('Browser table vertical alignment creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))) !== JSON.stringify([['vert', 'vert270'], ['wordArtVert', undefined]])) throw new Error('Browser table cell text direction creation failed');
if (JSON.stringify(createdTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))) !== JSON.stringify([['shrink', 'resize'], [undefined, undefined]])) throw new Error('Browser table cell text fit creation failed');
if (JSON.stringify(browserTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection)) !== JSON.stringify(['vert270', undefined])) throw new Error('Browser table text direction creation failed');
const browserMarginVector = (margins) => [margins?.top, margins?.right, margins?.bottom, margins?.left];
const browserInitialMargins = createdTable.rows.map(({ cells }) => cells.map(({ margins }) => browserMarginVector(margins)));
if (JSON.stringify(browserInitialMargins) !== JSON.stringify([[[4, 7.2, 3.6, 8], [1, 2, 3, 4]], [[0, 0, 0, 0], [9, 7.2, 3.6, 18]]])) throw new Error('Browser table cell margin creation failed');
if (createdTable.rows[0].cells[0].fill?.kind !== 'solid' || createdTable.rows[0].cells[0].fill.color.kind !== 'srgb' || createdTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || createdTable.rows[0].cells[0].fill.transparency !== 33.333 || createdTable.rows[0].cells[1].fill?.kind !== 'solid' || createdTable.rows[0].cells[1].fill.color.kind !== 'scheme' || createdTable.rows[0].cells[1].fill.color.value !== 'accent2' || createdTable.rows[0].cells[1].fill.transparency !== 25 || createdTable.rows[1].cells[1].fill?.kind !== 'none') throw new Error('Browser table cell fill creation failed');
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent4' || createdTable.rows[1].cells[0].fill.transparency !== 40) throw new Error('Browser table fill inheritance failed');
const browserCreationSides = ['top', 'right', 'bottom', 'left'];
const browserIsCreationLine = (border, colorKind, colorValue, width, style) => border?.kind === 'line' && border.color.kind === colorKind && border.color.value === colorValue && border.width === width && border.style === style;
const browserAllCreationLines = (borders, colorKind, colorValue, width, style) => borders !== undefined && browserCreationSides.every((side) => browserIsCreationLine(borders[side], colorKind, colorValue, width, style));
const browserAllCreationNone = (borders) => borders !== undefined && browserCreationSides.every((side) => borders[side]?.kind === 'none');
const browserScalarBorders = createdTable.rows[0].cells[0].borders;
const browserTupleBorders = createdTable.rows[0].cells[1].borders;
const browserNamedBorders = createdTable.rows[1].cells[0].borders;
if (!browserAllCreationLines(browserScalarBorders, 'srgb', 'C00000', 2, 'solid') || !browserIsCreationLine(browserTupleBorders?.top, 'scheme', 'accent1', 1.5, 'dash') || browserTupleBorders?.right?.kind !== 'none' || !browserIsCreationLine(browserTupleBorders?.bottom, 'srgb', '00FF00', 0, undefined) || browserTupleBorders?.left?.kind !== 'none' || !browserIsCreationLine(browserNamedBorders?.top, 'scheme', 'accent3', 1, 'dash') || browserNamedBorders?.right?.kind !== 'none' || browserNamedBorders?.bottom?.kind !== 'none' || browserNamedBorders?.left?.kind !== 'none' || !browserAllCreationNone(createdTable.rows[1].cells[1].borders)) throw new Error('Browser table cell border creation failed');
const browserInitialTableDefaultBorders = browserTableBorderDefault.rows[0].cells[0].borders;
if (!browserAllCreationLines(browserInitialTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') || !browserAllCreationNone(browserTableBorderDefault.rows[0].cells[1].borders)) throw new Error('Browser table border inheritance failed');
browserCreationColor.value = '000000';
browserCreationFill.transparency = 1;
browserTableFillColor.value = 'accent6';
browserTableFill.transparency = 1;
browserTableBorderColor.value = 'accent6';
browserTableBorder.width = 9;
browserCreationBorderColor.value = '000000';
browserCreationBorder.width = 9;
browserCreationMargin.top = 99;
browserCreationMargin.left = 99;
if (createdTable.rows[0].cells[0].fill?.kind !== 'solid' || createdTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || createdTable.rows[0].cells[0].fill.transparency !== 33.333) throw new Error('Browser table cell fill creation retained source state');
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent4' || createdTable.rows[1].cells[0].fill.transparency !== 40) throw new Error('Browser table fill creation retained source state');
if (!browserAllCreationLines(createdTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid')) throw new Error('Browser table cell border creation retained source state');
if (!browserAllCreationLines(browserTableBorderDefault.rows[0].cells[0].borders, 'scheme', 'accent4', 1.5, 'dash')) throw new Error('Browser table border creation retained source state');
if (JSON.stringify(browserMarginVector(createdTable.rows[0].cells[0].margins)) !== JSON.stringify([4, 7.2, 3.6, 8])) throw new Error('Browser table cell margin creation retained source state');
const createdTablePartXml = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
const browserCreatedTableCells = [...createdTablePartXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0]);
const browserCreatedTableHorizontalAlignments = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const browserCreatedTableDirections = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const browserCreatedTableFits = browserCreatedTableCells.map((cellXml) => cellXml.match(/<a:(normAutofit|spAutoFit)\\/>/)?.[1]);
const createdTableGrid = [...createdTablePartXml.matchAll(/<a:gridCol w="(\\d+)"\\/>/g)].map((match) => Number(match[1]));
const createdTableRows = [...createdTablePartXml.matchAll(/<a:tr h="(\\d+)">/g)].map((match) => Number(match[1]));
if (!createdTablePartXml.includes('<a:tcPr marL="228600" marR="91440" marT="114300" marB="45720" anchor="ctr">')) throw new Error('Browser table margin XML creation failed');
if (JSON.stringify(browserCreatedTableHorizontalAlignments) !== JSON.stringify(['l', 'ctr', 'r', 'just'])) throw new Error('Browser table horizontal alignment creation failed');
if (JSON.stringify(browserCreatedTableDirections) !== JSON.stringify(['vert', 'vert270', 'wordArtVert', undefined])) throw new Error('Browser table cell text direction XML creation failed');
if (JSON.stringify(browserCreatedTableFits) !== JSON.stringify(['normAutofit', 'spAutoFit', undefined, undefined])) throw new Error('Browser table cell text fit XML creation failed');
if (!(createdTable instanceof TableModel) || createdTable.transform.x !== inches(0.5) || createdTable.transform.y !== inches(0.5) || createdTable.transform.width !== inches(2.5) || createdTable.transform.height !== inches(1.5) || createdTableGrid.length !== 2 || createdTableGrid.some((width) => width !== inches(1.25)) || createdTableRows.length !== 2 || createdTableRows.some((height) => height !== inches(0.75)) || createdTable.rows[1].cells[1].margins?.top !== 9 || createdTable.rows[1].cells[1].margins?.left !== 18) throw new Error('Browser table sizing creation failed');
createdTable.setColumnWidths([inches(1), inches(1.5)]);
if (createdTable.columnWidths?.join(',') !== [inches(1), inches(1.5)].join(',') || createdTable.transform.width !== inches(2.5)) throw new Error('Browser table column-width editing failed');
createdTable.setRowHeights([inches(0.5), inches(1)]);
if (createdTable.rowHeights?.join(',') !== [inches(0.5), inches(1)].join(',') || createdTable.transform.height !== inches(1.5)) throw new Error('Browser table row-height editing failed');
createdTable.setCellText(1, 0, 'Edited West');
createdTable.setCellFill(1, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 50 });
if (createdTable.rows[1].cells[0].fill?.kind !== 'solid' || createdTable.rows[1].cells[0].fill.color.kind !== 'scheme' || createdTable.rows[1].cells[0].fill.color.value !== 'accent1' || createdTable.rows[1].cells[0].fill.transparency !== 50) throw new Error('Browser table fill override failed');
createdTable.setCellBorders(1, 0, { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, width: 1, style: 'solid' });
createdTable.setCellVerticalAlignment(0, 0, 'bottom');
createdTable.setCellVerticalAlignment(1, 1, undefined);
createdTable.setCellMargins(0, 0, { bottom: 9 });
createdTable.setCellMargins(1, 1, undefined);
createdTable.setCellFill(1, 1, undefined);
if (createdTable.rows[1].cells[1].fill !== undefined) throw new Error('Browser table fill clear re-inherited');
browserTableBorderDefault.setCellBorders(0, 0, undefined);
if (browserTableBorderDefault.rows[0].cells[0].borders !== undefined) throw new Error('Browser table border clear re-inherited');
const reopenedCreated = await PptxDocument.open(await created.write());
const reopenedCreatedTable = reopenedCreated.slides[0].shapes.find((shape) => shape.name === 'Created browser table');
const reopenedCreatedTableXml = new TextDecoder().decode(reopenedCreated.opcPackage.requirePart(reopenedCreated.slides[0].partUri).bytes);
const reopenedCreatedTableHorizontalAlignments = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:pPr[^>]*\\salgn="([^"]+)"/)?.[1]);
const reopenedCreatedTableDirections = [...reopenedCreatedTableXml.matchAll(/<a:tc(?:\\s[^>]*)?>[\\s\\S]*?<\\/a:tc>/g)].map((match) => match[0].match(/<a:tcPr[^>]*\\svert="([^"]+)"/)?.[1]);
const reopenedBrowserTableBorderDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created browser table border default');
const reopenedBrowserTableDirectionDefault = reopenedCreated.slides
  .flatMap(({ shapes }) => shapes)
  .find((shape) => shape.name === 'Created browser table direction default');
if (!(reopenedCreatedTable instanceof TableModel) || reopenedCreatedTable.columnWidths?.join(',') !== [inches(1), inches(1.5)].join(',') || reopenedCreatedTable.rowHeights?.join(',') !== [inches(0.5), inches(1)].join(',') || reopenedCreatedTable.transform.width !== inches(2.5) || reopenedCreatedTable.transform.height !== inches(1.5) || reopenedCreatedTable.rows[1].cells[0].text !== 'Edited West' || reopenedCreatedTable.rows[1].cells[1].text !== '' || reopenedCreatedTable.rows[0].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[0].cells[0].fill.color.value !== 'D9EAF7' || reopenedCreatedTable.rows[0].cells[1].fill?.kind !== 'solid' || reopenedCreatedTable.rows[0].cells[1].fill.color.kind !== 'scheme' || reopenedCreatedTable.rows[0].cells[1].fill.transparency !== 25 || reopenedCreatedTable.rows[1].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[1].cells[0].fill.color.value !== 'accent1' || reopenedCreatedTable.rows[1].cells[0].fill.transparency !== 50 || reopenedCreatedTable.rows[1].cells[1].fill !== undefined || !browserAllCreationLines(reopenedCreatedTable.rows[0].cells[0].borders, 'srgb', 'C00000', 2, 'solid') || !browserIsCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.top, 'scheme', 'accent1', 1.5, 'dash') || reopenedCreatedTable.rows[0].cells[1].borders?.right?.kind !== 'none' || !browserIsCreationLine(reopenedCreatedTable.rows[0].cells[1].borders?.bottom, 'srgb', '00FF00', 0, undefined) || reopenedCreatedTable.rows[0].cells[1].borders?.left?.kind !== 'none' || !browserAllCreationLines(reopenedCreatedTable.rows[1].cells[0].borders, 'srgb', 'FFFFFF', 1, 'solid') || !browserAllCreationNone(reopenedCreatedTable.rows[1].cells[1].borders)) throw new Error('Browser table creation round trip failed');
if (reopenedCreatedTable.rows[1].cells[0].fill?.kind !== 'solid' || reopenedCreatedTable.rows[1].cells[0].fill.color.kind !== 'scheme' || reopenedCreatedTable.rows[1].cells[0].fill.color.value !== 'accent1' || reopenedCreatedTable.rows[1].cells[0].fill.transparency !== 50 || reopenedCreatedTable.rows[1].cells[1].fill !== undefined) throw new Error('Browser table fill round trip failed');
if (!(reopenedBrowserTableBorderDefault instanceof TableModel) || reopenedBrowserTableBorderDefault.rows[0].cells[0].borders !== undefined || !browserAllCreationNone(reopenedBrowserTableBorderDefault.rows[0].cells[1].borders)) throw new Error('Browser table border round trip failed');
if (reopenedCreatedTable.rows[0].cells[0].margins?.top !== undefined || reopenedCreatedTable.rows[0].cells[0].margins?.right !== undefined || reopenedCreatedTable.rows[0].cells[0].margins?.bottom !== 9 || reopenedCreatedTable.rows[0].cells[0].margins?.left !== undefined) throw new Error('Browser table cell margin creation round trip failed');
if (reopenedCreatedTable.rows[1].cells[1].margins !== undefined) throw new Error('Browser table margin clear re-inherited');
if (reopenedCreatedTable.rows[0].cells[0].verticalAlignment !== 'bottom' || reopenedCreatedTable.rows[0].cells[1].verticalAlignment !== 'middle' || reopenedCreatedTable.rows[1].cells[0].verticalAlignment !== 'bottom' || reopenedCreatedTable.rows[1].cells[1].verticalAlignment !== undefined) throw new Error('Browser table cell vertical alignment creation round trip failed');
if (JSON.stringify(reopenedCreatedTableHorizontalAlignments) !== JSON.stringify(['l', 'ctr', 'r', 'just'])) throw new Error('Browser table horizontal alignment creation round trip failed');
if (JSON.stringify(reopenedCreatedTable.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))) !== JSON.stringify([['vert', 'vert270'], ['wordArtVert', undefined]]) || JSON.stringify(reopenedCreatedTableDirections) !== JSON.stringify(['vert', 'vert270', 'wordArtVert', undefined])) throw new Error('Browser table cell text direction creation round trip failed');
if (!(reopenedBrowserTableDirectionDefault instanceof TableModel) || JSON.stringify(reopenedBrowserTableDirectionDefault.rows[0].cells.map(({ textDirection }) => textDirection)) !== JSON.stringify(['vert270', undefined])) throw new Error('Browser table text direction creation round trip failed');
const tablePart = created.opcPackage.requirePart(tableSlide.partUri);
const tableXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Browser table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:bodyPr custom="TARGET"><a:noAutofit/></a:bodyPr><a:p><a:r><a:t>Browser target</a:t></a:r></a:p></a:txBody><a:tcPr vert="horz" anchor="ctr" marL="12700" marR="25400" marT="38100" marB="50800"><a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="sysDash"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr custom="NEIGHBOR"><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>Browser neighbor</a:t></a:r></a:p></a:txBody><a:tcPr vert="vert" anchor="b" marL="63500" marR="76200" marT="88900" marB="101600"><a:lnL w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
created.opcPackage.setPart(tableSlide.partUri, new TextDecoder().decode(tablePart.bytes).replace('</p:spTree>', tableXml + '</p:spTree>'), tablePart.contentType);
const table = tableSlide.shapes.find((shape) => shape.name === 'Browser table');
if (!(table instanceof TableModel) || table.rows[0].cells[0].textFit !== 'none' || table.rows[0].cells[1].textFit !== 'resize') throw new Error('Browser table-cell fit read failed');
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined) {
  throw new Error('Browser table-cell horizontal alignment initial read failed');
}
table.setCellHorizontalAlignment(0, 0, 'left');
if (table.rows[0].cells[0].horizontalAlignment !== 'left') {
  throw new Error('Browser table-cell left alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'center');
if (table.rows[0].cells[0].horizontalAlignment !== 'center') {
  throw new Error('Browser table-cell center alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'right');
if (table.rows[0].cells[0].horizontalAlignment !== 'right') {
  throw new Error('Browser table-cell right alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'justify');
if (table.rows[0].cells[0].horizontalAlignment !== 'justify') {
  throw new Error('Browser table-cell justify alignment failed');
}
table.setCellHorizontalAlignment(0, 0, undefined);
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined ||
    table.rows[0].cells[0].text !== 'Browser target' ||
    table.rows[0].cells[1].text !== 'Browser neighbor') {
  throw new Error('Browser table-cell horizontal alignment clear failed');
}
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'scheme' || table.rows[0].cells[0].borders.top.color.value !== 'accent2' || table.rows[0].cells[0].borders.top.width !== 1.5 || table.rows[0].cells[0].borders.top.style !== 'dash' || table.rows[0].cells[0].borders.right?.kind !== 'none' || table.rows[0].cells[0].borders.bottom?.kind !== 'none' || table.rows[0].cells[0].borders.left?.kind !== 'line') throw new Error('Browser table-cell border read failed');
table.setCellBorders(0, 0, { kind: 'line', color: { kind: 'srgb', value: '#0000FF' }, width: 2, style: 'solid' });
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'srgb' || table.rows[0].cells[0].borders.top.color.value !== '0000FF' || table.rows[0].cells[0].borders.top.width !== 2 || table.rows[0].cells[0].borders.top.style !== 'solid' || table.rows[0].cells[0].borders.right?.kind !== 'line' || table.rows[0].cells[0].borders.bottom?.kind !== 'line' || table.rows[0].cells[0].borders.left?.kind !== 'line') throw new Error('Browser table-cell scalar border failed');
table.setCellBorders(0, 0, [{ kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' }, { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 }, { kind: 'none' }, undefined]);
if (table.rows[0].cells[0].borders?.top?.kind !== 'line' || table.rows[0].cells[0].borders.top.color.kind !== 'scheme' || table.rows[0].cells[0].borders.top.color.value !== 'accent1' || table.rows[0].cells[0].borders.top.style !== 'dash' || table.rows[0].cells[0].borders.right?.kind !== 'line' || table.rows[0].cells[0].borders.right.width !== 0 || table.rows[0].cells[0].borders.right.style !== undefined || table.rows[0].cells[0].borders.bottom?.kind !== 'none' || table.rows[0].cells[0].borders.left !== undefined) throw new Error('Browser table-cell tuple border failed');
table.setCellBorders(0, 0, { left: { kind: 'none' } });
if (table.rows[0].cells[0].borders?.left?.kind !== 'none' || table.rows[0].cells[0].borders.top !== undefined || table.rows[0].cells[0].borders.right !== undefined || table.rows[0].cells[0].borders.bottom !== undefined) throw new Error('Browser table-cell partial border failed');
table.setCellBorders(0, 0, undefined);
if (table.rows[0].cells[0].borders !== undefined || table.rows[0].cells[1].borders?.left?.kind !== 'line' || table.rows[0].cells[1].borders.left.color.kind !== 'srgb' || table.rows[0].cells[1].borders.left.color.value !== '333333' || table.rows[0].cells[1].borders.left.width !== 2 || table.rows[0].cells[1].borders.left.style !== 'solid') throw new Error('Browser table-cell border clear failed');
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'scheme' || table.rows[0].cells[0].fill.color.value !== 'accent1' || table.rows[0].cells[0].fill.transparency !== 25) throw new Error('Browser table-cell fill read failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '#FF0000' } });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'srgb' || table.rows[0].cells[0].fill.color.value !== 'FF0000' || table.rows[0].cells[0].fill.transparency !== undefined) throw new Error('Browser table-cell opaque fill failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 0 });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'scheme' || table.rows[0].cells[0].fill.color.value !== 'accent2' || table.rows[0].cells[0].fill.transparency !== 0) throw new Error('Browser table-cell explicit opaque fill failed');
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'srgb', value: '112233' }, transparency: 33.333 });
if (table.rows[0].cells[0].fill?.kind !== 'solid' || table.rows[0].cells[0].fill.color.kind !== 'srgb' || table.rows[0].cells[0].fill.transparency !== 33.333) throw new Error('Browser table-cell fractional fill failed');
table.setCellFill(0, 0, { kind: 'none' });
if (table.rows[0].cells[0].fill?.kind !== 'none') throw new Error('Browser table-cell no fill failed');
table.setCellFill(0, 0, undefined);
if (table.rows[0].cells[0].fill !== undefined || table.rows[0].cells[1].fill?.kind !== 'solid' || table.rows[0].cells[1].fill.color.kind !== 'srgb' || table.rows[0].cells[1].fill.color.value !== '70AD47' || table.rows[0].cells[1].fill.transparency !== 50) throw new Error('Browser table-cell fill clear failed');
if (table.rows[0].cells[0].margins?.top !== 3 || table.rows[0].cells[0].margins?.right !== 2 || table.rows[0].cells[0].margins?.bottom !== 4 || table.rows[0].cells[0].margins?.left !== 1) throw new Error('Browser table-cell margin read failed');
table.setCellMargins(0, 0, 4);
if (table.rows[0].cells[0].margins?.top !== 4 || table.rows[0].cells[0].margins?.right !== 4 || table.rows[0].cells[0].margins?.bottom !== 4 || table.rows[0].cells[0].margins?.left !== 4) throw new Error('Browser table-cell scalar margin failed');
table.setCellMargins(0, 0, [1, 2, 3, 4]);
if (table.rows[0].cells[0].margins?.top !== 1 || table.rows[0].cells[0].margins?.right !== 2 || table.rows[0].cells[0].margins?.bottom !== 3 || table.rows[0].cells[0].margins?.left !== 4) throw new Error('Browser table-cell tuple margin failed');
table.setCellMargins(0, 0, { top: 5, left: 7 });
if (table.rows[0].cells[0].margins?.top !== 5 || table.rows[0].cells[0].margins?.right !== undefined || table.rows[0].cells[0].margins?.bottom !== undefined || table.rows[0].cells[0].margins?.left !== 7) throw new Error('Browser table-cell partial margin failed');
table.setCellMargins(0, 0, undefined);
if (table.rows[0].cells[0].margins !== undefined || table.rows[0].cells[1].margins?.top !== 7 || table.rows[0].cells[1].margins?.right !== 6 || table.rows[0].cells[1].margins?.bottom !== 8 || table.rows[0].cells[1].margins?.left !== 5) throw new Error('Browser table-cell margin clear failed');
table.setCellTextFit(0, 0, 'shrink');
const beforeSameFit = new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes);
table.setCellTextFit(0, 0, 'shrink');
if (table.rows[0].cells[0].textFit !== 'shrink' || new TextDecoder().decode(created.opcPackage.requirePart(tableSlide.partUri).bytes) !== beforeSameFit) throw new Error('Browser table-cell shrink fit failed');
table.setCellTextFit(0, 0, 'resize');
if (table.rows[0].cells[0].textFit !== 'resize') throw new Error('Browser table-cell resize fit failed');
table.setCellTextFit(0, 0, 'none');
if (table.rows[0].cells[0].textFit !== undefined) throw new Error('Browser table-cell none fit clear failed');
table.setCellTextFit(0, 0, 'shrink');
table.setCellTextFit(0, 0, undefined);
if (table.rows[0].cells[0].textFit !== undefined || table.rows[0].cells[1].textFit !== 'resize') throw new Error('Browser table-cell undefined fit clear failed');
if (table.rows[0].cells[0].verticalAlignment !== 'middle' || table.rows[0].cells[1].verticalAlignment !== 'bottom') throw new Error('Browser table-cell vertical alignment read failed');
table.setCellVerticalAlignment(0, 0, 'top');
if (table.rows[0].cells[0].verticalAlignment !== 'top') throw new Error('Browser table-cell top alignment failed');
table.setCellVerticalAlignment(0, 0, 'middle');
if (table.rows[0].cells[0].verticalAlignment !== 'middle') throw new Error('Browser table-cell middle alignment failed');
table.setCellVerticalAlignment(0, 0, 'bottom');
if (table.rows[0].cells[0].verticalAlignment !== 'bottom') throw new Error('Browser table-cell bottom alignment failed');
table.setCellVerticalAlignment(0, 0, undefined);
if (table.rows[0].cells[0].verticalAlignment !== undefined || table.rows[0].cells[1].verticalAlignment !== 'bottom') throw new Error('Browser table-cell vertical alignment clear failed');
table?.setCellTextDirection(0, 0, 'vert270');
if (!(table instanceof TableModel) || table.rows[0].cells[0].textDirection !== 'vert270' || table.rows[0].cells[1].textDirection !== 'vert') throw new Error('Browser table-cell direction edit failed');
table.setCellTextDirection(0, 0, 'wordArtVert');
if (table.rows[0].cells[0].textDirection !== 'wordArtVert') throw new Error('Browser table-cell stacked direction edit failed');
table.setCellTextDirection(0, 0, undefined);
if (table.rows[0].cells[0].textDirection !== undefined || table.rows[0].cells[0].text !== 'Browser target' || table.rows[0].cells[1].text !== 'Browser neighbor') throw new Error('Browser table-cell direction clear failed');
const reopenedBrowserEdited = await PptxDocument.open(await created.write());
const reopenedBrowserTable = reopenedBrowserEdited.slides[0].shapes.find(
  (shape) => shape.name === 'Browser table',
);
if (!(reopenedBrowserTable instanceof TableModel) ||
    reopenedBrowserTable.rows[0].cells[0].horizontalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].horizontalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[0].text !== 'Browser target' ||
    reopenedBrowserTable.rows[0].cells[1].text !== 'Browser neighbor' ||
    reopenedBrowserTable.rows[0].cells[0].textDirection !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].textDirection !== 'vert' ||
    reopenedBrowserTable.rows[0].cells[0].textFit !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].textFit !== 'resize' ||
    reopenedBrowserTable.rows[0].cells[0].verticalAlignment !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].verticalAlignment !== 'bottom' ||
    reopenedBrowserTable.rows[0].cells[0].margins !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].margins?.left !== 5 ||
    reopenedBrowserTable.rows[0].cells[0].borders !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].borders?.left?.kind !== 'line' ||
    reopenedBrowserTable.rows[0].cells[0].fill !== undefined ||
    reopenedBrowserTable.rows[0].cells[1].fill?.kind !== 'solid') {
  throw new Error('Browser table-cell horizontal alignment reopen failed');
}
if (created.rtlMode !== true) throw new Error('Browser presentation RTL create failed');
created.rtlMode = false;
if (created.rtlMode !== false || browserRich.rtl !== true) throw new Error('Browser presentation RTL edit failed');
created.rtlMode = undefined;
if (created.rtlMode !== undefined || browserRich.rtl !== true) throw new Error('Browser presentation RTL clear failed');
const browserMetadata = PptxDocument.create({ title: 'Browser title' });
if (browserMetadata.title !== 'Browser title') throw new Error('Browser presentation title create failed');
browserMetadata.title = 'Edited browser title';
if (browserMetadata.title !== 'Edited browser title') throw new Error('Browser presentation title edit failed');
const reopenedBrowserMetadata = await PptxDocument.open(await browserMetadata.write());
if (reopenedBrowserMetadata.title !== 'Edited browser title') throw new Error('Browser presentation title reopen failed');
browserMetadata.title = '';
if (browserMetadata.title !== '') throw new Error('Browser presentation title empty failed');
browserMetadata.title = undefined;
if (browserMetadata.title !== undefined) throw new Error('Browser presentation title clear failed');
const browserAuthorship = PptxDocument.create({ author: 'Browser author' });
if (browserAuthorship.author !== 'Browser author') throw new Error('Browser presentation author create failed');
browserAuthorship.author = 'Edited browser author';
if (browserAuthorship.author !== 'Edited browser author') throw new Error('Browser presentation author edit failed');
const reopenedBrowserAuthorship = await PptxDocument.open(await browserAuthorship.write());
if (reopenedBrowserAuthorship.author !== 'Edited browser author') throw new Error('Browser presentation author reopen failed');
browserAuthorship.author = '';
if (browserAuthorship.author !== '') throw new Error('Browser presentation author empty failed');
browserAuthorship.author = undefined;
if (browserAuthorship.author !== undefined) throw new Error('Browser presentation author clear failed');
const browserEditorship = PptxDocument.create({ lastModifiedBy: 'Browser editor' });
if (browserEditorship.lastModifiedBy !== 'Browser editor') throw new Error('Browser presentation lastModifiedBy create failed');
browserEditorship.lastModifiedBy = 'Edited browser editor';
if (browserEditorship.lastModifiedBy !== 'Edited browser editor') throw new Error('Browser presentation lastModifiedBy edit failed');
const reopenedBrowserEditorship = await PptxDocument.open(await browserEditorship.write());
if (reopenedBrowserEditorship.lastModifiedBy !== 'Edited browser editor') throw new Error('Browser presentation lastModifiedBy reopen failed');
browserEditorship.lastModifiedBy = '';
if (browserEditorship.lastModifiedBy !== '') throw new Error('Browser presentation lastModifiedBy empty failed');
browserEditorship.lastModifiedBy = undefined;
if (browserEditorship.lastModifiedBy !== undefined) throw new Error('Browser presentation lastModifiedBy clear failed');
const browserChronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
});
if (browserChronology.createdAt !== '2024-02-29T12:34:56.123+05:30') throw new Error('Browser presentation createdAt create failed');
if (browserChronology.modifiedAt !== '2024-03-01T01:02:03.456+08:00') throw new Error('Browser presentation modifiedAt create failed');
browserChronology.createdAt = '2026-07-30T00:00:00Z';
if (browserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation createdAt edit failed');
browserChronology.modifiedAt = '2026-07-30T01:02:03Z';
if (browserChronology.modifiedAt !== '2026-07-30T01:02:03Z') throw new Error('Browser presentation modifiedAt edit failed');
const reopenedBrowserChronology = await PptxDocument.open(await browserChronology.write());
if (reopenedBrowserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation createdAt reopen failed');
if (reopenedBrowserChronology.modifiedAt !== '2026-07-30T01:02:03Z') throw new Error('Browser presentation modifiedAt reopen failed');
browserChronology.modifiedAt = undefined;
if (browserChronology.modifiedAt !== undefined) throw new Error('Browser presentation modifiedAt clear failed');
if (browserChronology.createdAt !== '2026-07-30T00:00:00Z') throw new Error('Browser presentation modifiedAt changed createdAt');
browserChronology.createdAt = undefined;
if (browserChronology.createdAt !== undefined) throw new Error('Browser presentation createdAt clear failed');
const browserSubjectMatter = PptxDocument.create({ subject: 'Browser subject' });
if (browserSubjectMatter.subject !== 'Browser subject') throw new Error('Browser presentation subject create failed');
browserSubjectMatter.subject = 'Edited browser subject';
if (browserSubjectMatter.subject !== 'Edited browser subject') throw new Error('Browser presentation subject edit failed');
const reopenedBrowserSubjectMatter = await PptxDocument.open(await browserSubjectMatter.write());
if (reopenedBrowserSubjectMatter.subject !== 'Edited browser subject') throw new Error('Browser presentation subject reopen failed');
browserSubjectMatter.subject = '';
if (browserSubjectMatter.subject !== '') throw new Error('Browser presentation subject empty failed');
browserSubjectMatter.subject = undefined;
if (browserSubjectMatter.subject !== undefined) throw new Error('Browser presentation subject clear failed');
const browserRevisioned = PptxDocument.create({ revision: '007' });
if (browserRevisioned.revision !== '007') throw new Error('Browser presentation revision create failed');
browserRevisioned.revision = '42';
if (browserRevisioned.revision !== '42') throw new Error('Browser presentation revision edit failed');
const reopenedBrowserRevisioned = await PptxDocument.open(await browserRevisioned.write());
if (reopenedBrowserRevisioned.revision !== '42') throw new Error('Browser presentation revision reopen failed');
browserRevisioned.revision = undefined;
if (browserRevisioned.revision !== undefined) throw new Error('Browser presentation revision clear failed');
const browserOrganization = PptxDocument.create({ company: 'Browser company' });
if (browserOrganization.company !== 'Browser company') throw new Error('Browser presentation company create failed');
browserOrganization.company = 'Edited browser company';
if (browserOrganization.company !== 'Edited browser company') throw new Error('Browser presentation company edit failed');
const reopenedBrowserOrganization = await PptxDocument.open(await browserOrganization.write());
if (reopenedBrowserOrganization.company !== 'Edited browser company') throw new Error('Browser presentation company reopen failed');
browserOrganization.company = '';
if (browserOrganization.company !== '') throw new Error('Browser presentation company empty failed');
browserOrganization.company = undefined;
if (browserOrganization.company !== undefined) throw new Error('Browser presentation company clear failed');
const browserThemed = PptxDocument.create({
  theme: { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' },
});
const browserCreatedTheme = browserThemed.theme;
if (browserCreatedTheme?.headFontFace !== 'Aptos Display' || browserCreatedTheme.bodyFontFace !== 'Aptos') throw new Error('Browser presentation theme create failed');
browserCreatedTheme.headFontFace = 'Detached caller value';
if (browserThemed.theme?.headFontFace !== 'Aptos Display') throw new Error('Browser presentation theme snapshot was not detached');
browserThemed.theme = { headFontFace: 'Noto Sans Display' };
if (browserThemed.theme?.headFontFace !== 'Noto Sans Display' || browserThemed.theme.bodyFontFace !== 'Calibri') throw new Error('Browser presentation theme replacement failed');
browserThemed.masterLayoutTheme.presentationTheme.setFonts({ minorLatin: 'Noto Sans' });
const reopenedBrowserThemed = await PptxDocument.open(await browserThemed.writeBlob());
if (reopenedBrowserThemed.theme?.headFontFace !== 'Noto Sans Display' || reopenedBrowserThemed.theme.bodyFontFace !== 'Noto Sans') throw new Error('Browser presentation theme reopen failed');
const browserSectioned = PptxDocument.create();
const browserIntro = browserSectioned.addSection({ title: 'Browser intro' });
const browserSectionSlide = browserSectioned.addSlide({ sectionTitle: 'Browser intro' });
const browserData = browserSectioned.addSection({ title: 'Browser data' });
browserSectioned.assignSlideToSection(0, browserData.id);
browserSectioned.renameSection(browserIntro.id, 'Browser edited');
browserSectioned.moveSection(browserData.id, 0);
const reopenedBrowserSections = (await PptxDocument.open(await browserSectioned.write())).sections;
if (reopenedBrowserSections?.length !== 2 || reopenedBrowserSections[0].id !== browserData.id || reopenedBrowserSections[0].title !== 'Browser data' || reopenedBrowserSections[0].slideIds[0] !== browserSectionSlide.slideId || reopenedBrowserSections[1].id !== browserIntro.id || reopenedBrowserSections[1].title !== 'Browser edited' || reopenedBrowserSections[1].slideIds.length !== 0) throw new Error('Browser presentation sections failed');
const browserHiddenDeck = PptxDocument.create();
const browserHiddenSlide = browserHiddenDeck.addSlide();
browserHiddenSlide.hidden = true;
const reopenedBrowserHidden = await PptxDocument.open(await browserHiddenDeck.writeBlob());
if (reopenedBrowserHidden.slides[0]?.hidden !== true) throw new Error('Browser hidden slide failed');
const browserNotesDeck = PptxDocument.create();
const browserNotesSlide = browserNotesDeck.addSlide().addNotes('Browser notes');
if (browserNotesSlide.notes !== 'Browser notes') throw new Error('Browser notes immediate state failed');
const reopenedBrowserNotes = await PptxDocument.open(await browserNotesDeck.writeBlob());
if (reopenedBrowserNotes.slides[0]?.notes !== 'Browser notes') {
  throw new Error('Browser speaker notes failed');
}
PptxDocument.create({ slideSize: { width: inches(11.7), height: inches(8.3) } });
created.slideSize = { width: inches(10), height: inches(7.5) };
process.stdout.write(resolved);
`,
  );
  run(process.execPath, ['--conditions=browser', 'browser-smoke.mjs'], directory);

  await writeFile(
    join(directory, 'smoke.ts'),
    `import {
  degrees,
  PRESET_SHAPE_TYPES,
  PptxDocument,
  ShapeModel,
  TableModel,
  inches,
  type AddShapeOptions,
  type PresetShapeType,
  type SlideModel,
  type CustomSlideSize,
  type AddSectionOptions,
  type AddSlideOptions,
  type PresentationSection,
  type PresentationTheme,
  type PresentationThemeOptions,
  type ThemeFontSnapshot,
  type ThemeFontUpdate,
  type RichTextParagraph,
  type TextAlignment,
  type NumberingStyle,
  type ParagraphBullet,
  type ParagraphLineSpacing,
  type ParagraphSpacing,
  type ParagraphTabStop,
  type ParagraphTabStopAlignment,
  type RichTextUnderline,
  type RichTextUnderlineStyle,
  type RichTextStrikeStyle,
  type RichTextOutline,
  type RichTextGlow,
  type RichTextBaseline,
  type RichTextRunStyle,
  type TextBoxMarginInput,
  type TextBoxMargins,
  type TextBoxFit,
  type TextBoxTextDirection,
  type AddTableCell,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
  type TableCellTextDirection,
  type TableCellBorder,
  type TableCellBorderInput,
  type TableCellBorders,
  type TableCellBorderStyle,
  type TableCellFill,
  type TextBoxVerticalAlignment,
  GradientCodec,
  importPptxGenJS,
  transitions,
  animations,
  advancedCharts,
  smartArt,
} from '@jiayunxie/pptx';

const documentPromise: Promise<PptxDocument> = PptxDocument.open(new Uint8Array());
const createdDocument: PptxDocument = PptxDocument.create({ format: 'pptx', slideSize: 'wide' });
const typedPreset: PresetShapeType = 'foldedCorner';
const typedShapeOptions: AddShapeOptions = {
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(4),
  rotation: degrees(45),
  flipHorizontal: true,
  name: 'Typed shape',
};
const typedShape: ShapeModel = createdDocument.addSlide().addShape(
  typedPreset,
  typedShapeOptions,
);
const typedPresetRead: PresetShapeType | undefined = typedShape.presetType;
typedShape.presetType = 'rect';
const typedPresetCatalog: readonly PresetShapeType[] = PRESET_SHAPE_TYPES;
// @ts-expect-error folderCorner is not a canonical OOXML preset
createdDocument.addSlide().addShape('folderCorner');
// @ts-expect-error custGeom belongs to the custom-geometry API
createdDocument.addSlide().addShape('custGeom');
// @ts-expect-error unknown preset-shape options are rejected
createdDocument.addSlide().addShape('rect', { color: 'FF0000' });
// @ts-expect-error transforms use numeric native units
createdDocument.addSlide().addShape('rect', { width: '3', rotation: '45' });
const addSectionOptions: AddSectionOptions = { title: 'Typed', order: 0 };
const typedSection: PresentationSection = createdDocument.addSection(addSectionOptions);
const addSlideOptions: AddSlideOptions = { sectionTitle: typedSection.title };
createdDocument.addSlide(addSlideOptions);
const sectionSnapshot: readonly PresentationSection[] | undefined = createdDocument.sections;
createdDocument.renameSection(typedSection.id, 'Renamed');
createdDocument.moveSection(typedSection.id, 0);
createdDocument.assignSlideToSection(0, typedSection.id);
createdDocument.deleteSection(typedSection.id);
const typedVisibilitySlide = createdDocument.addSlide();
const hiddenSnapshot: boolean | undefined = typedVisibilitySlide.hidden;
typedVisibilitySlide.hidden = true;
typedVisibilitySlide.hidden = false;
const typedNotesSlide = createdDocument.addSlide();
const notesSnapshot: string | undefined = typedNotesSlide.notes;
typedNotesSlide.notes = 'Typed notes';
typedNotesSlide.notes = '';
typedNotesSlide.notes = undefined;
const returnedNotesSlide: SlideModel = typedNotesSlide.addNotes('Returned');
const globalRtl: PptxDocument = PptxDocument.create({ rtlMode: true });
const globalRtlSnapshot: boolean | undefined = globalRtl.rtlMode;
globalRtl.rtlMode = false;
globalRtl.rtlMode = undefined;
const titledDocument: PptxDocument = PptxDocument.create({ title: 'Typed title' });
const titleSnapshot: string | undefined = titledDocument.title;
titledDocument.title = 'Edited typed title';
titledDocument.title = '';
titledDocument.title = undefined;
const authoredDocument: PptxDocument = PptxDocument.create({ author: 'Typed author' });
const authorSnapshot: string | undefined = authoredDocument.author;
authoredDocument.author = 'Edited typed author';
authoredDocument.author = '';
authoredDocument.author = undefined;
const lastModifiedDocument: PptxDocument = PptxDocument.create({
  lastModifiedBy: 'Typed editor',
});
const lastModifiedSnapshot: string | undefined = lastModifiedDocument.lastModifiedBy;
lastModifiedDocument.lastModifiedBy = 'Edited typed editor';
lastModifiedDocument.lastModifiedBy = '';
lastModifiedDocument.lastModifiedBy = undefined;
const createdAtDocument: PptxDocument = PptxDocument.create({
  createdAt: '2026-07-30T00:00:00Z',
});
const createdAtSnapshot: string | undefined = createdAtDocument.createdAt;
createdAtDocument.createdAt = '2024-02-29T12:34:56.123+05:30';
createdAtDocument.createdAt = undefined;
const modifiedAtDocument: PptxDocument = PptxDocument.create({
  modifiedAt: '2026-07-30T01:02:03Z',
});
const modifiedAtSnapshot: string | undefined = modifiedAtDocument.modifiedAt;
modifiedAtDocument.modifiedAt = '2024-03-01T01:02:03.456+08:00';
modifiedAtDocument.modifiedAt = undefined;
const subjectDocument: PptxDocument = PptxDocument.create({ subject: 'Typed subject' });
const subjectSnapshot: string | undefined = subjectDocument.subject;
subjectDocument.subject = 'Edited typed subject';
subjectDocument.subject = '';
subjectDocument.subject = undefined;
const revisionDocument: PptxDocument = PptxDocument.create({ revision: '007' });
const revisionSnapshot: string | undefined = revisionDocument.revision;
revisionDocument.revision = '42';
revisionDocument.revision = undefined;
const companyDocument: PptxDocument = PptxDocument.create({ company: 'Typed company' });
const companySnapshot: string | undefined = companyDocument.company;
companyDocument.company = 'Edited typed company';
companyDocument.company = '';
companyDocument.company = undefined;
const themeOptions: PresentationThemeOptions = {
  headFontFace: 'Aptos Display',
};
const themedDocument = PptxDocument.create({ theme: themeOptions });
const themeSnapshot: PresentationTheme | undefined = themedDocument.theme;
const fontSnapshot: ThemeFontSnapshot | undefined =
  themedDocument.masterLayoutTheme.presentationTheme?.fonts;
const fontUpdate: ThemeFontUpdate = { minorLatin: 'Aptos' };
themedDocument.masterLayoutTheme.presentationTheme?.setFonts(fontUpdate);
const customSlideSize: CustomSlideSize = { width: inches(11.7), height: inches(8.3) };
const customDocument: PptxDocument = PptxDocument.create({ slideSize: customSlideSize });
customDocument.slideSize = { width: inches(10), height: inches(7.5) };
const alignment: TextAlignment = 'center';
const bullet: ParagraphBullet = { kind: 'bullet', character: '◆', indent: 18 };
const numbering: NumberingStyle = 'romanUcPeriod';
const lineSpacing: ParagraphLineSpacing = { kind: 'multiple', factor: 1.5 };
const spacing: ParagraphSpacing = { before: 4, after: 6, line: lineSpacing };
const tabAlignment: ParagraphTabStopAlignment = 'right';
const tabStops: readonly ParagraphTabStop[] = [{ position: 1.25 }, { position: 2.5, alignment: tabAlignment }];
const underlineStyle: RichTextUnderlineStyle = 'dotDashHeavy';
const underline: RichTextUnderline = { style: underlineStyle, color: { kind: 'srgb', value: 'FF0000' } };
const strike: RichTextStrikeStyle = 'dblStrike';
const outline: RichTextOutline = { color: { kind: 'scheme', value: 'accent1' }, size: 1.5 };
const glow: RichTextGlow = { color: { kind: 'scheme', value: 'accent2' }, opacity: 0.5, size: 8 };
const baseline: RichTextBaseline = 'superscript';
const characterStyle: RichTextRunStyle = { baseline, characterSpacing: 2.5, lang: 'de-DE' };
const margin: TextBoxMarginInput = [4, 8, 4, 8];
const verticalAlignment: TextBoxVerticalAlignment = 'top';
const direction: TextBoxTextDirection = 'vert270';
const cellDirection: TableCellTextDirection = 'vert270';
const fit: TextBoxFit = 'shrink';
const cellFit: TextBoxFit = 'shrink';
const cellAlignment: TextBoxVerticalAlignment = 'middle';
const cellHorizontalAlignment: TextAlignment = 'center';
const tableHorizontalAlignment: TextAlignment = 'center';
const cellMargins: TextBoxMarginInput = { top: 4, left: 8 };
const cellBorderStyle: TableCellBorderStyle = 'dash';
const cellBorder: TableCellBorder = { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: cellBorderStyle };
const cellBorderInput: TableCellBorderInput = [cellBorder, { kind: 'none' }, undefined, cellBorder];
const cellFill: TableCellFill = { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 };
const createdText = createdDocument.addSlide().addText('Typed\\ntext', { align: alignment, fit, valign: verticalAlignment, vert: direction, wrap: true, bullet, level: 2, margin, spacing, tabStops });
const creationBorder: TableCellBorderInput = [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'none' },
  undefined,
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 2 },
];
const creationMargin: TextBoxMarginInput = { top: 4, left: 8 };
const creationOptions: AddTableCellOptions = {
  align: cellHorizontalAlignment,
  border: creationBorder,
  fill: cellFill,
  fit: cellFit,
  margin: creationMargin,
  textDirection: cellDirection,
  valign: cellAlignment,
};
const objectCell: AddTableCell = { text: 'Revenue', options: creationOptions };
const tableRows: readonly (readonly AddTableCellInput[])[] = [['Region', objectCell], [{ text: 'East' }, { text: '' }]];
const tableOptions: AddTableOptions = { align: tableHorizontalAlignment, name: 'Typed table', x: inches(1), columnWidths: [inches(1), inches(3)], rowHeights: [inches(0.5), inches(1.5)], border: cellBorderInput, fill: cellFill, margin: cellMargins, textDirection: cellDirection, valign: cellAlignment };
const typedTable: TableModel = createdDocument.slides[0].addTable(tableRows, tableOptions);
const widthSnapshot: readonly number[] | undefined = typedTable.columnWidths;
const heightSnapshot: readonly number[] | undefined = typedTable.rowHeights;
typedTable.setColumnWidths(inches(2));
typedTable.setColumnWidths([inches(1.5), inches(2.5)]);
typedTable.setRowHeights(inches(1));
typedTable.setRowHeights([0, inches(1.5)]);
const table = createdDocument.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const snapshotDirection: TableCellTextDirection | undefined =
  table?.rows[0]?.cells[0]?.textDirection;
const snapshotFit: TextBoxFit | undefined = table?.rows[0]?.cells[0]?.textFit;
const snapshotAlignment: TextBoxVerticalAlignment | undefined =
  table?.rows[0]?.cells[0]?.verticalAlignment;
const snapshotHorizontalAlignment: TextAlignment | undefined =
  table?.rows[0]?.cells[0]?.horizontalAlignment;
const snapshotCellMargins: TextBoxMargins | undefined = table?.rows[0]?.cells[0]?.margins;
const snapshotCellBorders: TableCellBorders | undefined = table?.rows[0]?.cells[0]?.borders;
const snapshotCellFill: TableCellFill | undefined = table?.rows[0]?.cells[0]?.fill;
table?.setCellTextDirection(0, 0, cellDirection);
table?.setCellTextDirection(0, 0, undefined);
table?.setCellTextFit(0, 0, cellFit);
table?.setCellTextFit(0, 0, 'none');
table?.setCellTextFit(0, 0, undefined);
table?.setCellVerticalAlignment(0, 0, cellAlignment);
table?.setCellVerticalAlignment(0, 0, undefined);
table?.setCellHorizontalAlignment(0, 0, cellHorizontalAlignment);
table?.setCellHorizontalAlignment(0, 0, undefined);
table?.setCellMargins(0, 0, cellMargins);
table?.setCellMargins(0, 0, [3.6, 7.2, 10.8, 14.4]);
table?.setCellMargins(0, 0, undefined);
table?.setCellBorders(0, 0, cellBorder);
table?.setCellBorders(0, 0, cellBorderInput);
table?.setCellBorders(0, 0, { top: cellBorder, left: { kind: 'none' } });
table?.setCellBorders(0, 0, undefined);
table?.setCellFill(0, 0, cellFill);
table?.setCellFill(0, 0, { kind: 'none' });
table?.setCellFill(0, 0, undefined);
const marginSnapshot: TextBoxMargins | undefined = createdText.textMargins;
const wrapSnapshot: boolean | undefined = createdText.textWrap;
const directionSnapshot: TextBoxTextDirection | undefined = createdText.textDirection;
const fitSnapshot: TextBoxFit | undefined = createdText.textFit;
createdText.textMargins = { top: 3, left: 6 };
createdText.verticalAlignment = 'bottom';
createdText.verticalAlignment = undefined;
createdText.textWrap = false;
createdText.textWrap = undefined;
createdText.textDirection = 'wordArtVert';
createdText.textDirection = undefined;
createdText.textFit = 'resize';
createdText.textFit = 'none';
createdText.textFit = undefined;
createdText.text = 'Updated\\n\\ntyped text';
const paragraphs: readonly RichTextParagraph[] = [{ align: 'justify', bullet: { kind: 'number', style: numbering, startAt: 3 }, level: 3, spacing: { line: { kind: 'exact', points: 20 } }, tabStops, runs: [{ text: 'Typed rich', style: { ...characterStyle, fontSize: 12.5, bold: true, color: { kind: 'scheme', value: 'tx1' }, glow, highlight: { kind: 'srgb', value: 'FFFF00' }, outline, underline, strike } }] }];
const richText = createdDocument.slides[0].addRichText(paragraphs, { lang: 'fr-CA' });
richText.richText = paragraphs;
const transparentParagraphs: readonly RichTextParagraph[] = [{
  runs: [
    { text: 'Opaque', style: { transparency: 0 } },
    { text: 'Quarter', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 25 } },
    { text: 'Theme', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 } },
  ],
}];
createdDocument.addSlide().addRichText(transparentParagraphs);
const rtlParagraphs: readonly RichTextParagraph[] = [
  { rtl: true, runs: [{ text: 'RTL' }] },
  { rtl: false, runs: [{ text: 'LTR' }] },
];
createdDocument.addSlide().addRichText(rtlParagraphs, { rtlMode: true });
const paragraphMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginLeft: 12, runs: [{ text: 'Override' }] },
  { marginLeft: false, runs: [{ text: 'Suppressed' }] },
];
createdDocument.addSlide().addRichText(paragraphMargins, { paragraphMarginLeft: 24 });
const paragraphRightMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginRight: 12, runs: [{ text: 'Override' }] },
  { marginRight: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, marginRight: 18, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphRightMargins, { paragraphMarginRight: 24 });
const paragraphIndents: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { indent: 18, runs: [{ text: 'First-line' }] },
  { indent: -18, runs: [{ text: 'Hanging' }] },
  { indent: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, indent: false, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphIndents, { paragraphIndent: 24 });
const gradientConstructor: typeof GradientCodec = GradientCodec;
const adapter: typeof importPptxGenJS = importPptxGenJS;
const transition: transitions.SlideTransition = { effect: 'fade' };
const animationConstructor: typeof animations.AnimationTimingCodec = animations.AnimationTimingCodec;
const chartConstructor: typeof advancedCharts.AdvancedChartCodec = advancedCharts.AdvancedChartCodec;
const smartArtConstructor: typeof smartArt.SmartArtDiagramCodec = smartArt.SmartArtDiagramCodec;
documentPromise.then((document) => {
  transitions.installTransitionPlugin(document);
  animations.installAnimationPlugin(document);
  advancedCharts.installAdvancedChartPlugin(document);
  smartArt.installSmartArtPlugin(document);
});
void [typedNotesSlide, notesSnapshot, returnedNotesSlide];
void [typedPreset, typedShapeOptions, typedShape, typedPresetRead, typedPresetCatalog];
void [documentPromise, createdDocument, addSectionOptions, typedSection, addSlideOptions, sectionSnapshot, typedVisibilitySlide, hiddenSnapshot, globalRtl, globalRtlSnapshot, titledDocument, titleSnapshot, authoredDocument, authorSnapshot, lastModifiedDocument, lastModifiedSnapshot, createdAtDocument, createdAtSnapshot, modifiedAtDocument, modifiedAtSnapshot, subjectDocument, subjectSnapshot, revisionDocument, revisionSnapshot, companyDocument, companySnapshot, themedDocument, themeSnapshot, fontSnapshot, fontUpdate, customDocument, createdText, creationBorder, creationMargin, creationOptions, objectCell, tableRows, tableOptions, typedTable, widthSnapshot, heightSnapshot, table, snapshotDirection, snapshotFit, snapshotAlignment, snapshotHorizontalAlignment, snapshotCellMargins, snapshotCellBorders, snapshotCellFill, cellDirection, cellFit, cellAlignment, cellHorizontalAlignment, tableHorizontalAlignment, cellMargins, cellBorderStyle, cellBorder, cellBorderInput, cellFill, marginSnapshot, wrapSnapshot, directionSnapshot, fitSnapshot, fit, direction, verticalAlignment, richText, transparentParagraphs, rtlParagraphs, paragraphMargins, paragraphRightMargins, paragraphIndents, gradientConstructor, adapter, transition, animationConstructor, chartConstructor, smartArtConstructor];
`,
  );
  run(
    process.execPath,
    [
      join(repositoryRoot, 'node_modules/typescript/bin/tsc'),
      'smoke.ts',
      '--noEmit',
      '--strict',
      '--target',
      'ES2022',
      '--module',
      'NodeNext',
      '--moduleResolution',
      'NodeNext',
      '--typeRoots',
      join(repositoryRoot, 'node_modules/@types'),
      '--types',
      'node',
    ],
    directory,
  );

  const bin = join(directory, 'node_modules', '.bin', process.platform === 'win32' ? 'pptx-inspect.cmd' : 'pptx-inspect');
  const cliResult = run(bin, ['--json', 'doctor'], directory);
  const doctor = JSON.parse(cliResult.stdout);
  if (!doctor.ok || doctor.data?.version !== '0.1.0') throw new Error(`CLI smoke failed: ${cliResult.stdout}`);

  process.stdout.write(
    `${JSON.stringify({ ok: true, tarball: basename(tarball), api: apiChecks, presetShapes: apiChecks.presetShapes, types: true, cli: doctor.data.version })}\n`,
  );
} finally {
  await rm(directory, { recursive: true, force: true });
}

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8' });
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} failed (${result.status})\n${result.stdout}\n${result.stderr}`);
  }
  return result;
}

async function listFiles(directory) {
  const files = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await listFiles(path)));
    else files.push(path);
  }
  return files;
}
