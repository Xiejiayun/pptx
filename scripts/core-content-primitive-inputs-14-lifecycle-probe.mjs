const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function packageSnapshot(document) {
  return JSON.stringify({
    parts: document.opcPackage.parts.map(
      ({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes: [...bytes],
        relationships,
      }),
    ),
    mutations: [...document.opcPackage.mutations],
  });
}

function runState(shape) {
  return shape.richText.map(({ runs }) => runs.map(({ text, style }) => ({
    text,
    color: style?.color,
    bold: style?.bold,
    italic: style?.italic,
  })));
}

function cellState(table) {
  return table.rows[0].cells.map((cell) => ({
    text: cell.text,
    margins: cell.margins,
    runs: runState(cell),
  }));
}

function relationshipState(slide) {
  return slide.relationships.map(({ id, type, target, targetMode, resolvedTarget }) => ({
    id,
    type,
    target,
    targetMode,
    resolvedTarget,
  }));
}

function validRelationshipState(state) {
  return state.length === 1
    && state[0].type.endsWith('/slideLayout')
    && state[0].targetMode === 'Internal';
}

function slideXml(document, slide) {
  return new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
}

function ownerXml(xml, id, tag) {
  const offset = xml.indexOf(`<p:cNvPr id="${id}"`);
  if (offset < 0) return '';
  const start = xml.lastIndexOf(`<${tag}`, offset);
  const end = xml.indexOf(`</${tag}>`, offset);
  return start < 0 || end < offset ? '' : xml.slice(start, end + tag.length + 3);
}

export async function runCoreContentPrimitiveInputs14LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function'
      || typeof api.inches !== 'function') {
    throw new TypeError(
      'Core content/primitive inputs probe requires PptxDocument and inches',
    );
  }

  const document = api.PptxDocument.create({
    slideSize: { width: api.inches(10), height: api.inches(8) },
  });
  const textSlide = document.addSlide();
  const plain = textSlide.addText('Plain string', {
    name: 'Core plain string',
    x: api.inches(1),
    y: '10%',
    width: '30%',
    height: api.inches(1),
    margin: 10,
  });
  const richInput = [{
    runs: [
      {
        text: 'Hex run',
        style: { bold: true, color: { kind: 'srgb', value: '112233' } },
      },
      {
        text: ' Theme run',
        style: { italic: true, color: { kind: 'scheme', value: 'accent2' } },
      },
    ],
  }];
  const rich = textSlide.addRichText(richInput, {
    name: 'Core rich runs',
    x: '10%',
    y: api.inches(2),
    width: api.inches(5),
    height: '20%',
    margin: [1, 2, 3, 4],
  });
  const zero = textSlide.addText('Zero margin', {
    name: 'Core zero margin',
    x: api.inches(1),
    y: api.inches(4),
    width: api.inches(2),
    height: api.inches(0.5),
    margin: 0,
  });

  const tableSlide = document.addSlide();
  const richCellInput = [{
    runs: [
      {
        text: 'Rich cell hex',
        style: { bold: true, color: { kind: 'srgb', value: 'AABBCC' } },
      },
      {
        text: ' rich cell theme',
        style: { italic: true, color: { kind: 'scheme', value: 'accent3' } },
      },
    ],
  }];
  const table = tableSlide.addTable([[
    'Bare string cell',
    {
      text: 'Structured plain cell',
      options: { bold: true, color: { kind: 'srgb', value: '445566' }, margin: 0 },
    },
    {
      text: richCellInput,
      options: { margin: [1, 2, 3, 4] },
    },
  ]], {
    name: 'Core primitive table',
    x: api.inches(1),
    y: api.inches(1),
    width: api.inches(8),
    height: api.inches(1.5),
    margin: 0,
  });

  richInput[0].runs[0].text = 'MUTATED';
  richCellInput[0].runs[0].text = 'MUTATED';
  const created = {
    texts: [plain.text, rich.text, zero.text],
    transforms: [plain.transform, rich.transform],
    margins: [plain.textMargins, rich.textMargins, zero.textMargins],
    richRuns: runState(rich),
    cells: cellState(table),
  };
  const sourceIsolation = rich.text === 'Hex run Theme run'
    && table.rows[0].cells[2].text === 'Rich cell hex rich cell theme';
  const initialRelationships = [
    relationshipState(textSlide),
    relationshipState(tableSlide),
  ];
  const beforeNoOp = packageSnapshot(document);

  plain.text = plain.text;
  rich.richText = rich.richText;
  rich.textMargins = rich.textMargins;
  table.setCellRichText(0, 2, table.rows[0].cells[2].richText);
  const noOp = packageSnapshot(document) === beforeNoOp;

  const invalidMessages = [];
  for (const invoke of [
    () => { rich.textMargins = [1, 2, 3]; },
    () => { rich.richText = [{ runs: [{
      text: 'Invalid',
      style: { color: { kind: 'scheme', value: 'unknown' } },
    }] }]; },
    () => table.setCellRichText(9, 0, [{ runs: [{ text: 'Missing' }] }]),
  ]) {
    try {
      invoke();
    } catch (error) {
      invalidMessages.push(error instanceof Error ? error.message : String(error));
    }
  }
  const invalidIsolation = invalidMessages.length === 3
    && packageSnapshot(document) === beforeNoOp;

  let rollbackError;
  try {
    document.transaction(() => {
      plain.text = 'Rolled back plain';
      rich.richText = [{ runs: [{ text: 'Rolled back rich' }] }];
      table.setCellRichText(0, 2, [{ runs: [{ text: 'Rolled back cell' }] }]);
      throw new Error('restore core content primitives');
    });
  } catch (error) {
    rollbackError = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackError === 'restore core content primitives'
    && packageSnapshot(document) === beforeNoOp;

  plain.text = 'Plain edited';
  rich.richText = [{
    runs: [
      {
        text: 'Edited hex',
        style: { bold: true, color: { kind: 'srgb', value: 'ABCDEF' } },
      },
      {
        text: ' Edited theme',
        style: { italic: true, color: { kind: 'scheme', value: 'accent4' } },
      },
    ],
  }];
  table.setCellText(0, 0, 'Bare edited');
  table.setCellRichText(0, 2, [{
    runs: [
      {
        text: 'Edited table',
        style: { color: { kind: 'srgb', value: '123456' } },
      },
      {
        text: ' theme',
        style: { color: { kind: 'scheme', value: 'accent5' } },
      },
    ],
  }]);
  const edited = {
    texts: [plain.text, rich.text, zero.text],
    transforms: [plain.transform, rich.transform],
    margins: [plain.textMargins, rich.textMargins, zero.textMargins],
    richRuns: runState(rich),
    cells: cellState(table),
  };

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedPlain = reopened.slides[0]?.shapes[0];
  const reopenedRich = reopened.slides[0]?.shapes[1];
  const reopenedZero = reopened.slides[0]?.shapes[2];
  const reopenedTable = reopened.slides[1]?.shapes[0];
  const reopenedState = reopenedPlain && reopenedRich && reopenedZero && reopenedTable
    ? {
        texts: [reopenedPlain.text, reopenedRich.text, reopenedZero.text],
        transforms: [reopenedPlain.transform, reopenedRich.transform],
        margins: [
          reopenedPlain.textMargins,
          reopenedRich.textMargins,
          reopenedZero.textMargins,
        ],
        richRuns: runState(reopenedRich),
        cells: cellState(reopenedTable),
      }
    : undefined;
  const reopenedRelationships = reopened.slides.map(relationshipState);
  const relationshipStability = [
    ...initialRelationships,
    relationshipState(textSlide),
    relationshipState(tableSlide),
    ...reopenedRelationships,
  ].every(validRelationshipState)
    && JSON.stringify(reopenedRelationships) === JSON.stringify(initialRelationships);

  const textXml = reopened.slides[0] ? slideXml(reopened, reopened.slides[0]) : '';
  const tableXml = reopened.slides[1] ? slideXml(reopened, reopened.slides[1]) : '';
  const plainXml = reopenedPlain ? ownerXml(textXml, reopenedPlain.id, 'p:sp') : '';
  const richXml = reopenedRich ? ownerXml(textXml, reopenedRich.id, 'p:sp') : '';
  const zeroXml = reopenedZero ? ownerXml(textXml, reopenedZero.id, 'p:sp') : '';
  const tableOwnerXml = reopenedTable
    ? ownerXml(tableXml, reopenedTable.id, 'p:graphicFrame')
    : '';
  const exactOoxml = {
    plain: plainXml.includes(
      '<a:off x="914400" y="731520"/><a:ext cx="2743200" cy="914400"/>',
    ) && plainXml.includes(
      'lIns="127000" tIns="127000" rIns="127000" bIns="127000"',
    ) && plainXml.includes('Plain edited'),
    rich: richXml.includes(
      '<a:off x="914400" y="1828800"/><a:ext cx="4572000" cy="1463040"/>',
    ) && richXml.includes(
      'lIns="50800" tIns="12700" rIns="25400" bIns="38100"',
    ) && richXml.includes('<a:srgbClr val="ABCDEF"/>')
      && richXml.includes('<a:schemeClr val="accent4"/>'),
    zero: zeroXml.includes('lIns="0" tIns="0" rIns="0" bIns="0"'),
    tableCreationAndEdit: tableOwnerXml.includes('Structured plain cell')
      && tableOwnerXml.includes('Bare edited')
      && tableOwnerXml.includes('<a:srgbClr val="445566"/>')
      && tableOwnerXml.includes('<a:srgbClr val="123456"/>')
      && tableOwnerXml.includes('<a:schemeClr val="accent5"/>')
      && tableOwnerXml.includes(
        'marL="50800" marR="25400" marT="12700" marB="38100"',
      ),
  };
  const diagnostics = {
    createdErrors: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    createdWarnings: document.diagnostics
      .filter(({ severity }) => severity === 'warning').length,
    reopenedErrors: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopenedWarnings: reopened.diagnostics
      .filter(({ severity }) => severity === 'warning').length,
  };
  const state = {
    created,
    sourceIsolation,
    noOp,
    invalidIsolation,
    invalidMessages,
    rollback,
    edited,
    reopened: reopenedState,
    relationshipStability,
    relationships: reopenedRelationships,
    exactOoxml,
    diagnostics,
  };
  const ok = created.texts.join('|') === 'Plain string|Hex run Theme run|Zero margin'
    && created.transforms[0].y === 731_520
    && created.transforms[0].width === 2_743_200
    && created.transforms[1].x === 914_400
    && created.transforms[1].height === 1_463_040
    && created.margins[0].top === 10
    && created.margins[0].right === 10
    && created.margins[0].bottom === 10
    && created.margins[0].left === 10
    && created.margins[1].top === 1
    && created.margins[1].right === 2
    && created.margins[1].bottom === 3
    && created.margins[1].left === 4
    && created.margins[2].top === 0
    && created.margins[2].right === 0
    && created.margins[2].bottom === 0
    && created.margins[2].left === 0
    && created.cells.map(({ text }) => text).join('|')
      === 'Bare string cell|Structured plain cell|Rich cell hex rich cell theme'
    && sourceIsolation
    && noOp
    && invalidIsolation
    && rollback
    && edited.texts.join('|') === 'Plain edited|Edited hex Edited theme|Zero margin'
    && edited.cells.map(({ text }) => text).join('|')
      === 'Bare edited|Structured plain cell|Edited table theme'
    && JSON.stringify(reopenedState) === JSON.stringify(edited)
    && relationshipStability
    && Object.values(exactOoxml).every(Boolean)
    && Object.values(diagnostics).every((count) => count === 0);

  return {
    ok,
    state,
    explicitOutputBytes,
    mime: typeof Blob === 'function'
      ? new Blob([explicitOutputBytes], { type: PPTX_MIME }).type
      : PPTX_MIME,
  };
}
