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

function stableState(slide, shape, text) {
  return {
    shape: {
      transform: shape.transform,
      presetType: shape.presetType,
      adjustments: shape.adjustments,
    },
    text: {
      transform: text.transform,
      text: text.text,
      presetType: text.presetType,
      adjustments: text.adjustments,
      isTextBox: text.isTextBox,
      hyperlink: text.hyperlink,
    },
    ids: [shape.id, text.id],
    order: slide.shapes.map(({ id }) => id),
    relationships: slide.relationships,
  };
}

function slideXml(document, slide) {
  return new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
}

function shapeXmlById(xml, id) {
  const identityIndex = xml.indexOf(`<p:cNvPr id="${id}"`);
  if (identityIndex < 0) return '';
  const shapeStart = xml.lastIndexOf('<p:sp', identityIndex);
  const shapeEnd = xml.indexOf('</p:sp>', identityIndex);
  return shapeStart < 0 || shapeEnd < 0
    ? ''
    : xml.slice(shapeStart, shapeEnd + '</p:sp>'.length);
}

function relationshipState(slide) {
  const relationships = slide.relationships.map(
    ({ id, type, target, targetMode, resolvedTarget }) => ({
      id,
      type,
      target,
      targetMode,
      resolvedTarget,
    }),
  );
  return {
    relationships,
    layout: relationships.filter(({ type, targetMode }) =>
      type.endsWith('/slideLayout') && targetMode === 'Internal'),
    hyperlink: relationships.filter(({ type, targetMode }) =>
      type.endsWith('/hyperlink') && targetMode === 'External'),
  };
}

function validRelationshipState(state) {
  return state.relationships.length === 2
    && state.layout.length === 1
    && state.hyperlink.length === 1
    && state.hyperlink[0].target === 'https://example.com/shape-text-identity';
}

export async function runShapeTextTransformIdentity13LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function'
      || typeof api.inches !== 'function' || typeof api.degrees !== 'function') {
    throw new TypeError(
      'Shape/text transform identity probe requires PptxDocument, inches, and degrees',
    );
  }

  const document = api.PptxDocument.create();
  const sourceSlide = document.addSlide();
  const sourceShape = sourceSlide.addShape('roundRect', {
    name: 'Shape original',
    x: api.inches(1),
    y: api.inches(1.5),
    width: api.inches(3),
    height: api.inches(1.5),
    rotation: api.degrees(30),
    flipHorizontal: true,
    flipVertical: true,
    adjustments: [{ name: 'adj', value: 33_333 }],
  });
  const sourceText = sourceSlide.addText('Identity text', {
    name: 'Text original',
    x: api.inches(4.5),
    y: api.inches(1),
    width: api.inches(4),
    height: api.inches(2),
    rotation: api.degrees(-15),
    flipHorizontal: true,
    flipVertical: true,
    shape: 'roundRect',
    rectRadius: api.inches(0.25),
    isTextBox: true,
    hyperlink: { url: 'https://example.com/shape-text-identity', tooltip: 'Identity' },
  });
  const initialStableState = stableState(sourceSlide, sourceShape, sourceText);
  const initialRelationshipState = relationshipState(sourceSlide);
  const beforeNoOp = packageSnapshot(document);

  sourceShape.name = sourceShape.name;
  sourceText.name = sourceText.name;
  const noOp = packageSnapshot(document) === beforeNoOp;

  const invalidMessages = [];
  for (const [owner, value] of [[sourceShape, 42], [sourceText, 'bad\u0000name']]) {
    try {
      owner.name = value;
    } catch (error) {
      invalidMessages.push(error instanceof Error ? error.message : String(error));
    }
  }
  const invalidIsolation = invalidMessages.length === 2
    && invalidMessages.every((message) => message.includes('Shape name'))
    && packageSnapshot(document) === beforeNoOp;

  let rollbackError;
  try {
    document.opcPackage.transaction(() => {
      sourceShape.name = 'Rolled back shape';
      sourceText.name = '';
      throw new Error('restore shape/text names');
    });
  } catch (error) {
    rollbackError = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackError === 'restore shape/text names'
    && packageSnapshot(document) === beforeNoOp;

  sourceShape.name = 'Shape & <edited>';
  sourceText.name = '';
  const sourceIdentity = [sourceShape.name, sourceText.name];
  const sourceStable = JSON.stringify(stableState(sourceSlide, sourceShape, sourceText))
    === JSON.stringify(initialStableState);

  const duplicateSlide = document.duplicateSlide(0);
  const duplicateShape = duplicateSlide.shapes[0];
  const duplicateText = duplicateSlide.shapes[1];
  const initialDuplicateStableState = stableState(
    duplicateSlide,
    duplicateShape,
    duplicateText,
  );
  const initialDuplicateRelationshipState = relationshipState(duplicateSlide);
  duplicateShape.name = 'Duplicate shape';
  duplicateText.name = 'Duplicate text';
  const duplicateIdentity = [duplicateShape.name, duplicateText.name];
  const sourceIsolation = JSON.stringify([sourceShape.name, sourceText.name])
    === JSON.stringify(sourceIdentity)
    && JSON.stringify(stableState(sourceSlide, sourceShape, sourceText))
      === JSON.stringify(initialStableState);
  const duplicateStable = JSON.stringify(
    stableState(duplicateSlide, duplicateShape, duplicateText),
  ) === JSON.stringify(initialDuplicateStableState);

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedSourceSlide = reopened.slides[0];
  const reopenedDuplicateSlide = reopened.slides[1];
  const reopenedShape = reopenedSourceSlide?.shapes[0];
  const reopenedText = reopenedSourceSlide?.shapes[1];
  const reopenedDuplicateShape = reopenedDuplicateSlide?.shapes[0];
  const reopenedDuplicateText = reopenedDuplicateSlide?.shapes[1];
  const reopenedIdentity = [
    reopenedShape?.name,
    reopenedText?.name,
    reopenedDuplicateShape?.name,
    reopenedDuplicateText?.name,
  ];
  const reopenedStable = reopenedSourceSlide && reopenedShape && reopenedText
    ? JSON.stringify(stableState(reopenedSourceSlide, reopenedShape, reopenedText))
      === JSON.stringify(initialStableState)
    : false;
  const reopenedDuplicateStable = reopenedDuplicateSlide
      && reopenedDuplicateShape
      && reopenedDuplicateText
    ? JSON.stringify(stableState(
        reopenedDuplicateSlide,
        reopenedDuplicateShape,
        reopenedDuplicateText,
      )) === JSON.stringify(initialDuplicateStableState)
    : false;

  const editedRelationshipState = relationshipState(sourceSlide);
  const editedDuplicateRelationshipState = relationshipState(duplicateSlide);
  const reopenedRelationshipState = reopenedSourceSlide
    ? relationshipState(reopenedSourceSlide)
    : { relationships: [], layout: [], hyperlink: [] };
  const reopenedDuplicateRelationshipState = reopenedDuplicateSlide
    ? relationshipState(reopenedDuplicateSlide)
    : { relationships: [], layout: [], hyperlink: [] };
  const relationshipStability = [
    initialRelationshipState,
    initialDuplicateRelationshipState,
    editedRelationshipState,
    editedDuplicateRelationshipState,
    reopenedRelationshipState,
    reopenedDuplicateRelationshipState,
  ].every(validRelationshipState)
    && JSON.stringify(editedRelationshipState) === JSON.stringify(initialRelationshipState)
    && JSON.stringify(reopenedRelationshipState) === JSON.stringify(initialRelationshipState)
    && JSON.stringify(editedDuplicateRelationshipState)
      === JSON.stringify(initialDuplicateRelationshipState)
    && JSON.stringify(reopenedDuplicateRelationshipState)
      === JSON.stringify(initialDuplicateRelationshipState);

  const sourceXml = reopenedSourceSlide ? slideXml(reopened, reopenedSourceSlide) : '';
  const duplicateXml = reopenedDuplicateSlide ? slideXml(reopened, reopenedDuplicateSlide) : '';
  const sourceShapeXml = reopenedShape ? shapeXmlById(sourceXml, reopenedShape.id) : '';
  const sourceTextXml = reopenedText ? shapeXmlById(sourceXml, reopenedText.id) : '';
  const duplicateShapeXml = reopenedDuplicateShape
    ? shapeXmlById(duplicateXml, reopenedDuplicateShape.id)
    : '';
  const duplicateTextXml = reopenedDuplicateText
    ? shapeXmlById(duplicateXml, reopenedDuplicateText.id)
    : '';
  const exactOoxml = {
    sourceShape: sourceShapeXml.includes(
      `<p:cNvPr id="${reopenedShape?.id}" name="Shape &amp; &lt;edited&gt;"/>`,
    )
      && sourceShapeXml.includes('rot="1800000" flipH="1" flipV="1"')
      && sourceShapeXml.includes(
        '<a:off x="914400" y="1371600"/><a:ext cx="2743200" cy="1371600"/>',
      )
      && sourceShapeXml.includes('<a:prstGeom prst="roundRect">')
      && sourceShapeXml.includes('<a:gd name="adj" fmla="val 33333"/>'),
    sourceText: sourceTextXml.includes(
      `<p:cNvPr id="${reopenedText?.id}" name="">`,
    )
      && sourceTextXml.includes('<a:hlinkClick r:id="rId2" tooltip="Identity"/>')
      && sourceTextXml.includes('<p:cNvSpPr txBox="1"/>')
      && sourceTextXml.includes('rot="-900000" flipH="1" flipV="1"')
      && sourceTextXml.includes(
        '<a:off x="4114800" y="914400"/><a:ext cx="3657600" cy="1828800"/>',
      )
      && sourceTextXml.includes('<a:prstGeom prst="roundRect">')
      && sourceTextXml.includes('<a:gd name="adj" fmla="val 12500"/>')
      && sourceTextXml.includes('<a:t xml:space="preserve">Identity text</a:t>'),
    duplicateShape: duplicateShapeXml.includes(
      `<p:cNvPr id="${reopenedDuplicateShape?.id}" name="Duplicate shape"/>`,
    )
      && duplicateShapeXml.includes('rot="1800000" flipH="1" flipV="1"')
      && duplicateShapeXml.includes(
        '<a:off x="914400" y="1371600"/><a:ext cx="2743200" cy="1371600"/>',
      )
      && duplicateShapeXml.includes('<a:prstGeom prst="roundRect">')
      && duplicateShapeXml.includes('<a:gd name="adj" fmla="val 33333"/>'),
    duplicateText: duplicateTextXml.includes(
      `<p:cNvPr id="${reopenedDuplicateText?.id}" name="Duplicate text">`,
    )
      && duplicateTextXml.includes('<a:hlinkClick r:id="rId2" tooltip="Identity"/>')
      && duplicateTextXml.includes('<p:cNvSpPr txBox="1"/>')
      && duplicateTextXml.includes('rot="-900000" flipH="1" flipV="1"')
      && duplicateTextXml.includes(
        '<a:off x="4114800" y="914400"/><a:ext cx="3657600" cy="1828800"/>',
      )
      && duplicateTextXml.includes('<a:prstGeom prst="roundRect">')
      && duplicateTextXml.includes('<a:gd name="adj" fmla="val 12500"/>')
      && duplicateTextXml.includes('<a:t xml:space="preserve">Identity text</a:t>'),
  };
  const diagnostics = {
    created: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopened: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
  };
  const state = {
    noOp,
    invalidIsolation,
    rollback,
    sourceIdentity,
    duplicateIdentity,
    sourceStable,
    sourceIsolation,
    duplicateStable,
    reopenedIdentity,
    reopenedStable,
    reopenedDuplicateStable,
    relationshipStability,
    exactOoxml,
    diagnostics,
  };
  const ok = noOp
    && invalidIsolation
    && rollback
    && JSON.stringify(sourceIdentity) === JSON.stringify(['Shape & <edited>', ''])
    && JSON.stringify(duplicateIdentity) === JSON.stringify([
      'Duplicate shape',
      'Duplicate text',
    ])
    && sourceStable
    && sourceIsolation
    && duplicateStable
    && JSON.stringify(reopenedIdentity) === JSON.stringify([
      'Shape & <edited>',
      '',
      'Duplicate shape',
      'Duplicate text',
    ])
    && reopenedStable
    && reopenedDuplicateStable
    && relationshipStability
    && Object.values(exactOoxml).every(Boolean)
    && diagnostics.created === 0
    && diagnostics.reopened === 0;

  return {
    ok,
    state,
    explicitOutputBytes,
    mime: typeof Blob === 'function'
      ? new Blob([explicitOutputBytes], { type: PPTX_MIME }).type
      : PPTX_MIME,
  };
}
