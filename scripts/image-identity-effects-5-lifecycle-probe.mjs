const PNG_BYTES = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2,
  0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1,
  5, 1, 1, 39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174,
  66, 96, 130,
]);

const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function bytesEqual(left, right) {
  return left.length === right.length
    && left.every((value, index) => value === right[index]);
}

function packageSnapshot(document) {
  return JSON.stringify(document.opcPackage.parts.map(
    ({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: [...bytes],
      relationships,
    }),
  ));
}

function appearance(image) {
  return {
    name: image.name,
    altText: image.altText,
    rounding: image.rounding,
    transparency: image.transparency,
    shadow: image.shadow,
  };
}

function imageRelationships(slide, sourcePartUri) {
  return slide.relationships.filter(({ type, targetMode, resolvedTarget }) =>
    type.endsWith('/image')
    && targetMode === 'Internal'
    && resolvedTarget === sourcePartUri);
}

function slideXml(document, slide) {
  return new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
}

export async function runImageIdentityEffects5LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function') {
    throw new TypeError('Image identity/effects probe requires PptxDocument');
  }

  const document = api.PptxDocument.create();
  const sourceSlide = document.addSlide();
  const source = sourceSlide.addImage(PNG_BYTES.slice(), {
    contentType: 'image/png',
    name: 'Created & image',
    altText: 'Created <alt>',
    rounding: true,
    transparency: 25,
    shadow: {
      kind: 'outer',
      color: { kind: 'srgb', value: '123456' },
      opacity: 0.5,
      blur: 3,
      angle: 30,
      distance: 2,
      rotateWithShape: true,
    },
  });
  const sourcePartUri = source.sourcePartUri;
  const mediaBytes = document.opcPackage.requirePart(sourcePartUri).bytes.slice();
  const created = appearance(source);
  const beforeNoOp = packageSnapshot(document);

  source.name = created.name;
  source.altText = created.altText;
  source.rounding = created.rounding;
  source.transparency = created.transparency;
  source.shadow = created.shadow;
  const noOp = packageSnapshot(document) === beforeNoOp;

  let rollbackError;
  try {
    document.opcPackage.transaction(() => {
      source.name = 'Rolled back';
      source.altText = undefined;
      source.rounding = false;
      source.transparency = 100;
      source.shadow = undefined;
      throw new Error('restore image identity/effects');
    });
  } catch (error) {
    rollbackError = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackError === 'restore image identity/effects'
    && packageSnapshot(document) === beforeNoOp;

  const duplicateSlide = document.duplicateSlide(0);
  const duplicate = duplicateSlide.shapes[0];
  const sharedMedia = duplicate.sourcePartUri === sourcePartUri;
  duplicate.name = '';
  duplicate.altText = undefined;
  duplicate.rounding = false;
  duplicate.transparency = 100;
  duplicate.shadow = {
    kind: 'inner',
    color: { kind: 'scheme', value: 'accent3' },
    opacity: 0.4,
    blur: 1,
    angle: 90,
    distance: 1,
  };
  const duplicateAtHundred = appearance(duplicate);
  const duplicateAtHundredXml = slideXml(document, duplicateSlide);
  duplicate.transparency = 0;
  const editedDuplicate = appearance(duplicate);

  const sourceIsolation = JSON.stringify(appearance(source)) === JSON.stringify(created);
  const relationshipState = {
    source: imageRelationships(sourceSlide, sourcePartUri).length,
    duplicate: imageRelationships(duplicateSlide, sourcePartUri).length,
  };
  const mediaPartCount = document.opcPackage.parts.filter(
    ({ uri }) => uri.startsWith('/ppt/media/'),
  ).length;
  const mediaPreserved = bytesEqual(
    document.opcPackage.requirePart(sourcePartUri).bytes,
    mediaBytes,
  );

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedSource = reopened.slides[0]?.shapes[0];
  const reopenedDuplicate = reopened.slides[1]?.shapes[0];
  const reopenedSourceState = reopenedSource ? appearance(reopenedSource) : undefined;
  const reopenedDuplicateState = reopenedDuplicate ? appearance(reopenedDuplicate) : undefined;
  const reopenedRelationshipState = reopenedSource && reopenedDuplicate
    ? {
        source: imageRelationships(reopened.slides[0], sourcePartUri).length,
        duplicate: imageRelationships(reopened.slides[1], sourcePartUri).length,
      }
    : { source: 0, duplicate: 0 };
  const reopenedMediaPreserved = reopenedSource?.sourcePartUri === sourcePartUri
    && reopenedDuplicate?.sourcePartUri === sourcePartUri
    && bytesEqual(reopened.opcPackage.requirePart(sourcePartUri).bytes, mediaBytes);

  const sourceXml = slideXml(reopened, reopened.slides[0]);
  const duplicateXml = slideXml(reopened, reopened.slides[1]);
  const exactOoxml = {
    sourceIdentity: sourceXml.includes('name="Created &amp; image"')
      && sourceXml.includes('descr="Created &lt;alt&gt;"'),
    sourceRounding: sourceXml.includes('<a:prstGeom prst="ellipse">'),
    sourceTransparency: sourceXml.includes('<a:alphaModFix amt="75000"/>'),
    sourceShadow: sourceXml.includes('<a:outerShdw')
      && sourceXml.includes('rotWithShape="1"'),
    duplicateIdentity: duplicateXml.includes('name=""')
      && !duplicateXml.includes(' descr='),
    duplicateRounding: duplicateXml.includes('<a:prstGeom prst="rect">'),
    duplicateTransparency: !duplicateXml.includes('alphaModFix'),
    duplicateInnerShadow: duplicateXml.includes('<a:innerShdw'),
    hundredPercentIntermediate: duplicateAtHundredXml.includes(
      '<a:alphaModFix amt="0"/>',
    ),
  };
  const diagnostics = {
    created: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopened: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
  };
  const state = {
    created,
    noOp,
    rollback,
    duplicateAtHundred,
    editedDuplicate,
    sourceIsolation,
    sharedMedia,
    relationshipState,
    reopenedRelationshipState,
    mediaPartCount,
    mediaPreserved,
    reopenedMediaPreserved,
    reopenedSource: reopenedSourceState,
    reopenedDuplicate: reopenedDuplicateState,
    exactOoxml,
    diagnostics,
  };
  const ok = created.name === 'Created & image'
    && created.altText === 'Created <alt>'
    && created.rounding === true
    && created.transparency === 25
    && created.shadow?.kind === 'outer'
    && created.shadow.rotateWithShape === true
    && noOp
    && rollback
    && duplicateAtHundred.name === ''
    && duplicateAtHundred.altText === undefined
    && duplicateAtHundred.rounding === false
    && duplicateAtHundred.transparency === 100
    && duplicateAtHundred.shadow?.kind === 'inner'
    && editedDuplicate.transparency === 0
    && editedDuplicate.shadow?.kind === 'inner'
    && sourceIsolation
    && sharedMedia
    && relationshipState.source === 1
    && relationshipState.duplicate === 1
    && reopenedRelationshipState.source === 1
    && reopenedRelationshipState.duplicate === 1
    && mediaPartCount === 1
    && mediaPreserved
    && reopenedMediaPreserved
    && JSON.stringify(reopenedSourceState) === JSON.stringify(created)
    && JSON.stringify(reopenedDuplicateState) === JSON.stringify(editedDuplicate)
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
