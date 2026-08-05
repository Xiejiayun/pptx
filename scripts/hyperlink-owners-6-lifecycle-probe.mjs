const PNG_BYTES = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);
const SVG_BYTES = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"/>',
);
const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

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

function hyperlinkRelationships(slide) {
  return slide.relationships.filter(({ type }) =>
    type.endsWith('/hyperlink') || type.endsWith('/slide'));
}

function slideXml(document, slide) {
  return new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
}

function loaderPngBytes() {
  const bytes = new Uint8Array(24);
  bytes.set(PNG_BYTES);
  bytes.set([0, 0, 0, 13, 73, 72, 68, 82], 8);
  bytes[19] = 1;
  bytes[23] = 1;
  return bytes;
}

function deferredImageSource() {
  let release;
  let reads = 0;
  const gate = new Promise((resolve) => {
    release = resolve;
  });
  return {
    source: {
      async *[Symbol.asyncIterator]() {
        reads += 1;
        await gate;
        yield loaderPngBytes();
      },
    },
    release,
    reads: () => reads,
  };
}

async function probeAsyncTargetIdentity(api) {
  const reordered = api.PptxDocument.create();
  const owner = reordered.addSlide();
  const target = reordered.addSlide();
  reordered.addSlide();
  const reorderGate = deferredImageSource();
  const pendingReorder = reordered.addImage(0, reorderGate.source, {
    hyperlink: { slide: 2 },
  });
  reordered.moveSlide(1, 2);
  reorderGate.release();
  const reorderedImage = await pendingReorder;
  const movePreservesTarget = reorderedImage.hyperlink?.slide === 3
    && owner.relationships.some(({ resolvedTarget }) => resolvedTarget === target.partUri);

  const deletedTarget = api.PptxDocument.create();
  deletedTarget.addSlide();
  deletedTarget.addSlide();
  deletedTarget.addSlide();
  const targetGate = deferredImageSource();
  const pendingTarget = deletedTarget.addImage(0, targetGate.source, {
    hyperlink: { slide: 2 },
  });
  deletedTarget.deleteSlide(1);
  deletedTarget.addSlide();
  const beforeTargetReject = packageSnapshot(deletedTarget);
  targetGate.release();
  let targetRejected = false;
  try {
    await pendingTarget;
  } catch (error) {
    targetRejected = error instanceof Error && /target was deleted/.test(error.message);
  }
  const targetDeleteIsolation = targetRejected
    && packageSnapshot(deletedTarget) === beforeTargetReject;

  const deletedOwner = api.PptxDocument.create();
  deletedOwner.addSlide();
  deletedOwner.addSlide();
  const ownerGate = deferredImageSource();
  const pendingOwner = deletedOwner.addImage(0, ownerGate.source, {
    hyperlink: { slide: 2 },
  });
  deletedOwner.deleteSlide(0);
  deletedOwner.addSlide();
  const beforeOwnerReject = packageSnapshot(deletedOwner);
  ownerGate.release();
  let ownerRejected = false;
  try {
    await pendingOwner;
  } catch (error) {
    ownerRejected = error instanceof Error && /owner slide was deleted/.test(error.message);
  }
  const ownerDeleteIsolation = ownerRejected
    && packageSnapshot(deletedOwner) === beforeOwnerReject;

  const invalid = api.PptxDocument.create();
  invalid.addSlide();
  const invalidGate = deferredImageSource();
  const pendingInvalid = invalid.addImage(0, invalidGate.source, {
    hyperlink: { slide: 99 },
  });
  await Promise.resolve();
  const invalidReadsBeforeRelease = invalidGate.reads();
  invalidGate.release();
  let invalidRejected = false;
  try {
    await pendingInvalid;
  } catch (error) {
    invalidRejected = error instanceof Error && /out of range/.test(error.message);
  }

  return {
    movePreservesTarget,
    targetDeleteIsolation,
    ownerDeleteIsolation,
    invalidPreIo: invalidRejected && invalidReadsBeforeRelease === 0,
  };
}

export async function runHyperlinkOwners6LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function') {
    throw new TypeError('Hyperlink owners probe requires PptxDocument');
  }

  const document = api.PptxDocument.create();
  const sourceSlide = document.addSlide();
  document.addSlide();
  const targetSlide = document.addSlide();
  const rasterInput = {
    url: 'https://images.example?a=1&b=2',
    tooltip: 'Raster & tip',
  };
  const raster = sourceSlide.addImage(PNG_BYTES.slice(), {
    contentType: 'image/png',
    name: 'Hyperlink raster',
    hyperlink: rasterInput,
  });
  const vector = sourceSlide.addSvgImage(SVG_BYTES.slice(), PNG_BYTES.slice(), {
    name: 'Hyperlink SVG',
    hyperlink: { slide: 3, tooltip: '' },
  });
  rasterInput.url = 'https://mutated.example';
  rasterInput.tooltip = 'Mutated';

  const created = {
    raster: raster.hyperlink,
    vector: vector.hyperlink,
    callerDetached: rasterInput.url === 'https://mutated.example'
      && raster.hyperlink?.url === 'https://images.example?a=1&b=2',
  };
  const mediaState = {
    raster: raster.sourcePartUri,
    fallback: vector.fallbackPartUri,
    svg: vector.svgPartUri,
    parts: document.opcPackage.parts.filter(({ uri }) => uri.startsWith('/ppt/media/'))
      .map(({ uri, bytes }) => [uri, [...bytes]]),
  };
  const beforeNoOp = packageSnapshot(document);
  raster.hyperlink = { url: 'https://images.example?a=1&b=2', tooltip: 'Raster & tip' };
  vector.hyperlink = { slide: 3, tooltip: '' };
  const noOp = packageSnapshot(document) === beforeNoOp;

  const initialRelationships = hyperlinkRelationships(sourceSlide);
  const rasterRelationship = initialRelationships.find(({ target }) =>
    target === 'https://images.example?a=1&b=2');
  raster.hyperlink = { url: 'https://images.example?a=1&b=2', tooltip: '' };
  const tooltipRelationshipReuse = hyperlinkRelationships(sourceSlide)
    .some(({ id }) => id === rasterRelationship?.id);
  raster.hyperlink = { slide: 3 };
  const targetSwitchReuse = hyperlinkRelationships(sourceSlide)
    .some(({ id, resolvedTarget }) =>
      id === rasterRelationship?.id && resolvedTarget === targetSlide.partUri);

  let rollbackError;
  const beforeRollback = packageSnapshot(document);
  try {
    document.opcPackage.transaction(() => {
      raster.hyperlink = { url: 'https://rollback.example' };
      vector.hyperlink = undefined;
      throw new Error('restore hyperlink owners');
    });
  } catch (error) {
    rollbackError = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackError === 'restore hyperlink owners'
    && packageSnapshot(document) === beforeRollback;

  const duplicateSlide = document.duplicateSlide(0);
  const duplicateRaster = duplicateSlide.shapes[0];
  duplicateRaster.hyperlink = { url: 'https://duplicate.example' };
  const duplicateIsolation = raster.hyperlink?.slide === 3
    && duplicateRaster.hyperlink?.url === 'https://duplicate.example';
  const asyncTargetIdentity = await probeAsyncTargetIdentity(api);

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedRaster = reopened.slides[0]?.shapes[0];
  const reopenedVector = reopened.slides[0]?.shapes[1];
  const reopenedState = {
    raster: reopenedRaster?.hyperlink,
    vector: reopenedVector?.hyperlink,
    duplicate: reopened.slides[3]?.shapes[0]?.hyperlink,
  };
  const reopenedMediaState = {
    raster: reopenedRaster?.sourcePartUri,
    fallback: reopenedVector?.fallbackPartUri,
    svg: reopenedVector?.svgPartUri,
    parts: reopened.opcPackage.parts.filter(({ uri }) => uri.startsWith('/ppt/media/'))
      .map(({ uri, bytes }) => [uri, [...bytes]]),
  };
  const sourceXml = slideXml(reopened, reopened.slides[0]);
  const exactOoxml = {
    pictureOwners: (sourceXml.match(/<p:pic>/g) ?? []).length === 2,
    rasterClick: sourceXml.includes(
      `r:id="${rasterRelationship?.id}" action="ppaction://hlinksldjump"`,
    ),
    explicitEmptyTooltip: sourceXml.includes(
      'tooltip="" action="ppaction://hlinksldjump"',
    ),
    escapedRelationship: reopened.slides[3]?.relationships.some(
      ({ target }) => target === 'https://duplicate.example',
    ) === true,
    noUndefinedRelationship: !sourceXml.includes('rIdundefined'),
  };
  const diagnostics = {
    createdErrors: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    createdWarnings: document.diagnostics.filter(({ severity }) => severity === 'warning').length,
    reopenedErrors: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopenedWarnings: reopened.diagnostics.filter(({ severity }) => severity === 'warning').length,
  };
  const state = {
    created,
    noOp,
    tooltipRelationshipReuse,
    targetSwitchReuse,
    rollback,
    duplicateIsolation,
    asyncTargetIdentity,
    reopenedState,
    mediaPreserved: JSON.stringify(reopenedMediaState) === JSON.stringify(mediaState),
    exactOoxml,
    diagnostics,
  };
  const ok = created.raster?.url === 'https://images.example?a=1&b=2'
    && created.raster?.tooltip === 'Raster & tip'
    && created.vector?.slide === 3
    && created.vector?.tooltip === ''
    && created.callerDetached
    && noOp
    && tooltipRelationshipReuse
    && targetSwitchReuse
    && rollback
    && duplicateIsolation
    && Object.values(asyncTargetIdentity).every(Boolean)
    && reopenedState.raster?.slide === 3
    && reopenedState.vector?.slide === 3
    && reopenedState.vector?.tooltip === ''
    && reopenedState.duplicate?.url === 'https://duplicate.example'
    && state.mediaPreserved
    && Object.values(exactOoxml).every(Boolean)
    && diagnostics.createdErrors === 0
    && diagnostics.createdWarnings === 1
    && diagnostics.reopenedErrors === 0
    && diagnostics.reopenedWarnings === 0;

  return {
    ok,
    state,
    explicitOutputBytes,
    mime: typeof Blob === 'function'
      ? new Blob([explicitOutputBytes], { type: PPTX_MIME }).type
      : PPTX_MIME,
  };
}
