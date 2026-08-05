const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const FORMATS = ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'];

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

function ownerXml(xml, id) {
  const offset = xml.indexOf(`<p:cNvPr id="${id}"`);
  if (offset < 0) return '';
  const start = xml.lastIndexOf('<p:sp', offset);
  const end = xml.indexOf('</p:sp>', offset);
  return start < 0 || end < offset ? '' : xml.slice(start, end + 7);
}

function partXml(document, uri) {
  return new TextDecoder().decode(document.opcPackage.requirePart(uri).bytes);
}

function styleState(shape) {
  return {
    align: shape.richText.map(({ align }) => align),
    transparency: shape.richText.flatMap(({ runs }) =>
      runs.map(({ style }) => style?.transparency)),
    margins: shape.textMargins,
    valign: shape.verticalAlignment,
    placeholder: shape.placeholder,
  };
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

function layoutOnlyRelationships(state) {
  return state.length === 1
    && state[0].type.endsWith('/slideLayout')
    && state[0].targetMode === 'Internal';
}

export async function runPlaceholderTextStyle4LifecycleProbe(api) {
  if (!api || typeof api !== 'object' || typeof api.PptxDocument !== 'function'
      || typeof api.inches !== 'function') {
    throw new TypeError(
      'Placeholder text style probe requires PptxDocument and inches',
    );
  }

  const document = api.PptxDocument.create();
  const layout = document.layouts[0];
  const master = document.masters[0];

  const layoutPrompt = layout.addPlaceholder([{
    runs: [{
      text: 'Layout prompt',
      style: { color: { kind: 'srgb', value: '112233' }, transparency: 12.5 },
    }],
  }], {
    name: 'placeholder_style_layout',
    type: 'title',
    index: 240,
    x: api.inches(1),
    y: api.inches(1),
    width: api.inches(8),
    height: api.inches(1),
    align: 'center',
    margin: [1, 2, 3, 4],
    valign: 'bottom',
  });
  const masterPrompt = master.addPlaceholder([{
    runs: [{
      text: 'Master prompt',
      style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 },
    }],
  }], {
    name: 'placeholder_style_master',
    type: 'body',
    index: 241,
    align: 'right',
    margin: 0,
    valign: 'top',
  });

  const slide = document.addSlide({ masterName: layout.name });
  const materialized = slide.placeholders.find(
    ({ name }) => name === 'placeholder_style_layout',
  );
  const materializedIdentity = materialized && {
    id: materialized.id,
    name: materialized.name,
    transform: materialized.transform,
    placeholder: materialized.placeholder,
  };
  const populated = slide.addRichText([{
    runs: [{
      text: 'Populated owner',
      style: { color: { kind: 'srgb', value: '445566' }, transparency: 25 },
    }],
  }], {
    placeholder: 'placeholder_style_layout',
    align: 'justify',
    margin: 8,
    valign: 'top',
  });
  const directPrompt = slide.addPlaceholder([{
    runs: [
      { text: 'Omitted transparency' },
      {
        text: ' Zero transparency',
        style: { color: { kind: 'srgb', value: 'ABCDEF' }, transparency: 0 },
      },
      {
        text: ' Half transparency',
        style: { color: { kind: 'scheme', value: 'accent2' }, transparency: 50 },
      },
    ],
  }], {
    name: 'placeholder_style_slide',
    type: 'body',
    index: 242,
    align: 'left',
    margin: { top: 2, right: 3, bottom: 4, left: 5 },
    valign: 'middle',
  });

  const created = {
    layout: styleState(layoutPrompt),
    master: styleState(masterPrompt),
    populated: styleState(populated),
    slide: styleState(directPrompt),
    populatedIdentity: {
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    },
    materializedIdentity,
  };
  const ownerReuse = JSON.stringify(created.populatedIdentity)
    === JSON.stringify(materializedIdentity);
  const beforeNoOp = packageSnapshot(document);
  populated.richText = populated.richText;
  populated.textMargins = populated.textMargins;
  populated.verticalAlignment = populated.verticalAlignment;
  const noOp = packageSnapshot(document) === beforeNoOp;

  const invalidMessages = [];
  for (const invoke of [
    () => { populated.richText = [{ runs: [{ text: 'Invalid' }], align: 'centered' }]; },
    () => { populated.richText = [{ runs: [{
      text: 'Invalid', style: { transparency: 101 },
    }] }]; },
    () => { populated.textMargins = [1, 2, 3]; },
    () => { populated.verticalAlignment = 'center'; },
  ]) {
    try {
      invoke();
    } catch (error) {
      invalidMessages.push(error instanceof Error ? error.message : String(error));
    }
  }
  const invalidIsolation = invalidMessages.length === 4
    && packageSnapshot(document) === beforeNoOp;

  let rollbackMessage;
  try {
    document.transaction(() => {
      populated.richText = [{ runs: [{
        text: 'Rollback', style: { transparency: 75 },
      }], align: 'right' }];
      populated.textMargins = [9, 8, 7, 6];
      populated.verticalAlignment = 'bottom';
      throw new Error('restore placeholder text style');
    });
  } catch (error) {
    rollbackMessage = error instanceof Error ? error.message : String(error);
  }
  const rollback = rollbackMessage === 'restore placeholder text style'
    && packageSnapshot(document) === beforeNoOp;

  populated.richText = [{
    runs: [{
      text: 'Edited populated owner',
      style: { color: { kind: 'scheme', value: 'accent3' }, transparency: 50 },
    }],
    align: 'left',
  }];
  populated.textMargins = { top: 4, left: 8 };
  populated.verticalAlignment = 'bottom';
  const edited = styleState(populated);

  const duplicate = document.duplicateSlide(0);
  const duplicatePopulated = duplicate.shapes.find(
    ({ name }) => name === populated.name,
  );
  duplicatePopulated.richText = [{
    runs: [{ text: 'Duplicate', style: { transparency: 75 } }],
    align: 'right',
  }];
  duplicatePopulated.textMargins = [5, 6, 7, 8];
  duplicatePopulated.verticalAlignment = 'top';
  const duplicateState = styleState(duplicatePopulated);
  const duplicateIsolation = JSON.stringify(styleState(populated)) === JSON.stringify(edited);
  const initialRelationships = [relationshipState(slide), relationshipState(duplicate)];

  const explicitOutputBytes = await document.write({ compatibility: 'powerpoint-2010' });
  const reopened = await api.PptxDocument.open(explicitOutputBytes);
  const reopenedLayout = reopened.layouts.find(({ name }) => name === layout.name);
  const reopenedMaster = reopened.masters.find(({ name }) => name === master.name);
  const reopenedPopulated = reopened.slides[0]?.shapes.find(
    ({ name }) => name === populated.name,
  );
  const reopenedDirect = reopened.slides[0]?.shapes.find(
    ({ name }) => name === directPrompt.name,
  );
  const reopenedDuplicate = reopened.slides[1]?.shapes.find(
    ({ name }) => name === duplicatePopulated.name,
  );
  const reopenedState = reopenedLayout && reopenedMaster && reopenedPopulated
      && reopenedDirect && reopenedDuplicate
    ? {
        layout: styleState(reopenedLayout.placeholders.find(
          ({ name }) => name === layoutPrompt.name,
        )),
        master: styleState(reopenedMaster.placeholders.find(
          ({ name }) => name === masterPrompt.name,
        )),
        populated: styleState(reopenedPopulated),
        slide: styleState(reopenedDirect),
        duplicate: styleState(reopenedDuplicate),
      }
    : undefined;

  const reopenedRelationships = reopened.slides.map(relationshipState);
  const relationshipStability = [
    ...initialRelationships,
    ...reopenedRelationships,
  ].every(layoutOnlyRelationships)
    && JSON.stringify(reopenedRelationships) === JSON.stringify(initialRelationships);

  const layoutXml = reopenedLayout ? partXml(reopened, reopenedLayout.partUri) : '';
  const masterXml = reopenedMaster ? partXml(reopened, reopenedMaster.partUri) : '';
  const slideXml = reopened.slides[0] ? partXml(reopened, reopened.slides[0].partUri) : '';
  const duplicateXml = reopened.slides[1] ? partXml(reopened, reopened.slides[1].partUri) : '';
  const layoutOwnerXml = reopenedState
    ? ownerXml(layoutXml, reopenedLayout.placeholders.find(
      ({ name }) => name === layoutPrompt.name,
    ).id)
    : '';
  const masterOwnerXml = reopenedState
    ? ownerXml(masterXml, reopenedMaster.placeholders.find(
      ({ name }) => name === masterPrompt.name,
    ).id)
    : '';
  const populatedOwnerXml = reopenedPopulated ? ownerXml(slideXml, reopenedPopulated.id) : '';
  const directOwnerXml = reopenedDirect ? ownerXml(slideXml, reopenedDirect.id) : '';
  const duplicateOwnerXml = reopenedDuplicate
    ? ownerXml(duplicateXml, reopenedDuplicate.id)
    : '';
  const exactOoxml = {
    layoutPrompt: layoutOwnerXml.includes('algn="ctr"')
      && layoutOwnerXml.includes('<a:alpha val="87500"/>')
      && layoutOwnerXml.includes('lIns="50800" tIns="12700" rIns="25400" bIns="38100"')
      && layoutOwnerXml.includes('anchor="b"'),
    masterPrompt: masterOwnerXml.includes('algn="r"')
      && masterOwnerXml.includes('<a:alpha val="0"/>')
      && masterOwnerXml.includes('lIns="0" tIns="0" rIns="0" bIns="0"')
      && masterOwnerXml.includes('anchor="t"'),
    populatedOwner: populatedOwnerXml.includes('algn="l"')
      && populatedOwnerXml.includes('<a:alpha val="50000"/>')
      && populatedOwnerXml.includes('lIns="101600" tIns="50800"')
      && populatedOwnerXml.includes('anchor="b"'),
    directSlidePrompt: directOwnerXml.includes('algn="l"')
      && (directOwnerXml.match(/<a:alpha val="[0-9]+"\/>/gu) ?? []).length === 2
      && directOwnerXml.includes('<a:srgbClr val="ABCDEF"><a:alpha val="100000"/></a:srgbClr>')
      && directOwnerXml.includes('<a:alpha val="50000"/>')
      && directOwnerXml.includes('lIns="63500" tIns="25400" rIns="38100" bIns="50800"')
      && directOwnerXml.includes('anchor="ctr"'),
    duplicateOwner: duplicateOwnerXml.includes('algn="r"')
      && duplicateOwnerXml.includes('<a:alpha val="25000"/>')
      && duplicateOwnerXml.includes('lIns="101600"')
      && duplicateOwnerXml.includes('tIns="63500"')
      && duplicateOwnerXml.includes('rIns="76200"')
      && duplicateOwnerXml.includes('bIns="88900"')
      && duplicateOwnerXml.includes('anchor="t"'),
  };

  const formats = [];
  for (const format of FORMATS) {
    const formatted = api.PptxDocument.create({ format });
    const formattedLayout = formatted.layouts[0];
    formattedLayout.addPlaceholder([{
      runs: [{ text: format, style: { transparency: 25 } }],
    }], {
      name: `placeholder_style_${format}`,
      type: 'title',
      index: 250,
      align: 'justify',
      margin: [1, 2, 3, 4],
      valign: 'bottom',
    });
    const formattedSlide = formatted.addSlide({ masterName: formattedLayout.name });
    const formattedPopulated = formattedSlide.addText(`Populated ${format}`, {
      placeholder: `placeholder_style_${format}`,
      align: 'right',
      margin: 8,
      valign: 'top',
    });
    formattedPopulated.richText = [{
      runs: [{ text: `Edited ${format}`, style: { transparency: 50 } }],
      align: 'center',
    }];
    const formattedReopened = await api.PptxDocument.open(await formatted.write());
    const formattedOwner = formattedReopened.slides[0]?.shapes.find(
      ({ name }) => name === formattedPopulated.name,
    );
    formats.push({
      format,
      reopenedFormat: formattedReopened.format,
      state: formattedOwner && styleState(formattedOwner),
      errors: formattedReopened.diagnostics.filter(
        ({ severity }) => severity === 'error',
      ).length,
    });
  }
  const allFormats = formats.every(({ format, reopenedFormat, state, errors }) =>
    reopenedFormat === format
      && JSON.stringify(state?.align) === JSON.stringify(['center'])
      && JSON.stringify(state?.transparency) === JSON.stringify([50])
      && state?.margins?.top === 8
      && state?.margins?.right === 8
      && state?.margins?.bottom === 8
      && state?.margins?.left === 8
      && state?.valign === 'top'
      && errors === 0);

  const diagnostics = {
    createdErrors: document.diagnostics.filter(({ severity }) => severity === 'error').length,
    createdWarnings: document.diagnostics.filter(
      ({ severity }) => severity === 'warning',
    ).length,
    reopenedErrors: reopened.diagnostics.filter(({ severity }) => severity === 'error').length,
    reopenedWarnings: reopened.diagnostics.filter(
      ({ severity }) => severity === 'warning',
    ).length,
  };
  const state = {
    created,
    ownerReuse,
    noOp,
    invalidIsolation,
    invalidMessages,
    rollback,
    edited,
    duplicateState,
    duplicateIsolation,
    reopened: reopenedState,
    relationshipStability,
    exactOoxml,
    formats,
    allFormats,
    diagnostics,
  };
  const ok = ownerReuse
    && noOp
    && invalidIsolation
    && rollback
    && duplicateIsolation
    && JSON.stringify(reopenedState?.layout) === JSON.stringify(created.layout)
    && JSON.stringify(reopenedState?.master) === JSON.stringify(created.master)
    && JSON.stringify(reopenedState?.populated) === JSON.stringify(edited)
    && JSON.stringify(reopenedState?.slide) === JSON.stringify(created.slide)
    && JSON.stringify(reopenedState?.duplicate) === JSON.stringify(duplicateState)
    && relationshipStability
    && Object.values(exactOoxml).every(Boolean)
    && allFormats
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
