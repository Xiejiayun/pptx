async (page) => {
  const { readFile, writeFile } = await import('node:fs/promises');
  const { pathToFileURL } = await import('node:url');
  const chartPresentationWorkspace =
    process.env.PPTX_CHART_PRESENTATION_WORKSPACE ?? process.cwd();
  const chartPresentationFixtureModule = await import(pathToFileURL(
    `${chartPresentationWorkspace}/scripts/chart-presentation-91-fixture.mjs`,
  ).href);
  const chartPresentationFixtureBytes =
    await chartPresentationFixtureModule.createChartPresentation91Fixture();
  const chartPresentationProbeSource = await readFile(
    `${chartPresentationWorkspace}/scripts/chart-presentation-91-lifecycle-probe.mjs`,
    'utf8',
  );
  const imageIdentityEffects5ProbeSource = await readFile(
    `${chartPresentationWorkspace}/scripts/image-identity-effects-5-lifecycle-probe.mjs`,
    'utf8',
  );
  const chartPresentationFixtureBase64 = Buffer
    .from(chartPresentationFixtureBytes).toString('base64');
  const chartPresentationProbeBase64 = Buffer
    .from(chartPresentationProbeSource).toString('base64');
  const imageIdentityEffects5ProbeBase64 = Buffer
    .from(imageIdentityEffects5ProbeSource).toString('base64');
  const consoleErrors = [];
  const pageErrors = [];
  const networkErrors = [];
  const onConsole = (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  };
  const onPageError = (error) => pageErrors.push(error.message);
  const onRequestFailed = (request) => networkErrors.push(request.url());
  const onResponse = (response) => {
    if (response.status() >= 400) networkErrors.push(`${response.status()} ${response.url()}`);
  };
  page.on('console', onConsole);
  page.on('pageerror', onPageError);
  page.on('requestfailed', onRequestFailed);
  page.on('response', onResponse);
  const result = await page.evaluate(
    async ({
      moduleUrl,
      base64,
      chartPresentationFixtureBase64: presentationFixtureBase64,
      chartPresentationProbeBase64: presentationProbeBase64,
      imageIdentityEffects5ProbeBase64: imageEffectsProbeBase64,
    }) => {
      const binary = atob(base64);
      const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
      const api = await import(moduleUrl);
      const imageIdentityEffects5ProbeModule = await import(
        `data:text/javascript;base64,${imageEffectsProbeBase64}`
      );
      const imageIdentityEffects5Probe = await imageIdentityEffects5ProbeModule
        .runImageIdentityEffects5LifecycleProbe(api);
      const imageIdentityEffects5 = imageIdentityEffects5Probe.ok;
      const imageIdentityEffects5State = {
        noOp: imageIdentityEffects5Probe.state.noOp,
        rollback: imageIdentityEffects5Probe.state.rollback,
        sourceIsolation: imageIdentityEffects5Probe.state.sourceIsolation,
        sharedMedia: imageIdentityEffects5Probe.state.sharedMedia,
        relationships: Object.values(imageIdentityEffects5Probe.state.relationshipState)
          .every((count) => count === 1) &&
          Object.values(imageIdentityEffects5Probe.state.reopenedRelationshipState)
            .every((count) => count === 1),
        media: imageIdentityEffects5Probe.state.mediaPartCount === 1 &&
          imageIdentityEffects5Probe.state.mediaPreserved &&
          imageIdentityEffects5Probe.state.reopenedMediaPreserved,
        exactOoxml: Object.values(imageIdentityEffects5Probe.state.exactOoxml).every(Boolean),
        diagnostics: Object.values(imageIdentityEffects5Probe.state.diagnostics)
          .every((count) => count === 0),
      };
      globalThis.__pptxImageIdentityEffects5EvidenceBlob = new Blob(
        [imageIdentityEffects5Probe.explicitOutputBytes],
        { type: imageIdentityEffects5Probe.mime },
      );
      const chartPresentationProbeModule = await import(
        `data:text/javascript;base64,${presentationProbeBase64}`
      );
      const chartPresentationFixture = Uint8Array.from(
        atob(presentationFixtureBase64),
        (character) => character.charCodeAt(0),
      );
      const chartPresentation91Probe = await chartPresentationProbeModule
        .runChartPresentation91LifecycleProbe(api, chartPresentationFixture);
      const chartPresentation91 = chartPresentation91Probe.ok;
      const chartPresentation91State = chartPresentation91Probe.state;
      globalThis.__pptxChartPresentation91EvidenceBlob = new Blob(
        [chartPresentation91Probe.explicitOutputBytes],
        { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
      );
      const versionDocument = api.PptxDocument.create();
      const reopenedVersionDocument = await api.PptxDocument.open(
        await versionDocument.writeBlob(),
      );
      const presentationVersionState = {
        constant: api.PPTX_VERSION,
        created: versionDocument.version,
        reopened: reopenedVersionDocument.version,
      };
      const presentationVersion = Object.values(presentationVersionState)
        .every((value) => value === api.PPTX_VERSION)
        && versionDocument.diagnostics.filter(({ severity }) => severity === 'error').length === 0
        && reopenedVersionDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length === 0;
      const standardLayoutDocument = api.PptxDocument.create({ slideSize: '4:3' });
      const customLayoutDocument = api.PptxDocument.create({
        slideSize: { width: 10_698_480, height: 7_589_520 },
      });
      const detachedLayout = customLayoutDocument.presLayout;
      detachedLayout.width = 1;
      const customLayout = customLayoutDocument.presLayout;
      customLayoutDocument.slideSize = { width: 12_192_000, height: 6_858_000 };
      const editedLayout = customLayoutDocument.presLayout;
      const reopenedLayoutDocument = await api.PptxDocument.open(
        await customLayoutDocument.writeBlob(),
      );
      const presentationLayoutState = {
        standard: standardLayoutDocument.presLayout,
        custom: customLayout,
        edited: editedLayout,
        reopened: reopenedLayoutDocument.presLayout,
      };
      const presentationLayouts = JSON.stringify(presentationLayoutState) === JSON.stringify({
        standard: { name: 'screen4x3', width: 9_144_000, height: 6_858_000 },
        custom: { name: 'custom', width: 10_698_480, height: 7_589_520 },
        edited: { name: 'custom', width: 12_192_000, height: 6_858_000 },
        reopened: { name: 'custom', width: 12_192_000, height: 6_858_000 },
      })
        && customLayoutDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length === 0
        && reopenedLayoutDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length === 0;
      const horizontalAlignmentValues = [...api.TEXT_ALIGNMENTS];
      const horizontalAlignmentDocument = api.PptxDocument.create();
      const horizontalAlignmentSlide = horizontalAlignmentDocument.addSlide();
      horizontalAlignmentValues.forEach((alignment) => {
        horizontalAlignmentSlide.addText(alignment, { align: alignment });
      });
      const reopenedHorizontalAlignmentDocument = await api.PptxDocument.open(
        await horizontalAlignmentDocument.writeBlob(),
      );
      const horizontalAlignmentState = {
        values: horizontalAlignmentValues,
        reopened: reopenedHorizontalAlignmentDocument.slides[0].shapes
          .map(({ richText }) => richText[0]?.align),
        frozen: Object.isFrozen(api.TEXT_ALIGNMENTS),
      };
      const horizontalAlignments = JSON.stringify(horizontalAlignmentState) === JSON.stringify({
        values: ['left', 'center', 'right', 'justify'],
        reopened: ['left', 'center', 'right', 'justify'],
        frozen: true,
      })
        && reopenedHorizontalAlignmentDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length === 0;
      const verticalAlignmentValues = [...api.TEXT_VERTICAL_ALIGNMENTS];
      const verticalAlignmentDocument = api.PptxDocument.create();
      const verticalAlignmentSlide = verticalAlignmentDocument.addSlide();
      verticalAlignmentValues.forEach((alignment) => {
        verticalAlignmentSlide.addText(alignment, { valign: alignment });
      });
      verticalAlignmentSlide.addTable([
        verticalAlignmentValues.map((alignment) => ({
          text: alignment,
          options: { valign: alignment },
        })),
      ], { name: 'Browser vertical alignment table' });
      const reopenedVerticalAlignmentDocument = await api.PptxDocument.open(
        await verticalAlignmentDocument.writeBlob(),
      );
      const reopenedVerticalAlignmentTable = reopenedVerticalAlignmentDocument
        .slides[0].shapes.find(
          (shape) => shape.name === 'Browser vertical alignment table',
        );
      const verticalAlignmentState = {
        values: verticalAlignmentValues,
        textReopened: reopenedVerticalAlignmentDocument.slides[0].shapes
          .slice(0, 3).map(({ verticalAlignment }) => verticalAlignment),
        tableReopened: reopenedVerticalAlignmentTable instanceof api.TableModel
          ? reopenedVerticalAlignmentTable.rows[0].cells
            .map(({ verticalAlignment }) => verticalAlignment)
          : undefined,
        frozen: Object.isFrozen(api.TEXT_VERTICAL_ALIGNMENTS),
      };
      const verticalAlignments = JSON.stringify(verticalAlignmentState) === JSON.stringify({
        values: ['top', 'middle', 'bottom'],
        textReopened: ['top', 'middle', 'bottom'],
        tableReopened: ['top', 'middle', 'bottom'],
        frozen: true,
      })
        && reopenedVerticalAlignmentDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length === 0;
      const bulletNumberingStyles = [
        'alphaLcParenBoth',
        'alphaLcParenR',
        'alphaLcPeriod',
        'alphaUcParenBoth',
        'alphaUcParenR',
        'alphaUcPeriod',
        'arabicParenBoth',
        'arabicParenR',
        'arabicPeriod',
        'arabicPlain',
        'romanLcParenBoth',
        'romanLcParenR',
        'romanLcPeriod',
        'romanUcParenBoth',
        'romanUcParenR',
        'romanUcPeriod',
      ];
      const bulletSnapshot = (bullet) => {
        if (bullet === undefined) return null;
        return bullet.kind === 'bullet'
          ? { kind: bullet.kind, character: bullet.character, indent: bullet.indent }
          : {
              kind: bullet.kind,
              style: bullet.style,
              startAt: bullet.startAt,
              indent: bullet.indent,
            };
      };
      const namedOwnerXml = (xml, name, closingTag) => {
        const start = xml.indexOf(`name="${name}"`);
        const end = xml.indexOf(closingTag, start);
        return start >= 0 && end >= 0 ? xml.slice(start, end + closingTag.length) : '';
      };
      const bulletNumberingDocument = api.PptxDocument.create();
      const bulletNumberingLayout = bulletNumberingDocument.layouts[0];
      const bulletNumberingPlaceholder = bulletNumberingLayout.addPlaceholder(
        'Chrome bullet placeholder',
        {
          name: 'chrome_bullet_placeholder',
          type: 'body',
          index: 220,
          bullet: { kind: 'bullet', character: '→', indent: 19 },
        },
      );
      const bulletNumberingSlide = bulletNumberingDocument.addSlide({
        masterName: bulletNumberingLayout.name,
      });
      const bulletNumberingCatalogShape = bulletNumberingSlide.addRichText(
        bulletNumberingStyles.map((style, index) => ({
          runs: [{ text: style }],
          bullet: { kind: 'number', style, startAt: index + 1, indent: 20 },
        })),
        { name: 'Chrome bullet numbering catalog' },
      );
      const bulletNumberingTextShape = bulletNumberingSlide.addRichText([
        {
          runs: [{ text: 'Chrome custom bullet' }],
          bullet: { kind: 'bullet', character: '◆', indent: 18 },
        },
        {
          runs: [{ text: 'Chrome numbered paragraph' }],
          bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 7, indent: 22 },
        },
        { runs: [{ text: 'Chrome default bullet' }], bullet: true },
      ], { name: 'Chrome bullet text lifecycle' });
      const bulletNumberingTable = bulletNumberingSlide.addTable([[
        {
          text: [
            {
              runs: [{ text: 'Chrome table bullet' }],
              bullet: { kind: 'bullet', character: '→', indent: 17 },
            },
            {
              runs: [{ text: 'Chrome table number' }],
              bullet: { kind: 'number', style: 'arabicPlain', startAt: 6, indent: 19 },
            },
          ],
        },
      ]], { name: 'Chrome bullet table cell' });
      const initialBulletNumberingState = {
        text: bulletNumberingTextShape.richText.map(({ bullet }) =>
          bulletSnapshot(bullet)),
        placeholder: bulletSnapshot(bulletNumberingPlaceholder.richText[0]?.bullet),
        table: bulletNumberingTable.rows[0].cells[0].richText.map(({ bullet }) =>
          bulletSnapshot(bullet)),
      };
      const initialBulletNumberingSlideXml = new TextDecoder().decode(
        bulletNumberingDocument.opcPackage.requirePart(bulletNumberingSlide.partUri).bytes,
      );
      const initialBulletNumberingLayoutXml = new TextDecoder().decode(
        bulletNumberingDocument.opcPackage.requirePart(bulletNumberingLayout.partUri).bytes,
      );
      const initialCatalogXml = namedOwnerXml(
        initialBulletNumberingSlideXml,
        'Chrome bullet numbering catalog',
        '</p:sp>',
      );
      const initialTextLifecycleXml = namedOwnerXml(
        initialBulletNumberingSlideXml,
        'Chrome bullet text lifecycle',
        '</p:sp>',
      );
      const initialTableCellXml = namedOwnerXml(
        initialBulletNumberingSlideXml,
        'Chrome bullet table cell',
        '</p:graphicFrame>',
      );
      const initialPlaceholderXml = namedOwnerXml(
        initialBulletNumberingLayoutXml,
        'chrome_bullet_placeholder',
        '</p:sp>',
      );
      const initialCatalogTypes = [...initialCatalogXml.matchAll(
        /<a:buAutoNum type="([^"]+)" startAt="(\d+)"\/>/g,
      )].map((match) => ({ type: match[1], startAt: Number(match[2]) }));
      bulletNumberingTextShape.richText = [
        {
          runs: [{ text: 'Chrome edited number' }],
          bullet: { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
        },
        { runs: [{ text: 'Chrome cleared number' }], bullet: false },
        { runs: [{ text: 'Chrome retained default' }], bullet: true },
      ];
      bulletNumberingPlaceholder.richText = [{
        runs: [{ text: 'Chrome edited placeholder' }],
        bullet: { kind: 'number', style: 'alphaLcParenR', startAt: 5, indent: 20 },
      }];
      bulletNumberingTable.setCellRichText(0, 0, [
        {
          runs: [{ text: 'Chrome edited table bullet' }],
          bullet: { kind: 'bullet', character: '★', indent: 21 },
        },
        { runs: [{ text: 'Chrome cleared table number' }], bullet: false },
      ]);
      const editedBulletNumberingState = {
        text: bulletNumberingTextShape.richText.map(({ bullet }) =>
          bulletSnapshot(bullet)),
        placeholder: bulletSnapshot(bulletNumberingPlaceholder.richText[0]?.bullet),
        table: bulletNumberingTable.rows[0].cells[0].richText.map(({ bullet }) =>
          bulletSnapshot(bullet)),
      };
      const editedBulletNumberingSlideXml = new TextDecoder().decode(
        bulletNumberingDocument.opcPackage.requirePart(bulletNumberingSlide.partUri).bytes,
      );
      const editedBulletNumberingLayoutXml = new TextDecoder().decode(
        bulletNumberingDocument.opcPackage.requirePart(bulletNumberingLayout.partUri).bytes,
      );
      const editedTextLifecycleXml = namedOwnerXml(
        editedBulletNumberingSlideXml,
        'Chrome bullet text lifecycle',
        '</p:sp>',
      );
      const editedTableCellXml = namedOwnerXml(
        editedBulletNumberingSlideXml,
        'Chrome bullet table cell',
        '</p:graphicFrame>',
      );
      const editedPlaceholderXml = namedOwnerXml(
        editedBulletNumberingLayoutXml,
        'chrome_bullet_placeholder',
        '</p:sp>',
      );
      const bulletNumberingOutput = await bulletNumberingDocument.writeBlob();
      const reopenedBulletNumberingDocument = await api.PptxDocument.open(
        bulletNumberingOutput,
      );
      const reopenedBulletNumberingLayout = reopenedBulletNumberingDocument.layouts.find(
        ({ partUri }) => partUri === bulletNumberingLayout.partUri,
      );
      const reopenedBulletNumberingSlide = reopenedBulletNumberingDocument.slides.find(
        ({ partUri }) => partUri === bulletNumberingSlide.partUri,
      );
      const reopenedBulletNumberingCatalogShape = reopenedBulletNumberingSlide?.shapes.find(
        ({ name }) => name === 'Chrome bullet numbering catalog',
      );
      const reopenedBulletNumberingTextShape = reopenedBulletNumberingSlide?.shapes.find(
        ({ name }) => name === 'Chrome bullet text lifecycle',
      );
      const reopenedBulletNumberingTable = reopenedBulletNumberingSlide?.shapes.find(
        ({ name }) => name === 'Chrome bullet table cell',
      );
      const reopenedBulletNumberingPlaceholder = reopenedBulletNumberingLayout?.shapes.find(
        ({ name }) => name === 'chrome_bullet_placeholder',
      );
      const reopenedBulletNumberingSlideXml = reopenedBulletNumberingSlide === undefined
        ? ''
        : new TextDecoder().decode(
            reopenedBulletNumberingDocument.opcPackage
              .requirePart(reopenedBulletNumberingSlide.partUri).bytes,
          );
      const reopenedBulletNumberingLayoutXml = reopenedBulletNumberingLayout === undefined
        ? ''
        : new TextDecoder().decode(
            reopenedBulletNumberingDocument.opcPackage
              .requirePart(reopenedBulletNumberingLayout.partUri).bytes,
          );
      const reopenedCatalogXml = namedOwnerXml(
        reopenedBulletNumberingSlideXml,
        'Chrome bullet numbering catalog',
        '</p:sp>',
      );
      const reopenedTextLifecycleXml = namedOwnerXml(
        reopenedBulletNumberingSlideXml,
        'Chrome bullet text lifecycle',
        '</p:sp>',
      );
      const reopenedTableCellXml = namedOwnerXml(
        reopenedBulletNumberingSlideXml,
        'Chrome bullet table cell',
        '</p:graphicFrame>',
      );
      const reopenedPlaceholderXml = namedOwnerXml(
        reopenedBulletNumberingLayoutXml,
        'chrome_bullet_placeholder',
        '</p:sp>',
      );
      const reopenedCatalogTypes = [...reopenedCatalogXml.matchAll(
        /<a:buAutoNum type="([^"]+)" startAt="(\d+)"\/>/g,
      )].map((match) => ({ type: match[1], startAt: Number(match[2]) }));
      const reopenedBulletNumberingState = {
        text: reopenedBulletNumberingTextShape instanceof api.ShapeModel
          ? reopenedBulletNumberingTextShape.richText.map(({ bullet }) =>
              bulletSnapshot(bullet))
          : undefined,
        placeholder: reopenedBulletNumberingPlaceholder instanceof api.ShapeModel
          ? bulletSnapshot(reopenedBulletNumberingPlaceholder.richText[0]?.bullet)
          : undefined,
        table: reopenedBulletNumberingTable instanceof api.TableModel
          ? reopenedBulletNumberingTable.rows[0].cells[0].richText.map(({ bullet }) =>
              bulletSnapshot(bullet))
          : undefined,
      };
      const expectedCatalogTypes = bulletNumberingStyles.map((type, index) => ({
        type,
        startAt: index + 1,
      }));
      const bulletNumberingState = {
        mime: bulletNumberingOutput.type,
        catalog: bulletNumberingStyles,
        immediateCatalog: bulletNumberingCatalogShape.richText.map(
          ({ bullet }) => bullet?.kind === 'number' ? bullet.style : undefined,
        ),
        reopenedCatalog: reopenedBulletNumberingCatalogShape instanceof api.ShapeModel
          ? reopenedBulletNumberingCatalogShape.richText.map(
              ({ bullet }) => bullet?.kind === 'number' ? bullet.style : undefined,
            )
          : undefined,
        initial: initialBulletNumberingState,
        edited: editedBulletNumberingState,
        reopened: reopenedBulletNumberingState,
        exactOoxml: {
          catalog: JSON.stringify(initialCatalogTypes) === JSON.stringify(expectedCatalogTypes)
            && (initialCatalogXml.match(/indent="-254000" marL="254000"/g) ?? [])
              .length === bulletNumberingStyles.length,
          initialText: initialTextLifecycleXml.includes('<a:buChar char="◆"/>')
            && initialTextLifecycleXml.includes(
              '<a:buAutoNum type="romanUcPeriod" startAt="7"/>',
            )
            && initialTextLifecycleXml.includes('indent="-228600" marL="228600"')
            && initialTextLifecycleXml.includes('indent="-279400" marL="279400"'),
          initialPlaceholder: initialPlaceholderXml.includes('<a:buChar char="→"/>')
            && initialPlaceholderXml.includes('indent="-241300" marL="241300"'),
          initialTable: initialTableCellXml.includes('<a:buChar char="→"/>')
            && initialTableCellXml.includes(
              '<a:buAutoNum type="arabicPlain" startAt="6"/>',
            )
            && initialTableCellXml.includes('indent="-215900" marL="215900"')
            && initialTableCellXml.includes('indent="-241300" marL="241300"'),
          editedText: editedTextLifecycleXml.includes(
            '<a:buAutoNum type="romanLcParenR" startAt="4"/>',
          )
            && editedTextLifecycleXml.includes('indent="-304800" marL="304800"')
            && editedTextLifecycleXml.includes('<a:buChar char="•"/>')
            && editedTextLifecycleXml.includes('indent="-342900" marL="342900"')
            && (editedTextLifecycleXml.match(/<a:buNone\/>/g) ?? []).length === 1,
          editedPlaceholder: editedPlaceholderXml.includes(
            '<a:buAutoNum type="alphaLcParenR" startAt="5"/>',
          )
            && editedPlaceholderXml.includes('indent="-254000" marL="254000"'),
          editedTable: editedTableCellXml.includes('<a:buChar char="★"/>')
            && editedTableCellXml.includes('indent="-266700" marL="266700"')
            && (editedTableCellXml.match(/<a:buNone\/>/g) ?? []).length === 1,
          reopened: JSON.stringify(reopenedCatalogTypes) === JSON.stringify(expectedCatalogTypes)
            && reopenedTextLifecycleXml.includes(
              '<a:buAutoNum type="romanLcParenR" startAt="4"/>',
            )
            && reopenedTextLifecycleXml.includes('<a:buNone/>')
            && reopenedPlaceholderXml.includes(
              '<a:buAutoNum type="alphaLcParenR" startAt="5"/>',
            )
            && reopenedTableCellXml.includes('<a:buChar char="★"/>')
            && reopenedTableCellXml.includes('<a:buNone/>'),
        },
        diagnostics: bulletNumberingDocument.diagnostics.length
          + reopenedBulletNumberingDocument.diagnostics.length,
      };
      const expectedBulletNumberingParagraphs = {
        text: [
          { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
          null,
          { kind: 'bullet', character: '•', indent: 27 },
        ],
        placeholder: { kind: 'number', style: 'alphaLcParenR', startAt: 5, indent: 20 },
        table: [
          { kind: 'bullet', character: '★', indent: 21 },
          null,
        ],
      };
      const bulletNumbering = bulletNumberingState.mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && JSON.stringify(bulletNumberingState.immediateCatalog) ===
          JSON.stringify(bulletNumberingStyles)
        && JSON.stringify(bulletNumberingState.reopenedCatalog) ===
          JSON.stringify(bulletNumberingStyles)
        && JSON.stringify(bulletNumberingState.edited) ===
          JSON.stringify(expectedBulletNumberingParagraphs)
        && JSON.stringify(bulletNumberingState.reopened) ===
          JSON.stringify(expectedBulletNumberingParagraphs)
        && Object.values(bulletNumberingState.exactOoxml).every(Boolean)
        && bulletNumberingState.diagnostics === 0;
      const browserTabStops = [
        { position: 1.25, alignment: 'left' },
        { position: 2.5, alignment: 'center' },
        { position: 3.75, alignment: 'right' },
        { position: 4.5, alignment: 'decimal' },
      ];
      const browserTabStopsXml = '<a:tab pos="1143000" algn="l"/>'
        + '<a:tab pos="2286000" algn="ctr"/>'
        + '<a:tab pos="3429000" algn="r"/>'
        + '<a:tab pos="4114800" algn="dec"/>';
      const tabStopsDocument = api.PptxDocument.create();
      const tabStopsLayout = tabStopsDocument.layouts[0];
      const tabStopsPlaceholder = tabStopsLayout.addPlaceholder(
        'Chrome placeholder\t100',
        {
          name: 'chrome_tab_stops_placeholder',
          type: 'body',
          index: 221,
          tabStops: browserTabStops,
        },
      );
      const tabStopsSlide = tabStopsDocument.addSlide();
      const tabStopsText = tabStopsSlide.addText('Chrome text\t12.50', {
        name: 'Chrome tab stops text',
        tabStops: browserTabStops,
      });
      const tabStopsTable = tabStopsSlide.addTable([[{
        text: [{ runs: [{ text: 'Chrome cell\t42.75' }], tabStops: browserTabStops }],
      }]], { name: 'Chrome tab stops table' });
      const initialTabStopsSlideXml = new TextDecoder().decode(
        tabStopsDocument.opcPackage.requirePart(tabStopsSlide.partUri).bytes,
      );
      const initialTabStopsLayoutXml = new TextDecoder().decode(
        tabStopsDocument.opcPackage.requirePart(tabStopsLayout.partUri).bytes,
      );
      const tabStopsSnapshot = tabStopsText.richText;
      tabStopsSnapshot[0].tabStops[0].position = 99;
      const tabStopsSnapshotDetached =
        tabStopsText.richText[0].tabStops[0].position === 1.25;
      const beforeInvalidTabStops = tabStopsDocument.opcPackage
        .requirePart(tabStopsSlide.partUri).bytes.slice();
      const invalidTabStopsJournal = tabStopsDocument.opcPackage.mutations.length;
      const invalidTabStopsShapeCount = tabStopsSlide.shapes.length;
      let invalidTabStopsRejected = false;
      try {
        tabStopsSlide.addText('Invalid tab', {
          tabStops: [{ position: 1, alignment: 'tab' }],
        });
      } catch {
        invalidTabStopsRejected = true;
      }
      const afterInvalidTabStops = tabStopsDocument.opcPackage
        .requirePart(tabStopsSlide.partUri).bytes;
      const invalidTabStopsRollback = invalidTabStopsRejected
        && invalidTabStopsJournal === tabStopsDocument.opcPackage.mutations.length
        && invalidTabStopsShapeCount === tabStopsSlide.shapes.length
        && beforeInvalidTabStops.length === afterInvalidTabStops.length
        && beforeInvalidTabStops.every(
          (value, index) => value === afterInvalidTabStops[index],
        );
      tabStopsText.richText = [{
        runs: [{ text: 'Chrome edited\t12.50' }],
        tabStops: [{ position: 2.75, alignment: 'decimal' }],
      }];
      tabStopsPlaceholder.richText = [{
        runs: [{ text: 'Chrome explicit empty tabs' }],
        tabStops: [],
      }];
      tabStopsTable.setCellRichText(0, 0, [{
        runs: [{ text: 'Chrome cleared tabs' }],
        tabStops: false,
      }]);
      const editedTabStopsSlideXml = new TextDecoder().decode(
        tabStopsDocument.opcPackage.requirePart(tabStopsSlide.partUri).bytes,
      );
      const editedTabStopsLayoutXml = new TextDecoder().decode(
        tabStopsDocument.opcPackage.requirePart(tabStopsLayout.partUri).bytes,
      );
      const tabStopsOutput = await tabStopsDocument.writeBlob();
      const reopenedTabStopsDocument = await api.PptxDocument.open(tabStopsOutput);
      const reopenedTabStopsText = reopenedTabStopsDocument.slides[0].shapes.find(
        ({ name }) => name === tabStopsText.name,
      );
      const reopenedTabStopsPlaceholder = reopenedTabStopsDocument.layouts[0].shapes.find(
        ({ name }) => name === tabStopsPlaceholder.name,
      );
      const reopenedTabStopsTable = reopenedTabStopsDocument.slides[0].shapes.find(
        ({ name }) => name === tabStopsTable.name,
      );
      const tabStopsState = {
        snapshotDetached: tabStopsSnapshotDetached,
        immediate: initialTabStopsSlideXml.includes(browserTabStopsXml)
          && initialTabStopsLayoutXml.includes(browserTabStopsXml),
        initialOoxml: (initialTabStopsSlideXml.match(/<a:tab\b/g) ?? []).length === 8,
        invalidRollback: invalidTabStopsRollback,
        editedOoxml: editedTabStopsSlideXml.includes(
          '<a:tab pos="2514600" algn="dec"/>',
        ) && (editedTabStopsSlideXml.match(/<a:tab\b/g) ?? []).length === 1
          && editedTabStopsLayoutXml.includes('<a:tabLst></a:tabLst>'),
        reopenedText: reopenedTabStopsText instanceof api.ShapeModel
          && JSON.stringify(reopenedTabStopsText.richText[0].tabStops) ===
            JSON.stringify([{ position: 2.75, alignment: 'decimal' }]),
        reopenedPlaceholder: reopenedTabStopsPlaceholder instanceof api.ShapeModel
          && JSON.stringify(reopenedTabStopsPlaceholder.richText[0].tabStops) ===
            JSON.stringify([]),
        reopenedTable: reopenedTabStopsTable instanceof api.TableModel
          && reopenedTabStopsTable.rows[0].cells[0].richText[0].tabStops === undefined,
        diagnostics: tabStopsDocument.diagnostics.length === 0
          && reopenedTabStopsDocument.diagnostics.length === 0,
      };
      const tabStops = tabStopsOutput.type ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && Object.values(tabStopsState).every(Boolean);
      const browserUnderlineCommonStyles = [
        'dash',
        'dashHeavy',
        'dashLong',
        'dashLongHeavy',
        'dbl',
        'dotDash',
        'dotDotDash',
        'dotDotDashHeavy',
        'dotted',
        'dottedHeavy',
        'heavy',
        'sng',
        'wavy',
        'wavyDbl',
        'wavyHeavy',
      ];
      const browserUnderlineExpectedTokens = [
        ...browserUnderlineCommonStyles,
        'sng',
        'none',
        'dbl',
        'sng',
        'words',
        'dotDashHeavy',
      ];
      const browserUnderlineRuns = () => [
        ...browserUnderlineCommonStyles.map((style) => ({
          text: style,
          style: { underline: { style } },
        })),
        { text: 'true', style: { underline: true } },
        { text: 'none', style: { underline: false } },
        {
          text: 'srgb',
          style: {
            underline: {
              style: 'dbl',
              color: { kind: 'srgb', value: 'FF0000' },
            },
          },
        },
        {
          text: 'scheme',
          style: { underline: { color: { kind: 'scheme', value: 'accent2' } } },
        },
        { text: 'words', style: { underline: { style: 'words' } } },
        { text: 'dotDashHeavy', style: { underline: { style: 'dotDashHeavy' } } },
      ];
      const browserUnderlineSnapshot = (paragraphs) =>
        paragraphs[0].runs.map(({ style }) => style?.underline);
      const browserUnderlineTokens = (xml) => [...xml.matchAll(
        /<a:rPr\b[^>]*\bu="([^"]+)"/g,
      )].map((match) => match[1]);
      const underlineFamilyDocument = api.PptxDocument.create();
      const underlineFamilyLayout = underlineFamilyDocument.layouts[0];
      const browserUnderlinePlaceholder = underlineFamilyLayout.addPlaceholder(
        'Browser underline placeholder',
        { name: 'browser_underline_placeholder', type: 'body', index: 222 },
      );
      browserUnderlinePlaceholder.richText = [{ runs: browserUnderlineRuns() }];
      const underlineFamilySlide = underlineFamilyDocument.addSlide();
      const browserUnderlineText = underlineFamilySlide.addRichText(
        [{ runs: browserUnderlineRuns() }],
        { name: 'Browser underline text' },
      );
      const browserUnderlineCellTable = underlineFamilySlide.addTable([[{
        text: [{ runs: browserUnderlineRuns() }],
      }]], { name: 'Browser underline cell owner' });
      const browserUnderlineTableOwner = underlineFamilySlide.addTable([[{
        text: [{ runs: browserUnderlineRuns() }],
      }]], { name: 'Browser underline table owner equivalent' });
      const browserUnderlineInitialSnapshots = [
        browserUnderlineText.richText,
        browserUnderlinePlaceholder.richText,
        browserUnderlineCellTable.rows[0].cells[0].richText,
        browserUnderlineTableOwner.rows[0].cells[0].richText,
      ].map(browserUnderlineSnapshot);
      const browserUnderlineImmediate = browserUnderlineInitialSnapshots.every((snapshot) =>
        JSON.stringify(snapshot.slice(0, browserUnderlineCommonStyles.length)
          .map((underline) => typeof underline === 'object'
            ? underline.style : underline)) ===
          JSON.stringify(browserUnderlineCommonStyles)
        && JSON.stringify(snapshot.slice(browserUnderlineCommonStyles.length))
          === JSON.stringify([
            { style: 'sng' },
            false,
            { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
            { style: 'sng', color: { kind: 'scheme', value: 'accent2' } },
            { style: 'words' },
            { style: 'dotDashHeavy' },
          ]));
      const browserUnderlineDetached = browserUnderlineText.richText;
      browserUnderlineDetached[0].runs[0].style.underline.style = 'wavyDbl';
      browserUnderlineDetached[0].runs[browserUnderlineCommonStyles.length + 2]
        .style.underline.color.value = '00FF00';
      const browserUnderlineSnapshotDetached =
        browserUnderlineText.richText[0].runs[0].style.underline.style === 'dash'
        && browserUnderlineText.richText[0]
          .runs[browserUnderlineCommonStyles.length + 2]
          .style.underline.color.value === 'FF0000';
      const browserUnderlineInitialSlideXml = new TextDecoder().decode(
        underlineFamilyDocument.opcPackage.requirePart(underlineFamilySlide.partUri).bytes,
      );
      const browserUnderlineInitialLayoutXml = new TextDecoder().decode(
        underlineFamilyDocument.opcPackage.requirePart(underlineFamilyLayout.partUri).bytes,
      );
      const browserUnderlineInitialOoxml =
        JSON.stringify(browserUnderlineTokens(browserUnderlineInitialSlideXml))
          === JSON.stringify([
            ...browserUnderlineExpectedTokens,
            ...browserUnderlineExpectedTokens,
            ...browserUnderlineExpectedTokens,
          ])
        && JSON.stringify(browserUnderlineTokens(browserUnderlineInitialLayoutXml))
          === JSON.stringify(browserUnderlineExpectedTokens)
        && browserUnderlineInitialSlideXml.includes(
          '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
        )
        && browserUnderlineInitialSlideXml.includes(
          '<a:uFill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:uFill>',
        )
        && !browserUnderlineInitialSlideXml.includes('dotDashHeave"')
        && !browserUnderlineInitialLayoutXml.includes('dotDashHeave"');
      const browserUnderlineBeforeInvalid = underlineFamilyDocument.opcPackage
        .requirePart(underlineFamilySlide.partUri).bytes.slice();
      const browserUnderlineInvalidJournal = underlineFamilyDocument.opcPackage.mutations.length;
      let browserUnderlineInvalidRejected = false;
      try {
        browserUnderlineText.richText = [{ runs: [{
          text: 'Invalid upstream underline',
          style: { underline: { style: 'dotDashHeave' } },
        }] }];
      } catch {
        browserUnderlineInvalidRejected = true;
      }
      const browserUnderlineAfterInvalid = underlineFamilyDocument.opcPackage
        .requirePart(underlineFamilySlide.partUri).bytes;
      const browserUnderlineInvalidRollback = browserUnderlineInvalidRejected
        && browserUnderlineInvalidJournal === underlineFamilyDocument.opcPackage.mutations.length
        && browserUnderlineBeforeInvalid.length === browserUnderlineAfterInvalid.length
        && browserUnderlineBeforeInvalid.every(
          (value, index) => value === browserUnderlineAfterInvalid[index],
        );
      underlineFamilyDocument.duplicateSlide(0);
      const browserUnderlineEditedRuns = [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ];
      browserUnderlineText.richText = [{ runs: browserUnderlineEditedRuns }];
      browserUnderlinePlaceholder.richText = [{ runs: browserUnderlineEditedRuns }];
      browserUnderlineCellTable.setCellRichText(
        0,
        0,
        [{ runs: browserUnderlineEditedRuns }],
      );
      browserUnderlineTableOwner.setCellRichText(
        0,
        0,
        [{ runs: browserUnderlineEditedRuns }],
      );
      const browserUnderlineEditedSlideXml = new TextDecoder().decode(
        underlineFamilyDocument.opcPackage.requirePart(underlineFamilySlide.partUri).bytes,
      );
      const browserUnderlineEditedLayoutXml = new TextDecoder().decode(
        underlineFamilyDocument.opcPackage.requirePart(underlineFamilyLayout.partUri).bytes,
      );
      const browserUnderlineEditedOoxml =
        browserUnderlineTokens(browserUnderlineEditedSlideXml).every(
          (token) => token === 'none',
        )
        && browserUnderlineTokens(browserUnderlineEditedSlideXml).length === 3
        && JSON.stringify(browserUnderlineTokens(browserUnderlineEditedLayoutXml))
          === JSON.stringify(['none'])
        && !browserUnderlineEditedSlideXml.includes('<a:uFill>')
        && !browserUnderlineEditedLayoutXml.includes('<a:uFill>');
      const underlineFamilyOutput = await underlineFamilyDocument.writeBlob();
      const reopenedUnderlineFamilyDocument = await api.PptxDocument.open(
        underlineFamilyOutput,
      );
      const reopenedBrowserUnderlineText = reopenedUnderlineFamilyDocument
        .slides[0].shapes.find(({ name }) => name === browserUnderlineText.name);
      const reopenedBrowserUnderlinePlaceholder = reopenedUnderlineFamilyDocument
        .layouts[0].placeholders.find(({ name }) => name === browserUnderlinePlaceholder.name);
      const reopenedBrowserUnderlineCellTable = reopenedUnderlineFamilyDocument
        .slides[0].shapes.find(({ name }) => name === browserUnderlineCellTable.name);
      const reopenedBrowserUnderlineTableOwner = reopenedUnderlineFamilyDocument
        .slides[0].shapes.find(({ name }) => name === browserUnderlineTableOwner.name);
      const reopenedBrowserUnderlineDuplicate = reopenedUnderlineFamilyDocument
        .slides[1].shapes.find(({ name }) => name === browserUnderlineText.name);
      const browserUnderlineReopenedOwners = [
        reopenedBrowserUnderlineText.richText,
        reopenedBrowserUnderlinePlaceholder.richText,
        reopenedBrowserUnderlineCellTable.rows[0].cells[0].richText,
        reopenedBrowserUnderlineTableOwner.rows[0].cells[0].richText,
      ].every((paragraphs) => JSON.stringify(browserUnderlineSnapshot(paragraphs))
        === JSON.stringify([false, undefined]));
      const browserUnderlineDuplicateIsolation =
        JSON.stringify(browserUnderlineSnapshot(reopenedBrowserUnderlineDuplicate.richText)
          .slice(0, browserUnderlineCommonStyles.length).map((underline) =>
            typeof underline === 'object' ? underline.style : underline))
          === JSON.stringify(browserUnderlineCommonStyles)
        && browserUnderlineSnapshot(reopenedBrowserUnderlineDuplicate.richText)
          .at(-1).style === 'dotDashHeavy';
      const underlineFamilyState = {
        catalog: browserUnderlineInitialSnapshots.every((snapshot) =>
          JSON.stringify(snapshot.slice(0, browserUnderlineCommonStyles.length)
            .map((underline) => typeof underline === 'object'
              ? underline.style : underline)) === JSON.stringify(browserUnderlineCommonStyles)),
        trueAsSingle: browserUnderlineInitialSnapshots.every((snapshot) =>
          JSON.stringify(snapshot[browserUnderlineCommonStyles.length])
            === JSON.stringify({ style: 'sng' })),
        immediate: browserUnderlineImmediate,
        initialOoxml: browserUnderlineInitialOoxml,
        snapshotDetached: browserUnderlineSnapshotDetached,
        invalidRollback: browserUnderlineInvalidRollback,
        editedOoxml: browserUnderlineEditedOoxml,
        reopenedOwners: browserUnderlineReopenedOwners,
        duplicateIsolation: browserUnderlineDuplicateIsolation,
        diagnostics: underlineFamilyDocument.diagnostics.length === 0
          && reopenedUnderlineFamilyDocument.diagnostics.length === 0,
      };
      const underlineFamily = underlineFamilyOutput.type ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && Object.values(underlineFamilyState).every(Boolean);
      const presetShapeDocument = api.PptxDocument.create();
      const presetShapeSlide = presetShapeDocument.addSlide();
      const presetShapeModels = api.PRESET_SHAPE_TYPES.map(
        (type) => presetShapeSlide.addShape(type),
      );
      const presetShapeXml = new TextDecoder().decode(
        presetShapeDocument.opcPackage.requirePart(presetShapeSlide.partUri).bytes,
      );
      const serializedPresetShapeTypes = [...presetShapeXml.matchAll(
        /<a:prstGeom prst="([^"]+)"/g,
      )].map((match) => match[1]);
      const reopenedPresetShapeDocument = await api.PptxDocument.open(
        await presetShapeDocument.writeBlob(),
      );
      const reopenedPresetShapeTypes = reopenedPresetShapeDocument.slides[0].shapes.map(
        (shape) => shape instanceof api.ShapeModel ? shape.presetType : undefined,
      );
      const presetShapeState = {
        count: api.PRESET_SHAPE_TYPES.length,
        unique: new Set(api.PRESET_SHAPE_TYPES).size === api.PRESET_SHAPE_TYPES.length,
        frozen: Object.isFrozen(api.PRESET_SHAPE_TYPES),
        canonicalBoundary: api.PRESET_SHAPE_TYPES[0] === 'accentBorderCallout1'
          && api.PRESET_SHAPE_TYPES.at(-1) === 'wedgeRoundRectCallout'
          && api.PRESET_SHAPE_TYPES.includes('foldedCorner')
          && !api.PRESET_SHAPE_TYPES.includes('folderCorner'),
        created: presetShapeModels.every(
          ({ presetType }, index) => presetType === api.PRESET_SHAPE_TYPES[index],
        ),
        serialized: JSON.stringify(serializedPresetShapeTypes)
          === JSON.stringify(api.PRESET_SHAPE_TYPES),
        reopened: JSON.stringify(reopenedPresetShapeTypes)
          === JSON.stringify(api.PRESET_SHAPE_TYPES),
        validationErrors: reopenedPresetShapeDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length,
      };
      const presetShapes = JSON.stringify(presetShapeState) === JSON.stringify({
        count: 178,
        unique: true,
        frozen: true,
        canonicalBoundary: true,
        created: true,
        serialized: true,
        reopened: true,
        validationErrors: 0,
      });
      const browserCustomPathDocument = api.PptxDocument.create();
      const browserCustomPathSlide = browserCustomPathDocument.addSlide();
      const browserCustomPathSource = {
        paths: [{
          width: 400,
          height: 300,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'lineTo', point: { x: 100, y: 0 } },
            {
              kind: 'quadraticBezierTo',
              control: { x: 150, y: 0 },
              end: { x: 200, y: 100 },
            },
            {
              kind: 'cubicBezierTo',
              control1: { x: 225, y: 100 },
              control2: { x: 275, y: 200 },
              end: { x: 300, y: 200 },
            },
            {
              kind: 'arcTo',
              widthRadius: 100,
              heightRadius: 50,
              startAngle: 1800000,
              sweepAngle: 7200000,
            },
            { kind: 'moveTo', point: { x: 50, y: 50 } },
            { kind: 'lineTo', point: { x: 125, y: 125 } },
            { kind: 'close' },
          ],
        }],
      };
      const browserCustomPathExpected = structuredClone(browserCustomPathSource);
      const browserCustomPathShape = browserCustomPathSlide.addCustomShape(
        browserCustomPathSource,
        { name: 'Browser custom path' },
      );
      browserCustomPathSource.paths[0].width = 1;
      browserCustomPathSource.paths[0].commands.splice(0);
      const browserCustomPathSnapshot = browserCustomPathShape.customGeometry;
      const browserCustomPathXml = new TextDecoder().decode(
        browserCustomPathDocument.opcPackage.requirePart(browserCustomPathSlide.partUri).bytes,
      );
      const browserCustomPathCommandTags = [...browserCustomPathXml.matchAll(
        /<a:(moveTo|lnTo|quadBezTo|cubicBezTo|arcTo|close)(?:\s|\/?>)/g,
      )].map((match) => match[1]);
      const browserCustomPathBytes = browserCustomPathDocument.opcPackage
        .requirePart(browserCustomPathSlide.partUri).bytes.slice();
      const browserCustomPathJournal = JSON.stringify(
        browserCustomPathDocument.opcPackage.mutations,
      );
      let browserCustomPathRejected = false;
      try {
        browserCustomPathSlide.addCustomShape({
          paths: [{ width: 1, height: 1, commands: [{ kind: 'close' }] }],
        });
      } catch (error) {
        browserCustomPathRejected = error instanceof TypeError;
      }
      const browserCustomPathAfter = browserCustomPathDocument.opcPackage
        .requirePart(browserCustomPathSlide.partUri).bytes;
      const browserCustomPathOutput = await browserCustomPathDocument.writeBlob();
      globalThis.__pptxCustomPathEvidenceBlob = browserCustomPathOutput;
      const reopenedBrowserCustomPathDocument = await api.PptxDocument.open(
        browserCustomPathOutput,
      );
      const reopenedBrowserCustomPathShape = reopenedBrowserCustomPathDocument
        .slides[0].shapes[0];
      const browserCustomPathState = {
        live: browserCustomPathShape instanceof api.ShapeModel
          && browserCustomPathShape.presetType === undefined,
        detached: JSON.stringify(browserCustomPathSnapshot)
          === JSON.stringify(browserCustomPathExpected),
        frozen: Object.isFrozen(browserCustomPathSnapshot)
          && Object.isFrozen(browserCustomPathSnapshot?.paths)
          && Object.isFrozen(browserCustomPathSnapshot?.paths[0])
          && Object.isFrozen(browserCustomPathSnapshot?.paths[0]?.commands)
          && browserCustomPathSnapshot?.paths[0]?.commands.every((command) =>
            Object.isFrozen(command)
              && (!('point' in command) || Object.isFrozen(command.point))
              && (!('control' in command) || Object.isFrozen(command.control))
              && (!('control1' in command) || Object.isFrozen(command.control1))
              && (!('control2' in command) || Object.isFrozen(command.control2))
              && (!('end' in command) || Object.isFrozen(command.end))),
        serialized: JSON.stringify(browserCustomPathCommandTags) === JSON.stringify([
          'moveTo',
          'lnTo',
          'quadBezTo',
          'cubicBezTo',
          'arcTo',
          'moveTo',
          'lnTo',
          'close',
        ]),
        invalidRollback: browserCustomPathRejected
          && browserCustomPathBytes.length === browserCustomPathAfter.length
          && browserCustomPathBytes.every(
            (value, index) => value === browserCustomPathAfter[index],
          )
          && browserCustomPathJournal === JSON.stringify(
            browserCustomPathDocument.opcPackage.mutations,
          ),
        reopened: reopenedBrowserCustomPathShape instanceof api.ShapeModel
          && reopenedBrowserCustomPathShape.name === 'Browser custom path'
          && JSON.stringify(reopenedBrowserCustomPathShape.customGeometry)
            === JSON.stringify(browserCustomPathExpected),
        validationErrors: browserCustomPathDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length
          + reopenedBrowserCustomPathDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const browserCustomPaths = JSON.stringify(browserCustomPathState) === JSON.stringify({
        live: true,
        detached: true,
        frozen: true,
        serialized: true,
        invalidRollback: true,
        reopened: true,
        validationErrors: 0,
      });
      const shapeLineDashes = [
        'solid',
        'dash',
        'dashDot',
        'lgDash',
        'lgDashDot',
        'lgDashDotDot',
        'sysDash',
        'sysDot',
      ];
      const shapeArrowTypes = [
        'none',
        'arrow',
        'diamond',
        'oval',
        'stealth',
        'triangle',
      ];
      const shapeStyleDocument = api.PptxDocument.create();
      const shapeLineSlide = shapeStyleDocument.addSlide();
      const shapeLineModels = shapeLineDashes.map((dash) =>
        shapeLineSlide.addShape('rect', {
          line: {
            kind: 'line',
            color: { kind: 'scheme', value: 'accent2' },
            transparency: 25,
            width: 2.5,
            dash,
          },
        }));
      const shapeArrowSlide = shapeStyleDocument.addSlide();
      const shapeArrowModels = shapeArrowTypes.map((type) =>
        shapeArrowSlide.addShape('line', {
          arrows: { begin: type, end: type },
        }));
      const shapeStyleEditSlide = shapeStyleDocument.addSlide();
      const shapeStyleEditModel = shapeStyleEditSlide.addShape('line', {
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          dash: 'dashDot',
        },
        arrows: { begin: 'triangle', end: 'arrow' },
      });
      shapeStyleEditModel.line = {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
        width: 3,
        dash: 'sysDash',
      };
      const arrowsAfterLineEdit = shapeStyleEditModel.arrows;
      shapeStyleEditModel.arrows = { begin: 'diamond' };
      const lineAfterArrowEdit = shapeStyleEditModel.line;
      const shapeLineXml = new TextDecoder().decode(
        shapeStyleDocument.opcPackage.requirePart(shapeLineSlide.partUri).bytes,
      );
      const shapeArrowXml = new TextDecoder().decode(
        shapeStyleDocument.opcPackage.requirePart(shapeArrowSlide.partUri).bytes,
      );
      const serializedShapeLineDashes = [...shapeLineXml.matchAll(
        /<a:prstDash val="([^"]+)"\/>/g,
      )].map((match) => match[1]);
      const serializedShapeArrowBegins = [...shapeArrowXml.matchAll(
        /<a:headEnd type="([^"]+)"\/>/g,
      )].map((match) => match[1]);
      const serializedShapeArrowEnds = [...shapeArrowXml.matchAll(
        /<a:tailEnd type="([^"]+)"\/>/g,
      )].map((match) => match[1]);
      const reopenedShapeStyleDocument = await api.PptxDocument.open(
        await shapeStyleDocument.writeBlob(),
      );
      const reopenedShapeStyleEditModel = reopenedShapeStyleDocument.slides[2].shapes[0];
      const shapeLineState = {
        immediate: shapeLineModels.map(({ line }) => line?.dash),
        serialized: serializedShapeLineDashes,
        reopened: reopenedShapeStyleDocument.slides[0].shapes
          .map(({ line }) => line?.dash),
        preservedArrows: arrowsAfterLineEdit,
        edited: lineAfterArrowEdit,
        reopenedEdited: reopenedShapeStyleEditModel.line,
        validationErrors: reopenedShapeStyleDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length,
      };
      const shapeLines = JSON.stringify(shapeLineState) === JSON.stringify({
        immediate: shapeLineDashes,
        serialized: shapeLineDashes,
        reopened: shapeLineDashes,
        preservedArrows: { begin: 'triangle', end: 'arrow' },
        edited: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
          width: 3,
          dash: 'sysDash',
        },
        reopenedEdited: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
          width: 3,
          dash: 'sysDash',
        },
        validationErrors: 0,
      });
      const shapeArrowState = {
        immediateBegins: shapeArrowModels.map(({ arrows }) => arrows?.begin),
        immediateEnds: shapeArrowModels.map(({ arrows }) => arrows?.end),
        serializedBegins: serializedShapeArrowBegins,
        serializedEnds: serializedShapeArrowEnds,
        reopenedBegins: reopenedShapeStyleDocument.slides[1].shapes
          .map(({ arrows }) => arrows?.begin),
        reopenedEnds: reopenedShapeStyleDocument.slides[1].shapes
          .map(({ arrows }) => arrows?.end),
        preservedLine: lineAfterArrowEdit,
        reopenedEdited: reopenedShapeStyleEditModel.arrows,
        validationErrors: reopenedShapeStyleDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length,
      };
      const shapeArrows = JSON.stringify(shapeArrowState) === JSON.stringify({
        immediateBegins: shapeArrowTypes,
        immediateEnds: shapeArrowTypes,
        serializedBegins: shapeArrowTypes,
        serializedEnds: shapeArrowTypes,
        reopenedBegins: shapeArrowTypes,
        reopenedEnds: shapeArrowTypes,
        preservedLine: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
          width: 3,
          dash: 'sysDash',
        },
        reopenedEdited: { begin: 'diamond' },
        validationErrors: 0,
      });
      const tableVerticalAlignmentDocument = api.PptxDocument.create();
      const tableVerticalAlignmentSlide = tableVerticalAlignmentDocument.addSlide();
      const tableVerticalAlignmentTable = tableVerticalAlignmentSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], { name: 'Chrome table vertical alignment', valign: 'middle' });
      const tableVerticalAlignmentPart = () => tableVerticalAlignmentDocument.opcPackage
        .requirePart(tableVerticalAlignmentSlide.partUri).bytes;
      const tableVerticalAlignmentBytesEqual = (left, right) =>
        left.byteLength === right.byteLength &&
        left.every((value, index) => value === right[index]);
      const tableVerticalAlignmentReadBytes = tableVerticalAlignmentPart().slice();
      const tableVerticalAlignmentReadJournal = JSON.stringify(
        tableVerticalAlignmentDocument.opcPackage.mutations,
      );
      const tableVerticalAlignmentUniform = tableVerticalAlignmentTable.verticalAlignment;
      const tableVerticalAlignmentReadIsolation = tableVerticalAlignmentBytesEqual(
        tableVerticalAlignmentReadBytes,
        tableVerticalAlignmentPart(),
      ) && JSON.stringify(tableVerticalAlignmentDocument.opcPackage.mutations) ===
        tableVerticalAlignmentReadJournal;
      const tableVerticalAlignmentNoOpBytes = tableVerticalAlignmentPart().slice();
      const tableVerticalAlignmentNoOpJournal = JSON.stringify(
        tableVerticalAlignmentDocument.opcPackage.mutations,
      );
      tableVerticalAlignmentTable.verticalAlignment = 'middle';
      const tableVerticalAlignmentNoOp = tableVerticalAlignmentBytesEqual(
        tableVerticalAlignmentNoOpBytes,
        tableVerticalAlignmentPart(),
      ) && JSON.stringify(tableVerticalAlignmentDocument.opcPackage.mutations) ===
        tableVerticalAlignmentNoOpJournal;
      tableVerticalAlignmentTable.setCellVerticalAlignment(0, 1, 'top');
      const tableVerticalAlignmentMixed = tableVerticalAlignmentTable
        .verticalAlignment ?? null;
      tableVerticalAlignmentTable.verticalAlignment = 'bottom';
      const tableVerticalAlignmentOverwritten = tableVerticalAlignmentTable
        .verticalAlignment;
      const tableVerticalAlignmentOverwrittenCells = tableVerticalAlignmentTable.rows
        .flatMap(({ cells }) => cells.map(
          ({ verticalAlignment }) => verticalAlignment ?? null,
        ));
      tableVerticalAlignmentTable.verticalAlignment = undefined;
      const tableVerticalAlignmentCleared = tableVerticalAlignmentTable
        .verticalAlignment ?? null;
      const tableVerticalAlignmentClearedCells = tableVerticalAlignmentTable.rows
        .flatMap(({ cells }) => cells.map(
          ({ verticalAlignment }) => verticalAlignment ?? null,
        ));
      const tableVerticalAlignmentInvalidBytes = tableVerticalAlignmentPart().slice();
      const tableVerticalAlignmentInvalidJournal = JSON.stringify(
        tableVerticalAlignmentDocument.opcPackage.mutations,
      );
      let tableVerticalAlignmentInvalidError;
      try {
        tableVerticalAlignmentTable.verticalAlignment = 'distributed';
      } catch (error) {
        tableVerticalAlignmentInvalidError = {
          name: error.name,
          message: error.message,
        };
      }
      const tableVerticalAlignmentFailureIsolation = tableVerticalAlignmentBytesEqual(
        tableVerticalAlignmentInvalidBytes,
        tableVerticalAlignmentPart(),
      ) && JSON.stringify(tableVerticalAlignmentDocument.opcPackage.mutations) ===
        tableVerticalAlignmentInvalidJournal;
      tableVerticalAlignmentTable.verticalAlignment = 'top';
      const reopenedTableVerticalAlignmentDocument = await api.PptxDocument.open(
        await tableVerticalAlignmentDocument.writeBlob(),
      );
      const reopenedTableVerticalAlignmentTable = reopenedTableVerticalAlignmentDocument
        .slides[0].shapes.find(
          (shape) => shape.name === 'Chrome table vertical alignment',
        );
      const tableVerticalAlignmentState = {
        uniform: tableVerticalAlignmentUniform,
        readIsolation: tableVerticalAlignmentReadIsolation,
        noOp: tableVerticalAlignmentNoOp,
        mixed: tableVerticalAlignmentMixed,
        overwritten: tableVerticalAlignmentOverwritten,
        overwrittenCells: tableVerticalAlignmentOverwrittenCells,
        cleared: tableVerticalAlignmentCleared,
        clearedCells: tableVerticalAlignmentClearedCells,
        reopened: reopenedTableVerticalAlignmentTable instanceof api.TableModel
          ? reopenedTableVerticalAlignmentTable.verticalAlignment ?? null
          : null,
        reopenedCells: reopenedTableVerticalAlignmentTable instanceof api.TableModel
          ? reopenedTableVerticalAlignmentTable.rows.flatMap(({ cells }) => cells.map(
            ({ verticalAlignment }) => verticalAlignment ?? null,
          ))
          : [],
        invalidError: tableVerticalAlignmentInvalidError,
        failureIsolation: tableVerticalAlignmentFailureIsolation,
        validationErrors: tableVerticalAlignmentDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableVerticalAlignmentDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableVerticalAlignment = JSON.stringify(tableVerticalAlignmentState) ===
        JSON.stringify({
          uniform: 'middle',
          readIsolation: true,
          noOp: true,
          mixed: null,
          overwritten: 'bottom',
          overwrittenCells: ['bottom', 'bottom', 'bottom', 'bottom'],
          cleared: null,
          clearedCells: [null, null, null, null],
          reopened: 'top',
          reopenedCells: ['top', 'top', 'top', 'top'],
          invalidError: {
            name: 'TypeError',
            message: 'Table vertical alignment must be top, middle, or bottom',
          },
          failureIsolation: true,
          validationErrors: 0,
        });
      const tableTextDirectionDocument = api.PptxDocument.create();
      const tableTextDirectionSlide = tableTextDirectionDocument.addSlide();
      const tableTextDirectionTable = tableTextDirectionSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], { name: 'Chrome table text direction', textDirection: 'vert270' });
      const tableTextDirectionPart = () => tableTextDirectionDocument.opcPackage
        .requirePart(tableTextDirectionSlide.partUri).bytes;
      const tableTextDirectionBytesEqual = (left, right) =>
        left.byteLength === right.byteLength &&
        left.every((value, index) => value === right[index]);
      const tableTextDirectionReadBytes = tableTextDirectionPart().slice();
      const tableTextDirectionReadJournal = JSON.stringify(
        tableTextDirectionDocument.opcPackage.mutations,
      );
      const tableTextDirectionUniform = tableTextDirectionTable.textDirection;
      const tableTextDirectionReadIsolation = tableTextDirectionBytesEqual(
        tableTextDirectionReadBytes,
        tableTextDirectionPart(),
      ) && JSON.stringify(tableTextDirectionDocument.opcPackage.mutations) ===
        tableTextDirectionReadJournal;
      const tableTextDirectionNoOpBytes = tableTextDirectionPart().slice();
      const tableTextDirectionNoOpJournal = JSON.stringify(
        tableTextDirectionDocument.opcPackage.mutations,
      );
      tableTextDirectionTable.textDirection = 'vert270';
      const tableTextDirectionNoOp = tableTextDirectionBytesEqual(
        tableTextDirectionNoOpBytes,
        tableTextDirectionPart(),
      ) && JSON.stringify(tableTextDirectionDocument.opcPackage.mutations) ===
        tableTextDirectionNoOpJournal;
      tableTextDirectionTable.setCellTextDirection(0, 1, 'vert');
      const tableTextDirectionMixed = tableTextDirectionTable.textDirection ?? null;
      tableTextDirectionTable.textDirection = 'wordArtVert';
      const tableTextDirectionOverwritten = tableTextDirectionTable.textDirection;
      const tableTextDirectionOverwrittenCells = tableTextDirectionTable.rows
        .flatMap(({ cells }) => cells.map(({ textDirection }) => textDirection ?? null));
      tableTextDirectionTable.textDirection = 'horz';
      const tableTextDirectionHorizontal = tableTextDirectionTable.textDirection;
      const tableTextDirectionHorizontalCells = tableTextDirectionTable.rows
        .flatMap(({ cells }) => cells.map(({ textDirection }) => textDirection ?? null));
      tableTextDirectionTable.textDirection = undefined;
      const tableTextDirectionCleared = tableTextDirectionTable.textDirection ?? null;
      const tableTextDirectionClearedCells = tableTextDirectionTable.rows
        .flatMap(({ cells }) => cells.map(({ textDirection }) => textDirection ?? null));
      const tableTextDirectionInvalidBytes = tableTextDirectionPart().slice();
      const tableTextDirectionInvalidJournal = JSON.stringify(
        tableTextDirectionDocument.opcPackage.mutations,
      );
      let tableTextDirectionInvalidError;
      try {
        tableTextDirectionTable.textDirection = 'eaVert';
      } catch (error) {
        tableTextDirectionInvalidError = {
          name: error.name,
          message: error.message,
        };
      }
      const tableTextDirectionFailureIsolation = tableTextDirectionBytesEqual(
        tableTextDirectionInvalidBytes,
        tableTextDirectionPart(),
      ) && JSON.stringify(tableTextDirectionDocument.opcPackage.mutations) ===
        tableTextDirectionInvalidJournal;
      tableTextDirectionTable.textDirection = 'vert';
      const reopenedTableTextDirectionDocument = await api.PptxDocument.open(
        await tableTextDirectionDocument.writeBlob(),
      );
      const reopenedTableTextDirectionTable = reopenedTableTextDirectionDocument
        .slides[0].shapes.find(
          (shape) => shape.name === 'Chrome table text direction',
        );
      const tableTextDirectionState = {
        uniform: tableTextDirectionUniform,
        readIsolation: tableTextDirectionReadIsolation,
        noOp: tableTextDirectionNoOp,
        mixed: tableTextDirectionMixed,
        overwritten: tableTextDirectionOverwritten,
        overwrittenCells: tableTextDirectionOverwrittenCells,
        horizontal: tableTextDirectionHorizontal,
        horizontalCells: tableTextDirectionHorizontalCells,
        cleared: tableTextDirectionCleared,
        clearedCells: tableTextDirectionClearedCells,
        reopened: reopenedTableTextDirectionTable instanceof api.TableModel
          ? reopenedTableTextDirectionTable.textDirection ?? null
          : null,
        reopenedCells: reopenedTableTextDirectionTable instanceof api.TableModel
          ? reopenedTableTextDirectionTable.rows.flatMap(({ cells }) => cells.map(
            ({ textDirection }) => textDirection ?? null,
          ))
          : [],
        invalidError: tableTextDirectionInvalidError,
        failureIsolation: tableTextDirectionFailureIsolation,
        validationErrors: tableTextDirectionDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableTextDirectionDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableTextDirection = JSON.stringify(tableTextDirectionState) ===
        JSON.stringify({
          uniform: 'vert270',
          readIsolation: true,
          noOp: true,
          mixed: null,
          overwritten: 'wordArtVert',
          overwrittenCells: ['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert'],
          horizontal: 'horz',
          horizontalCells: ['horz', 'horz', 'horz', 'horz'],
          cleared: null,
          clearedCells: [null, null, null, null],
          reopened: 'vert',
          reopenedCells: ['vert', 'vert', 'vert', 'vert'],
          invalidError: {
            name: 'TypeError',
            message: 'Table text direction must be horz, vert, vert270, or wordArtVert',
          },
          failureIsolation: true,
          validationErrors: 0,
        });
      const browserTextDirectionValues = [
        'eaVert',
        'horz',
        'mongolianVert',
        'vert',
        'vert270',
        'wordArtVert',
        'wordArtVertRtl',
      ];
      const textDirectionFamilyDocument = api.PptxDocument.create();
      const textDirectionFamilySlide = textDirectionFamilyDocument.addSlide();
      const browserTextDirectionShapes = browserTextDirectionValues.map((vert) =>
        textDirectionFamilySlide.addText(`Browser ${vert}`, {
          name: `browser_text_direction_${vert}`,
          vert,
        }));
      const browserTextDirectionPart = () => textDirectionFamilyDocument.opcPackage
        .requirePart(textDirectionFamilySlide.partUri).bytes;
      const browserTextDirectionTokens = (bytes) => [
        ...new TextDecoder().decode(bytes).matchAll(
          /<a:bodyPr\b[^>]*\bvert="([^"]+)"/g,
        ),
      ].map((match) => match[1]);
      const browserTextDirectionCatalog = JSON.stringify(
        browserTextDirectionShapes.map(({ textDirection }) => textDirection),
      ) === JSON.stringify(browserTextDirectionValues);
      const browserTextDirectionInitialOoxml = JSON.stringify(
        browserTextDirectionTokens(browserTextDirectionPart()),
      ) === JSON.stringify(browserTextDirectionValues);
      const browserTextDirectionInvalidBytes = browserTextDirectionPart().slice();
      const browserTextDirectionInvalidJournal = JSON.stringify(
        textDirectionFamilyDocument.opcPackage.mutations,
      );
      let browserTextDirectionInvalidRejected = false;
      try {
        browserTextDirectionShapes[0].textDirection = 'vertical';
      } catch {
        browserTextDirectionInvalidRejected = true;
      }
      const browserTextDirectionInvalidRollback = browserTextDirectionInvalidRejected
        && tableTextDirectionBytesEqual(
          browserTextDirectionInvalidBytes,
          browserTextDirectionPart(),
        )
        && browserTextDirectionInvalidJournal === JSON.stringify(
          textDirectionFamilyDocument.opcPackage.mutations,
        );
      textDirectionFamilyDocument.duplicateSlide(0);
      browserTextDirectionShapes[0].textDirection = 'vert270';
      browserTextDirectionShapes[1].textDirection = undefined;
      const browserTextDirectionEdited = [
        'vert270',
        undefined,
        ...browserTextDirectionValues.slice(2),
      ];
      const browserTextDirectionEditedOoxml = JSON.stringify(
        browserTextDirectionTokens(browserTextDirectionPart()),
      ) === JSON.stringify(
        browserTextDirectionEdited.filter((value) => value !== undefined),
      );
      const reopenedTextDirectionFamilyDocument = await api.PptxDocument.open(
        await textDirectionFamilyDocument.writeBlob(),
      );
      const reopenedBrowserTextDirectionValues = reopenedTextDirectionFamilyDocument
        .slides[0].shapes.map(({ textDirection }) => textDirection);
      const duplicateBrowserTextDirectionValues = reopenedTextDirectionFamilyDocument
        .slides[1].shapes.map(({ textDirection }) => textDirection);
      const textDirectionFamilyState = {
        table: tableTextDirection,
        catalog: browserTextDirectionCatalog,
        initialOoxml: browserTextDirectionInitialOoxml,
        invalidRollback: browserTextDirectionInvalidRollback,
        editedOoxml: browserTextDirectionEditedOoxml,
        reopened: JSON.stringify(reopenedBrowserTextDirectionValues) ===
          JSON.stringify(browserTextDirectionEdited),
        duplicateIsolation: JSON.stringify(duplicateBrowserTextDirectionValues) ===
          JSON.stringify(browserTextDirectionValues),
        diagnostics: textDirectionFamilyDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0 && reopenedTextDirectionFamilyDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0,
      };
      const textDirectionFamily = Object.values(textDirectionFamilyState).every(Boolean);
      const textBoxFitFamilyDocument = api.PptxDocument.create();
      const textBoxFitFamilySlide = textBoxFitFamilyDocument.addSlide();
      const browserTextBoxFitShapes = [
        textBoxFitFamilySlide.addText('Browser fit omitted', {
          name: 'browser_text_box_fit_omitted',
        }),
        textBoxFitFamilySlide.addText('Browser fit none', {
          name: 'browser_text_box_fit_none',
          fit: 'none',
        }),
        textBoxFitFamilySlide.addText('Browser fit shrink', {
          name: 'browser_text_box_fit_shrink',
          fit: 'shrink',
        }),
        textBoxFitFamilySlide.addText('Browser fit resize', {
          name: 'browser_text_box_fit_resize',
          fit: 'resize',
        }),
      ];
      const browserTextBoxFitPart = () => textBoxFitFamilyDocument.opcPackage
        .requirePart(textBoxFitFamilySlide.partUri).bytes;
      const browserTextBoxFitBytesEqual = (left, right) =>
        left.byteLength === right.byteLength
        && left.every((value, index) => value === right[index]);
      const browserTextBoxFitXml = new TextDecoder().decode(browserTextBoxFitPart());
      const browserTextBoxFitOwners = Object.fromEntries(
        ['omitted', 'none', 'shrink', 'resize'].map((value) => [
          value,
          namedOwnerXml(
            browserTextBoxFitXml,
            `browser_text_box_fit_${value}`,
            '</p:sp>',
          ),
        ]),
      );
      const browserTextBoxFitInitial = JSON.stringify(
        browserTextBoxFitShapes.map(({ textFit }) => textFit),
      ) === JSON.stringify([undefined, undefined, 'shrink', 'resize']);
      const browserTextBoxFitInitialOoxml =
        !/<a:(?:noAutofit|normAutofit|spAutoFit)\b/u.test(
          browserTextBoxFitOwners.omitted,
        )
        && !/<a:(?:noAutofit|normAutofit|spAutoFit)\b/u.test(
          browserTextBoxFitOwners.none,
        )
        && /<a:normAutofit\/>/u.test(browserTextBoxFitOwners.shrink)
        && !/<a:(?:noAutofit|spAutoFit)\b/u.test(browserTextBoxFitOwners.shrink)
        && /<a:spAutoFit\/>/u.test(browserTextBoxFitOwners.resize)
        && !/<a:(?:noAutofit|normAutofit)\b/u.test(browserTextBoxFitOwners.resize);
      const browserTextBoxFitInvalidBytes = browserTextBoxFitPart().slice();
      const browserTextBoxFitInvalidJournal = JSON.stringify(
        textBoxFitFamilyDocument.opcPackage.mutations,
      );
      let browserTextBoxFitInvalidRejected = false;
      try {
        browserTextBoxFitShapes[1].textFit = 'SHRINK';
      } catch {
        browserTextBoxFitInvalidRejected = true;
      }
      const browserTextBoxFitInvalidRollback = browserTextBoxFitInvalidRejected
        && browserTextBoxFitBytesEqual(
          browserTextBoxFitInvalidBytes,
          browserTextBoxFitPart(),
        )
        && browserTextBoxFitInvalidJournal === JSON.stringify(
          textBoxFitFamilyDocument.opcPackage.mutations,
        );
      textBoxFitFamilyDocument.duplicateSlide(0);
      browserTextBoxFitShapes[0].textFit = 'shrink';
      browserTextBoxFitShapes[2].textFit = 'resize';
      browserTextBoxFitShapes[3].textFit = 'none';
      const browserTextBoxFitEditedXml = new TextDecoder().decode(browserTextBoxFitPart());
      const browserTextBoxFitEdited = [
        ['omitted', 'normAutofit'],
        ['none', undefined],
        ['shrink', 'spAutoFit'],
        ['resize', undefined],
      ].every(([name, token]) => {
        const owner = namedOwnerXml(
          browserTextBoxFitEditedXml,
          `browser_text_box_fit_${name}`,
          '</p:sp>',
        );
        return token === undefined
          ? !/<a:(?:noAutofit|normAutofit|spAutoFit)\b/u.test(owner)
          : new RegExp(`<a:${token}\\/>`, 'u').test(owner);
      });
      const textBoxFitFamilyOutput = await textBoxFitFamilyDocument.writeBlob();
      globalThis.__pptxTextBoxFitEvidenceBlob = textBoxFitFamilyOutput;
      const reopenedTextBoxFitFamilyDocument = await api.PptxDocument.open(
        textBoxFitFamilyOutput,
      );
      const textBoxFitFamilyState = {
        initial: browserTextBoxFitInitial,
        initialOoxml: browserTextBoxFitInitialOoxml,
        invalidRollback: browserTextBoxFitInvalidRollback,
        editedOoxml: browserTextBoxFitEdited,
        reopened: JSON.stringify(
          reopenedTextBoxFitFamilyDocument.slides[0].shapes.map(({ textFit }) => textFit),
        ) === JSON.stringify(['shrink', undefined, 'resize', undefined]),
        duplicateIsolation: JSON.stringify(
          reopenedTextBoxFitFamilyDocument.slides[1].shapes.map(({ textFit }) => textFit),
        ) === JSON.stringify([undefined, undefined, 'shrink', 'resize']),
        mime: textBoxFitFamilyOutput.type ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        diagnostics: textBoxFitFamilyDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0 && reopenedTextBoxFitFamilyDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0,
      };
      const textBoxFitFamily = Object.values(textBoxFitFamilyState).every(Boolean);
      const textParagraphLayoutDocument = api.PptxDocument.create();
      const textParagraphLayoutSlide = textParagraphLayoutDocument.addSlide();
      const textParagraphLayoutShape = textParagraphLayoutSlide.addText(
        'Browser paragraph layout',
        {
          name: 'browser_text_paragraph_layout',
          align: 'center',
          rtlMode: true,
          spacing: {
            before: 6.25,
            after: 8.5,
            line: { kind: 'exact', points: 28 },
          },
          margin: [1, 2, 3, 4],
          valign: 'bottom',
          wrap: false,
        },
      );
      const textParagraphLayoutCatalogShapes = [
        textParagraphLayoutSlide.addText('Browser paragraph layout left scalar middle', {
          name: 'browser_text_paragraph_layout_left_scalar_middle',
          align: 'left',
          rtlMode: false,
          margin: 8,
          valign: 'middle',
          wrap: true,
        }),
        textParagraphLayoutSlide.addText('Browser paragraph layout justify zero top', {
          name: 'browser_text_paragraph_layout_justify_zero_top',
          align: 'justify',
          rtlMode: true,
          margin: 0,
          valign: 'top',
          wrap: false,
        }),
      ];
      const textParagraphLayoutPart = () => textParagraphLayoutDocument.opcPackage
        .requirePart(textParagraphLayoutSlide.partUri).bytes;
      const textParagraphLayoutBytesEqual = (left, right) =>
        left.byteLength === right.byteLength
        && left.every((value, index) => value === right[index]);
      const textParagraphLayoutSnapshot = (shape) => ({
        paragraphs: shape.richText.map(({ align, rtl, spacing }) => ({
          align,
          rtl,
          spacing,
        })),
        margins: shape.textMargins,
        verticalAlignment: shape.verticalAlignment,
        textWrap: shape.textWrap,
      });
      const initialTextParagraphLayoutExpected = {
        paragraphs: [{
          align: 'center',
          rtl: true,
          spacing: {
            before: 6.25,
            after: 8.5,
            line: { kind: 'exact', points: 28 },
          },
        }],
        margins: { left: 4, top: 1, right: 2, bottom: 3 },
        verticalAlignment: 'bottom',
        textWrap: false,
      };
      const textParagraphLayoutCatalogExpected = [
        {
          paragraphs: [{ align: 'left', rtl: false }],
          margins: { left: 8, top: 8, right: 8, bottom: 8 },
          verticalAlignment: 'middle',
          textWrap: true,
        },
        {
          paragraphs: [{ align: 'justify', rtl: true }],
          margins: { left: 0, top: 0, right: 0, bottom: 0 },
          verticalAlignment: 'top',
          textWrap: false,
        },
      ];
      const textParagraphLayoutImmediate = JSON.stringify(
        textParagraphLayoutSnapshot(textParagraphLayoutShape),
      ) === JSON.stringify(initialTextParagraphLayoutExpected);
      const textParagraphLayoutCatalogImmediate = JSON.stringify(
        textParagraphLayoutCatalogShapes.map(textParagraphLayoutSnapshot),
      ) === JSON.stringify(textParagraphLayoutCatalogExpected);
      const initialTextParagraphLayoutSlideXml = new TextDecoder().decode(
        textParagraphLayoutPart(),
      );
      const initialTextParagraphLayoutXml = namedOwnerXml(
        initialTextParagraphLayoutSlideXml,
        'browser_text_paragraph_layout',
        '</p:sp>',
      );
      const textParagraphLayoutInitialOoxml =
        /<a:bodyPr(?=[^>]*wrap="none")(?=[^>]*lIns="50800")(?=[^>]*tIns="12700")(?=[^>]*rIns="25400")(?=[^>]*bIns="38100")(?=[^>]*anchor="b")[^>]*>/u
          .test(initialTextParagraphLayoutXml)
        && /<a:pPr(?=[^>]*algn="ctr")(?=[^>]*rtl="1")[^>]*>/u
          .test(initialTextParagraphLayoutXml)
        && initialTextParagraphLayoutXml.includes(
          '<a:lnSpc><a:spcPts val="2800"/></a:lnSpc>',
        )
        && initialTextParagraphLayoutXml.includes(
          '<a:spcBef><a:spcPts val="625"/></a:spcBef>',
        )
        && initialTextParagraphLayoutXml.includes(
          '<a:spcAft><a:spcPts val="850"/></a:spcAft>',
        );
      const initialTextParagraphLayoutLeftCatalogXml = namedOwnerXml(
        initialTextParagraphLayoutSlideXml,
        'browser_text_paragraph_layout_left_scalar_middle',
        '</p:sp>',
      );
      const initialTextParagraphLayoutJustifyCatalogXml = namedOwnerXml(
        initialTextParagraphLayoutSlideXml,
        'browser_text_paragraph_layout_justify_zero_top',
        '</p:sp>',
      );
      const textParagraphLayoutCatalogOoxml =
        /<a:bodyPr(?=[^>]*wrap="square")(?=[^>]*lIns="101600")(?=[^>]*tIns="101600")(?=[^>]*rIns="101600")(?=[^>]*bIns="101600")(?=[^>]*anchor="ctr")[^>]*>/u
          .test(initialTextParagraphLayoutLeftCatalogXml)
        && /<a:pPr(?=[^>]*algn="l")(?=[^>]*rtl="0")[^>]*>/u
          .test(initialTextParagraphLayoutLeftCatalogXml)
        && /<a:bodyPr(?=[^>]*wrap="none")(?=[^>]*lIns="0")(?=[^>]*tIns="0")(?=[^>]*rIns="0")(?=[^>]*bIns="0")(?=[^>]*anchor="t")[^>]*>/u
          .test(initialTextParagraphLayoutJustifyCatalogXml)
        && /<a:pPr(?=[^>]*algn="just")(?=[^>]*rtl="1")[^>]*>/u
          .test(initialTextParagraphLayoutJustifyCatalogXml);
      const textParagraphLayoutInvalidBytes = textParagraphLayoutPart().slice();
      const textParagraphLayoutInvalidJournal = JSON.stringify(
        textParagraphLayoutDocument.opcPackage.mutations,
      );
      let textParagraphLayoutInvalidRejected = false;
      try {
        textParagraphLayoutShape.textMargins = [1, 2, 3];
      } catch {
        textParagraphLayoutInvalidRejected = true;
      }
      const textParagraphLayoutInvalidMarginRollback = textParagraphLayoutInvalidRejected
        && textParagraphLayoutBytesEqual(
          textParagraphLayoutInvalidBytes,
          textParagraphLayoutPart(),
        )
        && textParagraphLayoutInvalidJournal === JSON.stringify(
          textParagraphLayoutDocument.opcPackage.mutations,
        );
      const textParagraphLayoutInvalidSpacingBytes = textParagraphLayoutPart().slice();
      const textParagraphLayoutInvalidSpacingJournal = JSON.stringify(
        textParagraphLayoutDocument.opcPackage.mutations,
      );
      let textParagraphLayoutInvalidSpacingRejected = false;
      try {
        textParagraphLayoutShape.richText = [{
          runs: [{ text: 'Invalid browser paragraph spacing' }],
          spacing: { line: { kind: 'multiple', factor: 0 } },
        }];
      } catch {
        textParagraphLayoutInvalidSpacingRejected = true;
      }
      const textParagraphLayoutInvalidSpacingRollback =
        textParagraphLayoutInvalidSpacingRejected
        && textParagraphLayoutBytesEqual(
          textParagraphLayoutInvalidSpacingBytes,
          textParagraphLayoutPart(),
        )
        && textParagraphLayoutInvalidSpacingJournal === JSON.stringify(
          textParagraphLayoutDocument.opcPackage.mutations,
        );
      textParagraphLayoutDocument.duplicateSlide(0);
      textParagraphLayoutShape.richText = [{
        runs: [{ text: 'Edited browser paragraph layout' }],
        align: 'right',
        rtl: false,
        spacing: {
          before: 2,
          after: 3,
          line: { kind: 'multiple', factor: 1.5 },
        },
      }];
      textParagraphLayoutShape.textMargins = { top: 5, left: 6 };
      textParagraphLayoutShape.verticalAlignment = 'top';
      textParagraphLayoutShape.textWrap = true;
      const editedTextParagraphLayoutExpected = {
        paragraphs: [{
          align: 'right',
          rtl: false,
          spacing: {
            before: 2,
            after: 3,
            line: { kind: 'multiple', factor: 1.5 },
          },
        }],
        margins: { left: 6, top: 5 },
        verticalAlignment: 'top',
        textWrap: true,
      };
      const textParagraphLayoutEdited = JSON.stringify(
        textParagraphLayoutSnapshot(textParagraphLayoutShape),
      ) === JSON.stringify(editedTextParagraphLayoutExpected);
      const editedTextParagraphLayoutXml = namedOwnerXml(
        new TextDecoder().decode(textParagraphLayoutPart()),
        'browser_text_paragraph_layout',
        '</p:sp>',
      );
      const textParagraphLayoutEditedOoxml =
        /<a:bodyPr(?=[^>]*wrap="square")(?=[^>]*lIns="76200")(?=[^>]*tIns="63500")(?=[^>]*anchor="t")(?![^>]*(?:rIns|bIns)=)[^>]*>/u
          .test(editedTextParagraphLayoutXml)
        && /<a:pPr(?=[^>]*algn="r")(?=[^>]*rtl="0")[^>]*>/u
          .test(editedTextParagraphLayoutXml)
        && editedTextParagraphLayoutXml.includes(
          '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
        )
        && editedTextParagraphLayoutXml.includes(
          '<a:spcBef><a:spcPts val="200"/></a:spcBef>',
        )
        && editedTextParagraphLayoutXml.includes(
          '<a:spcAft><a:spcPts val="300"/></a:spcAft>',
        );
      const textParagraphLayoutOutput = await textParagraphLayoutDocument.writeBlob();
      globalThis.__pptxTextParagraphLayoutEvidenceBlob = textParagraphLayoutOutput;
      const reopenedTextParagraphLayoutDocument = await api.PptxDocument.open(
        textParagraphLayoutOutput,
      );
      const reopenedTextParagraphLayoutSource = reopenedTextParagraphLayoutDocument
        .slides[0].shapes.find(({ name }) => name === 'browser_text_paragraph_layout');
      const reopenedTextParagraphLayoutDuplicate = reopenedTextParagraphLayoutDocument
        .slides[1].shapes.find(({ name }) => name === 'browser_text_paragraph_layout');
      const reopenedTextParagraphLayoutCatalog = reopenedTextParagraphLayoutDocument
        .slides[0].shapes
        .filter(({ name }) => name.startsWith('browser_text_paragraph_layout_'))
        .map(textParagraphLayoutSnapshot);
      const duplicateTextParagraphLayoutCatalog = reopenedTextParagraphLayoutDocument
        .slides[1].shapes
        .filter(({ name }) => name.startsWith('browser_text_paragraph_layout_'))
        .map(textParagraphLayoutSnapshot);
      const textParagraphLayoutFamilyState = {
        immediate: textParagraphLayoutImmediate,
        initialOoxml: textParagraphLayoutInitialOoxml,
        catalog: textParagraphLayoutCatalogImmediate
          && JSON.stringify(reopenedTextParagraphLayoutCatalog) ===
            JSON.stringify(textParagraphLayoutCatalogExpected)
          && JSON.stringify(duplicateTextParagraphLayoutCatalog) ===
            JSON.stringify(textParagraphLayoutCatalogExpected),
        catalogOoxml: textParagraphLayoutCatalogOoxml,
        invalidRollback: textParagraphLayoutInvalidMarginRollback
          && textParagraphLayoutInvalidSpacingRollback,
        edited: textParagraphLayoutEdited,
        editedOoxml: textParagraphLayoutEditedOoxml,
        reopened: reopenedTextParagraphLayoutSource !== undefined
          && JSON.stringify(textParagraphLayoutSnapshot(reopenedTextParagraphLayoutSource)) ===
            JSON.stringify(editedTextParagraphLayoutExpected),
        duplicateIsolation: reopenedTextParagraphLayoutDuplicate !== undefined
          && JSON.stringify(textParagraphLayoutSnapshot(reopenedTextParagraphLayoutDuplicate)) ===
            JSON.stringify(initialTextParagraphLayoutExpected),
        mime: textParagraphLayoutOutput.type ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        diagnostics: textParagraphLayoutDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0 && reopenedTextParagraphLayoutDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0,
      };
      const textParagraphLayoutFamily = Object.values(
        textParagraphLayoutFamilyState,
      ).every(Boolean);
      const richTextEffectsDocument = api.PptxDocument.create();
      const richTextEffectsSlide = richTextEffectsDocument.addSlide();
      const richTextEffectsShape = richTextEffectsSlide.addRichText([{
        runs: [
          {
            text: 'Primary',
            style: {
              baseline: 'superscript',
              characterSpacing: 2.5,
              color: { kind: 'srgb', value: '112233' },
              glow: {
                color: { kind: 'scheme', value: 'accent1' },
                opacity: 0.5,
                size: 8,
              },
              outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
              strike: 'sngStrike',
              transparency: 25,
            },
          },
          {
            text: ' Secondary',
            style: {
              baseline: 'subscript',
              characterSpacing: -1.25,
              glow: {
                color: { kind: 'srgb', value: '00FF00' },
                opacity: 0,
                size: 0,
              },
              outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0 },
              strike: 'dblStrike',
              transparency: 50.555,
            },
          },
          {
            text: ' Explicit zero',
            style: {
              baseline: 0,
              characterSpacing: 0,
              strike: false,
              transparency: 0,
            },
          },
        ],
      }], { name: 'browser_rich_text_effects_family' });
      const richTextEffectsPart = () => richTextEffectsDocument.opcPackage
        .requirePart(richTextEffectsSlide.partUri).bytes;
      const richTextEffectsBytesEqual = (left, right) =>
        left.byteLength === right.byteLength
        && left.every((value, index) => value === right[index]);
      const richTextEffectsSnapshot = (shape) => shape.richText[0].runs.map(
        ({ style }) => ({
          baseline: style?.baseline,
          characterSpacing: style?.characterSpacing,
          glow: style?.glow,
          outline: style?.outline,
          strike: style?.strike,
          transparency: style?.transparency,
        }),
      );
      const initialRichTextEffectsExpected = [
        {
          baseline: 'superscript',
          characterSpacing: 2.5,
          glow: {
            color: { kind: 'scheme', value: 'accent1' },
            opacity: 0.5,
            size: 8,
          },
          outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
          strike: 'sngStrike',
          transparency: 25,
        },
        {
          baseline: 'subscript',
          characterSpacing: -1.25,
          glow: {
            color: { kind: 'srgb', value: '00FF00' },
            opacity: 0,
            size: 0,
          },
          outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0 },
          strike: 'dblStrike',
          transparency: 50.555,
        },
        {
          baseline: 0,
          characterSpacing: 0,
          strike: false,
          transparency: 0,
        },
      ];
      const initialRichTextEffectsSnapshot = richTextEffectsSnapshot(
        richTextEffectsShape,
      );
      const richTextEffectsImmediate = JSON.stringify(
        initialRichTextEffectsSnapshot[0],
      ) === JSON.stringify(initialRichTextEffectsExpected[0]);
      const richTextEffectsCatalog = JSON.stringify(
        initialRichTextEffectsSnapshot.slice(1),
      ) === JSON.stringify(initialRichTextEffectsExpected.slice(1));
      const initialRichTextEffectsXml = namedOwnerXml(
        new TextDecoder().decode(richTextEffectsPart()),
        'browser_rich_text_effects_family',
        '</p:sp>',
      );
      const richTextEffectsInitialOoxml =
        /<a:rPr\b(?=[^>]*baseline="30000")(?=[^>]*spc="250")(?=[^>]*strike="sngStrike")[^>]*>/u
          .test(initialRichTextEffectsXml)
        && initialRichTextEffectsXml.includes(
          '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
        )
        && initialRichTextEffectsXml.includes(
          '<a:glow rad="101600"><a:schemeClr val="accent1"><a:alpha val="50000"/>',
        )
        && initialRichTextEffectsXml.includes('<a:alpha val="75000"/>');
      const richTextEffectsCatalogOoxml =
        /<a:rPr\b(?=[^>]*baseline="-40000")(?=[^>]*spc="-125")(?=[^>]*strike="dblStrike")[^>]*>/u
          .test(initialRichTextEffectsXml)
        && /<a:rPr\b(?=[^>]*baseline="0")(?=[^>]*spc="0")(?=[^>]*strike="noStrike")[^>]*>/u
          .test(initialRichTextEffectsXml)
        && initialRichTextEffectsXml.includes(
          '<a:ln w="0"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln>',
        )
        && initialRichTextEffectsXml.includes(
          '<a:glow rad="0"><a:srgbClr val="00FF00"><a:alpha val="0"/>',
        )
        && initialRichTextEffectsXml.includes('<a:alpha val="49445"/>');
      const richTextEffectsInvalidBytes = richTextEffectsPart().slice();
      const richTextEffectsInvalidJournal = JSON.stringify(
        richTextEffectsDocument.opcPackage.mutations,
      );
      let richTextEffectsInvalidRejected = false;
      try {
        richTextEffectsShape.richText = [{
          runs: [{
            text: 'Invalid browser rich text effect',
            style: { glow: { opacity: 2, size: 1 } },
          }],
        }];
      } catch {
        richTextEffectsInvalidRejected = true;
      }
      const richTextEffectsInvalidRollback = richTextEffectsInvalidRejected
        && richTextEffectsBytesEqual(
          richTextEffectsInvalidBytes,
          richTextEffectsPart(),
        )
        && richTextEffectsInvalidJournal === JSON.stringify(
          richTextEffectsDocument.opcPackage.mutations,
        );
      richTextEffectsDocument.duplicateSlide(0);
      richTextEffectsShape.richText = [{
        runs: [
          {
            text: 'Edited browser rich text effects',
            style: {
              baseline: 12.345,
              characterSpacing: 1,
              glow: {
                color: { kind: 'srgb', value: '0000FF' },
                opacity: 0.75,
                size: 3,
              },
              outline: { color: { kind: 'scheme', value: 'tx1' }, size: 2 },
              strike: false,
              transparency: 100,
            },
          },
          { text: ' Cleared browser rich text effects' },
        ],
      }];
      const editedRichTextEffectsExpected = [
        {
          baseline: 12.345,
          characterSpacing: 1,
          glow: {
            color: { kind: 'srgb', value: '0000FF' },
            opacity: 0.75,
            size: 3,
          },
          outline: { color: { kind: 'scheme', value: 'tx1' }, size: 2 },
          strike: false,
          transparency: 100,
        },
        {},
      ];
      const richTextEffectsEdited = JSON.stringify(
        richTextEffectsSnapshot(richTextEffectsShape),
      ) === JSON.stringify(editedRichTextEffectsExpected);
      const editedRichTextEffectsXml = namedOwnerXml(
        new TextDecoder().decode(richTextEffectsPart()),
        'browser_rich_text_effects_family',
        '</p:sp>',
      );
      const clearedRichTextEffectsRunXml = editedRichTextEffectsXml.match(
        /<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t xml:space="preserve"> Cleared browser rich text effects<\/a:t><\/a:r>/u,
      )?.[0] ?? '';
      const richTextEffectsEditedOoxml =
        /<a:rPr\b(?=[^>]*baseline="12345")(?=[^>]*spc="100")(?=[^>]*strike="noStrike")[^>]*>/u
          .test(editedRichTextEffectsXml)
        && editedRichTextEffectsXml.includes(
          '<a:ln w="25400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>',
        )
        && editedRichTextEffectsXml.includes(
          '<a:glow rad="38100"><a:srgbClr val="0000FF"><a:alpha val="75000"/>',
        )
        && editedRichTextEffectsXml.includes('<a:alpha val="0"/>')
        && clearedRichTextEffectsRunXml.length > 0
        && !/\bbaseline=|\bspc=|\bstrike=|<a:ln\b|<a:effectLst\b|<a:alpha\b/u
          .test(clearedRichTextEffectsRunXml);
      const richTextEffectsOutput = await richTextEffectsDocument.writeBlob();
      globalThis.__pptxRichTextEffectsEvidenceBlob = richTextEffectsOutput;
      const reopenedRichTextEffectsDocument = await api.PptxDocument.open(
        richTextEffectsOutput,
      );
      const reopenedRichTextEffectsSource = reopenedRichTextEffectsDocument
        .slides[0].shapes.find(
          ({ name }) => name === 'browser_rich_text_effects_family',
        );
      const reopenedRichTextEffectsDuplicate = reopenedRichTextEffectsDocument
        .slides[1].shapes.find(
          ({ name }) => name === 'browser_rich_text_effects_family',
        );
      const richTextEffectsFamilyState = {
        immediate: richTextEffectsImmediate,
        initialOoxml: richTextEffectsInitialOoxml,
        catalog: richTextEffectsCatalog,
        catalogOoxml: richTextEffectsCatalogOoxml,
        invalidRollback: richTextEffectsInvalidRollback,
        edited: richTextEffectsEdited,
        editedOoxml: richTextEffectsEditedOoxml,
        reopened: reopenedRichTextEffectsSource !== undefined
          && JSON.stringify(richTextEffectsSnapshot(reopenedRichTextEffectsSource)) ===
            JSON.stringify(editedRichTextEffectsExpected),
        duplicateIsolation: reopenedRichTextEffectsDuplicate !== undefined
          && JSON.stringify(richTextEffectsSnapshot(reopenedRichTextEffectsDuplicate)) ===
            JSON.stringify(initialRichTextEffectsExpected),
        mime: richTextEffectsOutput.type ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        diagnostics: richTextEffectsDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0 && reopenedRichTextEffectsDocument.diagnostics.filter(
          ({ severity }) => severity === 'error' || severity === 'warning',
        ).length === 0,
      };
      const richTextEffectsFamily = Object.values(
        richTextEffectsFamilyState,
      ).every(Boolean);
      const tableHorizontalAlignmentDocument = api.PptxDocument.create();
      const tableHorizontalAlignmentSlide = tableHorizontalAlignmentDocument.addSlide();
      const tableHorizontalAlignmentTable = tableHorizontalAlignmentSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], { name: 'Chrome table horizontal alignment', align: 'center' });
      const tableHorizontalAlignmentPart = () => tableHorizontalAlignmentDocument.opcPackage
        .requirePart(tableHorizontalAlignmentSlide.partUri).bytes;
      const tableHorizontalAlignmentBytesEqual = (left, right) =>
        left.byteLength === right.byteLength &&
        left.every((value, index) => value === right[index]);
      const tableHorizontalAlignmentReadBytes = tableHorizontalAlignmentPart().slice();
      const tableHorizontalAlignmentReadJournal = JSON.stringify(
        tableHorizontalAlignmentDocument.opcPackage.mutations,
      );
      const tableHorizontalAlignmentUniform = tableHorizontalAlignmentTable
        .horizontalAlignment;
      const tableHorizontalAlignmentReadIsolation = tableHorizontalAlignmentBytesEqual(
        tableHorizontalAlignmentReadBytes,
        tableHorizontalAlignmentPart(),
      ) && JSON.stringify(tableHorizontalAlignmentDocument.opcPackage.mutations) ===
        tableHorizontalAlignmentReadJournal;
      const tableHorizontalAlignmentNoOpBytes = tableHorizontalAlignmentPart().slice();
      const tableHorizontalAlignmentNoOpJournal = JSON.stringify(
        tableHorizontalAlignmentDocument.opcPackage.mutations,
      );
      tableHorizontalAlignmentTable.horizontalAlignment = 'center';
      const tableHorizontalAlignmentNoOp = tableHorizontalAlignmentBytesEqual(
        tableHorizontalAlignmentNoOpBytes,
        tableHorizontalAlignmentPart(),
      ) && JSON.stringify(tableHorizontalAlignmentDocument.opcPackage.mutations) ===
        tableHorizontalAlignmentNoOpJournal;
      tableHorizontalAlignmentTable.setCellHorizontalAlignment(0, 1, 'right');
      const tableHorizontalAlignmentMixed = tableHorizontalAlignmentTable
        .horizontalAlignment ?? null;
      tableHorizontalAlignmentTable.horizontalAlignment = 'justify';
      const tableHorizontalAlignmentOverwritten = tableHorizontalAlignmentTable
        .horizontalAlignment;
      const tableHorizontalAlignmentOverwrittenCells = tableHorizontalAlignmentTable.rows
        .flatMap(({ cells }) => cells.map(({ horizontalAlignment }) =>
          horizontalAlignment ?? null));
      tableHorizontalAlignmentTable.horizontalAlignment = 'left';
      const tableHorizontalAlignmentExplicitLeft = tableHorizontalAlignmentTable
        .horizontalAlignment;
      const tableHorizontalAlignmentExplicitLeftCells = tableHorizontalAlignmentTable.rows
        .flatMap(({ cells }) => cells.map(({ horizontalAlignment }) =>
          horizontalAlignment ?? null));
      tableHorizontalAlignmentTable.horizontalAlignment = undefined;
      const tableHorizontalAlignmentCleared = tableHorizontalAlignmentTable
        .horizontalAlignment ?? null;
      const tableHorizontalAlignmentClearedCells = tableHorizontalAlignmentTable.rows
        .flatMap(({ cells }) => cells.map(({ horizontalAlignment }) =>
          horizontalAlignment ?? null));
      const tableHorizontalAlignmentInvalidBytes = tableHorizontalAlignmentPart().slice();
      const tableHorizontalAlignmentInvalidJournal = JSON.stringify(
        tableHorizontalAlignmentDocument.opcPackage.mutations,
      );
      let tableHorizontalAlignmentInvalidError;
      try {
        tableHorizontalAlignmentTable.horizontalAlignment = 'dist';
      } catch (error) {
        tableHorizontalAlignmentInvalidError = {
          name: error.name,
          message: error.message,
        };
      }
      const tableHorizontalAlignmentFailureIsolation = tableHorizontalAlignmentBytesEqual(
        tableHorizontalAlignmentInvalidBytes,
        tableHorizontalAlignmentPart(),
      ) && JSON.stringify(tableHorizontalAlignmentDocument.opcPackage.mutations) ===
        tableHorizontalAlignmentInvalidJournal;
      tableHorizontalAlignmentTable.horizontalAlignment = 'right';
      const reopenedTableHorizontalAlignmentDocument = await api.PptxDocument.open(
        await tableHorizontalAlignmentDocument.writeBlob(),
      );
      const reopenedTableHorizontalAlignmentTable = reopenedTableHorizontalAlignmentDocument
        .slides[0].shapes.find(
          (shape) => shape.name === 'Chrome table horizontal alignment',
        );
      const tableHorizontalAlignmentState = {
        uniform: tableHorizontalAlignmentUniform,
        readIsolation: tableHorizontalAlignmentReadIsolation,
        noOp: tableHorizontalAlignmentNoOp,
        mixed: tableHorizontalAlignmentMixed,
        overwritten: tableHorizontalAlignmentOverwritten,
        overwrittenCells: tableHorizontalAlignmentOverwrittenCells,
        explicitLeft: tableHorizontalAlignmentExplicitLeft,
        explicitLeftCells: tableHorizontalAlignmentExplicitLeftCells,
        cleared: tableHorizontalAlignmentCleared,
        clearedCells: tableHorizontalAlignmentClearedCells,
        reopened: reopenedTableHorizontalAlignmentTable instanceof api.TableModel
          ? reopenedTableHorizontalAlignmentTable.horizontalAlignment ?? null
          : null,
        reopenedCells: reopenedTableHorizontalAlignmentTable instanceof api.TableModel
          ? reopenedTableHorizontalAlignmentTable.rows.flatMap(({ cells }) => cells.map(
            ({ horizontalAlignment }) => horizontalAlignment ?? null,
          ))
          : [],
        invalidError: tableHorizontalAlignmentInvalidError,
        failureIsolation: tableHorizontalAlignmentFailureIsolation,
        validationErrors: tableHorizontalAlignmentDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableHorizontalAlignmentDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableHorizontalAlignment = JSON.stringify(tableHorizontalAlignmentState) ===
        JSON.stringify({
          uniform: 'center',
          readIsolation: true,
          noOp: true,
          mixed: null,
          overwritten: 'justify',
          overwrittenCells: ['justify', 'justify', 'justify', 'justify'],
          explicitLeft: 'left',
          explicitLeftCells: ['left', 'left', 'left', 'left'],
          cleared: null,
          clearedCells: [null, null, null, null],
          reopened: 'right',
          reopenedCells: ['right', 'right', 'right', 'right'],
          invalidError: {
            name: 'TypeError',
            message: 'Table horizontal alignment must be left, center, right, or justify',
          },
          failureIsolation: true,
          validationErrors: 0,
        });
      const tableMarginsSnapshot = (value) => value === undefined
        ? null
        : {
            top: value.top,
            right: value.right,
            bottom: value.bottom,
            left: value.left,
          };
      const tableMarginsDocument = api.PptxDocument.create();
      const tableMarginsSlide = tableMarginsDocument.addSlide();
      const tableMarginsTable = tableMarginsSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], {
        name: 'Chrome table margins',
        margin: [3.6, 7.2, 10.8, 14.4],
      });
      const tableMarginsPart = () => tableMarginsDocument.opcPackage
        .requirePart(tableMarginsSlide.partUri).bytes;
      const tableMarginsBytesEqual = (left, right) =>
        left.byteLength === right.byteLength &&
        left.every((value, index) => value === right[index]);
      const tableMarginsReadBytes = tableMarginsPart().slice();
      const tableMarginsReadJournal = JSON.stringify(
        tableMarginsDocument.opcPackage.mutations,
      );
      const tableMarginsUniform = tableMarginsSnapshot(tableMarginsTable.margins);
      const tableMarginsDetached = tableMarginsTable.margins;
      if (tableMarginsDetached) tableMarginsDetached.top = 99;
      const tableMarginsReadIsolation = tableMarginsBytesEqual(
        tableMarginsReadBytes,
        tableMarginsPart(),
      ) && JSON.stringify(tableMarginsDocument.opcPackage.mutations) ===
        tableMarginsReadJournal &&
        JSON.stringify(tableMarginsSnapshot(tableMarginsTable.margins)) ===
          JSON.stringify({ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 });
      const tableMarginsNoOpBytes = tableMarginsPart().slice();
      const tableMarginsNoOpJournal = JSON.stringify(
        tableMarginsDocument.opcPackage.mutations,
      );
      tableMarginsTable.margins = [3.6, 7.2, 10.8, 14.4];
      const tableMarginsNoOp = tableMarginsBytesEqual(
        tableMarginsNoOpBytes,
        tableMarginsPart(),
      ) && JSON.stringify(tableMarginsDocument.opcPackage.mutations) ===
        tableMarginsNoOpJournal;
      tableMarginsTable.setCellMargins(0, 1, { top: 9 });
      const tableMarginsMixed = tableMarginsSnapshot(tableMarginsTable.margins);
      tableMarginsTable.margins = 6;
      const tableMarginsOverwritten = tableMarginsSnapshot(tableMarginsTable.margins);
      const tableMarginsOverwrittenCells = tableMarginsTable.rows
        .flatMap(({ cells }) => cells.map(
          ({ margins }) => tableMarginsSnapshot(margins),
        ));
      tableMarginsTable.margins = { top: 2, left: 4 };
      const tableMarginsPartial = tableMarginsSnapshot(tableMarginsTable.margins);
      const tableMarginsPartialCells = tableMarginsTable.rows
        .flatMap(({ cells }) => cells.map(
          ({ margins }) => tableMarginsSnapshot(margins),
        ));
      tableMarginsTable.margins = {};
      const tableMarginsCleared = tableMarginsSnapshot(tableMarginsTable.margins);
      const tableMarginsClearedCells = tableMarginsTable.rows
        .flatMap(({ cells }) => cells.map(
          ({ margins }) => tableMarginsSnapshot(margins),
        ));
      const tableMarginsInvalidBytes = tableMarginsPart().slice();
      const tableMarginsInvalidJournal = JSON.stringify(
        tableMarginsDocument.opcPackage.mutations,
      );
      let tableMarginsInvalidError;
      try {
        tableMarginsTable.margins = null;
      } catch (error) {
        tableMarginsInvalidError = { name: error.name, message: error.message };
      }
      const tableMarginsFailureIsolation = tableMarginsBytesEqual(
        tableMarginsInvalidBytes,
        tableMarginsPart(),
      ) && JSON.stringify(tableMarginsDocument.opcPackage.mutations) ===
        tableMarginsInvalidJournal;
      tableMarginsTable.margins = [1, 2, 3, 4];
      const reopenedTableMarginsDocument = await api.PptxDocument.open(
        await tableMarginsDocument.writeBlob(),
      );
      const reopenedTableMarginsTable = reopenedTableMarginsDocument.slides[0].shapes.find(
        (shape) => shape.name === 'Chrome table margins',
      );
      const tableMarginsState = {
        uniform: tableMarginsUniform,
        readIsolation: tableMarginsReadIsolation,
        noOp: tableMarginsNoOp,
        mixed: tableMarginsMixed,
        overwritten: tableMarginsOverwritten,
        overwrittenCells: tableMarginsOverwrittenCells,
        partial: tableMarginsPartial,
        partialCells: tableMarginsPartialCells,
        cleared: tableMarginsCleared,
        clearedCells: tableMarginsClearedCells,
        reopened: reopenedTableMarginsTable instanceof api.TableModel
          ? tableMarginsSnapshot(reopenedTableMarginsTable.margins)
          : null,
        reopenedCells: reopenedTableMarginsTable instanceof api.TableModel
          ? reopenedTableMarginsTable.rows.flatMap(({ cells }) => cells.map(
            ({ margins }) => tableMarginsSnapshot(margins),
          ))
          : [],
        invalidError: tableMarginsInvalidError,
        failureIsolation: tableMarginsFailureIsolation,
        validationErrors: tableMarginsDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableMarginsDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableMargins = JSON.stringify(tableMarginsState) === JSON.stringify({
        uniform: { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
        readIsolation: true,
        noOp: true,
        mixed: null,
        overwritten: { top: 6, right: 6, bottom: 6, left: 6 },
        overwrittenCells: Array(4).fill({ top: 6, right: 6, bottom: 6, left: 6 }),
        partial: { top: 2, left: 4 },
        partialCells: Array(4).fill({ top: 2, left: 4 }),
        cleared: null,
        clearedCells: [null, null, null, null],
        reopened: { top: 1, right: 2, bottom: 3, left: 4 },
        reopenedCells: Array(4).fill({ top: 1, right: 2, bottom: 3, left: 4 }),
        invalidError: {
          name: 'TypeError',
          message: 'Table margins must be a number, four-value tuple, or margin object',
        },
        failureIsolation: true,
        validationErrors: 0,
      });
      const tableFillSnapshot = (value) => {
        if (value === undefined) return null;
        if (value.kind === 'none') return { kind: 'none' };
        return {
          kind: 'solid',
          color: { kind: value.color.kind, value: value.color.value },
          ...(value.transparency !== undefined
            ? { transparency: value.transparency }
            : {}),
        };
      };
      const tableFillDocument = api.PptxDocument.create();
      const tableFillSlide = tableFillDocument.addSlide();
      const tableFillTable = tableFillSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], {
        name: 'Chrome table fill',
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
      });
      const tableFillPart = () => tableFillDocument.opcPackage
        .requirePart(tableFillSlide.partUri).bytes;
      const tableFillReadBytes = tableFillPart().slice();
      const tableFillReadJournal = JSON.stringify(
        tableFillDocument.opcPackage.mutations,
      );
      const tableFillUniform = tableFillSnapshot(tableFillTable.fill);
      const tableFillDetached = tableFillTable.fill;
      if (tableFillDetached?.kind === 'solid') {
        tableFillDetached.color.value = 'accent6';
      }
      const tableFillReadIsolation = tableMarginsBytesEqual(
        tableFillReadBytes,
        tableFillPart(),
      ) && JSON.stringify(tableFillDocument.opcPackage.mutations) ===
        tableFillReadJournal &&
        JSON.stringify(tableFillSnapshot(tableFillTable.fill)) === JSON.stringify({
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        });
      const tableFillNoOpBytes = tableFillPart().slice();
      const tableFillNoOpJournal = JSON.stringify(
        tableFillDocument.opcPackage.mutations,
      );
      tableFillTable.fill = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      };
      const tableFillNoOp = tableMarginsBytesEqual(
        tableFillNoOpBytes,
        tableFillPart(),
      ) && JSON.stringify(tableFillDocument.opcPackage.mutations) ===
        tableFillNoOpJournal;
      tableFillTable.setCellFill(0, 1, { kind: 'none' });
      const tableFillMixed = tableFillSnapshot(tableFillTable.fill);
      tableFillTable.fill = { kind: 'none' };
      const tableFillNone = tableFillSnapshot(tableFillTable.fill);
      const tableFillNoneCells = tableFillTable.rows.flatMap(({ cells }) =>
        cells.map(({ fill }) => tableFillSnapshot(fill)));
      tableFillTable.fill = {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      };
      const tableFillSolid = tableFillSnapshot(tableFillTable.fill);
      const tableFillSolidCells = tableFillTable.rows.flatMap(({ cells }) =>
        cells.map(({ fill }) => tableFillSnapshot(fill)));
      tableFillTable.fill = undefined;
      const tableFillCleared = tableFillSnapshot(tableFillTable.fill);
      const tableFillClearedCells = tableFillTable.rows.flatMap(({ cells }) =>
        cells.map(({ fill }) => tableFillSnapshot(fill)));
      const tableFillInvalidBytes = tableFillPart().slice();
      const tableFillInvalidJournal = JSON.stringify(
        tableFillDocument.opcPackage.mutations,
      );
      let tableFillInvalidError;
      try {
        tableFillTable.fill = null;
      } catch (error) {
        tableFillInvalidError = { name: error.name, message: error.message };
      }
      const tableFillFailureIsolation = tableMarginsBytesEqual(
        tableFillInvalidBytes,
        tableFillPart(),
      ) && JSON.stringify(tableFillDocument.opcPackage.mutations) ===
        tableFillInvalidJournal;
      tableFillTable.fill = { kind: 'none' };
      const reopenedTableFillDocument = await api.PptxDocument.open(
        await tableFillDocument.writeBlob(),
      );
      const reopenedTableFillTable = reopenedTableFillDocument.slides[0].shapes.find(
        (shape) => shape.name === 'Chrome table fill',
      );
      const tableFillState = {
        uniform: tableFillUniform,
        readIsolation: tableFillReadIsolation,
        noOp: tableFillNoOp,
        mixed: tableFillMixed,
        none: tableFillNone,
        noneCells: tableFillNoneCells,
        solid: tableFillSolid,
        solidCells: tableFillSolidCells,
        cleared: tableFillCleared,
        clearedCells: tableFillClearedCells,
        reopened: reopenedTableFillTable instanceof api.TableModel
          ? tableFillSnapshot(reopenedTableFillTable.fill)
          : null,
        reopenedCells: reopenedTableFillTable instanceof api.TableModel
          ? reopenedTableFillTable.rows.flatMap(({ cells }) =>
            cells.map(({ fill }) => tableFillSnapshot(fill)))
          : [],
        invalidError: tableFillInvalidError,
        failureIsolation: tableFillFailureIsolation,
        validationErrors: tableFillDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableFillDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableFill = JSON.stringify(tableFillState) === JSON.stringify({
        uniform: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        readIsolation: true,
        noOp: true,
        mixed: null,
        none: { kind: 'none' },
        noneCells: Array(4).fill({ kind: 'none' }),
        solid: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'D9EAF7' },
          transparency: 0,
        },
        solidCells: Array(4).fill({
          kind: 'solid',
          color: { kind: 'srgb', value: 'D9EAF7' },
          transparency: 0,
        }),
        cleared: null,
        clearedCells: [null, null, null, null],
        reopened: { kind: 'none' },
        reopenedCells: Array(4).fill({ kind: 'none' }),
        invalidError: {
          name: 'TypeError',
          message: 'Table fill must be an object',
        },
        failureIsolation: true,
        validationErrors: 0,
      });
      const tableTextDefaultsDocument = api.PptxDocument.create();
      const tableTextDefaultsSlide = tableTextDefaultsDocument.addSlide();
      const tableTextDefaultsTable = tableTextDefaultsSlide.addTable([[
        'Table defaults',
        {
          text: 'Cell replaced',
          options: {
            fontFamily: 'Courier New',
            fontSize: 10,
            bold: false,
            color: { kind: 'srgb', value: '00AA00' },
            spacing: { before: 3 },
          },
        },
        {
          text: 'Cell retained',
          options: {
            fontFamily: 'Courier New',
            fontSize: 10,
            bold: false,
            color: { kind: 'srgb', value: '00AA00' },
            spacing: { before: 3 },
          },
        },
      ]], {
        name: 'Chrome table text defaults',
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'scheme', value: 'accent1' },
        spacing: {
          before: 6,
          after: 8,
          line: { kind: 'multiple', factor: 1.5 },
        },
      });
      const tableTextDefaultsInitialTableParagraph =
        tableTextDefaultsTable.rows[0].cells[0].richText[0];
      const tableTextDefaultsInitialCellParagraph =
        tableTextDefaultsTable.rows[0].cells[1].richText[0];
      const tableTextDefaultsCreated =
        tableTextDefaultsInitialTableParagraph.spacing?.before === 6
        && tableTextDefaultsInitialTableParagraph.spacing?.after === 8
        && tableTextDefaultsInitialTableParagraph.spacing?.line?.kind === 'multiple'
        && tableTextDefaultsInitialTableParagraph.spacing.line.factor === 1.5
        && tableTextDefaultsInitialTableParagraph.runs[0].style?.fontFamily === 'Aptos'
        && tableTextDefaultsInitialTableParagraph.runs[0].style?.fontSize === 18.25
        && tableTextDefaultsInitialTableParagraph.runs[0].style?.bold === true
        && tableTextDefaultsInitialTableParagraph.runs[0].style?.color?.kind === 'scheme'
        && tableTextDefaultsInitialTableParagraph.runs[0].style.color.value === 'accent1'
        && tableTextDefaultsInitialCellParagraph.spacing?.before === 3
        && tableTextDefaultsInitialCellParagraph.spacing?.after === 8
        && tableTextDefaultsInitialCellParagraph.spacing?.line?.kind === 'multiple'
        && tableTextDefaultsInitialCellParagraph.spacing.line.factor === 1.5
        && tableTextDefaultsInitialCellParagraph.runs[0].style?.fontFamily ===
          'Courier New'
        && tableTextDefaultsInitialCellParagraph.runs[0].style?.fontSize === 10
        && tableTextDefaultsInitialCellParagraph.runs[0].style?.bold === false
        && tableTextDefaultsInitialCellParagraph.runs[0].style?.color?.kind === 'srgb'
        && tableTextDefaultsInitialCellParagraph.runs[0].style.color.value === '00AA00';
      tableTextDefaultsTable.setCellText(0, 0, 'Table edited');
      const tableTextDefaultsPlainEditParagraph =
        tableTextDefaultsTable.rows[0].cells[0].richText[0];
      const tableTextDefaultsPlainEdit =
        tableTextDefaultsPlainEditParagraph.runs[0].text === 'Table edited'
        && tableTextDefaultsPlainEditParagraph.runs[0].style?.fontFamily === 'Aptos'
        && tableTextDefaultsPlainEditParagraph.runs[0].style?.fontSize === 18.25
        && tableTextDefaultsPlainEditParagraph.runs[0].style?.bold === true
        && tableTextDefaultsPlainEditParagraph.spacing?.before === 6;
      tableTextDefaultsTable.setCellRichText(
        0,
        1,
        [{ runs: [{ text: 'Replacement' }] }],
      );
      const tableTextDefaultsReplacement =
        tableTextDefaultsTable.rows[0].cells[1].richText[0];
      const tableTextDefaultsRichReplacement =
        tableTextDefaultsReplacement.spacing === undefined
        && tableTextDefaultsReplacement.runs[0].text === 'Replacement'
        && tableTextDefaultsReplacement.runs[0].style?.fontFamily === '+mn-lt'
        && tableTextDefaultsReplacement.runs[0].style?.fontSize === undefined
        && tableTextDefaultsReplacement.runs[0].style?.bold === undefined
        && tableTextDefaultsReplacement.runs[0].style?.color?.kind === 'scheme'
        && tableTextDefaultsReplacement.runs[0].style.color.value === 'tx1';
      const reopenedTableTextDefaultsDocument = await api.PptxDocument.open(
        await tableTextDefaultsDocument.writeBlob(),
      );
      const reopenedTableTextDefaultsTable = reopenedTableTextDefaultsDocument
        .slides[0].shapes.find((shape) => shape.name === 'Chrome table text defaults');
      const reopenedTableTextDefaultsRetained =
        reopenedTableTextDefaultsTable instanceof api.TableModel
          ? reopenedTableTextDefaultsTable.rows[0].cells[2].richText[0]
          : undefined;
      const tableTextDefaultsReopened =
        reopenedTableTextDefaultsTable instanceof api.TableModel
        && reopenedTableTextDefaultsTable.rows[0].cells[0].text === 'Table edited'
        && reopenedTableTextDefaultsTable.rows[0].cells[1].text === 'Replacement'
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0].spacing === undefined
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0]
          .runs[0].style?.fontFamily === '+mn-lt'
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0]
          .runs[0].style?.fontSize === undefined
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0]
          .runs[0].style?.bold === undefined
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0]
          .runs[0].style?.color?.kind === 'scheme'
        && reopenedTableTextDefaultsTable.rows[0].cells[1].richText[0]
          .runs[0].style.color.value === 'tx1'
        && reopenedTableTextDefaultsRetained?.spacing?.before === 3
        && reopenedTableTextDefaultsRetained.spacing.after === 8
        && reopenedTableTextDefaultsRetained.runs[0].style?.fontFamily === 'Courier New'
        && reopenedTableTextDefaultsRetained.runs[0].style?.fontSize === 10
        && reopenedTableTextDefaultsRetained.runs[0].style?.bold === false
        && reopenedTableTextDefaultsRetained.runs[0].style?.color?.kind === 'srgb'
        && reopenedTableTextDefaultsRetained.runs[0].style.color.value === '00AA00';
      const tableTextDefaultsState = {
        created: tableTextDefaultsCreated,
        plainEdit: tableTextDefaultsPlainEdit,
        richReplacement: tableTextDefaultsRichReplacement,
        reopened: tableTextDefaultsReopened,
        validationErrors: tableTextDefaultsDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length
          + reopenedTableTextDefaultsDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableTextDefaults = Object.values(tableTextDefaultsState).every(
        (value) => value === true || value === 0,
      );
      const tableCellMergeSnapshot = (value) => value === undefined
        ? null
        : {
            rowIndex: value.rowIndex,
            columnIndex: value.columnIndex,
            rowspan: value.rowspan,
            colspan: value.colspan,
            ...(value.isAnchor !== undefined ? { isAnchor: value.isAnchor } : {}),
          };
      const tableCellMergesDocument = api.PptxDocument.create();
      const tableCellMergesSlide = tableCellMergesDocument.addSlide();
      const tableCellMergesTable = tableCellMergesSlide.addTable([
        [{
          text: 'Chrome merge anchor',
          options: { colspan: 2, rowspan: 2 },
        }, 'Chrome total'],
        ['Chrome tail'],
      ], { name: 'Chrome table cell merges' });
      const tableCellMergeRegion = {
        rowIndex: 0,
        columnIndex: 0,
        rowspan: 2,
        colspan: 2,
      };
      const tableCellMergeMembers = [
        { ...tableCellMergeRegion, isAnchor: true },
        { ...tableCellMergeRegion, isAnchor: false },
        null,
        { ...tableCellMergeRegion, isAnchor: false },
        { ...tableCellMergeRegion, isAnchor: false },
        null,
      ];
      const tableCellMergesCreated =
        tableCellMergesTable.rows.length === 2
        && tableCellMergesTable.rows.every(({ cells }) => cells.length === 3)
        && JSON.stringify(tableCellMergesTable.mergeRegions) ===
          JSON.stringify([tableCellMergeRegion]);
      const tableCellMergesRead = JSON.stringify(
        tableCellMergesTable.rows.flatMap(({ cells }) =>
          cells.map(({ merge }) => tableCellMergeSnapshot(merge))),
      ) === JSON.stringify(tableCellMergeMembers);
      const tableCellMergesSnapshotsFrozen =
        Object.isFrozen(tableCellMergesTable.mergeRegions)
        && Object.isFrozen(tableCellMergesTable.mergeRegions?.[0])
        && tableCellMergesTable.rows.flatMap(({ cells }) => cells)
          .filter(({ merge }) => merge !== undefined)
          .every(({ merge }) => Object.isFrozen(merge));
      tableCellMergesTable.unmergeCell(1, 1);
      const tableCellMergesUnmerged =
        JSON.stringify(tableCellMergesTable.mergeRegions) === '[]'
        && tableCellMergesTable.rows.flatMap(({ cells }) => cells)
          .every(({ merge }) => merge === undefined);
      tableCellMergesTable.setCellFill(1, 1, {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FCE4D6' },
      });
      const tableCellMergesEdited =
        tableCellMergesTable.rows[1].cells[1].fill?.kind === 'solid'
        && tableCellMergesTable.rows[1].cells[1].fill.color.kind === 'srgb'
        && tableCellMergesTable.rows[1].cells[1].fill.color.value === 'FCE4D6';
      tableCellMergesTable.mergeCells(0, 0, 2, 2);
      const tableCellMergesRemerged =
        JSON.stringify(tableCellMergesTable.mergeRegions) ===
          JSON.stringify([tableCellMergeRegion])
        && JSON.stringify(tableCellMergesTable.rows.flatMap(({ cells }) =>
          cells.map(({ merge }) => tableCellMergeSnapshot(merge)))) ===
          JSON.stringify(tableCellMergeMembers)
        && tableCellMergesTable.rows[1].cells[1].fill?.kind === 'solid'
        && tableCellMergesTable.rows[1].cells[1].fill.color.value === 'FCE4D6';
      const tableCellMergesEvidenceBlob = await tableCellMergesDocument.writeBlob();
      globalThis.__pptxTableCellMergesEvidenceBlob = tableCellMergesEvidenceBlob;
      const reopenedTableCellMergesDocument = await api.PptxDocument.open(
        tableCellMergesEvidenceBlob,
      );
      const reopenedTableCellMergesTable = reopenedTableCellMergesDocument
        .slides[0].shapes.find((shape) => shape.name === 'Chrome table cell merges');
      const tableCellMergesReopened = reopenedTableCellMergesTable instanceof api.TableModel
        && JSON.stringify(reopenedTableCellMergesTable.mergeRegions) ===
          JSON.stringify([tableCellMergeRegion])
        && JSON.stringify(reopenedTableCellMergesTable.rows.flatMap(({ cells }) =>
          cells.map(({ merge }) => tableCellMergeSnapshot(merge)))) ===
          JSON.stringify(tableCellMergeMembers)
        && reopenedTableCellMergesTable.rows[1].cells[1].fill?.kind === 'solid'
        && reopenedTableCellMergesTable.rows[1].cells[1].fill.color.value === 'FCE4D6';
      const tableCellMergesState = {
        created: tableCellMergesCreated,
        read: tableCellMergesRead,
        snapshotsFrozen: tableCellMergesSnapshotsFrozen,
        unmerged: tableCellMergesUnmerged,
        edited: tableCellMergesEdited,
        remerged: tableCellMergesRemerged,
        reopened: tableCellMergesReopened,
        validationErrors: tableCellMergesDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length
          + reopenedTableCellMergesDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableCellMerges = Object.values(tableCellMergesState).every(
        (value) => value === true || value === 0,
      );
      const tableStructureEditingDocument = api.PptxDocument.create();
      const tableStructureEditingSlide = tableStructureEditingDocument.addSlide();
      const tableStructureEditingSurvivorUrl =
        'https://table-structure.example?a=1&b=2';
      const tableStructureEditingDeletedRowUrl =
        'https://table-structure-row-deleted.example';
      const tableStructureEditingDeletedColumnUrl =
        'https://table-structure-column-deleted.example';
      const tableStructureEditingTable = tableStructureEditingSlide.addTable([
        [{
          text: 'Chrome structure anchor',
          options: {
            colspan: 2,
            rowspan: 2,
            fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
          },
        }, 'Chrome R0C2', 'Chrome R0C3'],
        ['Chrome R1C2', 'Chrome R1C3'],
        [
          'Chrome R2C0',
          'Chrome R2C1',
          {
            text: [{ runs: [
              {
                text: 'Styled ',
                style: {
                  bold: true,
                  color: { kind: 'srgb', value: '1F4E78' },
                },
              },
              { text: 'linked survivor', style: { italic: true } },
            ] }],
            options: {
              fill: { kind: 'solid', color: { kind: 'srgb', value: 'F4B183' } },
              hyperlink: {
                url: tableStructureEditingSurvivorUrl,
                tooltip: 'Chrome survivor',
              },
            },
          },
          {
            text: 'Deleted column link',
            options: { hyperlink: { url: tableStructureEditingDeletedColumnUrl } },
          },
        ],
        [
          'Chrome R3C0',
          'Chrome R3C1',
          {
            text: 'Deleted row link',
            options: { hyperlink: { url: tableStructureEditingDeletedRowUrl } },
          },
          'Chrome R3C3',
        ],
      ], {
        name: 'Chrome table structure editing',
        x: api.inches(1),
        y: api.inches(0.5),
        columnWidths: [api.inches(1), api.inches(2), api.inches(3), api.inches(4)],
        rowHeights: [
          api.inches(0.5),
          api.inches(1),
          api.inches(1.5),
          api.inches(2),
        ],
      });
      const tableStructureEditingHyperlinks = (slide) =>
        slide.relationships.filter(({ type }) => type.endsWith('/hyperlink'));
      const tableStructureEditingXml = (document, slide) => new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      const tableStructureEditingClickIds = (document, slide) => [
        ...tableStructureEditingXml(document, slide).matchAll(
          /<a:hlinkClick\b[^>]*\br:id="([^"]+)"/g,
        ),
      ].map((match) => match[1]);
      const tableStructureEditingInitialHyperlinks = tableStructureEditingHyperlinks(
        tableStructureEditingSlide,
      );
      const tableStructureEditingSurvivorRelationship =
        tableStructureEditingInitialHyperlinks.find(
          ({ target }) => target === tableStructureEditingSurvivorUrl,
        );
      const tableStructureEditingDeletedRowRelationship =
        tableStructureEditingInitialHyperlinks.find(
          ({ target }) => target === tableStructureEditingDeletedRowUrl,
        );
      const tableStructureEditingDeletedColumnRelationship =
        tableStructureEditingInitialHyperlinks.find(
          ({ target }) => target === tableStructureEditingDeletedColumnUrl,
        );
      const tableStructureEditingCreated =
        tableStructureEditingTable.rows.length === 4
        && tableStructureEditingTable.rows.every(({ cells }) => cells.length === 4)
        && JSON.stringify(tableStructureEditingTable.mergeRegions) === JSON.stringify([{
          rowIndex: 0,
          columnIndex: 0,
          rowspan: 2,
          colspan: 2,
        }])
        && tableStructureEditingInitialHyperlinks.length === 3
        && tableStructureEditingSurvivorRelationship !== undefined
        && tableStructureEditingDeletedRowRelationship !== undefined
        && tableStructureEditingDeletedColumnRelationship !== undefined
        && tableStructureEditingClickIds(
          tableStructureEditingDocument,
          tableStructureEditingSlide,
        ).filter((id) => id === tableStructureEditingSurvivorRelationship.id).length === 2;
      tableStructureEditingTable.insertRows(1, { rowHeights: api.inches(0.25) });
      const tableStructureEditingRowsInserted =
        tableStructureEditingTable.rows.length === 5
        && tableStructureEditingTable.rows.every(({ cells }) => cells.length === 4)
        && JSON.stringify(tableStructureEditingTable.rowHeights) === JSON.stringify([
          api.inches(0.5),
          api.inches(0.25),
          api.inches(1),
          api.inches(1.5),
          api.inches(2),
        ])
        && JSON.stringify(tableStructureEditingTable.mergeRegions) === JSON.stringify([{
          rowIndex: 0,
          columnIndex: 0,
          rowspan: 3,
          colspan: 2,
        }]);
      tableStructureEditingTable.insertColumns(1, { columnWidths: api.inches(0.5) });
      const tableStructureEditingColumnsInserted =
        tableStructureEditingTable.rows.length === 5
        && tableStructureEditingTable.rows.every(({ cells }) => cells.length === 5)
        && JSON.stringify(tableStructureEditingTable.columnWidths) === JSON.stringify([
          api.inches(1),
          api.inches(0.5),
          api.inches(2),
          api.inches(3),
          api.inches(4),
        ])
        && JSON.stringify(tableStructureEditingTable.mergeRegions) === JSON.stringify([{
          rowIndex: 0,
          columnIndex: 0,
          rowspan: 3,
          colspan: 3,
        }]);
      tableStructureEditingTable.setCellRichText(1, 1, [{ runs: [{
        text: 'Inserted hidden cell',
        style: {
          bold: true,
          color: { kind: 'srgb', value: '385723' },
        },
      }] }]);
      tableStructureEditingTable.setCellFill(1, 1, {
        kind: 'solid',
        color: { kind: 'srgb', value: 'E2F0D9' },
      });
      const tableStructureEditingHiddenCell = (table) => {
        const cell = table.rows[1]?.cells[1];
        const run = cell?.richText[0]?.runs[0];
        return cell?.text === 'Inserted hidden cell'
          && cell.merge?.isAnchor === false
          && cell.fill?.kind === 'solid'
          && cell.fill.color.kind === 'srgb'
          && cell.fill.color.value === 'E2F0D9'
          && run?.style?.bold === true
          && run.style.color?.kind === 'srgb'
          && run.style.color.value === '385723';
      };
      const tableStructureEditingNewCellEdited = tableStructureEditingHiddenCell(
        tableStructureEditingTable,
      );
      tableStructureEditingTable.deleteRows(4);
      const tableStructureEditingRowsDeleted =
        tableStructureEditingTable.rows.length === 4
        && tableStructureEditingTable.rows.every(({ cells }) => cells.length === 5)
        && !tableStructureEditingSlide.relationships.some(
          ({ id }) => id === tableStructureEditingDeletedRowRelationship?.id,
        )
        && tableStructureEditingSlide.relationships.some(
          ({ id }) => id === tableStructureEditingDeletedColumnRelationship?.id,
        )
        && tableStructureEditingSlide.relationships.some(
          ({ id }) => id === tableStructureEditingSurvivorRelationship?.id,
        );
      tableStructureEditingTable.deleteColumns(4);
      const tableStructureEditingColumnsDeleted =
        tableStructureEditingTable.rows.length === 4
        && tableStructureEditingTable.rows.every(({ cells }) => cells.length === 4)
        && !tableStructureEditingSlide.relationships.some(
          ({ id }) => id === tableStructureEditingDeletedColumnRelationship?.id,
        )
        && tableStructureEditingSlide.relationships.some(
          ({ id }) => id === tableStructureEditingSurvivorRelationship?.id,
        );
      const tableStructureEditingDimensionsMatch = (table) =>
        JSON.stringify(table.columnWidths) === JSON.stringify([
          api.inches(1), api.inches(0.5), api.inches(2), api.inches(3),
        ])
        && JSON.stringify(table.rowHeights) === JSON.stringify([
          api.inches(0.5), api.inches(0.25), api.inches(1), api.inches(1.5),
        ])
        && table.transform.width === api.inches(6.5)
        && table.transform.height === api.inches(3.25);
      const tableStructureEditingMergeMatch = (table) =>
        JSON.stringify(table.mergeRegions) === JSON.stringify([{
          rowIndex: 0,
          columnIndex: 0,
          rowspan: 3,
          colspan: 3,
        }]);
      const tableStructureEditingSurvivorMatch = (table) => {
        const survivor = table.rows[3]?.cells[3];
        const styles = survivor?.richText[0]?.runs.map(({ style }) => style);
        return survivor?.text === 'Styled linked survivor'
          && survivor.fill?.kind === 'solid'
          && survivor.fill.color.kind === 'srgb'
          && survivor.fill.color.value === 'F4B183'
          && styles?.[0]?.bold === true
          && styles[0].color?.kind === 'srgb'
          && styles[0].color.value === '1F4E78'
          && styles[0].hyperlink?.url === tableStructureEditingSurvivorUrl
          && styles[0].hyperlink.tooltip === 'Chrome survivor'
          && styles?.[1]?.italic === true
          && styles[1].hyperlink?.url === tableStructureEditingSurvivorUrl
          && styles[1].hyperlink.tooltip === 'Chrome survivor';
      };
      const tableStructureEditingDimensions = tableStructureEditingDimensionsMatch(
        tableStructureEditingTable,
      );
      const tableStructureEditingMerge = tableStructureEditingMergeMatch(
        tableStructureEditingTable,
      );
      const tableStructureEditingSurvivor = tableStructureEditingSurvivorMatch(
        tableStructureEditingTable,
      );
      const tableStructureEditingFinalHyperlinks = tableStructureEditingHyperlinks(
        tableStructureEditingSlide,
      );
      const tableStructureEditingFinalClickIds = tableStructureEditingClickIds(
        tableStructureEditingDocument,
        tableStructureEditingSlide,
      );
      const tableStructureEditingFinalRelationshipIds = new Set(
        tableStructureEditingSlide.relationships.map(({ id }) => id),
      );
      const tableStructureEditingRelationships =
        tableStructureEditingFinalHyperlinks.length === 1
        && tableStructureEditingFinalHyperlinks[0]?.id ===
          tableStructureEditingSurvivorRelationship?.id
        && tableStructureEditingFinalHyperlinks[0]?.target ===
          tableStructureEditingSurvivorUrl
        && tableStructureEditingFinalClickIds.length === 2
        && new Set(tableStructureEditingFinalClickIds).size === 1
        && tableStructureEditingFinalClickIds.every((id) =>
          id === tableStructureEditingSurvivorRelationship?.id
          && tableStructureEditingFinalRelationshipIds.has(id));
      const tableStructureEditingEvidenceBlob =
        await tableStructureEditingDocument.writeBlob();
      globalThis.__pptxTableStructureEditingEvidenceBlob =
        tableStructureEditingEvidenceBlob;
      const reopenedTableStructureEditingDocument = await api.PptxDocument.open(
        tableStructureEditingEvidenceBlob,
      );
      const reopenedTableStructureEditingSlide =
        reopenedTableStructureEditingDocument.slides[0];
      const reopenedTableStructureEditingTable =
        reopenedTableStructureEditingSlide.shapes.find(
          (shape) => shape.name === 'Chrome table structure editing',
        );
      const reopenedTableStructureEditingHyperlinks = tableStructureEditingHyperlinks(
        reopenedTableStructureEditingSlide,
      );
      const reopenedTableStructureEditingSurvivorRelationship =
        reopenedTableStructureEditingHyperlinks.find(
          ({ target }) => target === tableStructureEditingSurvivorUrl,
        );
      const reopenedTableStructureEditingClickIds = tableStructureEditingClickIds(
        reopenedTableStructureEditingDocument,
        reopenedTableStructureEditingSlide,
      );
      const tableStructureEditingReopened =
        reopenedTableStructureEditingTable instanceof api.TableModel
        && reopenedTableStructureEditingTable.rows.length === 4
        && reopenedTableStructureEditingTable.rows.every(
          ({ cells }) => cells.length === 4,
        )
        && tableStructureEditingDimensionsMatch(reopenedTableStructureEditingTable)
        && tableStructureEditingMergeMatch(reopenedTableStructureEditingTable)
        && tableStructureEditingHiddenCell(reopenedTableStructureEditingTable)
        && tableStructureEditingSurvivorMatch(reopenedTableStructureEditingTable)
        && reopenedTableStructureEditingHyperlinks.length === 1
        && reopenedTableStructureEditingClickIds.length === 2
        && reopenedTableStructureEditingClickIds.every(
          (id) => id === reopenedTableStructureEditingSurvivorRelationship?.id,
        );
      const tableStructureEditingState = {
        created: tableStructureEditingCreated,
        rowsInserted: tableStructureEditingRowsInserted,
        columnsInserted: tableStructureEditingColumnsInserted,
        newCellEdited: tableStructureEditingNewCellEdited,
        rowsDeleted: tableStructureEditingRowsDeleted,
        columnsDeleted: tableStructureEditingColumnsDeleted,
        dimensions: tableStructureEditingDimensions,
        merge: tableStructureEditingMerge,
        survivor: tableStructureEditingSurvivor,
        relationships: tableStructureEditingRelationships,
        reopened: tableStructureEditingReopened,
        validationErrors: tableStructureEditingDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length
          + reopenedTableStructureEditingDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableStructureEditing = Object.values(tableStructureEditingState).every(
        (value) => value === true || value === 0,
      );
      const tableAutoPageDocument = api.PptxDocument.create();
      const tableAutoPageLead = tableAutoPageDocument.addSlide();
      const tableAutoPageSource = tableAutoPageDocument.addSlide();
      const tableAutoPageSentinel = tableAutoPageDocument.addSlide();
      const tableAutoPageTarget = tableAutoPageDocument.addSlide();
      const tableAutoPageSection = tableAutoPageDocument.addSection({
        title: 'Chrome auto page',
      });
      tableAutoPageDocument.assignSlideToSection(
        tableAutoPageDocument.slides.indexOf(tableAutoPageSource),
        tableAutoPageSection.id,
      );
      const tableAutoPageTargetInput = tableAutoPageDocument.slides.indexOf(
        tableAutoPageTarget,
      ) + 1;
      const tableAutoPageHeaderOuterUrl =
        'https://chrome-auto-page.example/header-outer';
      const tableAutoPageHeaderRunUrl =
        'https://chrome-auto-page.example/header-run';
      const tableAutoPageBodyUrl = 'https://chrome-auto-page.example/body';
      const tableAutoPageEditedUrl = 'https://chrome-auto-page.example/edited';
      const tableAutoPageSourceTable = tableAutoPageSource.addTable([
        [
          {
            text: [{ runs: [
              {
                text: 'Chrome auto-page header ',
                style: {
                  bold: true,
                  color: { kind: 'scheme', value: 'accent1' },
                },
              },
              {
                text: 'external',
                style: {
                  hyperlink: {
                    url: tableAutoPageHeaderRunUrl,
                    tooltip: 'Chrome header run',
                  },
                  underline: true,
                },
              },
              {
                text: ' target',
                style: {
                  hyperlink: {
                    slide: tableAutoPageTargetInput,
                    tooltip: 'Chrome target',
                  },
                  italic: true,
                },
              },
            ] }],
            options: {
              rowspan: 2,
              colspan: 2,
              fill: {
                kind: 'solid',
                color: { kind: 'srgb', value: 'D9EAF7' },
              },
              hyperlink: {
                url: tableAutoPageHeaderOuterUrl,
                tooltip: 'Chrome header outer',
              },
            },
          },
          'Chrome header right',
        ],
        ['Chrome header lower right'],
        [
          {
            text: 'Chrome body A',
            options: {
              rowspan: 2,
              colspan: 2,
              fill: {
                kind: 'solid',
                color: { kind: 'srgb', value: 'FFF2CC' },
              },
              hyperlink: { url: tableAutoPageBodyUrl, tooltip: 'Chrome body' },
            },
          },
          'Chrome A right',
        ],
        ['Chrome A lower right'],
        [
          {
            text: 'Chrome body B',
            options: {
              rowspan: 2,
              bold: true,
              hyperlink: {
                slide: tableAutoPageTargetInput,
                tooltip: 'Chrome body target',
              },
            },
          },
          'Chrome B middle',
          'Chrome B right',
        ],
        ['Chrome B lower middle', 'Chrome B lower right'],
        ['Chrome tail', 'Chrome tail middle', 'Chrome tail right'],
      ], {
        name: 'Chrome table auto page',
        autoPage: true,
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 2,
        autoPageSlideStartY: api.inches(3.125),
        slideMargin: api.inches(0.5),
        x: api.inches(1),
        y: api.inches(3.125),
        columnWidths: [api.inches(1), api.inches(1), api.inches(1)],
        rowHeights: Array.from({ length: 7 }, () => api.inches(0.5)),
      });
      const tableAutoPageGenerated = tableAutoPageSource.newAutoPagedSlides;
      const tableAutoPageInitialSlides = [
        tableAutoPageLead,
        tableAutoPageSource,
        ...tableAutoPageGenerated,
        tableAutoPageSentinel,
        tableAutoPageTarget,
      ];
      const tableAutoPageTableFor = (slide) => slide.shapes.find(
        (shape) => shape instanceof api.TableModel &&
          shape.name === 'Chrome table auto page',
      );
      const tableAutoPageTables = [
        tableAutoPageSourceTable,
        ...tableAutoPageGenerated.map(tableAutoPageTableFor),
      ];
      const tableAutoPageXml = (document, slide) => new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      const tableAutoPageClickIds = (document, slide) => [
        ...tableAutoPageXml(document, slide).matchAll(
          /<a:hlinkClick\b[^>]*\br:id="([^"]+)"/g,
        ),
      ].map((match) => match[1]);
      const tableAutoPageOwnedLinks = (slide) => slide.relationships.filter(
        ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
      );
      const tableAutoPageLinksOwned = (document, slide) => {
        const clickIds = tableAutoPageClickIds(document, slide);
        const owned = tableAutoPageOwnedLinks(slide);
        const ownedIds = new Set(owned.map(({ id }) => id));
        return clickIds.length === owned.length
          && new Set(clickIds).size === owned.length
          && clickIds.every((id) => ownedIds.has(id));
      };
      const tableAutoPageLayoutTarget = (slide) => slide.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget;
      const tableAutoPageCreated =
        tableAutoPageGenerated.length === 2
        && Object.isFrozen(tableAutoPageGenerated)
        && JSON.stringify(tableAutoPageDocument.slides.map(({ partUri }) => partUri)) ===
          JSON.stringify(tableAutoPageInitialSlides.map(({ partUri }) => partUri))
        && tableAutoPageTables.every((table) => table instanceof api.TableModel)
        && JSON.stringify(tableAutoPageTables.map((table) => table.rows.map(
          (row) => row.cells[0]?.text,
        ))) === JSON.stringify([
          ['Chrome auto-page header external target', '', 'Chrome body A', ''],
          ['Chrome auto-page header external target', '', 'Chrome body B', ''],
          ['Chrome auto-page header external target', '', 'Chrome tail'],
        ])
        && JSON.stringify(tableAutoPageTables.map(({ rowHeights }) => rowHeights)) ===
          JSON.stringify([
            Array(4).fill(api.inches(0.5)),
            Array(4).fill(api.inches(0.5)),
            Array(3).fill(api.inches(0.5)),
          ])
        && JSON.stringify(tableAutoPageTables.map(({ mergeRegions }) => mergeRegions)) ===
          JSON.stringify([
            [
              { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
              { rowIndex: 2, columnIndex: 0, rowspan: 2, colspan: 2 },
            ],
            [
              { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
              { rowIndex: 2, columnIndex: 0, rowspan: 2, colspan: 1 },
            ],
            [{ rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 }],
          ])
        && new Set([
          tableAutoPageSource,
          ...tableAutoPageGenerated,
        ].map(tableAutoPageLayoutTarget)).size === 1
        && tableAutoPageDocument.sections?.find(
          ({ id }) => id === tableAutoPageSection.id,
        )?.slideIds.join(',') === [
          tableAutoPageSource,
          ...tableAutoPageGenerated,
        ].map(({ slideId }) => slideId).join(',')
        && [tableAutoPageSource, ...tableAutoPageGenerated].every(
          (slide) => tableAutoPageLinksOwned(tableAutoPageDocument, slide),
        )
        && [tableAutoPageSource, ...tableAutoPageGenerated].every(
          (slide) => tableAutoPageOwnedLinks(slide).some(
            ({ resolvedTarget }) => resolvedTarget === tableAutoPageTarget.partUri,
          ),
        );
      const tableAutoPageEditedTable = tableAutoPageTables[1];
      tableAutoPageEditedTable.setCellRichText(3, 2, [{ runs: [{
        text: 'Chrome auto-page edited',
        style: {
          bold: true,
          color: { kind: 'srgb', value: 'C00000' },
          hyperlink: { url: tableAutoPageEditedUrl, tooltip: 'Chrome edited' },
        },
      }] }]);
      const tableAutoPageEdited =
        tableAutoPageEditedTable.rows[3]?.cells[2]?.text === 'Chrome auto-page edited'
        && tableAutoPageEditedTable.rows[3]?.cells[2]?.richText[0]
          ?.runs[0]?.style?.bold === true
        && tableAutoPageEditedTable.rows[3]?.cells[2]?.richText[0]
          ?.runs[0]?.style?.color?.value === 'C00000'
        && tableAutoPageEditedTable.rows[3]?.cells[2]?.richText[0]
          ?.runs[0]?.style?.hyperlink?.url === tableAutoPageEditedUrl
        && tableAutoPageLinksOwned(
          tableAutoPageDocument,
          tableAutoPageGenerated[0],
        );
      tableAutoPageDocument.moveSlide(
        tableAutoPageDocument.slides.indexOf(tableAutoPageGenerated[1]),
        tableAutoPageDocument.slides.length - 1,
      );
      const tableAutoPageMovedAway =
        tableAutoPageSource.newAutoPagedSlides[0] === tableAutoPageGenerated[0]
        && tableAutoPageSource.newAutoPagedSlides[1] === tableAutoPageGenerated[1]
        && tableAutoPageOwnedLinks(tableAutoPageGenerated[1]).some(
          ({ resolvedTarget }) => resolvedTarget === tableAutoPageTarget.partUri,
        );
      tableAutoPageDocument.moveSlide(
        tableAutoPageDocument.slides.indexOf(tableAutoPageGenerated[1]),
        tableAutoPageDocument.slides.indexOf(tableAutoPageSentinel),
      );
      const tableAutoPageMoved = tableAutoPageMovedAway
        && JSON.stringify(tableAutoPageDocument.slides.map(({ partUri }) => partUri)) ===
          JSON.stringify(tableAutoPageInitialSlides.map(({ partUri }) => partUri));
      tableAutoPageDocument.deleteSlide(
        tableAutoPageDocument.slides.indexOf(tableAutoPageGenerated[1]),
      );
      const tableAutoPageFinalSlides = [
        tableAutoPageLead,
        tableAutoPageSource,
        tableAutoPageGenerated[0],
        tableAutoPageSentinel,
        tableAutoPageTarget,
      ];
      const tableAutoPageDeleted =
        tableAutoPageSource.newAutoPagedSlides.length === 1
        && tableAutoPageSource.newAutoPagedSlides[0] === tableAutoPageGenerated[0]
        && JSON.stringify(tableAutoPageDocument.slides.map(({ partUri }) => partUri)) ===
          JSON.stringify(tableAutoPageFinalSlides.map(({ partUri }) => partUri))
        && tableAutoPageDocument.sections?.find(
          ({ id }) => id === tableAutoPageSection.id,
        )?.slideIds.join(',') === [
          tableAutoPageSource.slideId,
          tableAutoPageGenerated[0].slideId,
        ].join(',');
      const tableAutoPageRelationships =
        [tableAutoPageSource, tableAutoPageGenerated[0]].every(
          (slide) => tableAutoPageLinksOwned(tableAutoPageDocument, slide),
        )
        && [tableAutoPageSource, tableAutoPageGenerated[0]].every(
          (slide) => tableAutoPageOwnedLinks(slide).some(
            ({ resolvedTarget }) => resolvedTarget === tableAutoPageTarget.partUri,
          ),
        )
        && tableAutoPageOwnedLinks(tableAutoPageSource).filter(
          ({ targetMode }) => targetMode === 'External',
        ).length === 3
        && tableAutoPageOwnedLinks(tableAutoPageGenerated[0]).filter(
          ({ targetMode }) => targetMode === 'External',
        ).length === 3;
      const tableAutoPageEvidenceBlob = await tableAutoPageDocument.writeBlob();
      globalThis.__pptxTableAutoPageEvidenceBlob = tableAutoPageEvidenceBlob;
      const reopenedTableAutoPageDocument = await api.PptxDocument.open(
        tableAutoPageEvidenceBlob,
      );
      const reopenedTableAutoPageSource = reopenedTableAutoPageDocument.slides.find(
        ({ partUri }) => partUri === tableAutoPageSource.partUri,
      );
      const reopenedTableAutoPageGenerated = reopenedTableAutoPageDocument.slides.find(
        ({ partUri }) => partUri === tableAutoPageGenerated[0].partUri,
      );
      const reopenedTableAutoPageTarget = reopenedTableAutoPageDocument.slides.find(
        ({ partUri }) => partUri === tableAutoPageTarget.partUri,
      );
      const reopenedTableAutoPageSourceTable = reopenedTableAutoPageSource === undefined
        ? undefined
        : tableAutoPageTableFor(reopenedTableAutoPageSource);
      const reopenedTableAutoPageGeneratedTable =
        reopenedTableAutoPageGenerated === undefined
          ? undefined
          : tableAutoPageTableFor(reopenedTableAutoPageGenerated);
      const reopenedTableAutoPageHeaderRuns = reopenedTableAutoPageGeneratedTable
        ?.rows[0]?.cells[0]?.richText[0]?.runs;
      const tableAutoPageReopened =
        reopenedTableAutoPageSource !== undefined
        && reopenedTableAutoPageGenerated !== undefined
        && reopenedTableAutoPageTarget !== undefined
        && reopenedTableAutoPageSourceTable instanceof api.TableModel
        && reopenedTableAutoPageGeneratedTable instanceof api.TableModel
        && reopenedTableAutoPageDocument.slides.every(
          (slide) => slide.newAutoPagedSlides.length === 0,
        )
        && JSON.stringify(reopenedTableAutoPageDocument.slides.map(
          ({ partUri }) => partUri,
        )) === JSON.stringify(tableAutoPageFinalSlides.map(({ partUri }) => partUri))
        && JSON.stringify(reopenedTableAutoPageSourceTable.rowHeights) ===
          JSON.stringify(Array(4).fill(api.inches(0.5)))
        && JSON.stringify(reopenedTableAutoPageGeneratedTable.rowHeights) ===
          JSON.stringify(Array(4).fill(api.inches(0.5)))
        && reopenedTableAutoPageGeneratedTable.rows[3]?.cells[2]?.text ===
          'Chrome auto-page edited'
        && reopenedTableAutoPageHeaderRuns?.[0]?.style?.hyperlink?.url ===
          tableAutoPageHeaderOuterUrl
        && reopenedTableAutoPageHeaderRuns?.[1]?.style?.hyperlink?.url ===
          tableAutoPageHeaderRunUrl
        && reopenedTableAutoPageHeaderRuns?.[2]?.style?.hyperlink?.slide ===
          reopenedTableAutoPageDocument.slides.indexOf(reopenedTableAutoPageTarget) + 1
        && tableAutoPageLinksOwned(
          reopenedTableAutoPageDocument,
          reopenedTableAutoPageSource,
        )
        && tableAutoPageLinksOwned(
          reopenedTableAutoPageDocument,
          reopenedTableAutoPageGenerated,
        );
      const tableAutoPageState = {
        created: tableAutoPageCreated,
        edited: tableAutoPageEdited,
        moved: tableAutoPageMoved,
        deleted: tableAutoPageDeleted,
        relationships: tableAutoPageRelationships,
        reopened: tableAutoPageReopened,
        validationErrors: tableAutoPageDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length
          + reopenedTableAutoPageDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableAutoPage = Object.values(tableAutoPageState).every(
        (value) => value === true || value === 0,
      );
      const tableContentMeasurementDocument = api.PptxDocument.create({
        slideSize: 'wide',
      });
      const tableContentMeasurementOwner = {
        x: api.inches(1),
        y: api.inches(1.25),
        width: api.inches(2),
        height: api.inches(2.5),
      };
      const tableContentMeasurementContinuationY = api.inches(1.5);
      const tableContentMeasurementLayout = await tableContentMeasurementDocument
        .defineSlideMaster({
          title: 'BROWSER-TABLE-CONTENT-MEASUREMENT',
          objects: [{
            kind: 'placeholder',
            text: 'Browser measured table prompt',
            options: {
              name: 'browser_measurement_table',
              type: 'tbl',
              index: 206,
              ...tableContentMeasurementOwner,
            },
          }],
        });
      const tableContentMeasurementLead = tableContentMeasurementDocument.addSlide();
      const tableContentMeasurementSource = tableContentMeasurementDocument.addSlide({
        masterName: tableContentMeasurementLayout.name,
      });
      const tableContentMeasurementSentinel = tableContentMeasurementDocument.addSlide();
      const tableContentMeasurementTarget = tableContentMeasurementDocument.addSlide();
      const tableContentMeasurementSection = tableContentMeasurementDocument.addSection({
        title: 'Browser measured pages',
      });
      tableContentMeasurementDocument.assignSlideToSection(
        tableContentMeasurementDocument.slides.indexOf(tableContentMeasurementSource),
        tableContentMeasurementSection.id,
      );
      const tableContentMeasurementTargetInput =
        tableContentMeasurementDocument.slides.indexOf(
          tableContentMeasurementTarget,
        ) + 1;
      const tableContentMeasurementHeaderUrl =
        'https://browser-measurement.example/header';
      const tableContentMeasurementOuterUrl =
        'https://browser-measurement.example/outer';
      const tableContentMeasurementRunUrl =
        'https://browser-measurement.example/run';
      const tableContentMeasurementEditedUrl =
        'https://browser-measurement.example/edited';
      const tableContentMeasurementFixedHeights = [
        api.inches(0.3),
        api.inches(0.3),
        api.inches(0.3),
      ];
      const tableContentMeasurementFixedTable = tableContentMeasurementLead.addTable([
        [
          {
            text: 'Browser fixed merge ' + 'F'.repeat(160),
            options: { rowspan: 2, colspan: 2, margin: 0 },
          },
          'Browser fixed right',
        ],
        ['Browser fixed lower right'],
        ['Browser fixed tail 1', 'Browser fixed tail 2', 'Browser fixed tail 3'],
      ], {
        name: 'Browser fixed measurement regression',
        autoPage: true,
        x: api.inches(0.5),
        y: api.inches(0.5),
        columnWidths: [api.inches(1), api.inches(1), api.inches(1)],
        rowHeights: tableContentMeasurementFixedHeights,
      });
      const tableContentMeasurementAutomaticTable =
        tableContentMeasurementLead.addTable([
          [{
            text: 'Browser automatic omitted row height ' + 'A'.repeat(36),
            options: { margin: 0 },
          }],
          [{ text: '瀏覽器自動行高漢字測量', options: { margin: 0 } }],
        ], {
          name: 'Browser omitted row-height measurement',
          autoPage: true,
          x: api.inches(0.5),
          y: api.inches(2),
          columnWidths: [api.inches(2)],
          margin: 0,
        });
      const tableContentMeasurementRichA = 'A'.repeat(180);
      const tableContentMeasurementRichB = 'B'.repeat(120);
      const tableContentMeasurementTailRows = Array.from(
        { length: 20 },
        (_, index) => [{
          text: 'Browser delete-only tail ' + String(index + 1),
          options: { colspan: 2, margin: 0 },
        }],
      );
      const tableContentMeasurementSourceTable = tableContentMeasurementSource.addTable([
        [{
          text: 'Browser measured header',
          options: {
            bold: true,
            margin: 0,
            hyperlink: {
              url: tableContentMeasurementHeaderUrl,
              tooltip: 'Browser repeated measured header',
            },
          },
        }, 'Browser measured header right'],
        [{
          text: [{ runs: [
            { text: 'Browser measured outer ', style: { bold: true } },
            {
              text: tableContentMeasurementRichA,
              style: {
                bold: true,
                color: { kind: 'srgb', value: 'C00000' },
                hyperlink: {
                  url: tableContentMeasurementRunUrl,
                  tooltip: 'Browser measured run URL',
                },
              },
            },
            {
              text: tableContentMeasurementRichB,
              softBreakBefore: true,
              style: {
                italic: true,
                color: { kind: 'scheme', value: 'accent2' },
                hyperlink: {
                  slide: tableContentMeasurementTargetInput,
                  tooltip: 'Browser measured target slide',
                },
              },
            },
          ] }],
          options: {
            colspan: 2,
            margin: 0,
            autoPageCharWeight: -1,
            autoPageLineWeight: 1,
            hyperlink: {
              url: tableContentMeasurementOuterUrl,
              tooltip: 'Browser measured default link',
            },
          },
        }],
        [{
          text: '瀏覽器漢字自動測量 CJK minimum',
          options: {
            colspan: 2,
            margin: 0,
            autoPageCharWeight: 1,
            autoPageLineWeight: -1,
          },
        }],
        [
          {
            text: 'Browser span',
            options: { rowspan: 2, margin: 0 },
          },
          'R1',
        ],
        ['R2'],
        ...tableContentMeasurementTailRows,
      ], {
        autoPage: true,
        autoPageCharWeight: 0,
        autoPageLineWeight: 0,
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageSlideStartY: tableContentMeasurementContinuationY,
        slideMargin: 0,
        placeholder: 'browser_measurement_table',
        columnWidths: [api.inches(1), api.inches(1)],
        rowHeights: [
          0,
          0,
          api.inches(0.35),
          0,
          0,
          ...tableContentMeasurementTailRows.map(() => 0),
        ],
      });
      const tableContentMeasurementGenerated =
        tableContentMeasurementSource.newAutoPagedSlides;
      const tableContentMeasurementInitialSlides = [
        tableContentMeasurementLead,
        tableContentMeasurementSource,
        ...tableContentMeasurementGenerated,
        tableContentMeasurementSentinel,
        tableContentMeasurementTarget,
      ];
      const tableContentMeasurementTableFor = (slide) => slide.shapes.find(
        (shape) => shape instanceof api.TableModel &&
          shape.name === 'browser_measurement_table',
      );
      const tableContentMeasurementPageTables = [
        tableContentMeasurementSourceTable,
        ...tableContentMeasurementGenerated.map(tableContentMeasurementTableFor),
      ];
      const tableContentMeasurementXml = (document, slide) =>
        new TextDecoder().decode(
          document.opcPackage.requirePart(slide.partUri).bytes,
        );
      const tableContentMeasurementClickIds = (document, slide) => new Set([
        ...tableContentMeasurementXml(document, slide).matchAll(
          /<a:hlinkClick\b[^>]*\br:id="([^"]+)"/g,
        ),
      ].map((match) => match[1]));
      const tableContentMeasurementOwnedLinks = (slide) =>
        slide.relationships.filter(({ type }) =>
          type.endsWith('/hyperlink') || type.endsWith('/slide'));
      const tableContentMeasurementLinksOwned = (document, slide) => {
        const clickIds = tableContentMeasurementClickIds(document, slide);
        const ownedIds = new Set(
          tableContentMeasurementOwnedLinks(slide).map(({ id }) => id),
        );
        return clickIds.size === ownedIds.size &&
          [...clickIds].every((id) => ownedIds.has(id));
      };
      const tableContentMeasurementRowsPositive = (table) =>
        table instanceof api.TableModel &&
        table.rowHeights?.every((height) => height > 0) === true &&
        table.transform.height === table.rowHeights.reduce(
          (sum, height) => sum + height,
          0,
        );
      const tableContentMeasurementCreated =
        tableContentMeasurementGenerated.length >= 4 &&
        Object.isFrozen(tableContentMeasurementGenerated) &&
        JSON.stringify(tableContentMeasurementDocument.slides.map(
          ({ partUri }) => partUri,
        )) === JSON.stringify(tableContentMeasurementInitialSlides.map(
          ({ partUri }) => partUri,
        )) &&
        tableContentMeasurementPageTables.every(
          tableContentMeasurementRowsPositive,
        );
      const tableContentMeasurementAutomatic =
        tableContentMeasurementRowsPositive(tableContentMeasurementAutomaticTable) &&
        tableContentMeasurementAutomaticTable.rowHeights.length === 2 &&
        tableContentMeasurementPageTables.flatMap(({ rowHeights }) => rowHeights)
          .every((height) => height > 0);
      const tableContentMeasurementMinimumRow = tableContentMeasurementPageTables
        .flatMap((table) => table.rows.map((row, rowIndex) => ({
          text: row.cells[0]?.text,
          height: table.rowHeights?.[rowIndex],
        })))
        .find(({ text }) => text === '瀏覽器漢字自動測量 CJK minimum');
      const tableContentMeasurementMinimum =
        tableContentMeasurementMinimumRow?.height >= api.inches(0.35);
      const tableContentMeasurementDeleteSlide =
        tableContentMeasurementGenerated.at(-1);
      const tableContentMeasurementDeleteTable =
        tableContentMeasurementDeleteSlide === undefined
          ? undefined
          : tableContentMeasurementTableFor(tableContentMeasurementDeleteSlide);
      const tableContentMeasurementDeleteOnly =
        tableContentMeasurementDeleteTable instanceof api.TableModel &&
        tableContentMeasurementDeleteTable.rows.length > 1 &&
        tableContentMeasurementDeleteTable.rows.slice(1).every(
          (row) => row.cells[0]?.text.startsWith(
            'Browser delete-only tail ',
          ) === true,
        );
      const tableContentMeasurementEditableSlide =
        tableContentMeasurementGenerated.find((slide) =>
          slide !== tableContentMeasurementDeleteSlide &&
          tableContentMeasurementTableFor(slide)?.rows.some(
            (row) => row.cells[0]?.text.startsWith(
              'Browser delete-only tail ',
            ) === true,
          ));
      const tableContentMeasurementEditableTable =
        tableContentMeasurementEditableSlide === undefined
          ? undefined
          : tableContentMeasurementTableFor(tableContentMeasurementEditableSlide);
      const tableContentMeasurementEditableRow =
        tableContentMeasurementEditableTable?.rows.findIndex(
          (row) => row.cells[0]?.text.startsWith(
            'Browser delete-only tail ',
          ) === true,
        ) ?? -1;
      if (tableContentMeasurementEditableTable instanceof api.TableModel &&
          tableContentMeasurementEditableRow >= 0) {
        tableContentMeasurementEditableTable.setCellRichText(
          tableContentMeasurementEditableRow,
          0,
          [{ runs: [{
            text: 'Browser measured edited tail',
            style: {
              bold: true,
              hyperlink: {
                url: tableContentMeasurementEditedUrl,
                tooltip: 'Browser measured edited tail',
              },
            },
          }] }],
        );
      }
      const tableContentMeasurementEdited =
        tableContentMeasurementEditableTable instanceof api.TableModel &&
        tableContentMeasurementEditableRow >= 0 &&
        tableContentMeasurementEditableTable.rows[
          tableContentMeasurementEditableRow
        ]?.cells[0]?.text === 'Browser measured edited tail' &&
        tableContentMeasurementOwnedLinks(tableContentMeasurementEditableSlide).some(
          ({ target }) => target === tableContentMeasurementEditedUrl,
        );
      const tableContentMeasurementDeleteOriginalIndex =
        tableContentMeasurementDeleteSlide === undefined
          ? -1
          : tableContentMeasurementDocument.slides.indexOf(
            tableContentMeasurementDeleteSlide,
          );
      if (tableContentMeasurementDeleteOriginalIndex >= 0) {
        tableContentMeasurementDocument.moveSlide(
          tableContentMeasurementDeleteOriginalIndex,
          tableContentMeasurementDocument.slides.length - 1,
        );
      }
      const tableContentMeasurementMovedAway =
        tableContentMeasurementDocument.slides.at(-1) ===
          tableContentMeasurementDeleteSlide;
      if (tableContentMeasurementDeleteSlide !== undefined) {
        tableContentMeasurementDocument.moveSlide(
          tableContentMeasurementDocument.slides.indexOf(
            tableContentMeasurementDeleteSlide,
          ),
          tableContentMeasurementDocument.slides.indexOf(
            tableContentMeasurementSentinel,
          ),
        );
      }
      const tableContentMeasurementMoved = tableContentMeasurementMovedAway &&
        tableContentMeasurementDocument.slides.indexOf(
          tableContentMeasurementDeleteSlide,
        ) === tableContentMeasurementDocument.slides.indexOf(
          tableContentMeasurementSentinel,
        ) - 1;
      if (tableContentMeasurementDeleteOnly &&
          tableContentMeasurementDeleteSlide !== undefined) {
        tableContentMeasurementDocument.deleteSlide(
          tableContentMeasurementDocument.slides.indexOf(
            tableContentMeasurementDeleteSlide,
          ),
        );
      }
      const tableContentMeasurementFinalPageSlides = [
        tableContentMeasurementSource,
        ...tableContentMeasurementSource.newAutoPagedSlides,
      ];
      const tableContentMeasurementFinalTables =
        tableContentMeasurementFinalPageSlides.map(
          tableContentMeasurementTableFor,
        );
      const tableContentMeasurementFragmentCells = tableContentMeasurementFinalTables
        .flatMap((table) => table.rows.flatMap(({ cells }) => cells))
        .filter((cell) => cell.hyperlink?.url === tableContentMeasurementOuterUrl ||
          cell.richText.some(({ runs }) => runs.some(({ style }) =>
            style?.hyperlink?.url === tableContentMeasurementRunUrl ||
            style?.hyperlink?.slide !== undefined)));
      const tableContentMeasurementFragmentText =
        tableContentMeasurementFragmentCells.map(({ text }) => text).join('')
          .split(String.fromCharCode(10)).join('');
      const tableContentMeasurementFragmentRuns = tableContentMeasurementFragmentCells
        .flatMap(({ richText }) => richText.flatMap(({ runs }) => runs));
      const tableContentMeasurementFragment =
        tableContentMeasurementFragmentCells.length >= 3 &&
        tableContentMeasurementFragmentText ===
          'Browser measured outer ' + tableContentMeasurementRichA +
            tableContentMeasurementRichB &&
        tableContentMeasurementFragmentRuns.some(({ style }) =>
          style?.bold === true &&
            style.hyperlink?.url === tableContentMeasurementRunUrl) &&
        tableContentMeasurementFragmentRuns.some(({ softBreakBefore, style }) =>
          softBreakBefore === true && style?.italic === true &&
            style.hyperlink?.slide !== undefined) &&
        tableContentMeasurementFinalTables.every(
          (table) => table.rows[0]?.cells[0]?.text ===
            'Browser measured header',
        ) &&
        tableContentMeasurementFinalTables.some(
          (table) => table.mergeRegions?.some(({ rowspan, colspan }) =>
            rowspan === 2 && colspan === 1) === true &&
            table.rows.some((row) => row.cells[0]?.text ===
              'Browser span'),
        );
      const tableContentMeasurementOwnerBottom =
        tableContentMeasurementOwner.y + tableContentMeasurementOwner.height;
      const tableContentMeasurementPlaceholder =
        tableContentMeasurementFinalTables.every((table, index) =>
          table.name === 'browser_measurement_table' &&
          table.placeholder?.type === 'tbl' && table.placeholder.index === 206 &&
          table.transform.x === tableContentMeasurementOwner.x &&
          table.transform.y === (index === 0
            ? tableContentMeasurementOwner.y
            : tableContentMeasurementContinuationY) &&
          table.transform.width === tableContentMeasurementOwner.width &&
          table.transform.y + table.transform.height <=
            tableContentMeasurementOwnerBottom &&
          tableContentMeasurementRowsPositive(table));
      const tableContentMeasurementFixed =
        JSON.stringify(tableContentMeasurementFixedTable.rowHeights) ===
          JSON.stringify(tableContentMeasurementFixedHeights) &&
        tableContentMeasurementFixedTable.transform.height ===
          tableContentMeasurementFixedHeights.reduce(
            (sum, height) => sum + height,
            0,
          ) &&
        JSON.stringify(tableContentMeasurementFixedTable.mergeRegions) ===
          JSON.stringify([
            { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
          ]);
      const tableContentMeasurementRelationships =
        tableContentMeasurementFinalPageSlides.every((slide) =>
          tableContentMeasurementLinksOwned(
            tableContentMeasurementDocument,
            slide,
          )) &&
        tableContentMeasurementFinalPageSlides.every((slide) =>
          tableContentMeasurementOwnedLinks(slide).some(
            ({ target }) => target === tableContentMeasurementHeaderUrl,
          )) &&
        tableContentMeasurementFinalPageSlides.some((slide) =>
          tableContentMeasurementOwnedLinks(slide).some(
            ({ target }) => target === tableContentMeasurementOuterUrl,
          )) &&
        tableContentMeasurementFinalPageSlides.some((slide) =>
          tableContentMeasurementOwnedLinks(slide).some(
            ({ target }) => target === tableContentMeasurementRunUrl,
          )) &&
        tableContentMeasurementFinalPageSlides.some((slide) =>
          tableContentMeasurementOwnedLinks(slide).some(
            ({ resolvedTarget }) => resolvedTarget ===
              tableContentMeasurementTarget.partUri,
          ));
      const tableContentMeasurementDeleted = tableContentMeasurementDeleteOnly &&
        tableContentMeasurementSource.newAutoPagedSlides.length ===
          tableContentMeasurementGenerated.length - 1 &&
        !tableContentMeasurementDocument.slides.includes(
          tableContentMeasurementDeleteSlide,
        ) &&
        tableContentMeasurementDocument.sections?.find(
          ({ id }) => id === tableContentMeasurementSection.id,
        )?.slideIds.join(',') === tableContentMeasurementFinalPageSlides
          .map(({ slideId }) => slideId).join(',');
      const tableContentMeasurementEvidenceBlob =
        await tableContentMeasurementDocument.writeBlob();
      globalThis.__pptxTableContentMeasurementEvidenceBlob =
        tableContentMeasurementEvidenceBlob;
      const reopenedTableContentMeasurementDocument =
        await api.PptxDocument.open(tableContentMeasurementEvidenceBlob);
      const reopenedTableContentMeasurementTables =
        reopenedTableContentMeasurementDocument.slides.flatMap((slide) =>
          slide.shapes.filter((shape) => shape instanceof api.TableModel));
      const reopenedTableContentMeasurementPageSlides =
        reopenedTableContentMeasurementDocument.slides.filter((slide) =>
          slide.shapes.some((shape) => shape instanceof api.TableModel &&
            shape.name === 'browser_measurement_table'));
      const reopenedTableContentMeasurementPages =
        reopenedTableContentMeasurementPageSlides.map((slide) =>
          slide.shapes.find((shape) => shape instanceof api.TableModel &&
            shape.name === 'browser_measurement_table'));
      const tableContentMeasurementReopened =
        reopenedTableContentMeasurementTables.length ===
          tableContentMeasurementFinalTables.length + 2 &&
        reopenedTableContentMeasurementPages.every(
          (table) => tableContentMeasurementRowsPositive(table) &&
            table.placeholder?.type === 'tbl' && table.placeholder.index === 206,
        ) &&
        reopenedTableContentMeasurementPages.every(
          (table) => table.rows[0]?.cells[0]?.text ===
            'Browser measured header',
        ) &&
        reopenedTableContentMeasurementPages.some((table) =>
          table.rows.some((row) => row.cells[0]?.text ===
            'Browser measured edited tail')) &&
        reopenedTableContentMeasurementPages.some((table) =>
          table.mergeRegions?.some(({ rowspan, colspan }) =>
            rowspan === 2 && colspan === 1) === true) &&
        reopenedTableContentMeasurementPageSlides.every((slide) =>
          tableContentMeasurementLinksOwned(
            reopenedTableContentMeasurementDocument,
            slide,
          ));
      const tableContentMeasurementState = {
        created: tableContentMeasurementCreated,
        automatic: tableContentMeasurementAutomatic,
        minimum: tableContentMeasurementMinimum,
        fragment: tableContentMeasurementFragment,
        placeholder: tableContentMeasurementPlaceholder,
        fixed: tableContentMeasurementFixed,
        relationships: tableContentMeasurementRelationships,
        edited: tableContentMeasurementEdited,
        moved: tableContentMeasurementMoved,
        deleted: tableContentMeasurementDeleted,
        reopened: tableContentMeasurementReopened,
        validationErrors: tableContentMeasurementDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableContentMeasurementDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableContentMeasurement = Object.values(
        tableContentMeasurementState,
      ).every((value) => value === true || value === 0);
      const tableToSlidesElement = globalThis.document.createElement('table');
      tableToSlidesElement.id = 'browser-table-to-slides';
      tableToSlidesElement.style.borderCollapse = 'collapse';
      tableToSlidesElement.style.tableLayout = 'fixed';
      tableToSlidesElement.style.width = '600px';
      tableToSlidesElement.style.margin = '16px';
      tableToSlidesElement.style.background = 'white';
      const tableToSlidesAppendCell = (row, text, options = {}) => {
        const element = globalThis.document.createElement(
          options.header === true ? 'th' : 'td',
        );
        element.innerText = text;
        if (options.width !== undefined) element.style.width = options.width + 'px';
        if (options.colspan !== undefined) element.colSpan = options.colspan;
        if (options.rowspan !== undefined) element.rowSpan = options.rowspan;
        for (const [name, value] of Object.entries(options.attributes ?? {})) {
          element.setAttribute(name, value);
        }
        Object.assign(element.style, options.style ?? {});
        row.append(element);
        return element;
      };
      const tableToSlidesHeaderStyle = {
        color: 'rgb(1, 2, 3)',
        backgroundColor: 'rgb(240, 241, 242)',
        fontFamily: 'Aptos, Arial, sans-serif',
        fontSize: '16px',
        fontWeight: '600',
        textAlign: 'center',
        verticalAlign: 'middle',
        direction: 'ltr',
        padding: '4px 6px',
        borderTop: '2px solid rgb(10, 20, 30)',
        borderRight: '1px dashed rgb(40, 50, 60)',
        borderBottom: '1px solid rgb(70, 80, 90)',
        borderLeft: '1px solid rgb(100, 110, 120)',
      };
      const tableToSlidesHead = tableToSlidesElement.createTHead();
      const tableToSlidesHeaderRow = tableToSlidesHead.insertRow();
      tableToSlidesAppendCell(tableToSlidesHeaderRow, 'Browser HTML header A', {
        header: true,
        width: 150,
        attributes: { 'data-pptx-min-width': '0.8' },
        style: tableToSlidesHeaderStyle,
      });
      tableToSlidesAppendCell(tableToSlidesHeaderRow, 'Browser HTML header B', {
        header: true,
        width: 300,
        attributes: { 'data-pptx-width': '2.5' },
        style: {
          ...tableToSlidesHeaderStyle,
          backgroundColor: 'transparent',
          textAlign: 'right',
        },
      });
      tableToSlidesAppendCell(tableToSlidesHeaderRow, 'Browser HTML header C', {
        header: true,
        width: 150,
        style: tableToSlidesHeaderStyle,
      });
      const tableToSlidesSubheadRow = tableToSlidesHead.insertRow();
      tableToSlidesAppendCell(tableToSlidesSubheadRow, 'Browser HTML subhead A', {
        header: true,
        width: 150,
      });
      tableToSlidesAppendCell(tableToSlidesSubheadRow, 'Browser HTML subhead B', {
        header: true,
        width: 300,
      });
      tableToSlidesAppendCell(tableToSlidesSubheadRow, 'Browser HTML subhead C', {
        header: true,
        width: 150,
      });
      const tableToSlidesFirstBody = tableToSlidesElement.createTBody();
      const tableToSlidesFragmentRow = tableToSlidesFirstBody.insertRow();
      tableToSlidesAppendCell(
        tableToSlidesFragmentRow,
        'Browser fragmented HTML ' + 'fragment '.repeat(60),
        {
          colspan: 3,
          style: {
            color: 'rgb(128, 0, 32)',
            fontSize: '14px',
            fontWeight: '500',
            padding: '2px',
          },
        },
      );
      const tableToSlidesMergeRow = tableToSlidesFirstBody.insertRow();
      tableToSlidesAppendCell(tableToSlidesMergeRow, 'Browser rowspan', {
        rowspan: 2,
      });
      tableToSlidesAppendCell(tableToSlidesMergeRow, 'Browser colspan', {
        colspan: 2,
      });
      const tableToSlidesMergeLower = tableToSlidesFirstBody.insertRow();
      tableToSlidesAppendCell(tableToSlidesMergeLower, 'Browser merge lower B');
      tableToSlidesAppendCell(tableToSlidesMergeLower, 'Browser merge lower C');
      const tableToSlidesSecondBody = tableToSlidesElement.createTBody();
      for (let index = 0; index < 4; index += 1) {
        const row = tableToSlidesSecondBody.insertRow();
        tableToSlidesAppendCell(row, 'Browser tail ' + String(index + 1));
        tableToSlidesAppendCell(row, 'Browser value ' + String(index + 1));
        tableToSlidesAppendCell(
          row,
          'Browser detail ' + String(index + 1) + ' ' + 'content '.repeat(4),
        );
      }
      const tableToSlidesFoot = tableToSlidesElement.createTFoot();
      const tableToSlidesFootRow = tableToSlidesFoot.insertRow();
      tableToSlidesAppendCell(tableToSlidesFootRow, 'Browser footer A');
      tableToSlidesAppendCell(tableToSlidesFootRow, 'Browser footer B');
      tableToSlidesAppendCell(tableToSlidesFootRow, 'Browser footer C');
      globalThis.document.body.append(tableToSlidesElement);
      void tableToSlidesElement.offsetWidth;
      const tableToSlidesDocument = api.PptxDocument.create({
        firstSlideNumber: 31,
        slideSize: 'wide',
      });
      const tableToSlidesLayout = await tableToSlidesDocument.defineSlideMaster({
        title: 'BROWSER-HTML-TABLE',
        margin: [
          api.inches(0.4),
          api.inches(0.5),
          api.inches(0.5),
          api.inches(0.5),
        ],
        background: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'F7F9FC' },
        },
        slideNumber: {
          x: api.inches(12.2),
          y: api.inches(7),
          width: api.inches(0.5),
          height: api.inches(0.25),
        },
      });
      const tableToSlidesTargetSection = tableToSlidesDocument.addSection({
        title: 'Browser HTML target',
      });
      const tableToSlidesTarget = tableToSlidesDocument.addSlide({
        masterName: tableToSlidesLayout.name,
        sectionTitle: tableToSlidesTargetSection.title,
      });
      tableToSlidesTarget.addText('Browser HTML internal target', {
        x: api.inches(0.5),
        y: api.inches(0.5),
        width: api.inches(4),
        height: api.inches(0.5),
      });
      const tableToSlidesPngBytes = Uint8Array.from([
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
        0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
        39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
      ]);
      const tableToSlidesExternalUrl =
        'https://browser-table-to-slides.example/external';
      const tableToSlidesTableUrl =
        'https://browser-table-to-slides.example/table';
      const tableToSlidesPages = await tableToSlidesDocument.tableToSlides(
        tableToSlidesElement.id,
        {
          name: 'browser_html_table',
          masterSlideName: tableToSlidesLayout.name,
          autoPage: true,
          autoPageCharWeight: -1,
          autoPageLineWeight: 1,
          autoPageRepeatHeader: true,
          autoPageHeaderRows: 2,
          x: api.inches(0.5),
          y: api.inches(0.6),
          width: api.inches(6),
          height: api.inches(6.2),
          slideMargin: api.inches(0.5),
          addImage: {
            source: tableToSlidesPngBytes,
            options: {
              name: 'browser_html_image',
              x: api.inches(7.2),
              y: api.inches(0.6),
              width: api.inches(1),
              height: api.inches(0.5),
            },
          },
          addShape: {
            type: 'roundRect',
            options: {
              name: 'browser_html_shape',
              x: api.inches(7.2),
              y: api.inches(1.3),
              width: api.inches(1.2),
              height: api.inches(0.5),
              fill: {
                kind: 'solid',
                color: { kind: 'scheme', value: 'accent2' },
              },
            },
          },
          addTable: {
            rows: [[{
              text: 'Browser addition table',
              options: {
                hyperlink: {
                  url: tableToSlidesTableUrl,
                  tooltip: 'Browser table addition',
                },
              },
            }]],
            options: {
              name: 'browser_html_addition_table',
              x: api.inches(7.2),
              y: api.inches(2),
              width: api.inches(3),
            },
          },
          addText: {
            text: [{ runs: [
              {
                text: 'Browser external',
                style: {
                  bold: true,
                  hyperlink: {
                    url: tableToSlidesExternalUrl,
                    tooltip: 'Browser external addition',
                  },
                },
              },
              {
                text: ' / internal',
                style: {
                  italic: true,
                  hyperlink: {
                    slide: tableToSlidesDocument.slides.indexOf(
                      tableToSlidesTarget,
                    ) + 1,
                    tooltip: 'Browser internal addition',
                  },
                },
              },
            ] }],
            options: {
              name: 'browser_html_text',
              x: api.inches(7.2),
              y: api.inches(3),
              width: api.inches(4.8),
              height: api.inches(0.7),
            },
          },
        },
      );
      const tableToSlidesPageTables = tableToSlidesPages.map((page) =>
        page.shapes.find((shape) =>
          shape instanceof api.TableModel && shape.name === 'browser_html_table'));
      const tableToSlidesAdditionNames = [
        'browser_html_table',
        'browser_html_image',
        'browser_html_shape',
        'browser_html_addition_table',
        'browser_html_text',
      ];
      const tableToSlidesCreated = tableToSlidesPages.length > 2 &&
        Object.isFrozen(tableToSlidesPages) &&
        tableToSlidesPageTables.every((table) => table instanceof api.TableModel);
      const tableToSlidesFirstCell = tableToSlidesPageTables[0]?.rows[0]?.cells[0];
      const tableToSlidesSecondCell = tableToSlidesPageTables[0]?.rows[0]?.cells[1];
      const tableToSlidesFirstRunStyle =
        tableToSlidesFirstCell?.richText[0]?.runs[0]?.style;
      const tableToSlidesStyles =
        tableToSlidesFirstRunStyle?.color?.kind === 'srgb' &&
        tableToSlidesFirstRunStyle.color.value === '010203' &&
        tableToSlidesFirstCell.fill?.kind === 'solid' &&
        tableToSlidesFirstCell.fill.color.kind === 'srgb' &&
        tableToSlidesFirstCell.fill.color.value === 'F0F1F2' &&
        tableToSlidesFirstRunStyle.fontFamily === 'Aptos' &&
        tableToSlidesFirstRunStyle.fontSize === 16 &&
        tableToSlidesFirstRunStyle.bold === true &&
        tableToSlidesFirstCell.horizontalAlignment === 'center' &&
        tableToSlidesFirstCell.verticalAlignment === 'middle' &&
        tableToSlidesFirstCell.borders?.top?.kind === 'line' &&
        tableToSlidesFirstCell.borders.top.width === 2 &&
        tableToSlidesSecondCell?.fill?.kind === 'solid' &&
        tableToSlidesSecondCell.fill.color.kind === 'srgb' &&
        tableToSlidesSecondCell.fill.color.value === 'FFFFFF' &&
        tableToSlidesSecondCell.horizontalAlignment === 'right';
      const tableToSlidesWidths = tableToSlidesPageTables.every((table) =>
        table.columnWidths?.length === 3 &&
        table.columnWidths[1] === api.inches(2.5) &&
        table.columnWidths.every((width) => width > 0) &&
        table.columnWidths.reduce((sum, width) => sum + width, 0) ===
          api.inches(6) &&
        table.transform.width === api.inches(6));
      const tableToSlidesHeaders = tableToSlidesPageTables.every((table) =>
        table.rows[0]?.cells[0]?.text === 'Browser HTML header A' &&
        table.rows[1]?.cells[0]?.text === 'Browser HTML subhead A');
      const tableToSlidesGeneratedSection = tableToSlidesDocument.sections?.find(
        ({ slideIds }) => tableToSlidesPages.every((page) =>
          slideIds.includes(page.slideId)),
      );
      const tableToSlidesLayoutState = tableToSlidesPages.every((page) =>
        page.relationships.some(({ type, resolvedTarget }) =>
          type.endsWith('/slideLayout') &&
            resolvedTarget === tableToSlidesLayout.partUri) &&
        page.slideNumber !== undefined) &&
        tableToSlidesGeneratedSection?.slideIds.join(',') ===
          tableToSlidesPages.map(({ slideId }) => slideId).join(',');
      const tableToSlidesAdditions = tableToSlidesPages.every((page) =>
        JSON.stringify(page.shapes.filter(({ name }) =>
          tableToSlidesAdditionNames.includes(name)).map(({ name }) => name)) ===
            JSON.stringify(tableToSlidesAdditionNames));
      const tableToSlidesImageTargets = new Set(tableToSlidesPages.flatMap((page) =>
        page.relationships.filter(({ type }) => type.endsWith('/image'))
          .map(({ resolvedTarget }) => resolvedTarget)));
      const tableToSlidesRelationships = tableToSlidesImageTargets.size === 1 &&
        !tableToSlidesImageTargets.has(undefined) &&
        tableToSlidesPages.every((page) =>
          page.relationships.some(({ type, target }) =>
            type.endsWith('/hyperlink') && target === tableToSlidesExternalUrl) &&
          page.relationships.some(({ type, target }) =>
            type.endsWith('/hyperlink') && target === tableToSlidesTableUrl) &&
          page.relationships.some(({ type, resolvedTarget }) =>
            type.endsWith('/slide') &&
              resolvedTarget === tableToSlidesTarget.partUri));
      const tableToSlidesEditableTable = tableToSlidesPageTables.find((table) =>
        table.rows.some((row) => row.cells[0]?.text.startsWith(
          'Browser tail ',
        ) === true));
      const tableToSlidesEditableRow = tableToSlidesEditableTable?.rows.findIndex(
        (row) => row.cells[0]?.text.startsWith('Browser tail ') === true,
      ) ?? -1;
      if (tableToSlidesEditableTable instanceof api.TableModel &&
          tableToSlidesEditableRow >= 0) {
        tableToSlidesEditableTable.setCellText(
          tableToSlidesEditableRow,
          0,
          'Browser HTML edited',
        );
      }
      const tableToSlidesEditablePage = tableToSlidesPages[
        tableToSlidesPageTables.indexOf(tableToSlidesEditableTable)
      ];
      const tableToSlidesLifecyclePage = tableToSlidesPages.find(
        (page, index) => page !== tableToSlidesEditablePage &&
          tableToSlidesPageTables[index]?.rows.some((row) =>
            row.cells[0]?.text.startsWith('Browser tail ') === true) === true &&
          index > 0 && index < tableToSlidesPages.length - 1,
      ) ?? tableToSlidesPages.find((page) => page !== tableToSlidesEditablePage);
      const tableToSlidesLifecycleOriginalIndex = tableToSlidesDocument.slides.indexOf(
        tableToSlidesLifecyclePage,
      );
      tableToSlidesDocument.moveSlide(
        tableToSlidesLifecycleOriginalIndex,
        tableToSlidesDocument.slides.length - 1,
      );
      const tableToSlidesMovedAway = tableToSlidesDocument.slides.at(-1) ===
        tableToSlidesLifecyclePage;
      tableToSlidesDocument.moveSlide(
        tableToSlidesDocument.slides.indexOf(tableToSlidesLifecyclePage),
        tableToSlidesLifecycleOriginalIndex,
      );
      const tableToSlidesMovedBack = tableToSlidesDocument.slides[
        tableToSlidesLifecycleOriginalIndex
      ] === tableToSlidesLifecyclePage;
      tableToSlidesDocument.deleteSlide(
        tableToSlidesDocument.slides.indexOf(tableToSlidesLifecyclePage),
      );
      const tableToSlidesDeleted = !tableToSlidesDocument.slides.includes(
        tableToSlidesLifecyclePage,
      );
      const tableToSlidesEdited =
        tableToSlidesEditableTable instanceof api.TableModel &&
        tableToSlidesEditableRow >= 0 &&
        tableToSlidesEditableTable.rows[tableToSlidesEditableRow]?.cells[0]?.text ===
          'Browser HTML edited' &&
        tableToSlidesMovedAway && tableToSlidesMovedBack && tableToSlidesDeleted;
      const tableToSlidesEvidenceBlob = await tableToSlidesDocument.writeBlob();
      globalThis.__pptxTableToSlidesEvidenceBlob = tableToSlidesEvidenceBlob;
      const reopenedTableToSlidesDocument = await api.PptxDocument.open(
        tableToSlidesEvidenceBlob,
      );
      const reopenedTableToSlidesPages = reopenedTableToSlidesDocument.slides.filter(
        (page) => page.shapes.some((shape) =>
          shape instanceof api.TableModel && shape.name === 'browser_html_table'),
      );
      const reopenedTableToSlidesTables = reopenedTableToSlidesPages.map((page) =>
        page.shapes.find((shape) =>
          shape instanceof api.TableModel && shape.name === 'browser_html_table'));
      const tableToSlidesReopened =
        reopenedTableToSlidesPages.length === tableToSlidesPages.length - 1 &&
        reopenedTableToSlidesTables.every((table) =>
          table instanceof api.TableModel &&
            table.columnWidths?.length === 3 &&
            table.columnWidths[1] === api.inches(2.5) &&
            table.rows[0]?.cells[0]?.text === 'Browser HTML header A' &&
            table.rows[1]?.cells[0]?.text === 'Browser HTML subhead A') &&
        reopenedTableToSlidesTables.some((table) =>
          table.rows.some((row) => row.cells[0]?.text === 'Browser HTML edited')) &&
        reopenedTableToSlidesPages.every((page) =>
          tableToSlidesAdditionNames.every((name) =>
            page.shapes.some((shape) => shape.name === name))) &&
        reopenedTableToSlidesPages.every((page) =>
          page.relationships.some(({ type, target }) =>
            type.endsWith('/hyperlink') && target === tableToSlidesExternalUrl) &&
          page.relationships.some(({ type, resolvedTarget }) =>
            type.endsWith('/slide') &&
              resolvedTarget === reopenedTableToSlidesDocument.slides[0]?.partUri));
      const tableToSlidesState = {
        created: tableToSlidesCreated,
        styles: tableToSlidesStyles,
        widths: tableToSlidesWidths,
        headers: tableToSlidesHeaders,
        layout: tableToSlidesLayoutState,
        additions: tableToSlidesAdditions,
        relationships: tableToSlidesRelationships,
        edited: tableToSlidesEdited,
        reopened: tableToSlidesReopened,
        validationErrors: tableToSlidesDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableToSlidesDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableToSlides = Object.values(tableToSlidesState).every(
        (value) => value === true || value === 0,
      );
      const tableBorderSideSnapshot = (value) => {
        if (value.kind === 'none') return { kind: 'none' };
        return {
          kind: 'line',
          color: { kind: value.color.kind, value: value.color.value },
          width: value.width,
          ...(value.style !== undefined ? { style: value.style } : {}),
        };
      };
      const tableBordersSnapshot = (value) => value === undefined
        ? null
        : {
            ...(value.top !== undefined
              ? { top: tableBorderSideSnapshot(value.top) }
              : {}),
            ...(value.right !== undefined
              ? { right: tableBorderSideSnapshot(value.right) }
              : {}),
            ...(value.bottom !== undefined
              ? { bottom: tableBorderSideSnapshot(value.bottom) }
              : {}),
            ...(value.left !== undefined
              ? { left: tableBorderSideSnapshot(value.left) }
              : {}),
          };
      const tableBordersDocument = api.PptxDocument.create();
      const tableBordersSlide = tableBordersDocument.addSlide();
      const tableBordersInitialLine = {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      };
      const tableBordersTable = tableBordersSlide.addTable([
        ['North', 'South'],
        ['East', 'West'],
      ], {
        name: 'Chrome table borders',
        border: tableBordersInitialLine,
      });
      const tableBordersPart = () => tableBordersDocument.opcPackage
        .requirePart(tableBordersSlide.partUri).bytes;
      const tableBordersReadBytes = tableBordersPart().slice();
      const tableBordersReadJournal = JSON.stringify(
        tableBordersDocument.opcPackage.mutations,
      );
      const tableBordersUniform = tableBordersSnapshot(tableBordersTable.borders);
      const tableBordersDetached = tableBordersTable.borders;
      if (tableBordersDetached?.top?.kind === 'line') {
        tableBordersDetached.top.color.value = 'accent6';
        tableBordersDetached.top.width = 99;
      }
      const tableBordersReadIsolation = tableMarginsBytesEqual(
        tableBordersReadBytes,
        tableBordersPart(),
      ) && JSON.stringify(tableBordersDocument.opcPackage.mutations) ===
        tableBordersReadJournal &&
        JSON.stringify(tableBordersSnapshot(tableBordersTable.borders)) ===
          JSON.stringify({
            top: tableBordersInitialLine,
            right: tableBordersInitialLine,
            bottom: tableBordersInitialLine,
            left: tableBordersInitialLine,
          });
      const tableBordersNoOpBytes = tableBordersPart().slice();
      const tableBordersNoOpJournal = JSON.stringify(
        tableBordersDocument.opcPackage.mutations,
      );
      tableBordersTable.borders = tableBordersInitialLine;
      const tableBordersNoOp = tableMarginsBytesEqual(
        tableBordersNoOpBytes,
        tableBordersPart(),
      ) && JSON.stringify(tableBordersDocument.opcPackage.mutations) ===
        tableBordersNoOpJournal;
      tableBordersTable.setCellBorders(0, 1, { kind: 'none' });
      const tableBordersMixed = tableBordersSnapshot(tableBordersTable.borders);
      const tableBordersPartialInput = {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: 'D9EAF7' },
          width: 2,
        },
        bottom: { kind: 'none' },
      };
      tableBordersTable.borders = tableBordersPartialInput;
      const tableBordersPartial = tableBordersSnapshot(tableBordersTable.borders);
      const tableBordersPartialCells = tableBordersTable.rows.flatMap(({ cells }) =>
        cells.map(({ borders }) => tableBordersSnapshot(borders)));
      tableBordersTable.borders = { kind: 'none' };
      const tableBordersNone = tableBordersSnapshot(tableBordersTable.borders);
      const tableBordersNoneCells = tableBordersTable.rows.flatMap(({ cells }) =>
        cells.map(({ borders }) => tableBordersSnapshot(borders)));
      tableBordersTable.borders = undefined;
      const tableBordersCleared = tableBordersSnapshot(tableBordersTable.borders);
      const tableBordersClearedCells = tableBordersTable.rows.flatMap(({ cells }) =>
        cells.map(({ borders }) => tableBordersSnapshot(borders)));
      const tableBordersInvalidBytes = tableBordersPart().slice();
      const tableBordersInvalidJournal = JSON.stringify(
        tableBordersDocument.opcPackage.mutations,
      );
      let tableBordersInvalidError;
      try {
        tableBordersTable.borders = null;
      } catch (error) {
        tableBordersInvalidError = { name: error.name, message: error.message };
      }
      const tableBordersFailureIsolation = tableMarginsBytesEqual(
        tableBordersInvalidBytes,
        tableBordersPart(),
      ) && JSON.stringify(tableBordersDocument.opcPackage.mutations) ===
        tableBordersInvalidJournal;
      tableBordersTable.borders = { kind: 'none' };
      const reopenedTableBordersDocument = await api.PptxDocument.open(
        await tableBordersDocument.writeBlob(),
      );
      const reopenedTableBordersTable = reopenedTableBordersDocument.slides[0]
        .shapes.find((shape) => shape.name === 'Chrome table borders');
      const tableBordersState = {
        uniform: tableBordersUniform,
        readIsolation: tableBordersReadIsolation,
        noOp: tableBordersNoOp,
        mixed: tableBordersMixed,
        partial: tableBordersPartial,
        partialCells: tableBordersPartialCells,
        none: tableBordersNone,
        noneCells: tableBordersNoneCells,
        cleared: tableBordersCleared,
        clearedCells: tableBordersClearedCells,
        reopened: reopenedTableBordersTable instanceof api.TableModel
          ? tableBordersSnapshot(reopenedTableBordersTable.borders)
          : null,
        reopenedCells: reopenedTableBordersTable instanceof api.TableModel
          ? reopenedTableBordersTable.rows.flatMap(({ cells }) =>
            cells.map(({ borders }) => tableBordersSnapshot(borders)))
          : [],
        invalidError: tableBordersInvalidError,
        failureIsolation: tableBordersFailureIsolation,
        validationErrors: tableBordersDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableBordersDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableBordersNoneSnapshot = {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      };
      const tableBorders = JSON.stringify(tableBordersState) === JSON.stringify({
        uniform: {
          top: tableBordersInitialLine,
          right: tableBordersInitialLine,
          bottom: tableBordersInitialLine,
          left: tableBordersInitialLine,
        },
        readIsolation: true,
        noOp: true,
        mixed: null,
        partial: tableBordersPartialInput,
        partialCells: Array(4).fill(tableBordersPartialInput),
        none: tableBordersNoneSnapshot,
        noneCells: Array(4).fill(tableBordersNoneSnapshot),
        cleared: null,
        clearedCells: [null, null, null, null],
        reopened: tableBordersNoneSnapshot,
        reopenedCells: Array(4).fill(tableBordersNoneSnapshot),
        invalidError: {
          name: 'TypeError',
          message: 'Table borders must be an object',
        },
        failureIsolation: true,
        validationErrors: 0,
      });
      const schemeColorIsolationDocument = api.PptxDocument.create();
      const schemeColorIsolationJournal = JSON.stringify(
        schemeColorIsolationDocument.opcPackage.mutations,
      );
      const schemeColorDocument = api.PptxDocument.create();
      schemeColorDocument.addSlide().addRichText([{
        runs: [{
          text: 'Chrome scheme colors',
          style: { color: { kind: 'scheme', value: api.SCHEME_COLORS.text1 } },
        }],
      }], {
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: api.SCHEME_COLORS.accent1 },
        },
      });
      const reopenedSchemeColorDocument = await api.PptxDocument.open(
        await schemeColorDocument.writeBlob(),
      );
      const reopenedSchemeColorShape = reopenedSchemeColorDocument.slides[0].shapes[0];
      const schemeColorState = {
        entries: Object.entries(api.SCHEME_COLORS),
        frozen: Object.isFrozen(api.SCHEME_COLORS),
        mutationIsolation: JSON.stringify(
          schemeColorIsolationDocument.opcPackage.mutations,
        ) === schemeColorIsolationJournal,
        textColor: reopenedSchemeColorShape instanceof api.ShapeModel
          ? reopenedSchemeColorShape.richText[0]?.runs[0]?.style?.color
          : undefined,
        fill: reopenedSchemeColorShape instanceof api.ShapeModel
          ? reopenedSchemeColorShape.fill
          : undefined,
        validationErrors: reopenedSchemeColorDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length,
      };
      const schemeColors = JSON.stringify(schemeColorState) === JSON.stringify({
        entries: [
          ['text1', 'tx1'],
          ['text2', 'tx2'],
          ['background1', 'bg1'],
          ['background2', 'bg2'],
          ['accent1', 'accent1'],
          ['accent2', 'accent2'],
          ['accent3', 'accent3'],
          ['accent4', 'accent4'],
          ['accent5', 'accent5'],
          ['accent6', 'accent6'],
        ],
        frozen: true,
        mutationIsolation: true,
        textColor: { kind: 'scheme', value: 'tx1' },
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
        validationErrors: 0,
      });
      const outputTypeDocument = api.PptxDocument.create();
      const outputTypeJournal = JSON.stringify(outputTypeDocument.opcPackage.mutations);
      const outputTypeState = {
        values: [...api.OUTPUT_TYPES],
        frozen: Object.isFrozen(api.OUTPUT_TYPES),
        mutationIsolation: JSON.stringify(outputTypeDocument.opcPackage.mutations) ===
          outputTypeJournal,
      };
      const outputTypes = JSON.stringify(outputTypeState) === JSON.stringify({
        values: [
          'arraybuffer',
          'base64',
          'binarystring',
          'blob',
          'nodebuffer',
          'uint8array',
        ],
        frozen: true,
        mutationIsolation: true,
      });
      const writeOutputDocument = api.PptxDocument.create();
      writeOutputDocument.addSlide().addText('Browser output types 你好');
      const writeOutputJournal = JSON.stringify(writeOutputDocument.opcPackage.mutations);
      const defaultWriteOutput = await writeOutputDocument.write();
      const emptyWriteOutput = await writeOutputDocument.write({});
      const arrayBufferWriteOutput = await writeOutputDocument.write({
        outputType: 'arraybuffer',
      });
      const base64WriteOutput = await writeOutputDocument.write({ outputType: 'base64' });
      const binaryStringWriteOutput = await writeOutputDocument.write({
        outputType: 'binarystring',
      });
      const blobWriteOutput = await writeOutputDocument.write({ outputType: 'blob' });
      const uint8ArrayWriteOutput = await writeOutputDocument.write({
        outputType: 'uint8array',
      });
      const convenienceWriteBlob = await writeOutputDocument.writeBlob();
      const decodeWriteOutput = async (outputType, value) => {
        if (outputType === 'arraybuffer') return new Uint8Array(value);
        if (outputType === 'base64') {
          return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
        }
        if (outputType === 'binarystring') {
          return Uint8Array.from(value, (character) => character.charCodeAt(0));
        }
        if (outputType === 'blob') return new Uint8Array(await value.arrayBuffer());
        return new Uint8Array(value);
      };
      const writeOutputValues = [
        ['arraybuffer', arrayBufferWriteOutput],
        ['base64', base64WriteOutput],
        ['binarystring', binaryStringWriteOutput],
        ['blob', blobWriteOutput],
        ['uint8array', uint8ArrayWriteOutput],
      ];
      const decodedWriteOutputs = await Promise.all(
        writeOutputValues.map(([outputType, value]) => decodeWriteOutput(outputType, value)),
      );
      const equalWriteOutputBytes = (left, right) =>
        left.byteLength === right.byteLength
          && left.every((value, index) => value === right[index]);
      const reopenedWriteOutputTitles = [];
      for (const outputBytes of decodedWriteOutputs) {
        const reopenedWriteOutput = await api.PptxDocument.open(outputBytes);
        const outputShape = reopenedWriteOutput.slides[0].shapes[0];
        reopenedWriteOutputTitles.push(
          outputShape instanceof api.ShapeModel ? outputShape.text : undefined,
        );
      }
      const failureDiagnostics = JSON.stringify(writeOutputDocument.diagnostics);
      const failureJournal = JSON.stringify(writeOutputDocument.opcPackage.mutations);
      let nodebufferError;
      try {
        await writeOutputDocument.write({ outputType: 'nodebuffer' });
      } catch (error) {
        nodebufferError = { name: error.name, message: error.message };
      }
      const writeOutputTypeState = {
        defaultKind: defaultWriteOutput instanceof Uint8Array ? 'uint8array' : typeof defaultWriteOutput,
        emptyKind: emptyWriteOutput instanceof Uint8Array ? 'uint8array' : typeof emptyWriteOutput,
        arraybufferKind: arrayBufferWriteOutput instanceof ArrayBuffer
          ? 'arraybuffer'
          : typeof arrayBufferWriteOutput,
        base64Kind: typeof base64WriteOutput,
        binarystringKind: typeof binaryStringWriteOutput,
        blobKind: blobWriteOutput instanceof Blob ? 'blob' : typeof blobWriteOutput,
        blobType: blobWriteOutput.type,
        uint8arrayKind: uint8ArrayWriteOutput instanceof Uint8Array
          ? 'uint8array'
          : typeof uint8ArrayWriteOutput,
        byteEquality: decodedWriteOutputs.every((outputBytes) =>
          equalWriteOutputBytes(outputBytes, defaultWriteOutput)),
        reopenTitles: reopenedWriteOutputTitles,
        writeBlobType: convenienceWriteBlob.type,
        nodebufferError,
        failureIsolation: JSON.stringify(writeOutputDocument.diagnostics) === failureDiagnostics
          && JSON.stringify(writeOutputDocument.opcPackage.mutations) === failureJournal,
        mutationIsolation: JSON.stringify(writeOutputDocument.opcPackage.mutations) ===
          writeOutputJournal,
      };
      const writeOutputTypes = JSON.stringify(writeOutputTypeState) === JSON.stringify({
        defaultKind: 'uint8array',
        emptyKind: 'uint8array',
        arraybufferKind: 'arraybuffer',
        base64Kind: 'string',
        binarystringKind: 'string',
        blobKind: 'blob',
        blobType: 'application/zip',
        uint8arrayKind: 'uint8array',
        byteEquality: true,
        reopenTitles: Array(5).fill('Browser output types 你好'),
        writeBlobType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        nodebufferError: {
          name: 'Error',
          message: 'nodebuffer is not supported by this platform',
        },
        failureIsolation: true,
        mutationIsolation: true,
      }) && writeOutputDocument.diagnostics
        .filter(({ severity }) => severity === 'error').length === 0;
      const nodeReadableStreamDocument = api.PptxDocument.create();
      nodeReadableStreamDocument.addSlide().addText('Browser node stream boundary');
      const nodeReadableStreamDiagnostics = JSON.stringify(
        nodeReadableStreamDocument.diagnostics,
      );
      const nodeReadableStreamJournal = JSON.stringify(
        nodeReadableStreamDocument.opcPackage.mutations,
      );
      let nodeReadableStreamError;
      try {
        await nodeReadableStreamDocument.stream();
      } catch (error) {
        nodeReadableStreamError = { name: error.name, message: error.message };
      }
      const nodeReadableStreamBytes = await nodeReadableStreamDocument.write();
      const nodeReadableStreamReopened = await api.PptxDocument.open(
        nodeReadableStreamBytes,
      );
      const nodeReadableStreamShape = nodeReadableStreamReopened.slides[0].shapes[0];
      const nodeReadableStreamState = {
        error: nodeReadableStreamError,
        failureIsolation: JSON.stringify(nodeReadableStreamDocument.diagnostics) ===
          nodeReadableStreamDiagnostics &&
          JSON.stringify(nodeReadableStreamDocument.opcPackage.mutations) ===
          nodeReadableStreamJournal,
        laterWriteTitle: nodeReadableStreamShape instanceof api.ShapeModel
          ? nodeReadableStreamShape.text
          : undefined,
      };
      const nodeReadableStream = JSON.stringify(nodeReadableStreamState) === JSON.stringify({
        error: {
          name: 'Error',
          message: 'PptxDocument.stream() is only supported in Node.js',
        },
        failureIsolation: true,
        laterWriteTitle: 'Browser node stream boundary',
      });
      const compressionMethods = (zipBytes) => {
        const view = new DataView(
          zipBytes.buffer,
          zipBytes.byteOffset,
          zipBytes.byteLength,
        );
        let eocd = zipBytes.byteLength - 22;
        while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
        if (eocd < 0) throw new Error('ZIP EOCD not found');
        const entries = view.getUint16(eocd + 10, true);
        let offset = view.getUint32(eocd + 16, true);
        const methods = [];
        for (let index = 0; index < entries; index += 1) {
          if (view.getUint32(offset, true) !== 0x02014b50) {
            throw new Error('ZIP central directory entry not found');
          }
          const nameLength = view.getUint16(offset + 28, true);
          const extraLength = view.getUint16(offset + 30, true);
          const commentLength = view.getUint16(offset + 32, true);
          const name = new TextDecoder().decode(
            zipBytes.subarray(offset + 46, offset + 46 + nameLength),
          );
          if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
          offset += 46 + nameLength + extraLength + commentLength;
        }
        return [...new Set(methods)];
      };
      const compressionEqual = (left, right) =>
        left.byteLength === right.byteLength &&
        left.every((value, index) => value === right[index]);
      const compressionDocument = api.PptxDocument.create();
      compressionDocument.addSlide().addText('Chrome compression policy');
      compressionDocument.opcPackage.setPart(
        '/custom/chrome-compression.bin',
        new Uint8Array(65_536).fill(0x41),
        'application/octet-stream',
      );
      const compressionDiagnostics = JSON.stringify(compressionDocument.diagnostics);
      const compressionJournal = JSON.stringify(compressionDocument.opcPackage.mutations);
      const compressionDefault = await compressionDocument.write();
      const compressionStore = await compressionDocument.write({ compression: false });
      const compressionDeflate = await compressionDocument.write({ compression: true });
      const compressionBlob = new Uint8Array(
        await (await compressionDocument.writeBlob({ compression: true })).arrayBuffer(),
      );
      const compressionReopened = await api.PptxDocument.open(compressionDeflate);
      const compressionShape = compressionReopened.slides[0].shapes[0];
      const originalCompressionPackageWrite = compressionDocument.opcPackage.write.bind(
        compressionDocument.opcPackage,
      );
      let compressionPackageWrites = 0;
      compressionDocument.opcPackage.write = (...args) => {
        compressionPackageWrites += 1;
        return originalCompressionPackageWrite(...args);
      };
      let compressionError;
      try {
        await compressionDocument.write({ compression: 'true' });
      } catch (error) {
        compressionError = { name: error.name, message: error.message };
      }
      compressionDocument.opcPackage.write = originalCompressionPackageWrite;
      const compressionLaterReopened = await api.PptxDocument.open(
        await compressionDocument.write({ compression: false }),
      );
      const compressionLaterShape = compressionLaterReopened.slides[0].shapes[0];
      const compressionPolicyState = {
        defaultEqualsFalse: compressionEqual(compressionDefault, compressionStore),
        storeMethods: compressionMethods(compressionStore),
        deflateMethods: compressionMethods(compressionDeflate),
        storeBytes: compressionStore.byteLength,
        deflateBytes: compressionDeflate.byteLength,
        deflateSmaller: compressionDeflate.byteLength < compressionStore.byteLength,
        blobEquality: compressionEqual(compressionBlob, compressionDeflate),
        reopenTitle: compressionShape instanceof api.ShapeModel
          ? compressionShape.text
          : undefined,
        invalidError: compressionError,
        invalidEarly: compressionPackageWrites === 0,
        laterWriteTitle: compressionLaterShape instanceof api.ShapeModel
          ? compressionLaterShape.text
          : undefined,
        failureIsolation: JSON.stringify(compressionDocument.diagnostics) ===
          compressionDiagnostics &&
          JSON.stringify(compressionDocument.opcPackage.mutations) === compressionJournal,
      };
      const compressionPolicy = JSON.stringify(compressionPolicyState) === JSON.stringify({
        defaultEqualsFalse: true,
        storeMethods: [0],
        deflateMethods: [8],
        storeBytes: compressionStore.byteLength,
        deflateBytes: compressionDeflate.byteLength,
        deflateSmaller: true,
        blobEquality: true,
        reopenTitle: 'Chrome compression policy',
        invalidError: {
          name: 'TypeError',
          message: 'PptxDocument output compression must be a boolean',
        },
        invalidEarly: true,
        laterWriteTitle: 'Chrome compression policy',
        failureIsolation: true,
      });
      const fromBlob = await api.PptxDocument.open(new Blob([bytes.buffer]));
      const stream = new ReadableStream({
        start(controller) {
          controller.enqueue(bytes.slice(0, 40));
          controller.enqueue(bytes.slice(40));
          controller.close();
        },
      });
      const document = await api.PptxDocument.open(stream);
      document.slides[0].title.text = 'Browser updated';
      const slideNumberDocument = api.PptxDocument.create({ firstSlideNumber: -2 });
      const slideNumberSource = slideNumberDocument.addSlide();
      slideNumberSource.slideNumber = {
        align: 'center',
        rtl: true,
        valign: 'middle',
        margin: [1, 2, 3, 4],
        style: {
          italic: true,
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
      };
      slideNumberDocument.layouts[0].slideNumber = { x: 200, align: 'center' };
      slideNumberDocument.masters[0].slideNumber = { x: 300, align: 'right' };
      const slideNumberDuplicate = slideNumberDocument.duplicateSlide(0);
      slideNumberDocument.moveSlide(slideNumberDocument.slides.indexOf(slideNumberDuplicate), 0);
      const slideNumberOutput = await slideNumberDocument.writeBlob();
      const reopenedSlideNumbers = await api.PptxDocument.open(slideNumberOutput);
      await reopenedSlideNumbers.write({ compatibility: 'powerpoint-current' });
      const slideNumberXml = (partUri) => new TextDecoder().decode(
        reopenedSlideNumbers.opcPackage.requirePart(partUri).bytes,
      );
      const slideNumberCache = (partUri) => {
        const xml = slideNumberXml(partUri);
        const fieldStart = xml.indexOf('type="slidenum"');
        const textStart = xml.indexOf('<a:t>', fieldStart);
        const textEnd = xml.indexOf('</a:t>', textStart);
        return fieldStart < 0 || textStart < 0 || textEnd < 0
          ? undefined
          : xml.slice(textStart + 5, textEnd);
      };
      const slideNumberOwnerCount = (partUri) =>
        slideNumberXml(partUri).split('type="sldNum"').length - 1;
      const slideNumberLayout = reopenedSlideNumbers.layouts[0];
      const slideNumberMaster = reopenedSlideNumbers.masters[0];
      const slideNumberState = {
        firstSlideNumber: reopenedSlideNumbers.firstSlideNumber,
        mime: slideNumberOutput.type,
        slideCount: reopenedSlideNumbers.slides.length,
        values: reopenedSlideNumbers.slides.map(({ slideNumber }) => ({
          width: slideNumber?.width,
          height: slideNumber?.height,
          align: slideNumber?.align,
          rtl: slideNumber?.rtl,
          valign: slideNumber?.valign,
          margin: slideNumber?.margin,
          italic: slideNumber?.style.italic,
          color: slideNumber?.style.color,
          transparency: slideNumber?.style.transparency,
        })),
        caches: reopenedSlideNumbers.slides.map(({ partUri }) => slideNumberCache(partUri)),
        layoutX: slideNumberLayout.slideNumber?.x,
        masterX: slideNumberMaster.slideNumber?.x,
        layoutCache: slideNumberCache(slideNumberLayout.partUri),
        masterCache: slideNumberCache(slideNumberMaster.partUri),
        masterEnabled: slideNumberXml(slideNumberMaster.partUri).includes('sldNum="1"'),
        ownerCounts: [
          ...reopenedSlideNumbers.slides.map(({ partUri }) => slideNumberOwnerCount(partUri)),
          slideNumberOwnerCount(slideNumberLayout.partUri),
          slideNumberOwnerCount(slideNumberMaster.partUri),
        ],
        diagnostics: reopenedSlideNumbers.diagnostics
          .filter(({ code }) => code.startsWith('SLIDE_NUMBER_'))
          .map(({ code }) => code),
      };
      const browserMasterPngBytes = Uint8Array.from([
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
        0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
        39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
      ]);
      const browserMasterPngDataUri =
        'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC'
        + 'AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
      const masterLayoutDocument = api.PptxDocument.create({ slideSize: 'wide' });
      masterLayoutDocument.masters[0].background = {
        kind: 'solid',
        color: { kind: 'srgb', value: 'F3F6FA' },
      };
      const browserMasterLayout = await masterLayoutDocument.defineSlideMaster({
        title: 'BROWSER-MASTER-LAYOUT',
        background: {
          kind: 'image-source',
          source: new Blob([browserMasterPngBytes], { type: 'image/png' }),
        },
        margin: [api.inches(0.1), api.inches(0.2), api.inches(0.3), api.inches(0.4)],
        objects: [
          {
            kind: 'placeholder', text: 'Title prompt',
            options: {
              name: 'browser_title', type: 'title', index: 101,
              x: api.inches(0.5), y: api.inches(0.5),
              width: api.inches(12), height: api.inches(0.6),
            },
          },
          {
            kind: 'placeholder', text: 'Body prompt',
            options: {
              name: 'browser_body', type: 'body', index: 102,
              x: api.inches(0.5), y: api.inches(1.3),
              width: api.inches(4), height: api.inches(0.8),
            },
          },
          {
            kind: 'placeholder', text: 'Picture prompt',
            options: {
              name: 'browser_picture', type: 'pic', index: 103,
              x: api.inches(0.5), y: api.inches(2.4),
              width: api.inches(3), height: api.inches(2),
            },
          },
          {
            kind: 'placeholder', text: 'Chart prompt',
            options: {
              name: 'browser_chart', type: 'chart', index: 104,
              x: api.inches(3.75), y: api.inches(2.4),
              width: api.inches(4), height: api.inches(2),
            },
          },
          {
            kind: 'placeholder', text: 'Table prompt',
            options: {
              name: 'browser_table', type: 'tbl', index: 105,
              x: api.inches(8), y: api.inches(2.4),
              width: api.inches(4.8), height: api.inches(2),
            },
          },
          {
            kind: 'placeholder', text: 'Media prompt',
            options: {
              name: 'browser_media', type: 'media', index: 106,
              x: api.inches(0.5), y: api.inches(4.8),
              width: api.inches(2.5), height: api.inches(1.4),
            },
          },
          {
            kind: 'image',
            source: browserMasterPngDataUri,
            options: { name: 'Browser layout image', width: api.inches(1), height: api.inches(1) },
          },
        ],
      });
      const browserMasterMargin = browserMasterLayout.margin;
      const browserMasterLiveIdentity =
        browserMasterLayout instanceof api.SlideLayoutModel
        && masterLayoutDocument.masters[0] instanceof api.SlideMasterModel
        && masterLayoutDocument.layouts.find(
          ({ partUri }) => partUri === browserMasterLayout.partUri,
        ) === browserMasterLayout
        && masterLayoutDocument.masters[0].layouts.some(
          (layout) => layout === browserMasterLayout,
        );
      const browserMasterSlide = masterLayoutDocument.addSlide({
        masterName: browserMasterLayout.name,
      });
      browserMasterSlide.addText('Browser master layout', { placeholder: 'browser_title' });
      browserMasterSlide.addRichText([{
        runs: [{ text: 'Browser ', style: { bold: true } }, { text: 'placeholder body' }],
      }], { placeholder: { type: 'body', index: 102 } });
      await masterLayoutDocument.addImage(0, browserMasterPngBytes, {
        contentType: 'image/png',
        placeholder: { type: 'pic', index: 103 },
      });
      await masterLayoutDocument.addChart(0, 'bar', [{
        name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
      }], { placeholder: 'browser_chart' });
      browserMasterSlide.addTable([
        ['Quarter', 'Revenue'],
        ['Q1', '10'],
      ], { placeholder: { type: 'tbl', index: 105 } });
      await masterLayoutDocument.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
        placeholder: 'browser_media',
        poster: browserMasterPngBytes,
        posterContentType: 'image/png',
      });
      const browserMasterSelectedTarget = browserMasterSlide.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget;
      const browserMasterOutput = await masterLayoutDocument.writeBlob();
      const reopenedMasterLayoutDocument = await api.PptxDocument.open(browserMasterOutput);
      await reopenedMasterLayoutDocument.write({ compatibility: 'powerpoint-current' });
      const reopenedBrowserMasterLayout = reopenedMasterLayoutDocument.layouts.find(
        ({ name }) => name === 'BROWSER-MASTER-LAYOUT',
      );
      const reopenedBrowserMasterSlide = reopenedMasterLayoutDocument.slides[0];
      const reopenedBrowserMasterImage = reopenedBrowserMasterSlide.shapes.find(
        (shape) => shape instanceof api.ImageModel,
      );
      const reopenedBrowserMasterChart = reopenedBrowserMasterSlide.shapes.find(
        (shape) => shape instanceof api.ChartModel,
      );
      const reopenedBrowserMasterMedia = reopenedBrowserMasterSlide.shapes.find(
        (shape) => shape instanceof api.MediaModel,
      );
      const hashBytes = async (value) => Array.from(
        new Uint8Array(await crypto.subtle.digest('SHA-256', value)),
        (byte) => byte.toString(16).padStart(2, '0'),
      ).join('');
      const masterLayoutState = {
        mime: browserMasterOutput.type,
        liveWrapperIdentity: browserMasterLiveIdentity,
        layoutNames: reopenedMasterLayoutDocument.layouts.map(({ name }) => name),
        backgroundKinds: {
          master: reopenedMasterLayoutDocument.masters[0].background?.kind,
          layout: reopenedBrowserMasterLayout.background?.kind,
        },
        marginBeforeWrite: browserMasterMargin,
        placeholderTypes: [...api.PLACEHOLDER_TYPES],
        layoutPlaceholders: reopenedBrowserMasterLayout.placeholders.map(
          ({ name, placeholder }) => ({ name, placeholder }),
        ),
        slidePlaceholders: reopenedBrowserMasterSlide.placeholders.map(
          ({ name, kind, placeholder }) => ({ name, kind, placeholder }),
        ),
        placeholderGeometryPreserved: reopenedBrowserMasterSlide.placeholders.every((shape) => {
          const source = reopenedBrowserMasterLayout.placeholders.find(({ placeholder }) =>
            placeholder?.type === shape.placeholder?.type
            && placeholder?.index === shape.placeholder?.index);
          return source !== undefined
            && source.transform.x === shape.transform.x
            && source.transform.y === shape.transform.y
            && source.transform.width === shape.transform.width
            && source.transform.height === shape.transform.height
            && source.transform.rotation === shape.transform.rotation
            && source.transform.flipHorizontal === shape.transform.flipHorizontal
            && source.transform.flipVertical === shape.transform.flipVertical;
        }),
        selectedTargets: [
          browserMasterSelectedTarget,
          reopenedBrowserMasterSlide.relationships.find(
            ({ type }) => type.endsWith('/slideLayout'),
          )?.resolvedTarget,
        ],
        reopenedMargin: reopenedBrowserMasterLayout.margin ?? null,
        payloadHashes: {
          background: await hashBytes(reopenedBrowserMasterLayout.background.bytes),
          image: await hashBytes(reopenedMasterLayoutDocument.opcPackage.requirePart(
            reopenedBrowserMasterImage.sourcePartUri,
          ).bytes),
          media: await hashBytes(reopenedMasterLayoutDocument.opcPackage.requirePart(
            reopenedBrowserMasterMedia.mediaPartUri,
          ).bytes),
        },
        chartDefinition: reopenedBrowserMasterChart.definition.groups.map(
          ({ type, axis, series }) => ({
            type,
            axis,
            series: series.map(({ name, categories, values }) => ({ name, categories, values })),
          }),
        ),
        validationErrors: reopenedMasterLayoutDocument.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const slideDefaultColorDocument = api.PptxDocument.create();
      const slideDefaultColorSource = slideDefaultColorDocument.addSlide();
      slideDefaultColorSource.color = { kind: 'scheme', value: 'accent1' };
      slideDefaultColorSource.addRichText([{
        runs: [
          { text: 'Browser inherited' },
          { text: 'Browser override', style: { color: { kind: 'srgb', value: '00AA00' } } },
          { text: 'Browser alpha', style: { transparency: 25 } },
        ],
      }]);
      const slideDefaultColorDuplicate = slideDefaultColorDocument.duplicateSlide(0);
      const slideDefaultColorDuplicateIdentity =
        slideDefaultColorDuplicate.color === slideDefaultColorSource.color;
      slideDefaultColorDuplicate.addText('Browser duplicate inherited');
      const slideDefaultColorLive = slideDefaultColorDocument.slides.map(({ color }) => color);
      const slideDefaultColorOutput = await slideDefaultColorDocument.writeBlob();
      const reopenedSlideDefaultColors = await api.PptxDocument.open(slideDefaultColorOutput);
      await reopenedSlideDefaultColors.write({ compatibility: 'powerpoint-current' });
      const slideDefaultColorState = {
        mime: slideDefaultColorOutput.type,
        live: slideDefaultColorLive,
        duplicateIdentity: slideDefaultColorDuplicateIdentity,
        materialized: reopenedSlideDefaultColors.slides.map((slide) => slide.shapes
          .filter((shape) => shape instanceof api.ShapeModel)
          .map(({ richText }) => richText.flatMap(({ runs }) => runs.map(({ style }) => ({
            color: style?.color,
            transparency: style?.transparency,
          }))))),
        reopened: reopenedSlideDefaultColors.slides.map(({ color }) => color ?? null),
        validationErrors: reopenedSlideDefaultColors.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const shapeFillDocument = api.PptxDocument.create();
      const shapeFillSlide = shapeFillDocument.addSlide();
      const shapeFillSource = {
        kind: 'solid',
        color: { kind: 'srgb', value: '#AB12CD' },
        transparency: 25,
      };
      const browserNoneShapeFill = shapeFillSlide.addShape('rect', {
        fill: { kind: 'none' },
      });
      const browserSrgbShapeFill = shapeFillSlide.addShape('ellipse', {
        fill: shapeFillSource,
      });
      const browserSchemeShapeFill = shapeFillSlide.addShape('star5', {
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
      });
      const shapeFillImmediate = [
        browserNoneShapeFill.fill,
        browserSrgbShapeFill.fill,
        browserSchemeShapeFill.fill,
      ];
      shapeFillSource.color.value = 'FFFFFF';
      shapeFillSource.transparency = 90;
      const shapeFillDetached = browserSrgbShapeFill.fill;
      const shapeFillPartCount = shapeFillDocument.opcPackage.parts.length;
      const shapeFillRelationshipCount = shapeFillSlide.relationships.length;
      browserSrgbShapeFill.fill = { kind: 'none' };
      const shapeFillReplacedNone = browserSrgbShapeFill.fill;
      browserSrgbShapeFill.fill = undefined;
      const shapeFillCleared = browserSrgbShapeFill.fill;
      browserSrgbShapeFill.fill = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
      };
      const shapeFillXml = new TextDecoder().decode(
        shapeFillDocument.opcPackage.requirePart(shapeFillSlide.partUri).bytes,
      );
      const shapeFillOutput = await shapeFillDocument.writeBlob();
      const reopenedShapeFills = await api.PptxDocument.open(shapeFillOutput);
      const shapeFillState = {
        mime: shapeFillOutput.type,
        immediate: shapeFillImmediate,
        detached: shapeFillDetached,
        replacedNone: shapeFillReplacedNone,
        cleared: shapeFillCleared ?? null,
        edited: browserSrgbShapeFill.fill,
        preserved: browserNoneShapeFill.fill,
        serialized: {
          none: shapeFillXml.includes('<a:noFill/>'),
          edited: shapeFillXml.includes(
            '<a:schemeClr val="accent4"><a:alpha val="50000"/></a:schemeClr>',
          ),
          explicitZero: shapeFillXml.includes(
            '<a:schemeClr val="accent2"><a:alpha val="100000"/></a:schemeClr>',
          ),
        },
        partIsolation: shapeFillDocument.opcPackage.parts.length === shapeFillPartCount,
        relationshipIsolation: shapeFillSlide.relationships.length ===
          shapeFillRelationshipCount,
        reopened: reopenedShapeFills.slides[0].shapes.map(({ fill }) => fill),
        validationErrors: reopenedShapeFills.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const shapeFills = JSON.stringify(shapeFillState) === JSON.stringify({
        mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        immediate: [
          { kind: 'none' },
          {
            kind: 'solid',
            color: { kind: 'srgb', value: 'AB12CD' },
            transparency: 25,
          },
          {
            kind: 'solid',
            color: { kind: 'scheme', value: 'accent2' },
            transparency: 0,
          },
        ],
        detached: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
        },
        replacedNone: { kind: 'none' },
        cleared: null,
        edited: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
        },
        preserved: { kind: 'none' },
        serialized: { none: true, edited: true, explicitZero: true },
        partIsolation: true,
        relationshipIsolation: true,
        reopened: [
          { kind: 'none' },
          {
            kind: 'solid',
            color: { kind: 'scheme', value: 'accent4' },
            transparency: 50,
          },
          {
            kind: 'solid',
            color: { kind: 'scheme', value: 'accent2' },
            transparency: 0,
          },
        ],
        validationErrors: 0,
      });
      const textShapeFillDocument = api.PptxDocument.create();
      const textShapeFillLayout = textShapeFillDocument.layouts[0];
      textShapeFillLayout.addPlaceholder('Browser text fill prompt', {
        name: 'browser_text_fill_placeholder',
        type: 'title',
        index: 190,
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 100,
        },
      });
      const textShapeFillSlide = textShapeFillDocument.addSlide({
        masterName: textShapeFillLayout.name,
      });
      const textShapeFillSource = {
        kind: 'solid',
        color: { kind: 'srgb', value: '#AB12CD' },
        transparency: 25,
      };
      const browserPlainTextFill = textShapeFillSlide.addText('Browser plain text fill', {
        name: 'browser_plain_text_fill',
        fill: textShapeFillSource,
      });
      const browserRichTextFill = textShapeFillSlide.addRichText([{
        runs: [{ text: 'Browser rich text fill' }],
      }], {
        name: 'browser_rich_text_fill',
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
      });
      const browserPopulatedTextFill = textShapeFillSlide.addText(
        'Browser populated text fill',
        { placeholder: 'browser_text_fill_placeholder', fill: { kind: 'none' } },
      );
      const textShapeFillImmediate = [
        browserPlainTextFill.fill,
        browserRichTextFill.fill,
        browserPopulatedTextFill.fill,
      ];
      textShapeFillSource.color.value = 'FFFFFF';
      textShapeFillSource.transparency = 90;
      const textShapeFillDetached = browserPlainTextFill.fill;
      const textShapeFillOutput = await textShapeFillDocument.writeBlob();
      const reopenedTextShapeFills = await api.PptxDocument.open(textShapeFillOutput);
      await reopenedTextShapeFills.write({ compatibility: 'powerpoint-current' });
      const textShapeFillByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeFillState = {
        mime: textShapeFillOutput.type,
        immediate: textShapeFillImmediate,
        detached: textShapeFillDetached,
        reopened: [
          textShapeFillByName(
            reopenedTextShapeFills.slides[0],
            'browser_plain_text_fill',
          ).fill,
          textShapeFillByName(
            reopenedTextShapeFills.slides[0],
            'browser_rich_text_fill',
          ).fill,
          textShapeFillByName(
            reopenedTextShapeFills.slides[0],
            'browser_text_fill_placeholder',
          ).fill,
        ],
        layout: textShapeFillByName(
          reopenedTextShapeFills.layouts[0],
          'browser_text_fill_placeholder',
        ).fill,
        validationErrors: reopenedTextShapeFills.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeLineDocument = api.PptxDocument.create();
      const textShapeLineLayout = textShapeLineDocument.layouts[0];
      textShapeLineLayout.addPlaceholder('Browser text line prompt', {
        name: 'browser_text_line_placeholder',
        type: 'title',
        index: 191,
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 100,
          width: 0,
          dash: 'sysDot',
        },
      });
      const textShapeLineSlide = textShapeLineDocument.addSlide({
        masterName: textShapeLineLayout.name,
      });
      const textShapeLineSource = {
        kind: 'line',
        color: { kind: 'srgb', value: '#AB12CD' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      };
      const browserPlainTextLine = textShapeLineSlide.addText('Browser plain text line', {
        name: 'browser_plain_text_line',
        line: textShapeLineSource,
      });
      const browserRichTextLine = textShapeLineSlide.addRichText([{
        runs: [{ text: 'Browser rich text line' }],
      }], {
        name: 'browser_rich_text_line',
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
      });
      const browserPopulatedTextLine = textShapeLineSlide.addText(
        'Browser populated text line',
        { placeholder: 'browser_text_line_placeholder', line: { kind: 'none' } },
      );
      const textShapeLineImmediate = [
        browserPlainTextLine.line,
        browserRichTextLine.line,
        browserPopulatedTextLine.line,
      ];
      textShapeLineSource.color.value = 'FFFFFF';
      textShapeLineSource.transparency = 90;
      textShapeLineSource.width = 9;
      textShapeLineSource.dash = 'solid';
      const textShapeLineDetached = browserPlainTextLine.line;
      const textShapeLineOutput = await textShapeLineDocument.writeBlob();
      const reopenedTextShapeLines = await api.PptxDocument.open(textShapeLineOutput);
      await reopenedTextShapeLines.write({ compatibility: 'powerpoint-current' });
      const textShapeLineByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeLineState = {
        mime: textShapeLineOutput.type,
        immediate: textShapeLineImmediate,
        detached: textShapeLineDetached,
        reopened: [
          textShapeLineByName(
            reopenedTextShapeLines.slides[0],
            'browser_plain_text_line',
          ).line,
          textShapeLineByName(
            reopenedTextShapeLines.slides[0],
            'browser_rich_text_line',
          ).line,
          textShapeLineByName(
            reopenedTextShapeLines.slides[0],
            'browser_text_line_placeholder',
          ).line,
        ],
        layout: textShapeLineByName(
          reopenedTextShapeLines.layouts[0],
          'browser_text_line_placeholder',
        ).line,
        validationErrors: reopenedTextShapeLines.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeArrowDocument = api.PptxDocument.create();
      const textShapeArrowLayout = textShapeArrowDocument.layouts[0];
      textShapeArrowLayout.addPlaceholder('Browser text arrow prompt', {
        name: 'browser_text_arrow_placeholder',
        type: 'title',
        index: 192,
        arrows: { begin: 'none', end: 'stealth' },
      });
      const textShapeArrowSlide = textShapeArrowDocument.addSlide({
        masterName: textShapeArrowLayout.name,
      });
      const textShapeArrowSource = { begin: 'triangle', end: 'arrow' };
      const browserPlainTextArrow = textShapeArrowSlide.addText('Browser plain text arrow', {
        name: 'browser_plain_text_arrow',
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          width: 2,
          dash: 'dashDot',
        },
        arrows: textShapeArrowSource,
      });
      const browserRichTextArrow = textShapeArrowSlide.addRichText([{
        runs: [{ text: 'Browser rich text arrow' }],
      }], {
        name: 'browser_rich_text_arrow',
        arrows: { end: 'diamond' },
      });
      const browserPopulatedTextArrow = textShapeArrowSlide.addText(
        'Browser populated text arrow',
        { placeholder: 'browser_text_arrow_placeholder', arrows: { begin: 'arrow' } },
      );
      const textShapeArrowImmediate = [
        browserPlainTextArrow.arrows,
        browserRichTextArrow.arrows,
        browserPopulatedTextArrow.arrows,
      ];
      textShapeArrowSource.begin = 'oval';
      textShapeArrowSource.end = 'triangle';
      const textShapeArrowDetached = browserPlainTextArrow.arrows;
      const textShapeArrowOutput = await textShapeArrowDocument.writeBlob();
      const reopenedTextShapeArrows = await api.PptxDocument.open(textShapeArrowOutput);
      await reopenedTextShapeArrows.write({ compatibility: 'powerpoint-current' });
      const textShapeArrowByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeArrowState = {
        mime: textShapeArrowOutput.type,
        immediate: textShapeArrowImmediate,
        detached: textShapeArrowDetached,
        reopened: [
          textShapeArrowByName(
            reopenedTextShapeArrows.slides[0],
            'browser_plain_text_arrow',
          ).arrows,
          textShapeArrowByName(
            reopenedTextShapeArrows.slides[0],
            'browser_rich_text_arrow',
          ).arrows,
          textShapeArrowByName(
            reopenedTextShapeArrows.slides[0],
            'browser_text_arrow_placeholder',
          ).arrows,
        ],
        layout: textShapeArrowByName(
          reopenedTextShapeArrows.layouts[0],
          'browser_text_arrow_placeholder',
        ).arrows,
        line: textShapeArrowByName(
          reopenedTextShapeArrows.slides[0],
          'browser_plain_text_arrow',
        ).line,
        validationErrors: reopenedTextShapeArrows.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeShadowDocument = api.PptxDocument.create();
      const textShapeShadowLayout = textShapeShadowDocument.layouts[0];
      textShapeShadowLayout.addPlaceholder('Browser text shadow prompt', {
        name: 'browser_text_shadow_placeholder',
        type: 'title',
        index: 193,
        shadow: {
          kind: 'outer',
          color: { kind: 'scheme', value: 'accent1' },
          rotateWithShape: true,
        },
      });
      const textShapeShadowSlide = textShapeShadowDocument.addSlide({
        masterName: textShapeShadowLayout.name,
      });
      const textShapeShadowColor = { kind: 'scheme', value: 'accent4' };
      const textShapeShadowSource = {
        kind: 'outer',
        color: textShapeShadowColor,
        opacity: 0.4,
        blur: 2,
        angle: 45,
        distance: 3,
        rotateWithShape: true,
      };
      const browserPlainTextShadow = textShapeShadowSlide.addText(
        'Browser plain text shadow',
        {
          name: 'browser_plain_text_shadow',
          line: {
            kind: 'line',
            color: { kind: 'scheme', value: 'accent2' },
            width: 2,
            dash: 'dashDot',
          },
          arrows: { begin: 'triangle', end: 'arrow' },
          shadow: textShapeShadowSource,
        },
      );
      const browserRichTextShadow = textShapeShadowSlide.addRichText([{
        runs: [{ text: 'Browser rich text shadow' }],
      }], {
        name: 'browser_rich_text_shadow',
        shadow: {
          kind: 'inner',
          color: { kind: 'srgb', value: '667788' },
          opacity: 0,
          blur: 0,
          angle: 0,
          distance: 0,
        },
      });
      const browserPopulatedTextShadow = textShapeShadowSlide.addText(
        'Browser populated text shadow',
        { placeholder: 'browser_text_shadow_placeholder', shadow: { kind: 'outer' } },
      );
      const textShapeShadowImmediate = [
        browserPlainTextShadow.shadow,
        browserRichTextShadow.shadow,
        browserPopulatedTextShadow.shadow,
      ];
      textShapeShadowColor.value = 'accent6';
      textShapeShadowSource.opacity = 0.9;
      textShapeShadowSource.rotateWithShape = false;
      const textShapeShadowDetached = browserPlainTextShadow.shadow;
      const textShapeShadowOutput = await textShapeShadowDocument.writeBlob();
      const reopenedTextShapeShadows = await api.PptxDocument.open(textShapeShadowOutput);
      await reopenedTextShapeShadows.write({ compatibility: 'powerpoint-current' });
      const textShapeShadowByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeShadowState = {
        mime: textShapeShadowOutput.type,
        immediate: textShapeShadowImmediate,
        detached: textShapeShadowDetached,
        reopened: [
          textShapeShadowByName(
            reopenedTextShapeShadows.slides[0],
            'browser_plain_text_shadow',
          ).shadow,
          textShapeShadowByName(
            reopenedTextShapeShadows.slides[0],
            'browser_rich_text_shadow',
          ).shadow,
          textShapeShadowByName(
            reopenedTextShapeShadows.slides[0],
            'browser_text_shadow_placeholder',
          ).shadow,
        ],
        layout: textShapeShadowByName(
          reopenedTextShapeShadows.layouts[0],
          'browser_text_shadow_placeholder',
        ).shadow,
        line: textShapeShadowByName(
          reopenedTextShapeShadows.slides[0],
          'browser_plain_text_shadow',
        ).line,
        arrows: textShapeShadowByName(
          reopenedTextShapeShadows.slides[0],
          'browser_plain_text_shadow',
        ).arrows,
        validationErrors: reopenedTextShapeShadows.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapePresetGeometryDocument = api.PptxDocument.create();
      const textShapePresetGeometryLayout = textShapePresetGeometryDocument.layouts[0];
      const browserTextPresetGeometryPlaceholder =
        textShapePresetGeometryLayout.addPlaceholder('Browser geometry prompt', {
          name: 'browser_text_geometry_placeholder',
          type: 'title',
          index: 198,
          shape: 'foldedCorner',
        });
      const textShapePresetGeometrySlide = textShapePresetGeometryDocument.addSlide({
        masterName: textShapePresetGeometryLayout.name,
      });
      const browserDefaultTextPresetGeometry = textShapePresetGeometrySlide.addText(
        'Browser default geometry',
        { name: 'browser_default_text_geometry' },
      );
      const browserEllipseTextPresetGeometry = textShapePresetGeometrySlide.addText(
        'Browser ellipse geometry',
        { name: 'browser_ellipse_text_geometry', shape: 'ellipse' },
      );
      const browserRichTextPresetGeometry = textShapePresetGeometrySlide.addRichText([{
        runs: [{ text: 'Browser rich ' }, { text: 'line', style: { bold: true } }],
      }], { name: 'browser_rich_text_geometry', shape: 'line' });
      const browserPopulatedTextPresetGeometry = textShapePresetGeometrySlide.addText(
        'Browser populated geometry',
        { placeholder: 'browser_text_geometry_placeholder', shape: 'roundRect' },
      );
      const textShapePresetGeometryImmediate = [
        browserDefaultTextPresetGeometry.presetType,
        browserEllipseTextPresetGeometry.presetType,
        browserRichTextPresetGeometry.presetType,
        browserPopulatedTextPresetGeometry.presetType,
        browserTextPresetGeometryPlaceholder.presetType,
      ];
      browserEllipseTextPresetGeometry.presetType = 'star5';
      const textShapePresetGeometryOutput =
        await textShapePresetGeometryDocument.writeBlob();
      const reopenedTextShapePresetGeometry = await api.PptxDocument.open(
        textShapePresetGeometryOutput,
      );
      await reopenedTextShapePresetGeometry.write({
        compatibility: 'powerpoint-current',
      });
      const textShapePresetGeometryByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapePresetGeometryState = {
        mime: textShapePresetGeometryOutput.type,
        immediate: textShapePresetGeometryImmediate,
        edited: browserEllipseTextPresetGeometry.presetType,
        reopened: [
          textShapePresetGeometryByName(
            reopenedTextShapePresetGeometry.slides[0],
            'browser_default_text_geometry',
          ).presetType,
          textShapePresetGeometryByName(
            reopenedTextShapePresetGeometry.slides[0],
            'browser_ellipse_text_geometry',
          ).presetType,
          textShapePresetGeometryByName(
            reopenedTextShapePresetGeometry.slides[0],
            'browser_rich_text_geometry',
          ).presetType,
          textShapePresetGeometryByName(
            reopenedTextShapePresetGeometry.slides[0],
            'browser_text_geometry_placeholder',
          ).presetType,
        ],
        layout: textShapePresetGeometryByName(
          reopenedTextShapePresetGeometry.layouts[0],
          'browser_text_geometry_placeholder',
        ).presetType,
        texts: reopenedTextShapePresetGeometry.slides[0].shapes.map(({ text }) => text),
        validationErrors: reopenedTextShapePresetGeometry.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapePresetGeometry =
        textShapePresetGeometryState.mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && textShapePresetGeometryState.immediate.join(',') ===
          'rect,ellipse,line,roundRect,foldedCorner'
        && textShapePresetGeometryState.edited === 'star5'
        && textShapePresetGeometryState.reopened.join(',') ===
          'rect,star5,line,roundRect'
        && textShapePresetGeometryState.layout === 'foldedCorner'
        && textShapePresetGeometryState.validationErrors === 0;
      const textShapeRectRadiusDocument = api.PptxDocument.create();
      const textShapeRectRadiusLayout = textShapeRectRadiusDocument.layouts[0];
      const browserTextRectRadiusPlaceholder = textShapeRectRadiusLayout.addPlaceholder(
        'Browser rounded prompt',
        {
          name: 'browser_text_rect_radius_placeholder',
          type: 'title',
          index: 397,
          shape: 'roundRect',
          rectRadius: api.inches(0.25),
          width: api.inches(4),
          height: api.inches(2),
        },
      );
      const textShapeRectRadiusSlide = textShapeRectRadiusDocument.addSlide({
        masterName: textShapeRectRadiusLayout.name,
      });
      const browserOmittedTextRectRadius = textShapeRectRadiusSlide.addText(
        'Browser omitted radius',
        { name: 'browser_omitted_text_rect_radius', shape: 'roundRect' },
      );
      const browserZeroTextRectRadius = textShapeRectRadiusSlide.addText(
        'Browser zero radius',
        {
          name: 'browser_zero_text_rect_radius',
          shape: 'roundRect',
          rectRadius: api.inches(0),
          width: api.inches(2),
          height: api.inches(1),
        },
      );
      const browserTwoByOneTextRectRadius = textShapeRectRadiusSlide.addText(
        'Browser two by one radius',
        {
          name: 'browser_two_by_one_text_rect_radius',
          shape: 'roundRect',
          rectRadius: api.inches(0.5),
          width: api.inches(2),
          height: api.inches(1),
        },
      );
      const browserRichTextRectRadius = textShapeRectRadiusSlide.addRichText([{
        runs: [{ text: 'Browser rich ' }, { text: 'radius', style: { bold: true } }],
      }], {
        name: 'browser_rich_text_rect_radius',
        shape: 'roundRect',
        rectRadius: api.inches(0.5),
        width: api.inches(4),
        height: api.inches(2),
      });
      const browserPopulatedTextRectRadius = textShapeRectRadiusSlide.addText(
        'Browser populated radius',
        {
          placeholder: 'browser_text_rect_radius_placeholder',
          shape: 'roundRect',
          rectRadius: api.inches(0.5),
          width: api.inches(1),
          height: api.inches(1),
        },
      );
      const textShapeRectRadiusImmediate = [
        browserOmittedTextRectRadius.adjustments,
        browserZeroTextRectRadius.adjustments,
        browserTwoByOneTextRectRadius.adjustments,
        browserRichTextRectRadius.adjustments,
        browserPopulatedTextRectRadius.adjustments,
        browserTextRectRadiusPlaceholder.adjustments,
      ];
      browserTwoByOneTextRectRadius.setTransform({
        width: api.inches(4),
        height: api.inches(2),
      });
      browserRichTextRectRadius.adjustments = [{ name: 'adj', value: 12_500 }];
      browserOmittedTextRectRadius.adjustments = [{ name: 'adj', value: 75_000 }];
      browserOmittedTextRectRadius.adjustments = [];
      const textShapeRectRadiusOutput = await textShapeRectRadiusDocument.writeBlob();
      const reopenedTextShapeRectRadius = await api.PptxDocument.open(
        textShapeRectRadiusOutput,
      );
      await reopenedTextShapeRectRadius.write({ compatibility: 'powerpoint-current' });
      const textShapeRectRadiusByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeRectRadiusState = {
        mime: textShapeRectRadiusOutput.type,
        immediate: textShapeRectRadiusImmediate,
        resizeStable: browserTwoByOneTextRectRadius.adjustments,
        edited: browserRichTextRectRadius.adjustments,
        cleared: browserOmittedTextRectRadius.adjustments,
        reopened: [
          textShapeRectRadiusByName(
            reopenedTextShapeRectRadius.slides[0],
            'browser_omitted_text_rect_radius',
          ).adjustments,
          textShapeRectRadiusByName(
            reopenedTextShapeRectRadius.slides[0],
            'browser_zero_text_rect_radius',
          ).adjustments,
          textShapeRectRadiusByName(
            reopenedTextShapeRectRadius.slides[0],
            'browser_two_by_one_text_rect_radius',
          ).adjustments,
          textShapeRectRadiusByName(
            reopenedTextShapeRectRadius.slides[0],
            'browser_rich_text_rect_radius',
          ).adjustments,
          textShapeRectRadiusByName(
            reopenedTextShapeRectRadius.slides[0],
            'browser_text_rect_radius_placeholder',
          ).adjustments,
        ],
        layout: textShapeRectRadiusByName(
          reopenedTextShapeRectRadius.layouts[0],
          'browser_text_rect_radius_placeholder',
        ).adjustments,
        validationErrors: reopenedTextShapeRectRadius.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeRectRadius =
        textShapeRectRadiusState.mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && JSON.stringify(textShapeRectRadiusState.immediate) === JSON.stringify([
          [],
          [{ name: 'adj', value: 0 }],
          [{ name: 'adj', value: 50_000 }],
          [{ name: 'adj', value: 25_000 }],
          [{ name: 'adj', value: 25_000 }],
          [{ name: 'adj', value: 12_500 }],
        ])
        && textShapeRectRadiusState.resizeStable?.[0]?.value === 50_000
        && textShapeRectRadiusState.edited?.[0]?.value === 12_500
        && textShapeRectRadiusState.cleared?.length === 0
        && JSON.stringify(textShapeRectRadiusState.reopened) === JSON.stringify([
          [],
          [{ name: 'adj', value: 0 }],
          [{ name: 'adj', value: 50_000 }],
          [{ name: 'adj', value: 12_500 }],
          [{ name: 'adj', value: 25_000 }],
        ])
        && textShapeRectRadiusState.layout?.[0]?.value === 12_500
        && textShapeRectRadiusState.validationErrors === 0;
      const textShapeIsTextBoxDocument = api.PptxDocument.create();
      const textShapeIsTextBoxLayout = textShapeIsTextBoxDocument.layouts[0];
      const textShapeIsTextBoxMaster = textShapeIsTextBoxDocument.masters[0];
      const browserLayoutShapeText = textShapeIsTextBoxLayout.addText(
        'Browser layout shape text',
        { name: 'browser_layout_shape_text', isTextBox: false },
      );
      const browserMasterTextBox = textShapeIsTextBoxMaster.addRichText([{
        runs: [{ text: 'Browser master text box' }],
      }], {
        name: 'browser_master_text_box',
        isTextBox: true,
      });
      const browserFalseTextBoxSource = textShapeIsTextBoxLayout.addPlaceholder(
        'Browser false text box prompt',
        {
          name: 'browser_false_text_box_source',
          type: 'title',
          index: 400,
          isTextBox: false,
        },
      );
      const browserTrueTextBoxSource = textShapeIsTextBoxLayout.addPlaceholder(
        'Browser true text box prompt',
        {
          name: 'browser_true_text_box_source',
          type: 'body',
          index: 401,
          isTextBox: true,
        },
      );
      const textShapeIsTextBoxSlide = textShapeIsTextBoxDocument.addSlide({
        masterName: textShapeIsTextBoxLayout.name,
      });
      const textShapeIsTextBoxMaterialized = textShapeIsTextBoxSlide.placeholders.map(
        ({ isTextBox }) => isTextBox,
      );
      const browserPlainTextBox = textShapeIsTextBoxSlide.addText(
        'Browser plain shape text',
        { name: 'browser_plain_text_box' },
      );
      const browserRichTextBox = textShapeIsTextBoxSlide.addRichText([{
        runs: [{ text: 'Browser rich text box', style: { bold: true } }],
      }], {
        name: 'browser_rich_text_box',
        isTextBox: true,
      });
      const browserPopulatedFalseTextBox = textShapeIsTextBoxSlide.addText(
        'Browser population keeps false source',
        { placeholder: browserFalseTextBoxSource.name, isTextBox: true },
      );
      const browserPopulatedTrueTextBox = textShapeIsTextBoxSlide.addText(
        'Browser population keeps true source',
        { placeholder: browserTrueTextBoxSource.name, isTextBox: false },
      );
      const textShapeIsTextBoxImmediate = [
        browserLayoutShapeText.isTextBox,
        browserMasterTextBox.isTextBox,
        browserFalseTextBoxSource.isTextBox,
        browserTrueTextBoxSource.isTextBox,
        browserPlainTextBox.isTextBox,
        browserRichTextBox.isTextBox,
        browserPopulatedFalseTextBox.isTextBox,
        browserPopulatedTrueTextBox.isTextBox,
      ];
      const textShapeIsTextBoxNoOpBytes = textShapeIsTextBoxDocument.opcPackage
        .requirePart(textShapeIsTextBoxSlide.partUri).bytes.slice();
      const textShapeIsTextBoxNoOpJournal =
        textShapeIsTextBoxDocument.opcPackage.mutations.length;
      browserRichTextBox.isTextBox = true;
      const textShapeIsTextBoxNoOpCurrent = textShapeIsTextBoxDocument.opcPackage
        .requirePart(textShapeIsTextBoxSlide.partUri).bytes;
      const textShapeIsTextBoxNoOp =
        textShapeIsTextBoxNoOpJournal ===
          textShapeIsTextBoxDocument.opcPackage.mutations.length
        && textShapeIsTextBoxNoOpBytes.length === textShapeIsTextBoxNoOpCurrent.length
        && textShapeIsTextBoxNoOpBytes.every(
          (value, index) => value === textShapeIsTextBoxNoOpCurrent[index],
        );
      browserPlainTextBox.isTextBox = true;
      browserRichTextBox.isTextBox = false;
      const browserDeclarativeTextBoxLayout =
        await textShapeIsTextBoxDocument.defineSlideMaster({
          title: 'BROWSER-TEXT-BOX-STATE',
          objects: [
            {
              kind: 'text',
              text: 'Browser declarative text box',
              options: { name: 'browser_declarative_text_box', isTextBox: true },
            },
            {
              kind: 'placeholder',
              text: 'Browser declarative shape prompt',
              options: {
                name: 'browser_declarative_shape_prompt',
                type: 'title',
                index: 402,
                isTextBox: false,
              },
            },
            {
              kind: 'placeholder',
              text: 'Browser declarative text box prompt',
              options: {
                name: 'browser_declarative_text_box_prompt',
                type: 'body',
                index: 403,
                isTextBox: true,
              },
            },
          ],
        });
      const browserDeclarativeTextBoxSlide = textShapeIsTextBoxDocument.addSlide({
        masterName: browserDeclarativeTextBoxLayout.name,
      });
      const textShapeIsTextBoxDeclarative = [
        browserDeclarativeTextBoxLayout.shapes.find(
          ({ name }) => name === 'browser_declarative_text_box',
        ).isTextBox,
        ...browserDeclarativeTextBoxLayout.placeholders.map(({ isTextBox }) => isTextBox),
        ...browserDeclarativeTextBoxSlide.placeholders.map(({ isTextBox }) => isTextBox),
      ];
      const duplicateTextShapeIsTextBoxSlide = textShapeIsTextBoxDocument.duplicateSlide(
        textShapeIsTextBoxDocument.slides.indexOf(textShapeIsTextBoxSlide),
      );
      const duplicatePlainTextBox = duplicateTextShapeIsTextBoxSlide.shapes.find(
        ({ name }) => name === browserPlainTextBox.name,
      );
      duplicatePlainTextBox.isTextBox = false;
      const textShapeIsTextBoxDuplicate = [
        browserPlainTextBox.isTextBox,
        duplicatePlainTextBox.isTextBox,
      ];
      const textShapeIsTextBoxOutput = await textShapeIsTextBoxDocument.writeBlob();
      const reopenedTextShapeIsTextBox = await api.PptxDocument.open(
        textShapeIsTextBoxOutput,
      );
      await reopenedTextShapeIsTextBox.write({ compatibility: 'powerpoint-current' });
      const textShapeIsTextBoxByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const reopenedTextShapeIsTextBoxSlide = reopenedTextShapeIsTextBox.slides.find(
        ({ partUri }) => partUri === textShapeIsTextBoxSlide.partUri,
      );
      const reopenedTextShapeIsTextBoxLayout = reopenedTextShapeIsTextBox.layouts.find(
        ({ partUri }) => partUri === textShapeIsTextBoxLayout.partUri,
      );
      const reopenedTextShapeIsTextBoxMaster = reopenedTextShapeIsTextBox.masters.find(
        ({ partUri }) => partUri === textShapeIsTextBoxMaster.partUri,
      );
      const reopenedDeclarativeTextBoxLayout = reopenedTextShapeIsTextBox.layouts.find(
        ({ name }) => name === browserDeclarativeTextBoxLayout.name,
      );
      const textShapeIsTextBoxState = {
        mime: textShapeIsTextBoxOutput.type,
        immediate: textShapeIsTextBoxImmediate,
        materialized: textShapeIsTextBoxMaterialized,
        noOp: textShapeIsTextBoxNoOp,
        edited: [browserPlainTextBox.isTextBox, browserRichTextBox.isTextBox],
        duplicate: textShapeIsTextBoxDuplicate,
        declarative: textShapeIsTextBoxDeclarative,
        reopened: [
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxSlide,
            browserPlainTextBox.name,
          ).isTextBox,
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxSlide,
            browserRichTextBox.name,
          ).isTextBox,
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxSlide,
            browserFalseTextBoxSource.name,
          ).isTextBox,
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxSlide,
            browserTrueTextBoxSource.name,
          ).isTextBox,
        ],
        layout: [
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxLayout,
            browserLayoutShapeText.name,
          ).isTextBox,
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxLayout,
            browserFalseTextBoxSource.name,
          ).isTextBox,
          textShapeIsTextBoxByName(
            reopenedTextShapeIsTextBoxLayout,
            browserTrueTextBoxSource.name,
          ).isTextBox,
        ],
        master: textShapeIsTextBoxByName(
          reopenedTextShapeIsTextBoxMaster,
          browserMasterTextBox.name,
        ).isTextBox,
        reopenedDeclarative: [
          textShapeIsTextBoxByName(
            reopenedDeclarativeTextBoxLayout,
            'browser_declarative_text_box',
          ).isTextBox,
          ...reopenedDeclarativeTextBoxLayout.placeholders.map(({ isTextBox }) => isTextBox),
        ],
        validationErrors: reopenedTextShapeIsTextBox.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeIsTextBox =
        textShapeIsTextBoxState.mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && JSON.stringify(textShapeIsTextBoxState.immediate) ===
          JSON.stringify([false, true, false, true, false, true, false, true])
        && JSON.stringify(textShapeIsTextBoxState.materialized) ===
          JSON.stringify([false, true])
        && textShapeIsTextBoxState.noOp
        && JSON.stringify(textShapeIsTextBoxState.edited) === JSON.stringify([true, false])
        && JSON.stringify(textShapeIsTextBoxState.duplicate) === JSON.stringify([true, false])
        && JSON.stringify(textShapeIsTextBoxState.declarative) ===
          JSON.stringify([true, false, true, false, true])
        && JSON.stringify(textShapeIsTextBoxState.reopened) ===
          JSON.stringify([true, false, false, true])
        && JSON.stringify(textShapeIsTextBoxState.layout) ===
          JSON.stringify([false, false, true])
        && textShapeIsTextBoxState.master === true
        && JSON.stringify(textShapeIsTextBoxState.reopenedDeclarative) ===
          JSON.stringify([true, false, true])
        && textShapeIsTextBoxState.validationErrors === 0;
      const richTextBreakLineDocument = api.PptxDocument.create();
      const richTextBreakLineLayout = richTextBreakLineDocument.layouts[0];
      const richTextBreakLineMaster = richTextBreakLineDocument.masters[0];
      const browserBreakLineInput = (prefix, hyperlink) => [{
        align: 'center',
        marginLeft: 12,
        spacing: { before: 3, after: 5 },
        tabStops: [{ position: 1.5, alignment: 'right' }],
        runs: [
          {
            text: `${prefix} first`,
            breakLine: true,
            ...(hyperlink === undefined ? {} : { style: { hyperlink } }),
          },
          { text: '', breakLine: true },
          { text: `${prefix} soft`, softBreakBefore: true },
          { text: `${prefix} trailing`, breakLine: true },
        ],
      }];
      const browserBreakLineParagraphs = (shape) => shape.richText.map(({ runs }) =>
        runs.map(({ text }) => text));
      const browserLayoutBreakLine = richTextBreakLineLayout.addRichText(
        browserBreakLineInput('Browser layout'),
        { name: 'browser_layout_break_line' },
      );
      const browserMasterBreakLine = richTextBreakLineMaster.addRichText(
        browserBreakLineInput('Browser master'),
        { name: 'browser_master_break_line' },
      );
      const browserLayoutBreakLinePrompt = richTextBreakLineLayout.addPlaceholder(
        browserBreakLineInput('Browser layout prompt'),
        {
          name: 'browser_layout_break_line_prompt',
          type: 'body',
          index: 410,
        },
      );
      const browserMasterBreakLinePrompt = richTextBreakLineMaster.addPlaceholder(
        browserBreakLineInput('Browser master prompt'),
        {
          name: 'browser_master_break_line_prompt',
          type: 'body',
          index: 411,
        },
      );
      const richTextBreakLineSource = richTextBreakLineDocument.addSlide({
        masterName: richTextBreakLineLayout.name,
      });
      const richTextBreakLineTarget = richTextBreakLineDocument.addSlide({
        masterName: richTextBreakLineLayout.name,
      });
      const browserMaterializedBreakLine = browserBreakLineParagraphs(
        richTextBreakLineSource.placeholders.find(
          ({ name }) => name === browserLayoutBreakLinePrompt.name,
        ),
      );
      const richTextBreakLineTargetIndex =
        richTextBreakLineDocument.slides.indexOf(richTextBreakLineTarget) + 1;
      const browserDirectBreakLineInput = browserBreakLineInput(
        'Browser direct',
        { slide: richTextBreakLineTargetIndex, tooltip: '' },
      );
      const browserDirectBreakLine = richTextBreakLineSource.addRichText(
        browserDirectBreakLineInput,
        { name: 'browser_break_line_source' },
      );
      browserDirectBreakLineInput[0].runs[0].text = 'Changed browser caller text';
      browserDirectBreakLineInput[0].runs[0].breakLine = false;
      const browserPopulatedBreakLine = richTextBreakLineSource.addRichText(
        browserBreakLineInput('Browser populated'),
        { placeholder: browserLayoutBreakLinePrompt.name },
      );
      const browserEditedBreakLine = richTextBreakLineSource.addRichText([{
        runs: [{ text: 'Browser before edit' }],
      }], { name: 'browser_break_line_edited' });
      browserEditedBreakLine.richText = browserBreakLineInput('Browser edited');
      const browserDeclarativeBreakLineLayout =
        await richTextBreakLineDocument.defineSlideMaster({
          title: 'BROWSER-RICH-TEXT-BREAK-LINE',
          objects: [
            {
              kind: 'text',
              text: browserBreakLineInput('Browser declarative'),
              options: { name: 'browser_declarative_break_line' },
            },
            {
              kind: 'placeholder',
              text: browserBreakLineInput('Browser declarative prompt'),
              options: {
                name: 'browser_declarative_break_line_prompt',
                type: 'body',
                index: 412,
              },
            },
          ],
        });
      const browserDeclarativeBreakLineSlide = richTextBreakLineDocument.addSlide({
        masterName: browserDeclarativeBreakLineLayout.name,
      });
      const browserDeclarativePopulatedBreakLine =
        browserDeclarativeBreakLineSlide.addRichText(
          browserBreakLineInput('Browser declarative populated'),
          { placeholder: 'browser_declarative_break_line_prompt' },
        );
      const browserBreakLineDuplicate = richTextBreakLineDocument.duplicateSlide(
        richTextBreakLineDocument.slides.indexOf(richTextBreakLineSource),
      );
      const browserDuplicateBreakLineShape = browserBreakLineDuplicate.shapes.find(
        ({ name }) => name === browserDirectBreakLine.name,
      );
      browserDuplicateBreakLineShape.richText = [{
        runs: [{ text: 'Browser duplicate only' }],
      }];
      const richTextBreakLineSourceXml = new TextDecoder().decode(
        richTextBreakLineDocument.opcPackage
          .requirePart(richTextBreakLineSource.partUri).bytes,
      );
      const browserBreakLineIdOffset = richTextBreakLineSourceXml.indexOf(
        `<p:cNvPr id="${browserDirectBreakLine.id}"`,
      );
      const browserBreakLineShapeStart = richTextBreakLineSourceXml.lastIndexOf(
        '<p:sp',
        browserBreakLineIdOffset,
      );
      const browserBreakLineShapeEnd = richTextBreakLineSourceXml.indexOf(
        '</p:sp>',
        browserBreakLineIdOffset,
      );
      const browserBreakLineShapeXml = richTextBreakLineSourceXml.slice(
        browserBreakLineShapeStart,
        browserBreakLineShapeEnd + '</p:sp>'.length,
      );
      const browserBreakLineParagraphXml = browserBreakLineShapeXml.match(
        /<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g,
      ) ?? [];
      const richTextBreakLineOutput = await richTextBreakLineDocument.writeBlob();
      const reopenedRichTextBreakLine = await api.PptxDocument.open(
        richTextBreakLineOutput,
      );
      await reopenedRichTextBreakLine.write({ compatibility: 'powerpoint-current' });
      const browserBreakLineByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const reopenedRichTextBreakLineSource = reopenedRichTextBreakLine.slides.find(
        ({ partUri }) => partUri === richTextBreakLineSource.partUri,
      );
      const reopenedRichTextBreakLineDuplicate = reopenedRichTextBreakLine.slides.find(
        ({ partUri }) => partUri === browserBreakLineDuplicate.partUri,
      );
      const reopenedRichTextBreakLineLayout = reopenedRichTextBreakLine.layouts.find(
        ({ partUri }) => partUri === richTextBreakLineLayout.partUri,
      );
      const reopenedRichTextBreakLineMaster = reopenedRichTextBreakLine.masters.find(
        ({ partUri }) => partUri === richTextBreakLineMaster.partUri,
      );
      const reopenedDeclarativeBreakLineLayout = reopenedRichTextBreakLine.layouts.find(
        ({ name }) => name === browserDeclarativeBreakLineLayout.name,
      );
      const richTextBreakLineState = {
        mime: richTextBreakLineOutput.type,
        direct: browserBreakLineParagraphs(browserDirectBreakLine),
        materialized: browserMaterializedBreakLine,
        owners: [
          browserBreakLineParagraphs(browserLayoutBreakLine),
          browserBreakLineParagraphs(browserMasterBreakLine),
          browserBreakLineParagraphs(browserLayoutBreakLinePrompt),
          browserBreakLineParagraphs(browserMasterBreakLinePrompt),
        ],
        populated: browserBreakLineParagraphs(browserPopulatedBreakLine),
        edited: browserBreakLineParagraphs(browserEditedBreakLine),
        duplicate: [
          browserBreakLineParagraphs(browserDirectBreakLine),
          browserBreakLineParagraphs(browserDuplicateBreakLineShape),
        ],
        declarative: [
          browserBreakLineParagraphs(browserBreakLineByName(
            browserDeclarativeBreakLineLayout,
            'browser_declarative_break_line',
          )),
          browserBreakLineParagraphs(browserBreakLineByName(
            browserDeclarativeBreakLineLayout,
            'browser_declarative_break_line_prompt',
          )),
          browserBreakLineParagraphs(browserDeclarativePopulatedBreakLine),
        ],
        reopened: [
          browserBreakLineParagraphs(browserBreakLineByName(
            reopenedRichTextBreakLineSource,
            browserDirectBreakLine.name,
          )),
          browserBreakLineParagraphs(browserBreakLineByName(
            reopenedRichTextBreakLineDuplicate,
            browserDirectBreakLine.name,
          )),
          browserBreakLineParagraphs(browserBreakLineByName(
            reopenedRichTextBreakLineLayout,
            browserLayoutBreakLine.name,
          )),
          browserBreakLineParagraphs(browserBreakLineByName(
            reopenedRichTextBreakLineMaster,
            browserMasterBreakLine.name,
          )),
          browserBreakLineParagraphs(browserBreakLineByName(
            reopenedDeclarativeBreakLineLayout,
            'browser_declarative_break_line',
          )),
        ],
        properties: browserDirectBreakLine.richText.every(
          ({ align, marginLeft, spacing, tabStops }) =>
            align === 'center' && marginLeft === 12 && spacing?.before === 3 &&
            spacing.after === 5 && tabStops?.[0]?.position === 1.5 &&
            tabStops[0].alignment === 'right',
        ),
        canonical: browserDirectBreakLine.richText.every(({ runs }) =>
          runs.every((run) => !Object.hasOwn(run, 'breakLine'))),
        softBreak: browserDirectBreakLine.richText[2].runs[0].softBreakBefore === true,
        relationshipTarget: richTextBreakLineSource.relationships.some(
          ({ resolvedTarget }) => resolvedTarget === richTextBreakLineTarget.partUri,
        ),
        xml: {
          paragraphCount: browserBreakLineParagraphXml.length,
          emptyParagraph: !browserBreakLineParagraphXml[1]?.includes('<a:r>') &&
            !/<a:t(?:\s|>)/.test(browserBreakLineParagraphXml[1] ?? ''),
          softBreak: browserBreakLineParagraphXml[2]?.includes('<a:br/>') === true,
          privateMarker: browserBreakLineShapeXml.includes('breakLine'),
        },
        validationErrors: reopenedRichTextBreakLine.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
        validationWarnings: reopenedRichTextBreakLine.diagnostics.filter(
          ({ severity }) => severity === 'warning',
        ).length,
      };
      const richTextBreakLine =
        richTextBreakLineState.mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        && richTextBreakLineState.direct.length === 3
        && richTextBreakLineState.direct[1].length === 0
        && richTextBreakLineState.direct[2].length === 2
        && JSON.stringify(richTextBreakLineState.materialized) ===
          JSON.stringify([[]])
        && richTextBreakLineState.owners.every((paragraphs) => paragraphs.length === 3)
        && richTextBreakLineState.populated.length === 3
        && richTextBreakLineState.edited.length === 3
        && richTextBreakLineState.duplicate[0].length === 3
        && richTextBreakLineState.duplicate[1][0][0] === 'Browser duplicate only'
        && richTextBreakLineState.declarative.every(
          (paragraphs) => paragraphs.length === 3,
        )
        && richTextBreakLineState.reopened[0].length === 3
        && richTextBreakLineState.reopened[1][0][0] === 'Browser duplicate only'
        && richTextBreakLineState.reopened.slice(2).every(
          (paragraphs) => paragraphs.length === 3,
        )
        && richTextBreakLineState.properties
        && richTextBreakLineState.canonical
        && richTextBreakLineState.softBreak
        && richTextBreakLineState.relationshipTarget
        && richTextBreakLineState.xml.paragraphCount === 3
        && richTextBreakLineState.xml.emptyParagraph
        && richTextBreakLineState.xml.softBreak
        && !richTextBreakLineState.xml.privateMarker
        && richTextBreakLineState.validationErrors === 0
        && richTextBreakLineState.validationWarnings === 0;
      const textShapeHyperlinkDocument = api.PptxDocument.create();
      const textShapeHyperlinkLayout = textShapeHyperlinkDocument.layouts[0];
      textShapeHyperlinkLayout.addPlaceholder('Browser text hyperlink prompt', {
        name: 'browser_text_hyperlink_placeholder',
        type: 'title',
        index: 194,
        hyperlink: { url: 'https://layout-browser.example', tooltip: '' },
      });
      const textShapeHyperlinkSlide = textShapeHyperlinkDocument.addSlide({
        masterName: textShapeHyperlinkLayout.name,
      });
      const textShapeHyperlinkTargetLayout = await textShapeHyperlinkDocument.defineSlideMaster({
        title: 'BROWSER-TEXT-HYPERLINK-TARGET',
        objects: [],
      });
      textShapeHyperlinkDocument.addSlide({
        masterName: textShapeHyperlinkTargetLayout.name,
      });
      const textShapeHyperlinkInput = {
        url: 'https://browser-text.example/path?a=1&b=2',
        tooltip: 'Browser & text',
      };
      const browserPlainTextHyperlink = textShapeHyperlinkSlide.addText(
        'Browser plain text hyperlink\nSecond line',
        {
          name: 'browser_plain_text_hyperlink',
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
          line: {
            kind: 'line',
            color: { kind: 'scheme', value: 'accent2' },
            width: 2,
            dash: 'dashDot',
          },
          arrows: { begin: 'triangle', end: 'arrow' },
          shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent4' } },
          hyperlink: textShapeHyperlinkInput,
        },
      );
      const browserRichTextHyperlink = textShapeHyperlinkSlide.addRichText([{
        runs: [
          { text: 'Browser rich one' },
          { text: ' and two', style: { underline: false } },
        ],
      }], {
        name: 'browser_rich_text_hyperlink',
        hyperlink: { slide: 2, tooltip: '' },
      });
      const browserPopulatedTextHyperlink = textShapeHyperlinkSlide.addText(
        'Browser populated text hyperlink',
        {
          placeholder: 'browser_text_hyperlink_placeholder',
          hyperlink: { slide: 1 },
        },
      );
      const textShapeHyperlinkImmediate = [
        browserPlainTextHyperlink.hyperlink,
        browserRichTextHyperlink.hyperlink,
        browserPopulatedTextHyperlink.hyperlink,
      ];
      textShapeHyperlinkInput.url = 'https://changed.browser-text.example';
      textShapeHyperlinkInput.tooltip = 'Changed';
      const textShapeHyperlinkDetached = browserPlainTextHyperlink.hyperlink;
      const textShapeHyperlinkOutput = await textShapeHyperlinkDocument.writeBlob();
      const reopenedTextShapeHyperlinks = await api.PptxDocument.open(
        textShapeHyperlinkOutput,
      );
      await reopenedTextShapeHyperlinks.write({ compatibility: 'powerpoint-current' });
      const textShapeHyperlinkByName = (owner, name) => owner.shapes.find(
        (shape) => shape instanceof api.ShapeModel && shape.name === name,
      );
      const textShapeHyperlinkShapeXml = (owner, name) => {
        const shape = textShapeHyperlinkByName(owner, name);
        const xml = new TextDecoder().decode(
          reopenedTextShapeHyperlinks.opcPackage.requirePart(owner.partUri).bytes,
        );
        const idOffset = xml.indexOf('<p:cNvPr id="' + shape.id + '"');
        const shapeStart = xml.lastIndexOf('<p:sp', idOffset);
        const shapeEnd = xml.indexOf('</p:sp>', idOffset);
        return xml.slice(shapeStart, shapeEnd + '</p:sp>'.length);
      };
      const textShapeHyperlinkClickIds = (xml) => xml.split('<a:hlinkClick').slice(1).map(
        (fragment) => fragment.split('r:id="')[1]?.split('"')[0],
      );
      const browserPlainTextHyperlinkXml = textShapeHyperlinkShapeXml(
        reopenedTextShapeHyperlinks.slides[0],
        'browser_plain_text_hyperlink',
      );
      const browserRichTextHyperlinkXml = textShapeHyperlinkShapeXml(
        reopenedTextShapeHyperlinks.slides[0],
        'browser_rich_text_hyperlink',
      );
      const browserPlainTextHyperlinkIds = textShapeHyperlinkClickIds(
        browserPlainTextHyperlinkXml,
      );
      const browserRichTextHyperlinkIds = textShapeHyperlinkClickIds(
        browserRichTextHyperlinkXml,
      );
      const textShapeHyperlinkState = {
        mime: textShapeHyperlinkOutput.type,
        immediate: textShapeHyperlinkImmediate,
        detached: textShapeHyperlinkDetached,
        reopened: [
          textShapeHyperlinkByName(
            reopenedTextShapeHyperlinks.slides[0],
            'browser_plain_text_hyperlink',
          ).hyperlink,
          textShapeHyperlinkByName(
            reopenedTextShapeHyperlinks.slides[0],
            'browser_rich_text_hyperlink',
          ).hyperlink,
          textShapeHyperlinkByName(
            reopenedTextShapeHyperlinks.slides[0],
            'browser_text_hyperlink_placeholder',
          ).hyperlink,
        ],
        layout: textShapeHyperlinkByName(
          reopenedTextShapeHyperlinks.layouts[0],
          'browser_text_hyperlink_placeholder',
        ).hyperlink,
        fill: textShapeHyperlinkByName(
          reopenedTextShapeHyperlinks.slides[0],
          'browser_plain_text_hyperlink',
        ).fill,
        line: textShapeHyperlinkByName(
          reopenedTextShapeHyperlinks.slides[0],
          'browser_plain_text_hyperlink',
        ).line,
        arrows: textShapeHyperlinkByName(
          reopenedTextShapeHyperlinks.slides[0],
          'browser_plain_text_hyperlink',
        ).arrows,
        shadowKind: textShapeHyperlinkByName(
          reopenedTextShapeHyperlinks.slides[0],
          'browser_plain_text_hyperlink',
        ).shadow?.kind,
        clickCounts: [
          browserPlainTextHyperlinkIds.length,
          browserRichTextHyperlinkIds.length,
        ],
        sharedIds: [
          new Set(browserPlainTextHyperlinkIds).size === 1,
          new Set(browserRichTextHyperlinkIds).size === 1,
        ],
        internalActions: browserRichTextHyperlinkXml.split('<a:hlinkClick').slice(1).every(
          (fragment) => fragment.includes('action="ppaction://hlinksldjump"'),
        ),
        validationErrors: reopenedTextShapeHyperlinks.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const textShapeHyperlinks =
        textShapeHyperlinkState.clickCounts[0] === 3 &&
        textShapeHyperlinkState.clickCounts[1] === 3 &&
        textShapeHyperlinkState.sharedIds.every(Boolean) &&
        textShapeHyperlinkState.internalActions &&
        textShapeHyperlinkState.validationErrors === 0;
      const richTextRunHyperlinkDocument = api.PptxDocument.create();
      const richTextRunHyperlinkSource = richTextRunHyperlinkDocument.addSlide();
      const richTextRunHyperlinkTarget = richTextRunHyperlinkDocument.addSlide();
      const richTextRunHyperlinkShape = richTextRunHyperlinkSource.addRichText([{
        runs: [
          { text: 'Browser inherited' },
          {
            text: ' browser local one',
            style: {
              hyperlink: { url: 'https://browser-run.example', tooltip: 'One' },
            },
          },
          {
            text: ' browser local two',
            style: {
              hyperlink: { url: 'https://browser-run.example', tooltip: 'Two' },
            },
          },
          { text: ' browser suppressed', style: { hyperlink: false } },
          {
            text: ' browser target',
            style: { hyperlink: { slide: 2, tooltip: '' }, underline: false },
          },
          { text: ' browser self', style: { hyperlink: { slide: 1 } } },
        ],
      }], {
        name: 'browser_run_hyperlinks',
        hyperlink: { url: 'https://browser-outer-run.example', tooltip: 'Outer' },
      });
      const browserRunHyperlinkValues = (shape) => shape.richText[0].runs.map(
        (run) => run.style?.hyperlink,
      );
      const browserRunHyperlinkShapeXml = (document, owner, shape) => {
        const xml = new TextDecoder().decode(
          document.opcPackage.requirePart(owner.partUri).bytes,
        );
        const idOffset = xml.indexOf('<p:cNvPr id="' + shape.id + '"');
        const shapeStart = xml.lastIndexOf('<p:sp', idOffset);
        const shapeEnd = xml.indexOf('</p:sp>', idOffset);
        return xml.slice(shapeStart, shapeEnd + '</p:sp>'.length);
      };
      const browserRunHyperlinkClickIds = (xml) => xml.split('<a:hlinkClick').slice(1).map(
        (fragment) => fragment.split('r:id="')[1]?.split('"')[0],
      );
      const richTextRunHyperlinkInitial = browserRunHyperlinkValues(
        richTextRunHyperlinkShape,
      );
      const richTextRunHyperlinkInitialXml = browserRunHyperlinkShapeXml(
        richTextRunHyperlinkDocument,
        richTextRunHyperlinkSource,
        richTextRunHyperlinkShape,
      );
      const richTextRunHyperlinkInitialIds = browserRunHyperlinkClickIds(
        richTextRunHyperlinkInitialXml,
      );
      const richTextRunHyperlinkInitialRelationships = richTextRunHyperlinkInitialIds.map(
        (id) => richTextRunHyperlinkSource.relationships.find(
          (relationship) => relationship.id === id,
        ),
      );
      const richTextRunHyperlinkTargetTextOffset = richTextRunHyperlinkInitialXml.indexOf(
        '> browser target</a:t>',
      );
      const richTextRunHyperlinkTargetRunStart = richTextRunHyperlinkInitialXml.lastIndexOf(
        '<a:r>',
        richTextRunHyperlinkTargetTextOffset,
      );
      const richTextRunHyperlinkTargetRunXml = richTextRunHyperlinkInitialXml.slice(
        richTextRunHyperlinkTargetRunStart,
        richTextRunHyperlinkTargetTextOffset,
      );
      const richTextRunHyperlinkRelationshipIndependence =
        richTextRunHyperlinkInitialIds.length === 6
        && new Set(richTextRunHyperlinkInitialIds).size === 5
        && richTextRunHyperlinkInitialIds[0] === richTextRunHyperlinkInitialIds[1]
        && richTextRunHyperlinkInitialIds[2] !== richTextRunHyperlinkInitialIds[3]
        && richTextRunHyperlinkInitialRelationships[2]?.target ===
          'https://browser-run.example'
        && richTextRunHyperlinkInitialRelationships[3]?.target ===
          'https://browser-run.example'
        && richTextRunHyperlinkTargetRunXml.includes('u="none"')
        && richTextRunHyperlinkTargetRunXml.includes('<a:hlinkClick')
        && richTextRunHyperlinkInitialXml.split('<a:hlinkClick').slice(5).every(
          (fragment) => fragment.includes('action="ppaction://hlinksldjump"'),
        );
      const setBrowserRunHyperlink = (shape, runIndex, hyperlink) => {
        shape.richText = shape.richText.map((paragraph, paragraphIndex) => ({
          ...paragraph,
          runs: paragraph.runs.map((run, candidateIndex) => {
            if (paragraphIndex !== 0 || candidateIndex !== runIndex) return run;
            const { hyperlink: _oldHyperlink, ...style } = run.style ?? {};
            return {
              ...run,
              style: hyperlink === undefined ? style : { ...style, hyperlink },
            };
          }),
        }));
      };
      setBrowserRunHyperlink(richTextRunHyperlinkShape, 1, {
        url: 'https://browser-run-edited.example',
        tooltip: '',
      });
      setBrowserRunHyperlink(richTextRunHyperlinkShape, 2, false);
      setBrowserRunHyperlink(richTextRunHyperlinkShape, 4, false);
      const richTextRunHyperlinkAfterEdit = browserRunHyperlinkValues(
        richTextRunHyperlinkShape,
      );
      const richTextRunHyperlinkDuplicate = richTextRunHyperlinkDocument.duplicateSlide(0);
      const richTextRunHyperlinkDuplicateShape = richTextRunHyperlinkDuplicate.shapes.find(
        ({ name }) => name === 'browser_run_hyperlinks',
      );
      richTextRunHyperlinkDocument.moveSlide(
        richTextRunHyperlinkDocument.slides.indexOf(richTextRunHyperlinkTarget),
        0,
      );
      const richTextRunHyperlinkAfterMove = [
        browserRunHyperlinkValues(richTextRunHyperlinkShape)[5],
        browserRunHyperlinkValues(richTextRunHyperlinkDuplicateShape)[5],
      ];
      richTextRunHyperlinkDocument.deleteSlide(0);
      const richTextRunHyperlinkOutput = await richTextRunHyperlinkDocument.writeBlob();
      const reopenedRichTextRunHyperlinks = await api.PptxDocument.open(
        richTextRunHyperlinkOutput,
      );
      await reopenedRichTextRunHyperlinks.write({ compatibility: 'powerpoint-current' });
      const reopenedRichTextRunHyperlinkShapes = reopenedRichTextRunHyperlinks.slides.map(
        (slide) => slide.shapes.find(({ name }) => name === 'browser_run_hyperlinks'),
      );
      const reopenedRichTextRunHyperlinkXml = reopenedRichTextRunHyperlinkShapes.map(
        (shape, index) => browserRunHyperlinkShapeXml(
          reopenedRichTextRunHyperlinks,
          reopenedRichTextRunHyperlinks.slides[index],
          shape,
        ),
      );
      const reopenedRichTextRunHyperlinkIds = reopenedRichTextRunHyperlinkXml.map(
        browserRunHyperlinkClickIds,
      );
      const richTextRunHyperlinkState = {
        mime: richTextRunHyperlinkOutput.type,
        immediate: richTextRunHyperlinkInitial,
        relationshipIndependence: richTextRunHyperlinkRelationshipIndependence,
        afterEdit: richTextRunHyperlinkAfterEdit,
        afterMove: richTextRunHyperlinkAfterMove,
        reopened: reopenedRichTextRunHyperlinkShapes.map(browserRunHyperlinkValues),
        clickCounts: reopenedRichTextRunHyperlinkIds.map(({ length }) => length),
        independentIds: reopenedRichTextRunHyperlinkIds.map(
          (ids) => new Set(ids).size,
        ),
        internalActions: reopenedRichTextRunHyperlinkXml.every(
          (xml) => xml.split('<a:hlinkClick').at(-1).includes(
            'action="ppaction://hlinksldjump"',
          ),
        ),
        validationErrors: reopenedRichTextRunHyperlinks.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
      };
      const richTextRunHyperlinks =
        richTextRunHyperlinkState.relationshipIndependence
        && richTextRunHyperlinkState.clickCounts.every((count) => count === 4)
        && richTextRunHyperlinkState.independentIds.every((count) => count === 3)
        && richTextRunHyperlinkState.internalActions
        && richTextRunHyperlinkState.validationErrors === 0;
      const tableCellHyperlinkDocument = api.PptxDocument.create();
      const tableCellHyperlinkSource = tableCellHyperlinkDocument.addSlide();
      const tableCellHyperlinkTarget = tableCellHyperlinkDocument.addSlide();
      const tableCellHyperlinkAlternate = tableCellHyperlinkDocument.addSlide();
      const tableCellHyperlinkUrlInput = {
        url: 'https://shared.example?a=1&b=2',
        tooltip: 'Visit & learn',
      };
      const tableCellHyperlinkEmptyTooltipInput = {
        url: 'https://second.example',
        tooltip: '',
      };
      const tableCellHyperlinkTable = tableCellHyperlinkSource.addTable([[
        { text: 'URL', options: { hyperlink: tableCellHyperlinkUrlInput } },
        { text: 'Empty', options: { hyperlink: tableCellHyperlinkEmptyTooltipInput } },
        { text: 'Shared', options: { hyperlink: { url: 'https://third.example' } } },
        { text: 'Slide', options: { hyperlink: { slide: 2 } } },
        { text: 'Self', options: { hyperlink: { slide: 1, tooltip: '' } } },
        'Plain',
      ]], { name: 'Chrome table-cell hyperlinks' });
      const tableCellHyperlinkSnapshot = (table) => table.rows[0].cells
        .map(({ hyperlink }) => hyperlink ?? null);
      const tableCellHyperlinkImmediate = tableCellHyperlinkSnapshot(
        tableCellHyperlinkTable,
      );
      const tableCellHyperlinkSnapshotsFrozen = tableCellHyperlinkTable.rows[0].cells
        .every(({ hyperlink }) => hyperlink === undefined || Object.isFrozen(hyperlink));
      tableCellHyperlinkUrlInput.url = 'https://changed.example';
      tableCellHyperlinkUrlInput.tooltip = 'Changed';
      tableCellHyperlinkEmptyTooltipInput.url = 'https://changed.example';
      tableCellHyperlinkEmptyTooltipInput.tooltip = 'Changed';
      const tableCellHyperlinkInputDetached =
        JSON.stringify(tableCellHyperlinkSnapshot(tableCellHyperlinkTable)) ===
          JSON.stringify(tableCellHyperlinkImmediate)
        && !Object.hasOwn(tableCellHyperlinkUrlInput, '_rId')
        && !Object.hasOwn(tableCellHyperlinkEmptyTooltipInput, '_rId');
      const readTableCellHyperlinkXml = () => new TextDecoder().decode(
        tableCellHyperlinkDocument.opcPackage
          .requirePart(tableCellHyperlinkSource.partUri).bytes,
      );
      const readTableCellHyperlinkClickIds = () => [...readTableCellHyperlinkXml().matchAll(
        /<a:hlinkClick[^>]*r:id="([^"]+)"/g,
      )].map((match) => match[1]);
      const tableCellHyperlinkXml = readTableCellHyperlinkXml();
      const tableCellHyperlinkClickIds = [...tableCellHyperlinkXml.matchAll(
        /<a:hlinkClick[^>]*r:id="([^"]+)"/g,
      )].map((match) => match[1]);
      const tableCellHyperlinkRelationships = tableCellHyperlinkSource.relationships
        .filter(({ type }) => type.endsWith('/hyperlink'));
      const tableCellSlideRelationships = tableCellHyperlinkSource.relationships
        .filter(({ type }) => type.endsWith('/slide'));
      const tableCellHyperlinkNonVisualEnd = tableCellHyperlinkXml
        .indexOf('</p:nvGraphicFramePr>');
      const tableCellHyperlinkIndependentRelationships =
        tableCellHyperlinkClickIds.length === 5
        && new Set(tableCellHyperlinkClickIds).size === 5
        && tableCellHyperlinkRelationships.length === 3
        && new Set(tableCellHyperlinkRelationships.map(({ id }) => id)).size === 3
        && tableCellHyperlinkRelationships.every(
          ({ targetMode }) => targetMode === 'External',
        )
        && tableCellSlideRelationships.length === 2
        && tableCellSlideRelationships[0].resolvedTarget === tableCellHyperlinkTarget.partUri
        && tableCellSlideRelationships[1].resolvedTarget === tableCellHyperlinkSource.partUri
        && tableCellHyperlinkXml.split('u="sng"').length - 1 === 5
        && tableCellHyperlinkNonVisualEnd >= 0
        && !tableCellHyperlinkXml.slice(0, tableCellHyperlinkNonVisualEnd)
          .includes('<a:hlinkClick');

      const tableCellHyperlinkSharedId = tableCellHyperlinkClickIds[0];
      const tableCellHyperlinkReplacedIds = tableCellHyperlinkClickIds.slice(1, 3);
      let tableCellHyperlinkSharedXml = tableCellHyperlinkXml;
      for (const id of tableCellHyperlinkReplacedIds) {
        tableCellHyperlinkSharedXml = tableCellHyperlinkSharedXml.replace(
          'r:id="' + id + '"',
          'r:id="' + tableCellHyperlinkSharedId + '"',
        );
      }
      tableCellHyperlinkDocument.transaction(() => {
        tableCellHyperlinkDocument.opcPackage.setPart(
          tableCellHyperlinkSource.partUri,
          tableCellHyperlinkSharedXml,
          tableCellHyperlinkDocument.opcPackage
            .requirePart(tableCellHyperlinkSource.partUri).contentType,
        );
        for (const id of tableCellHyperlinkReplacedIds) {
          tableCellHyperlinkDocument.opcPackage.removeRelationship(
            tableCellHyperlinkSource.partUri,
            id,
          );
        }
      });
      const tableCellHyperlinkShared = tableCellHyperlinkSnapshot(
        tableCellHyperlinkTable,
      );

      const tableCellHyperlinkNoOpBytes = tableCellHyperlinkDocument.opcPackage
        .requirePart(tableCellHyperlinkSource.partUri).bytes.slice();
      const tableCellHyperlinkNoOpRelationships = JSON.stringify(
        tableCellHyperlinkSource.relationships,
      );
      const tableCellHyperlinkNoOpJournal = JSON.stringify(
        tableCellHyperlinkDocument.opcPackage.mutations,
      );
      tableCellHyperlinkTable.setCellHyperlink(0, 2, {
        url: 'https://shared.example?a=1&b=2',
      });
      const tableCellHyperlinkNoOp = tableVerticalAlignmentBytesEqual(
        tableCellHyperlinkNoOpBytes,
        tableCellHyperlinkDocument.opcPackage
          .requirePart(tableCellHyperlinkSource.partUri).bytes,
      ) && JSON.stringify(tableCellHyperlinkSource.relationships) ===
        tableCellHyperlinkNoOpRelationships
        && JSON.stringify(tableCellHyperlinkDocument.opcPackage.mutations) ===
          tableCellHyperlinkNoOpJournal;

      tableCellHyperlinkTable.setCellHyperlink(0, 1, {
        url: 'https://shared.example?a=1&b=2',
        tooltip: 'Peer',
      });
      const tableCellHyperlinkTooltipIdReuse = readTableCellHyperlinkClickIds()[1] ===
        tableCellHyperlinkSharedId;
      tableCellHyperlinkTable.setCellHyperlink(0, 0, {
        url: 'https://edited.example?a=1&b=2',
        tooltip: 'Edited',
      });
      const tableCellHyperlinkClonedId = readTableCellHyperlinkClickIds()[0];
      const tableCellHyperlinkCloneOnWrite =
        tableCellHyperlinkClonedId !== tableCellHyperlinkSharedId
        && readTableCellHyperlinkClickIds().slice(1, 3).every(
          (id) => id === tableCellHyperlinkSharedId,
        )
        && tableCellHyperlinkSource.relationships.find(
          ({ id }) => id === tableCellHyperlinkClonedId,
        )?.target === 'https://edited.example?a=1&b=2'
        && tableCellHyperlinkSource.relationships.find(
          ({ id }) => id === tableCellHyperlinkSharedId,
        )?.target === 'https://shared.example?a=1&b=2';

      const tableCellHyperlinkUniqueInternalId = readTableCellHyperlinkClickIds()[3];
      tableCellHyperlinkTable.setCellHyperlink(0, 3, { slide: 3, tooltip: '' });
      const tableCellHyperlinkUniqueIdReuse = readTableCellHyperlinkClickIds()[3] ===
        tableCellHyperlinkUniqueInternalId
        && tableCellHyperlinkSource.relationships.find(
          ({ id }) => id === tableCellHyperlinkUniqueInternalId,
        )?.resolvedTarget === tableCellHyperlinkAlternate.partUri;

      tableCellHyperlinkTable.setCellHyperlink(0, 5, {
        url: 'https://added.example',
      });
      const tableCellHyperlinkAddedId = readTableCellHyperlinkClickIds()[5];
      tableCellHyperlinkTable.setCellHyperlink(0, 5, undefined);
      const tableCellHyperlinkAddClear =
        tableCellHyperlinkTable.rows[0].cells[5].hyperlink === undefined
        && !tableCellHyperlinkSource.relationships.some(
          ({ id }) => id === tableCellHyperlinkAddedId,
        ) && readTableCellHyperlinkXml().includes('u="sng"');

      tableCellHyperlinkTable.setCellHyperlink(0, 1, undefined);
      const tableCellHyperlinkSharedRetained = tableCellHyperlinkSource.relationships.some(
        ({ id }) => id === tableCellHyperlinkSharedId,
      );
      tableCellHyperlinkTable.setCellHyperlink(0, 2, undefined);
      const tableCellHyperlinkSharedCollected = !tableCellHyperlinkSource.relationships.some(
        ({ id }) => id === tableCellHyperlinkSharedId,
      );
      const tableCellHyperlinkSharedGc = tableCellHyperlinkSharedRetained
        && tableCellHyperlinkSharedCollected;

      tableCellHyperlinkTable.setCellText(0, 0, 'URL edited');
      tableCellHyperlinkTable.setCellFill(0, 0, { kind: 'none' });
      tableCellHyperlinkTable.setCellMargins(0, 0, { top: 2 });
      const tableCellHyperlinkTextEditPreserved = tableCellHyperlinkTable
        .rows[0].cells[0].hyperlink;
      const tableCellHyperlinkAfterEdit = tableCellHyperlinkSnapshot(
        tableCellHyperlinkTable,
      );
      tableCellHyperlinkDocument.moveSlide(
        tableCellHyperlinkDocument.slides.indexOf(tableCellHyperlinkAlternate),
        0,
      );
      const tableCellHyperlinkMovedInternal = tableCellHyperlinkTable
        .rows[0].cells[3].hyperlink;
      const tableCellHyperlinkMovedSelf = tableCellHyperlinkTable
        .rows[0].cells[4].hyperlink;
      tableCellHyperlinkDocument.moveSlide(0, tableCellHyperlinkDocument.slides.length - 1);
      const tableCellHyperlinkRestoredInternal = tableCellHyperlinkTable
        .rows[0].cells[3].hyperlink;
      const tableCellHyperlinkDuplicate = tableCellHyperlinkDocument.duplicateSlide(
        tableCellHyperlinkDocument.slides.indexOf(tableCellHyperlinkSource),
      );
      const tableCellHyperlinkDuplicateTable = tableCellHyperlinkDuplicate.shapes.find(
        (shape) => shape instanceof api.TableModel,
      );
      const tableCellHyperlinkDuplicateSelf = tableCellHyperlinkDuplicateTable
        .rows[0].cells[4].hyperlink;
      tableCellHyperlinkDocument.deleteSlide(
        tableCellHyperlinkDocument.slides.indexOf(tableCellHyperlinkAlternate),
      );
      const tableCellHyperlinkAfterTargetDeletion = {
        source: tableCellHyperlinkTable.rows[0].cells[3].hyperlink ?? null,
        duplicate: tableCellHyperlinkDuplicateTable.rows[0].cells[3].hyperlink ?? null,
      };
      const tableCellHyperlinkInvalidBytes = tableCellHyperlinkDocument.opcPackage
        .requirePart(tableCellHyperlinkSource.partUri).bytes.slice();
      const tableCellHyperlinkInvalidRelationships = JSON.stringify(
        tableCellHyperlinkSource.relationships,
      );
      const tableCellHyperlinkInvalidJournal = JSON.stringify(
        tableCellHyperlinkDocument.opcPackage.mutations,
      );
      let tableCellHyperlinkInvalidError;
      try {
        tableCellHyperlinkTable.setCellHyperlink(0, 0, { slide: 99 });
      } catch (error) {
        tableCellHyperlinkInvalidError = { name: error.name, message: error.message };
      }
      const tableCellHyperlinkFailureIsolation = tableVerticalAlignmentBytesEqual(
        tableCellHyperlinkInvalidBytes,
        tableCellHyperlinkDocument.opcPackage
          .requirePart(tableCellHyperlinkSource.partUri).bytes,
      ) && JSON.stringify(tableCellHyperlinkSource.relationships) ===
        tableCellHyperlinkInvalidRelationships
        && JSON.stringify(tableCellHyperlinkDocument.opcPackage.mutations) ===
          tableCellHyperlinkInvalidJournal;
      const reopenedTableCellHyperlinkDocument = await api.PptxDocument.open(
        await tableCellHyperlinkDocument.writeBlob(),
      );
      const reopenedTableCellHyperlinkTable = reopenedTableCellHyperlinkDocument
        .slides[0].shapes.find(({ name }) => name === 'Chrome table-cell hyperlinks');
      const reopenedTableCellHyperlinkDuplicateTable = reopenedTableCellHyperlinkDocument
        .slides.at(-1).shapes.find(({ name }) => name === 'Chrome table-cell hyperlinks');
      const tableCellHyperlinkFinalXml = readTableCellHyperlinkXml();
      const tableCellHyperlinkFinalClickIds = readTableCellHyperlinkClickIds();
      const tableCellHyperlinkFinalRelationshipIds = tableCellHyperlinkSource.relationships
        .filter(({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'))
        .map(({ id }) => id);
      const tableCellHyperlinkState = {
        immediate: tableCellHyperlinkImmediate,
        inputDetached: tableCellHyperlinkInputDetached,
        snapshotsFrozen: tableCellHyperlinkSnapshotsFrozen,
        independentRelationships: tableCellHyperlinkIndependentRelationships,
        shared: tableCellHyperlinkShared,
        noOp: tableCellHyperlinkNoOp,
        tooltipIdReuse: tableCellHyperlinkTooltipIdReuse,
        cloneOnWrite: tableCellHyperlinkCloneOnWrite,
        uniqueIdReuse: tableCellHyperlinkUniqueIdReuse,
        addClear: tableCellHyperlinkAddClear,
        sharedGc: tableCellHyperlinkSharedGc,
        textEditPreserved: tableCellHyperlinkTextEditPreserved,
        afterEdit: tableCellHyperlinkAfterEdit,
        movedInternal: tableCellHyperlinkMovedInternal,
        movedSelf: tableCellHyperlinkMovedSelf,
        restoredInternal: tableCellHyperlinkRestoredInternal,
        duplicateSelf: tableCellHyperlinkDuplicateSelf,
        targetDeletion: tableCellHyperlinkAfterTargetDeletion,
        finalRelationshipOwnership: {
          clickCount: tableCellHyperlinkFinalClickIds.length,
          relationshipCount: tableCellHyperlinkFinalRelationshipIds.length,
          uniqueClickIds: new Set(tableCellHyperlinkFinalClickIds).size,
          idsMatch: JSON.stringify([...tableCellHyperlinkFinalClickIds].sort()) ===
            JSON.stringify([...tableCellHyperlinkFinalRelationshipIds].sort()),
          underlineCount: tableCellHyperlinkFinalXml.split('u="sng"').length - 1,
        },
        reopened: tableCellHyperlinkSnapshot(reopenedTableCellHyperlinkTable),
        reopenedDuplicate: tableCellHyperlinkSnapshot(
          reopenedTableCellHyperlinkDuplicateTable,
        ),
        invalidError: tableCellHyperlinkInvalidError,
        failureIsolation: tableCellHyperlinkFailureIsolation,
        validationErrors: tableCellHyperlinkDocument.diagnostics
          .filter(({ severity }) => severity === 'error').length +
          reopenedTableCellHyperlinkDocument.diagnostics
            .filter(({ severity }) => severity === 'error').length,
      };
      const tableCellHyperlinks = JSON.stringify(tableCellHyperlinkState) ===
        JSON.stringify({
          immediate: [
            { url: 'https://shared.example?a=1&b=2', tooltip: 'Visit & learn' },
            { url: 'https://second.example', tooltip: '' },
            { url: 'https://third.example' },
            { slide: 2 },
            { slide: 1, tooltip: '' },
            null,
          ],
          inputDetached: true,
          snapshotsFrozen: true,
          independentRelationships: true,
          shared: [
            { url: 'https://shared.example?a=1&b=2', tooltip: 'Visit & learn' },
            { url: 'https://shared.example?a=1&b=2', tooltip: '' },
            { url: 'https://shared.example?a=1&b=2' },
            { slide: 2 },
            { slide: 1, tooltip: '' },
            null,
          ],
          noOp: true,
          tooltipIdReuse: true,
          cloneOnWrite: true,
          uniqueIdReuse: true,
          addClear: true,
          sharedGc: true,
          textEditPreserved: {
            url: 'https://edited.example?a=1&b=2',
            tooltip: 'Edited',
          },
          afterEdit: [
            { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
            null,
            null,
            { slide: 3, tooltip: '' },
            { slide: 1, tooltip: '' },
            null,
          ],
          movedInternal: { slide: 1, tooltip: '' },
          movedSelf: { slide: 2, tooltip: '' },
          restoredInternal: { slide: 3, tooltip: '' },
          duplicateSelf: { slide: 4, tooltip: '' },
          targetDeletion: { source: null, duplicate: null },
          finalRelationshipOwnership: {
            clickCount: 2,
            relationshipCount: 2,
            uniqueClickIds: 2,
            idsMatch: true,
            underlineCount: 6,
          },
          reopened: [
            { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
            null,
            null,
            null,
            { slide: 1, tooltip: '' },
            null,
          ],
          reopenedDuplicate: [
            { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
            null,
            null,
            null,
            { slide: 3, tooltip: '' },
            null,
          ],
          invalidError: {
            name: 'RangeError',
            message: 'Table cell 0,0 hyperlink slide 99 is out of range',
          },
          failureIsolation: true,
          validationErrors: 0,
        });
      const tableCellHyperlinkEditing = tableCellHyperlinks;
      const svgDocument = api.PptxDocument.create();
      svgDocument.addSlide();
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360">'
        + '<rect width="640" height="360" fill="#4472C4"/></svg>';
      const svgDataUri = 'data:image/svg+xml;base64,' + btoa(svg);
      const blobSvg = await svgDocument.addImage(0, new Blob([svg], { type: 'image/svg+xml' }), {
        name: 'Browser Blob SVG',
        altText: 'Canvas fallback from Blob',
        x: api.inches(1),
        y: api.inches(1),
        sizing: { type: 'cover', width: api.inches(4), height: api.inches(3) },
        rotation: api.degrees(15),
        flipHorizontal: true,
      });
      const dataSvg = await svgDocument.addImage(0, svgDataUri, {
        name: 'Browser data SVG',
        altText: 'Canvas fallback from data URI',
        x: api.inches(5.5),
        y: api.inches(1),
        sizing: { type: 'contain', width: api.inches(4), height: api.inches(3) },
        flipVertical: true,
      });
      const relativePngPath = new URL(
        '../../../docs/images/wp4-gradient-codecs.png',
        moduleUrl,
      ).pathname;
      const relativePngResponse = await fetch(relativePngPath);
      const relativePngInfo = api.inspectRasterImage(
        new Uint8Array(await relativePngResponse.arrayBuffer()),
      );
      const relativeRaster = await svgDocument.addImage(0, relativePngPath, {
        name: 'Browser relative PNG',
        altText: 'Relative path with pixel crop',
        x: api.inches(1),
        y: api.inches(4.5),
        sizing: {
          type: 'crop',
          width: api.inches(4),
          height: api.inches(3),
          source: {
            x: relativePngInfo.width / 4,
            y: relativePngInfo.height / 4,
            width: relativePngInfo.width / 2,
            height: relativePngInfo.height / 2,
          },
        },
        rotation: api.degrees(-10),
        flipHorizontal: true,
        flipVertical: true,
      });
      const backgroundDocument = api.PptxDocument.create();
      const backgroundPngBytes = Uint8Array.from([
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
        0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
        39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
      ]);
      backgroundDocument.addSlide();
      backgroundDocument.addSlide();
      const solidBackgroundSlide = backgroundDocument.addSlide();
      const gradientBackgroundSlide = backgroundDocument.addSlide();
      await backgroundDocument.setSlideBackgroundImage(
        0,
        new Blob([backgroundPngBytes], { type: 'image/png' }),
      );
      await backgroundDocument.setSlideBackgroundImage(
        1,
        'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      );
      solidBackgroundSlide.background = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      };
      gradientBackgroundSlide.background = {
        kind: 'linear-gradient',
        angle: 45,
        stops: [
          { offset: 0, color: 'FF0000' },
          { offset: 1, color: '0000FF', alpha: 0.5 },
        ],
      };
      const backgroundOutput = await backgroundDocument.writeBlob({
        compatibility: 'powerpoint-2010',
      });
      const reopenedBackgroundDocument = await api.PptxDocument.open(backgroundOutput);
      await reopenedBackgroundDocument.write({ compatibility: 'powerpoint-2010' });
      const hexDigest = async (payload) => Array.from(
        new Uint8Array(await crypto.subtle.digest('SHA-256', payload)),
        (value) => value.toString(16).padStart(2, '0'),
      ).join('');
      const backgroundPayloadHashes = await Promise.all(
        reopenedBackgroundDocument.slides
          .map(({ background }) => background)
          .filter((background) => background?.kind === 'image')
          .map(({ bytes: payload }) => hexDigest(payload)),
      );
      const backgroundRelationshipCounts = reopenedBackgroundDocument.slides.map((slide) =>
        slide.relationships.filter(
          ({ type, targetMode }) => type.endsWith('/image') && targetMode === 'Internal',
        ).length);
      const mediaDocument = api.PptxDocument.create();
      mediaDocument.addSlide();
      const mediaPngPoster = Uint8Array.from([
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
        0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
        39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
      ]);
      const posterStream = new ReadableStream({
        start(controller) {
          controller.enqueue(mediaPngPoster.slice(0, 20));
          controller.enqueue(mediaPngPoster.slice(20));
          controller.close();
        },
      });
      const browserAudio = await mediaDocument.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
        name: 'Browser MP3 narration',
        altText: 'Browser data URI audio',
        poster: posterStream,
        posterContentType: 'image/png',
        x: api.inches(1),
        y: api.inches(1),
        width: api.inches(2),
        height: api.inches(1),
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0.5,
      });
      const browserVideo = await mediaDocument.addVideo(
        0,
        new Blob([Uint8Array.of(5, 6, 7, 8)], { type: 'video/mp4' }),
        {
          name: 'Browser Blob video',
          altText: 'Browser Blob video with JPEG poster',
          poster: Uint8Array.of(255, 216, 255, 224, 0, 16, 74, 70, 73, 70, 255, 217),
          posterContentType: 'image/jpeg',
          x: api.inches(4),
          y: api.inches(1),
          width: api.inches(4),
          height: api.inches(2.25),
        },
      );
      browserVideo.settings = {
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0.25,
      };
      browserVideo.settings = undefined;
      const browserVideoSettingsCleared = Object.keys(browserVideo.settings).length === 0;
      const browserMediaIdentity = mediaDocument.media(0)[0] === browserAudio
        && mediaDocument.slides[0].media[0] === browserAudio
        && mediaDocument.slides[0].shapes[0] === browserAudio;
      browserAudio.name = 'Browser MP3 narration edited';
      browserAudio.altText = undefined;
      browserAudio.settings = { play: 'click', loop: false, volume: 1 };
      browserAudio.setTransform({
        x: api.inches(2),
        y: api.inches(1),
        width: api.inches(3),
        height: api.inches(1),
      });
      await browserAudio.replaceSource('https://example.com/browser-audio.mp3');
      await browserAudio.replaceSource(
        new Blob([Uint8Array.of(9, 10)], { type: 'audio/mpeg' }),
        { contentType: 'audio/mpeg', fileName: 'browser-replaced.mp3' },
      );
      await browserAudio.replacePoster(
        new Blob([Uint8Array.of(71, 73, 70, 56, 57, 97)], { type: 'image/gif' }),
        { contentType: 'image/gif' },
      );
      const browserPosterReplacement = mediaDocument.opcPackage
        .requirePart(browserAudio.posterPartUri).contentType === 'image/gif';
      await browserAudio.replacePoster();
      const browserDuplicate = mediaDocument.duplicateSlide(0);
      const browserDuplicateAudio = browserDuplicate.media[0];
      const browserDuplicateVideo = browserDuplicate.media[1];
      const browserMediaShared = browserDuplicateAudio.mediaPartUri === browserAudio.mediaPartUri
        && browserDuplicateAudio.posterPartUri === browserAudio.posterPartUri
        && browserDuplicateVideo.mediaPartUri === browserVideo.mediaPartUri;
      browserDuplicateAudio.settings = {
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0.25,
      };
      await browserDuplicateAudio.replaceSource(
        new ReadableStream({
          start(controller) {
            controller.enqueue(Uint8Array.of(11));
            controller.enqueue(Uint8Array.of(12));
            controller.close();
          },
        }),
        { contentType: 'audio/wav' },
      );
      await browserDuplicateAudio.replacePoster(
        new Blob([Uint8Array.of(255, 216, 255, 217)], { type: 'image/jpeg' }),
        { contentType: 'image/jpeg' },
      );
      const browserMediaCloneOnWrite = browserDuplicateAudio.mediaPartUri !== browserAudio.mediaPartUri
        && browserDuplicateAudio.posterPartUri !== browserAudio.posterPartUri;
      const browserVideoTarget = browserVideo.mediaPartUri;
      browserDuplicateVideo.remove();
      const browserMediaRemovalIsolation = mediaDocument.opcPackage.hasPart(browserVideoTarget)
        && mediaDocument.media(0)[1] === browserVideo && browserDuplicate.media.length === 1;
      mediaDocument.moveSlide(1, 0);
      const browserMediaMoveIdentity = mediaDocument.slides[0] === browserDuplicate
        && browserDuplicate.media[0] === browserDuplicateAudio;
      mediaDocument.moveSlide(0, 1);
      const browserOnlineVideoLink =
        'https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0&controls=0';
      await mediaDocument.addVideo(0, browserOnlineVideoLink, {
        name: 'Browser online video',
        altText: 'External online video relationship',
        poster: mediaPngPoster,
        posterContentType: 'image/png',
        x: api.inches(6),
        y: api.inches(4),
        width: api.inches(4),
        height: api.inches(2.25),
      });
      const output = await document.writeBlob();
      const reopened = await api.PptxDocument.open(output);
      const reopenedSvgDocument = await api.PptxDocument.open(await svgDocument.writeBlob());
      const reopenedSlide = reopenedSvgDocument.slides[0];
      const reopenedSvgImages = reopenedSlide.shapes.filter((shape) => shape.isSvg);
      const svgState = reopenedSvgImages.map((image) => {
        const fallback = reopenedSvgDocument.opcPackage.requirePart(image.fallbackPartUri);
        const vector = reopenedSvgDocument.opcPackage.requirePart(image.svgPartUri);
        const targets = [image.fallbackPartUri, image.svgPartUri];
        return {
          name: image.name,
          fallbackType: fallback.contentType,
          svgType: vector.contentType,
          pngSignature: Array.from(fallback.bytes.slice(0, 8)),
          transform: image.transform,
          sourceRectangle: image.sourceRectangle,
          internalTargets: targets.filter((target) => reopenedSlide.relationships.some(
            ({ type, targetMode, resolvedTarget }) => type.endsWith('/image')
              && targetMode === 'Internal' && resolvedTarget === target,
          )).length,
        };
      });
      const reopenedRelativeRaster = reopenedSlide.shapes.find(
        ({ name }) => name === 'Browser relative PNG',
      );
      const relativeRasterPart = reopenedSvgDocument.opcPackage.requirePart(
        reopenedRelativeRaster.sourcePartUri,
      );
      const imageSourceSizingTransformState = {
        live: svgDocument.slides[0].shapes.includes(blobSvg)
          && svgDocument.slides[0].shapes.includes(dataSvg)
          && svgDocument.slides[0].shapes.includes(relativeRaster),
        relativeRaster: {
          contentType: relativeRasterPart.contentType,
          transform: reopenedRelativeRaster.transform,
          sourceRectangle: reopenedRelativeRaster.sourceRectangle,
          internalTargets: reopenedSlide.relationships.filter(
            ({ type, targetMode, resolvedTarget }) => type.endsWith('/image')
              && targetMode === 'Internal'
              && resolvedTarget === reopenedRelativeRaster.sourcePartUri,
          ).length,
        },
      };
      const mediaOutput = await mediaDocument.writeBlob({ compatibility: 'powerpoint-2010' });
      const reopenedMediaDocument = await api.PptxDocument.open(mediaOutput);
      await reopenedMediaDocument.write({ mode: 'permissive', compatibility: 'powerpoint-2010' });
      const reopenedMediaSlide = reopenedMediaDocument.slides[0];
      const reopenedMedia = reopenedMediaDocument.media(0);
      const reopenedDuplicateMedia = reopenedMediaDocument.media(1);
      const mediaXml = new TextDecoder().decode(
        reopenedMediaDocument.opcPackage.requirePart(reopenedMediaSlide.partUri).bytes,
      );
      const mediaRelationshipsPartUri = reopenedMediaSlide.partUri.replace(
        /\/([^/]+)$/u,
        '/_rels/$1.rels',
      );
      const mediaRelationshipsXml = new TextDecoder().decode(
        reopenedMediaDocument.opcPackage.requirePart(mediaRelationshipsPartUri).bytes,
      );
      const mediaState = reopenedMedia.map((model) => {
        const posterPart = reopenedMediaDocument.opcPackage.requirePart(model.posterPartUri);
        const posterRelationships = reopenedMediaSlide.relationships.filter(
          ({ resolvedTarget }) => resolvedTarget === model.posterPartUri,
        );
        if (model.externalUrl) {
          const externalRelationships = reopenedMediaSlide.relationships.filter(
            ({ target }) => target === model.externalUrl,
          );
          return {
            kind: model.kind,
            externalUrl: model.externalUrl,
            mediaPartUri: model.mediaPartUri ?? null,
            posterType: posterPart.contentType,
            posterExtension: model.posterPartUri.slice(model.posterPartUri.lastIndexOf('.')),
            transform: model.transform,
            roles: [
              externalRelationships.some(({ type, targetMode }) =>
                type.endsWith('/video') && targetMode === 'External'),
              posterRelationships.some(({ type, targetMode }) =>
                type.endsWith('/image') && targetMode === 'Internal'),
            ],
            posterSignature: Array.from(posterPart.bytes.slice(0, 4)),
          };
        }
        const mediaPart = reopenedMediaDocument.opcPackage.requirePart(model.mediaPartUri);
        const mediaRelationships = reopenedMediaSlide.relationships.filter(
          ({ resolvedTarget }) => resolvedTarget === model.mediaPartUri,
        );
        return {
          kind: model.kind,
          mediaType: mediaPart.contentType,
          mediaExtension: model.mediaPartUri.slice(model.mediaPartUri.lastIndexOf('.')),
          posterType: posterPart.contentType,
          posterExtension: model.posterPartUri.slice(model.posterPartUri.lastIndexOf('.')),
          roles: [
            mediaRelationships.some(({ type }) => type.endsWith('/' + model.kind)),
            mediaRelationships.some(({ type }) =>
              type === 'http://schemas.microsoft.com/office/2007/relationships/media'),
            posterRelationships.some(({ type }) => type.endsWith('/image')),
          ],
          posterSignature: Array.from(posterPart.bytes.slice(0, 4)),
        };
      });
      const mediaNames = reopenedMedia.map((model) => {
        const match = mediaXml.match(new RegExp(
          '<p:cNvPr\\b[^>]*\\bid="' + model.shapeId + '"[^>]*\\bname="([^"]*)"',
        ));
        return match?.[1];
      });
      const reopenedBrowserOnlineVideo = reopenedMedia.find(
        ({ externalUrl }) => externalUrl === browserOnlineVideoLink,
      );
      const browserOnlineVideoState = {
        externalUrl: reopenedBrowserOnlineVideo?.externalUrl,
        mediaPartUri: reopenedBrowserOnlineVideo?.mediaPartUri ?? null,
        name: reopenedBrowserOnlineVideo?.name,
        transform: reopenedBrowserOnlineVideo?.transform,
        posterType: reopenedBrowserOnlineVideo === undefined
          ? undefined
          : reopenedMediaDocument.opcPackage
              .requirePart(reopenedBrowserOnlineVideo.posterPartUri).contentType,
        externalRelationship: reopenedMediaSlide.relationships.some(
          ({ type, target, targetMode }) => type.endsWith('/video')
            && target === browserOnlineVideoLink && targetMode === 'External',
        ),
        posterRelationship: reopenedBrowserOnlineVideo !== undefined &&
          reopenedMediaSlide.relationships.some(
            ({ type, resolvedTarget, targetMode }) => type.endsWith('/image')
              && resolvedTarget === reopenedBrowserOnlineVideo.posterPartUri
              && targetMode === 'Internal',
          ),
        escapedQuery: mediaRelationshipsXml.includes('rel=0&amp;controls=0'),
      };
      const mediaOrphanCount = reopenedMediaDocument.opcPackage.parts
        .filter(({ uri }) => uri.startsWith('/ppt/media/'))
        .filter(({ uri }) =>
          (reopenedMediaDocument.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) === 0)
        .length;
      const timingSummaries = reopenedMediaDocument.slides.map((slide, slideIndex) => {
        const source = new TextDecoder().decode(
          reopenedMediaDocument.opcPackage.requirePart(slide.partUri).bytes,
        );
        const ids = [...source.matchAll(/<p:cTn\b[^>]*\bid="([0-9]+)"/g)]
          .map((match) => Number(match[1]));
        const targets = [...source.matchAll(/<p:spTgt\b[^>]*\bspid="([0-9]+)"/g)]
          .map((match) => Number(match[1]));
        const configuredModels = reopenedMediaDocument.media(slideIndex)
          .filter((model) => Object.keys(model.settings).length > 0);
        const configured = configuredModels.map(({ shapeId }) => shapeId);
        return {
          timing: (source.match(/<p:timing>/g) ?? []).length,
          media: (source.match(/<p:cMediaNode\b/g) ?? []).length,
          commands: (source.match(/<p:cmd\b/g) ?? []).length,
          playback: (source.match(/<px:playback\b/g) ?? []).length,
          configured: configured.length,
          expectedCommands: configuredModels.reduce(
            (count, { kind }) => count + (kind === 'video' ? 2 : 1),
            0,
          ),
          ids: ids.length,
          uniqueIds: new Set(ids).size,
          isolatedTargets: targets.length > 0 && targets.every((target) => configured.includes(target))
            && new Set(targets).size === configured.length,
        };
      });
      const settingsAfterReopen = reopenedMediaDocument.slides.map((_slide, slideIndex) =>
        reopenedMediaDocument.media(slideIndex).map(({ settings }) => settings));
      const timingDiagnostics = reopenedMediaDocument.diagnostics
        .filter(({ code }) => code.startsWith('MEDIA_TIMING_'))
        .map(({ code }) => code);
      const nativeMediaTiming = browserVideoSettingsCleared
        && timingSummaries.every((summary) => summary.timing === (summary.configured > 0 ? 1 : 0)
          && summary.media === summary.configured
          && summary.commands === summary.expectedCommands
          && summary.playback === summary.configured
          && summary.ids === summary.uniqueIds && summary.isolatedTargets)
        && timingDiagnostics.length === 0
        && settingsAfterReopen[0][0].play === 'click'
        && Object.keys(settingsAfterReopen[0][1]).length === 0
        && settingsAfterReopen[0][2].play === 'click'
        && settingsAfterReopen[0][2].volume === 1
        && settingsAfterReopen[1][0].play === 'auto'
        && settingsAfterReopen[1][0].loop === true
        && settingsAfterReopen[1][0].hideWhenStopped === true
        && settingsAfterReopen[1][0].volume === 0.25;
      const stableMediaLifecycle = browserMediaIdentity && browserPosterReplacement
        && browserMediaShared && browserMediaCloneOnWrite && browserMediaRemovalIsolation
        && browserMediaMoveIdentity && reopenedDuplicateMedia.length === 1
        && reopenedDuplicateMedia[0].mediaPartUri.endsWith('.wav')
        && reopenedDuplicateMedia[0].posterPartUri.endsWith('.jpg')
        && mediaOrphanCount === 0;
      const chartDocument = api.PptxDocument.create();
      const chartModels = [];
      for (const type of api.CHART_TYPES) {
        const slide = chartDocument.addSlide();
        const series = type === 'scatter'
          ? [{ name: 'Forecast', xValues: [1, 2, 3], values: [120, 150, 135] }]
          : type === 'bubble'
            ? [{ name: 'Portfolio', xValues: [1, 2, 3], values: [120, 150, 135], sizes: [8, 12, 10] }]
            : [{ name: 'Revenue', categories: ['North', 'South', 'West'], values: [120, 150, 135] }];
        const chart = await slide.addChart(type, series, {
          name: `Browser ${type} chart`,
          altText: `${type} chart with regional business data`,
          x: api.inches(0.5),
          y: api.inches(0.5),
          width: api.inches(9),
          height: api.inches(6.5),
        });
        if (type === 'area') {
          await chart.replaceDefinition({
            groups: chart.definition.groups,
            options: {
              ...chart.definition.options,
              roundedCorners: true,
              chartArea: {
                fill: {
                  kind: 'solid',
                  color: { kind: 'scheme', value: 'accent2' },
                  transparency: 25,
                },
                line: {
                  kind: 'line',
                  color: { kind: 'srgb', value: '112233' },
                  width: 2,
                },
              },
              plotArea: {
                fill: {
                  kind: 'solid',
                  color: { kind: 'srgb', value: '445566' },
                  transparency: 50,
                },
                line: {
                  kind: 'line',
                  color: { kind: 'scheme', value: 'accent1' },
                  width: 1,
                },
              },
              categoryAxis: {
                visible: false,
                labelPosition: 'high',
                labelRotation: -45,
                line: { kind: 'none' },
                majorGridLine: {
                  kind: 'line',
                  color: { kind: 'srgb', value: '667788' },
                  width: 0.75,
                  dash: 'sysDot',
                },
                majorTickMark: 'inside',
                minorTickMark: 'outside',
              },
              valueAxis: {
                visible: false,
                labelPosition: 'none',
                labelRotation: 45,
                line: {
                  kind: 'line',
                  color: { kind: 'srgb', value: '778899' },
                  width: 1.25,
                  dash: 'dash',
                },
                majorGridLine: { kind: 'none' },
                majorTickMark: 'cross',
                minorTickMark: 'none',
              },
            },
          });
        }
        chartModels.push(chart);
      }
      const comboSlide = chartDocument.addSlide();
      const combo = await comboSlide.addChart([
        {
          type: 'bar',
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [100, 130] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{ name: 'Margin', categories: ['Q1', 'Q2'], values: [24, 28] }],
        },
      ]);
      await chartModels[0].replaceSeries([{
        name: 'Revenue edited',
        categories: ['North', 'South', 'West'],
        values: [125, 155, 140],
      }]);
      await chartModels[1].replaceDefinition({ groups: [{
        type: 'line',
        series: [{ name: 'Converted', categories: ['Q1', 'Q2'], values: [11, 22] }],
      }] });
      const comboChartPartUri = combo.chartPartUri;
      const chartAxisAdvancedDisplayUnits = [
        'billions',
        'hundredMillions',
        'hundredThousands',
        'hundreds',
        'millions',
        'tenMillions',
        'tenThousands',
        'thousands',
        'trillions',
      ];
      for (const [index, displayUnit] of chartAxisAdvancedDisplayUnits.entries()) {
        const slide = chartDocument.addSlide();
        const chart = await slide.addChart(index === 0 ? 'bar3D' : 'bar', [{
          name: displayUnit,
          categories: index === 0 ? [46_023, 46_054] : ['Q1', 'Q2'],
          values: [10, 20],
        }], { name: `Browser advanced axis ${displayUnit}` });
        await chart.replaceDefinition({
          groups: chart.definition.groups,
          options: {
            valueAxis: {
              crossesAt: index === 0 ? 4 : 'autoZero',
              displayUnit,
              displayUnitLabel: index % 2 === 0,
            },
            ...(index === 0 ? {
              categoryAxis: {
                kind: 'date',
                crossesAt: 2,
                minimum: 46_000,
                maximum: 47_000,
                majorUnit: 2,
                minorUnit: 1,
                baseTimeUnit: 'days',
                majorTimeUnit: 'months',
                minorTimeUnit: 'days',
                labelFrequency: 4,
                multiLevelLabels: true,
                numberFormat: 'yyyy-mm-dd',
              },
              seriesAxis: {
                visible: false,
                title: { text: 'Browser series axis', rotation: 25 },
                labelPosition: 'high',
                labelFrequency: 3,
                majorUnit: 2,
                minorUnit: 1,
                numberFormat: '0.00',
                orientation: 'maxMin',
                line: { kind: 'line', color: { kind: 'srgb', value: '112233' } },
                majorGridLine: {
                  kind: 'line', color: { kind: 'srgb', value: '445566' },
                },
              },
            } : {}),
          },
        });
      }
      const duplicateChartSlide = chartDocument.duplicateSlide(chartDocument.slides.length - 1);
      const duplicateChart = duplicateChartSlide.shapes.find(
        (shape) => shape instanceof api.ChartModel,
      );
      const duplicateChartPartUri = duplicateChart.chartPartUri;
      duplicateChart.remove();
      const duplicateChartRemoved = !chartDocument.opcPackage.hasPart(duplicateChartPartUri);
      const chartOutput = await chartDocument.writeBlob({ compatibility: 'powerpoint-2010' });
      const reopenedChartDocument = await api.PptxDocument.open(chartOutput);
      await reopenedChartDocument.write({ compatibility: 'powerpoint-2010' });
      const reopenedCharts = reopenedChartDocument.slides.flatMap(({ shapes }) => shapes)
        .filter((shape) => shape instanceof api.ChartModel);
      const chartWorkbookResults = await Promise.all(reopenedCharts.map((chart) =>
        api.chartWorkbookMatches(
          reopenedChartDocument.opcPackage.requirePart(chart.workbookPartUri).bytes,
          chart.definition,
          chart.xml,
        )));
      const reopenedChartTypes = new Set(reopenedCharts.flatMap(({ definition }) =>
        definition.groups.map(({ type }) => type)));
      const chartIdsUnique = reopenedChartDocument.slides.every((slide) => {
        const ids = slide.shapes.map(({ id }) => id);
        return new Set(ids).size === ids.length;
      });
      const chartOrphanCount = reopenedChartDocument.opcPackage.parts
        .filter(({ contentType }) =>
          contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
          || contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        .filter(({ uri }) =>
          (reopenedChartDocument.opcPackage.graph.find((node) => node.uri === uri)?.incoming.length ?? 0) === 0)
        .length;
      const reopenedAreaChart = reopenedCharts.find(
        ({ definition }) => definition.groups[0]?.type === 'area',
      );
      const chartAreaFillLineState = reopenedAreaChart?.definition.options;
      const chartAreaFillLine = chartAreaFillLineState?.roundedCorners === true
        && chartAreaFillLineState.chartArea?.fill?.kind === 'solid'
        && chartAreaFillLineState.chartArea.fill.color.kind === 'scheme'
        && chartAreaFillLineState.chartArea.fill.color.value === 'accent2'
        && chartAreaFillLineState.chartArea.fill.transparency === 25
        && chartAreaFillLineState.chartArea.line?.kind === 'line'
        && chartAreaFillLineState.chartArea.line.color.kind === 'srgb'
        && chartAreaFillLineState.chartArea.line.color.value === '112233'
        && chartAreaFillLineState.chartArea.line.width === 2
        && chartAreaFillLineState.plotArea?.fill?.kind === 'solid'
        && chartAreaFillLineState.plotArea.fill.color.kind === 'srgb'
        && chartAreaFillLineState.plotArea.fill.color.value === '445566'
        && chartAreaFillLineState.plotArea.fill.transparency === 50
        && chartAreaFillLineState.plotArea.line?.kind === 'line'
        && chartAreaFillLineState.plotArea.line.color.kind === 'scheme'
        && chartAreaFillLineState.plotArea.line.color.value === 'accent1'
        && chartAreaFillLineState.plotArea.line.width === 1;
      const chartAxisFoundationXml = reopenedAreaChart?.xml ?? '';
      const chartAxisFoundation = chartAreaFillLineState?.categoryAxis?.visible === false
        && chartAreaFillLineState.categoryAxis.labelPosition === 'high'
        && chartAreaFillLineState.categoryAxis.labelRotation === -45
        && chartAreaFillLineState.categoryAxis.line?.kind === 'none'
        && chartAreaFillLineState.categoryAxis.majorGridLine?.kind === 'line'
        && chartAreaFillLineState.categoryAxis.majorGridLine.color.kind === 'srgb'
        && chartAreaFillLineState.categoryAxis.majorGridLine.color.value === '667788'
        && chartAreaFillLineState.categoryAxis.majorGridLine.width === 0.75
        && chartAreaFillLineState.categoryAxis.majorGridLine.dash === 'sysDot'
        && chartAreaFillLineState.categoryAxis.majorTickMark === 'inside'
        && chartAreaFillLineState.categoryAxis.minorTickMark === 'outside'
        && chartAreaFillLineState.valueAxis?.visible === false
        && chartAreaFillLineState.valueAxis.labelPosition === 'none'
        && chartAreaFillLineState.valueAxis.labelRotation === 45
        && chartAreaFillLineState.valueAxis.line?.kind === 'line'
        && chartAreaFillLineState.valueAxis.line.color.kind === 'srgb'
        && chartAreaFillLineState.valueAxis.line.color.value === '778899'
        && chartAreaFillLineState.valueAxis.line.width === 1.25
        && chartAreaFillLineState.valueAxis.line.dash === 'dash'
        && chartAreaFillLineState.valueAxis.majorGridLine?.kind === 'none'
        && chartAreaFillLineState.valueAxis.majorTickMark === 'cross'
        && chartAreaFillLineState.valueAxis.minorTickMark === undefined
        && chartAxisFoundationXml.includes('<c:delete val="1"/>')
        && chartAxisFoundationXml.includes('<c:tickLblPos val="high"/>')
        && chartAxisFoundationXml.includes('<c:tickLblPos val="none"/>')
        && chartAxisFoundationXml.includes('<c:majorTickMark val="in"/>')
        && chartAxisFoundationXml.includes('<c:minorTickMark val="out"/>')
        && chartAxisFoundationXml.includes('<c:majorTickMark val="cross"/>')
        && chartAxisFoundationXml.includes('<c:minorTickMark val="none"/>')
        && chartAxisFoundationXml.includes('rot="-2700000"')
        && chartAxisFoundationXml.includes('rot="2700000"')
        && chartAxisFoundationXml.includes('<a:prstDash val="sysDot"/>')
        && chartAxisFoundationXml.includes('<a:prstDash val="dash"/>');
      const reopenedChartAxisAdvanced = reopenedCharts.find(
        ({ name }) => name === 'Browser advanced axis billions',
      );
      const chartAxisAdvancedOptions = reopenedChartAxisAdvanced?.definition.options;
      const chartAxisAdvancedXml = reopenedChartAxisAdvanced?.xml ?? '';
      const chartAxisAdvanced = chartAxisAdvancedOptions?.categoryAxis?.kind === 'date'
        && chartAxisAdvancedOptions.categoryAxis.crossesAt === 2
        && chartAxisAdvancedOptions.categoryAxis.baseTimeUnit === 'days'
        && chartAxisAdvancedOptions.categoryAxis.majorTimeUnit === 'months'
        && chartAxisAdvancedOptions.categoryAxis.minorTimeUnit === 'days'
        && chartAxisAdvancedOptions.categoryAxis.labelFrequency === 4
        && chartAxisAdvancedOptions.categoryAxis.multiLevelLabels === true
        && chartAxisAdvancedOptions.valueAxis?.crossesAt === 4
        && chartAxisAdvancedOptions.valueAxis.displayUnit === 'billions'
        && chartAxisAdvancedOptions.valueAxis.displayUnitLabel === true
        && chartAxisAdvancedOptions.seriesAxis?.visible === false
        && chartAxisAdvancedOptions.seriesAxis.labelPosition === 'high'
        && chartAxisAdvancedOptions.seriesAxis.labelFrequency === 3
        && chartAxisAdvancedOptions.seriesAxis.majorUnit === 2
        && chartAxisAdvancedOptions.seriesAxis.minorUnit === 1
        && chartAxisAdvancedOptions.seriesAxis.orientation === 'maxMin'
        && chartAxisAdvancedXml.includes('<c:dateAx>')
        && chartAxisAdvancedXml.includes('<c:serAx>')
        && chartAxisAdvancedXml.includes('<c:dispUnitsLbl>');
      const chartAxisDisplayUnitCatalog = chartAxisAdvancedDisplayUnits.every((displayUnit) => {
        const chart = reopenedCharts.find(
          ({ name }) => name === `Browser advanced axis ${displayUnit}`,
        );
        return chart?.definition.options.valueAxis?.displayUnit === displayUnit
          && chart.xml.includes(`<c:builtInUnit val="${displayUnit}"/>`);
      });
      const chartPresentationFrame = reopenedAreaChart?.name === 'Browser area chart'
        && reopenedAreaChart.altText === 'area chart with regional business data'
        && reopenedAreaChart.transform.x === api.inches(0.5)
        && reopenedAreaChart.transform.y === api.inches(0.5)
        && reopenedAreaChart.transform.width === api.inches(9)
        && reopenedAreaChart.transform.height === api.inches(6.5);
      const nativeCharts = reopenedCharts.length === 19
        && api.CHART_TYPES.every((type) => reopenedChartTypes.has(type))
        && chartWorkbookResults.every(Boolean)
        && chartPresentationFrame
        && chartAreaFillLine
        && chartAxisFoundation
        && chartAxisAdvanced
        && chartAxisDisplayUnitCatalog
        && chartIdsUnique
        && chartOrphanCount === 0
        && duplicateChartRemoved
        && chartDocument.opcPackage.hasPart(comboChartPartUri)
        && reopenedChartDocument.diagnostics.filter(({ code }) => code.startsWith('CHART_')).length === 0;
      const textRunScalarProperties = [
        'bold',
        'breakLine',
        'color',
        'fontFace',
        'fontSize',
        'highlight',
        'italic',
        'lang',
        'softBreakBefore',
      ];
      const textRunScalarOwners = [
        'PlaceholderProps',
        'SlideNumberProps',
        'TableCellProps',
        'TableProps',
        'TableToSlidesProps',
        'TextPropsOptions',
      ];
      const browserTextRunScalarStyle = {
        fontFamily: 'Aptos Display',
        fontSize: 17.25,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'srgb', value: 'C00000' },
        highlight: { kind: 'scheme', value: 'accent2' },
      };
      const textRunScalarDocument = api.PptxDocument.create();
      const textRunScalarSlide = textRunScalarDocument.addSlide();
      const textRunScalarShape = textRunScalarSlide.addRichText([{
        runs: [
          { text: 'Browser scalar styled', style: browserTextRunScalarStyle },
          {
            text: 'Browser scalar soft',
            softBreakBefore: true,
            style: { bold: false, italic: false, lang: 'fr-CA' },
          },
          { text: 'Browser scalar paragraph', breakLine: true },
          { text: 'Browser scalar tail' },
        ],
      }], { name: 'browser_text_run_scalar_family' });
      const textRunScalarInitial =
        textRunScalarShape.richText.length === 2
        && textRunScalarShape.richText[0].runs[0].style.fontFamily === 'Aptos Display'
        && textRunScalarShape.richText[0].runs[0].style.fontSize === 17.25
        && textRunScalarShape.richText[0].runs[0].style.bold === true
        && textRunScalarShape.richText[0].runs[0].style.italic === true
        && textRunScalarShape.richText[0].runs[0].style.color?.value === 'C00000'
        && textRunScalarShape.richText[0].runs[0].style.highlight?.value === 'accent2'
        && textRunScalarShape.richText[0].runs[1].softBreakBefore === true
        && textRunScalarShape.richText[0].runs[1].style.bold === false
        && textRunScalarShape.richText[0].runs[1].style.italic === false;
      const textRunScalarDuplicate = textRunScalarDocument.duplicateSlide(0);
      textRunScalarShape.richText = [{ runs: [{
        text: 'Browser scalar edited',
        style: {
          fontFamily: 'Courier New',
          fontSize: 24,
          lang: 'ja-JP',
          bold: false,
          italic: false,
          color: { kind: 'scheme', value: 'tx2' },
          highlight: { kind: 'srgb', value: '00FF00' },
        },
      }] }];
      const textRunScalarEdited =
        textRunScalarShape.richText[0].runs[0].style.fontFamily === 'Courier New'
        && textRunScalarShape.richText[0].runs[0].style.fontSize === 24
        && textRunScalarShape.richText[0].runs[0].style.lang === 'ja-JP'
        && textRunScalarShape.richText[0].runs[0].style.bold === false
        && textRunScalarShape.richText[0].runs[0].style.italic === false
        && textRunScalarShape.richText[0].runs[0].style.color?.value === 'tx2'
        && textRunScalarShape.richText[0].runs[0].style.highlight?.value === '00FF00';
      const textRunScalarRollbackBytes = textRunScalarDocument.opcPackage
        .requirePart(textRunScalarSlide.partUri).bytes.slice();
      let textRunScalarRollbackCaught = false;
      try {
        textRunScalarDocument.transaction(() => {
          textRunScalarShape.richText = [{
            runs: [{ text: 'Browser scalar rollback' }],
          }];
          throw new Error('restore browser scalar styles');
        });
      } catch (error) {
        textRunScalarRollbackCaught = error.message === 'restore browser scalar styles';
      }
      const textRunScalarRollback = textRunScalarRollbackCaught
        && tableVerticalAlignmentBytesEqual(
          textRunScalarRollbackBytes,
          textRunScalarDocument.opcPackage.requirePart(textRunScalarSlide.partUri).bytes,
        )
        && textRunScalarShape.richText[0].runs[0].style.fontSize === 24;
      const textRunScalarXml = new TextDecoder().decode(
        textRunScalarDocument.opcPackage.requirePart(textRunScalarSlide.partUri).bytes,
      );
      const textRunScalarDuplicateShape = textRunScalarDuplicate.shapes.find(
        ({ name }) => name === textRunScalarShape.name,
      );
      const textRunScalarDuplicatePreserved =
        textRunScalarDuplicateShape?.richText[0]?.runs[0]?.style?.fontFamily ===
          'Aptos Display'
        && textRunScalarDuplicateShape.richText[0].runs[0].style.fontSize === 17.25
        && textRunScalarDuplicateShape.richText[0].runs[0].style.bold === true
        && textRunScalarDuplicateShape.richText[0].runs[0].style.highlight?.value ===
          'accent2';
      const reopenedTextRunScalarDocument = await api.PptxDocument.open(
        await textRunScalarDocument.writeBlob(),
      );
      const reopenedTextRunScalarSource = reopenedTextRunScalarDocument
        .slides[0].shapes.find(({ name }) => name === textRunScalarShape.name);
      const reopenedTextRunScalarDuplicate = reopenedTextRunScalarDocument
        .slides[1].shapes.find(({ name }) => name === textRunScalarShape.name);
      const textRunScalarReopened =
        reopenedTextRunScalarSource?.richText[0]?.runs[0]?.style?.fontFamily === 'Courier New'
        && reopenedTextRunScalarSource.richText[0].runs[0].style.fontSize === 24
        && reopenedTextRunScalarSource.richText[0].runs[0].style.bold === false
        && reopenedTextRunScalarSource.richText[0].runs[0].style.highlight?.value ===
          '00FF00'
        && reopenedTextRunScalarDuplicate?.richText[0]?.runs[0]?.style?.fontFamily ===
          'Aptos Display'
        && reopenedTextRunScalarDuplicate.richText[0].runs[0].style.highlight?.value ===
          'accent2';
      const textRunScalarFormatStates = [];
      for (const format of ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm']) {
        const formatDocument = api.PptxDocument.create({ format });
        formatDocument.addSlide().addRichText([{
          runs: [{ text: `Browser scalar ${format}`, style: browserTextRunScalarStyle }],
        }]);
        const reopenedFormat = await api.PptxDocument.open(
          await formatDocument.writeBlob(),
        );
        textRunScalarFormatStates.push(
          reopenedFormat.format === format
          && reopenedFormat.slides[0].shapes[0].richText[0].runs[0].style.fontSize ===
            17.25
          && reopenedFormat.slides[0].shapes[0].richText[0].runs[0].style.highlight
            ?.value === 'accent2'
          && reopenedFormat.diagnostics.length === 0,
        );
      }
      const textRunScalarFamilyState = {
        catalog: textRunScalarProperties.length === 9 && textRunScalarOwners.length === 6,
        owners: richTextBreakLine && tableTextDefaults && tableToSlides
          && slideNumberState.diagnostics.length === 0,
        initial: textRunScalarInitial,
        edited: textRunScalarEdited,
        rollback: textRunScalarRollback,
        duplicate: textRunScalarDuplicatePreserved,
        reopened: textRunScalarReopened,
        sixFormats: textRunScalarFormatStates.every(Boolean),
        xml: textRunScalarXml.includes('lang="ja-JP"')
          && textRunScalarXml.includes('sz="2400"')
          && textRunScalarXml.includes('b="0" i="0"')
          && textRunScalarXml.includes('<a:schemeClr val="tx2"/>')
          && textRunScalarXml.includes(
            '<a:highlight><a:srgbClr val="00FF00"/></a:highlight>',
          )
          && textRunScalarXml.includes('<a:latin typeface="Courier New"/>'),
        diagnostics: textRunScalarDocument.diagnostics.length === 0
          && reopenedTextRunScalarDocument.diagnostics.length === 0,
      };
      const textRunScalarFamily = Object.values(textRunScalarFamilyState).every(Boolean);
      return {
        textRunScalarFamily,
        textRunScalarFamilyState,
        presentationVersion,
        presentationVersionState,
        presentationLayouts,
        presentationLayoutState,
        horizontalAlignments,
        horizontalAlignmentState,
        verticalAlignments,
        verticalAlignmentState,
        underlineFamily,
        underlineFamilyState,
        tabStops,
        tabStopsState,
        bulletNumbering,
        bulletNumberingState,
        presetShapes,
        presetShapeState,
        browserCustomPaths,
        browserCustomPathState,
        shapeLines,
        shapeLineState,
        shapeArrows,
        shapeArrowState,
        tableVerticalAlignment,
        tableVerticalAlignmentState,
        textDirectionFamily,
        textDirectionFamilyState,
        textBoxFitFamily,
        textBoxFitFamilyState,
        textParagraphLayoutFamily,
        textParagraphLayoutFamilyState,
        richTextEffectsFamily,
        richTextEffectsFamilyState,
        tableTextDirection,
        tableTextDirectionState,
        tableHorizontalAlignment,
        tableHorizontalAlignmentState,
        tableMargins,
        tableMarginsState,
        tableCellHyperlinks,
        tableCellHyperlinkEditing,
        tableCellHyperlinkState,
        tableBorders,
        tableBordersState,
        tableFill,
        tableFillState,
        tableTextDefaults,
        tableTextDefaultsState,
        tableCellMerges,
        tableCellMergesState,
        tableStructureEditing,
        tableStructureEditingState,
        tableAutoPage,
        tableAutoPageState,
        tableContentMeasurement,
        tableContentMeasurementState,
        tableToSlides,
        tableToSlidesState,
        schemeColors,
        schemeColorState,
        outputTypes,
        outputTypeState,
        writeOutputTypes,
        writeOutputTypeState,
        nodeReadableStream,
        nodeReadableStreamState,
        compressionPolicy,
        compressionPolicyState,
        format: reopened.format,
        title: reopened.slides[0].title.text,
        mime: output.type,
        transition: typeof api.transitions.TransitionCodec,
        smartArt: typeof api.smartArt.SmartArtDiagramCodec,
        blobInputTitle: fromBlob.slides[0].title.text,
        slideNumbers: slideNumberState,
        masterLayouts: masterLayoutState,
        slideDefaultColor: slideDefaultColorState,
        shapeFills,
        shapeFillState,
        textShapeFills: textShapeFillState,
        textShapeLines: textShapeLineState,
        textShapeArrows: textShapeArrowState,
        textShapeShadows: textShapeShadowState,
        textShapePresetGeometry,
        textShapePresetGeometryState,
        textShapeRectRadius,
        textShapeRectRadiusState,
        textShapeIsTextBox,
        textShapeIsTextBoxState,
        richTextBreakLine,
        richTextBreakLineState,
        textShapeHyperlinks,
        textShapeHyperlinkState,
        richTextRunHyperlinks,
        richTextRunHyperlinkState,
        svgCreatedLive: imageSourceSizingTransformState.live,
        svgState,
        imageSourceSizingTransformState,
        backgroundMime: backgroundOutput.type,
        slideBackgroundKinds: reopenedBackgroundDocument.slides.map(
          ({ background }) => background?.kind,
        ),
        backgroundPayloadHashes,
        backgroundRelationshipCounts,
        backgroundValidationErrors: reopenedBackgroundDocument.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
        mediaMime: mediaOutput.type,
        mediaNames,
        mediaElementCounts: {
          audio: (mediaXml.match(/<a:audioFile\b/g) ?? []).length,
          video: (mediaXml.match(/<a:videoFile\b/g) ?? []).length,
        },
        mediaState,
        browserOnlineVideoState,
        mediaValidationErrors: mediaDocument.diagnostics.filter(
          ({ severity }) => severity === 'error',
        ).length,
        mediaTimingElementCounts: timingSummaries.map(({ timing, media, commands, playback }) => ({
          timing,
          media,
          commands,
          playback,
        })),
        mediaTimingUniqueIdCount: timingSummaries.map(({ uniqueIds }) => uniqueIds),
        settingsAfterReopen,
        timingDiagnostics,
        nativeMediaTiming,
        nativeCharts,
        imageIdentityEffects5,
        imageIdentityEffects5State,
        chartPresentation91,
        chartPresentation91State,
        chartPresentation: chartPresentation91,
        chartPresentationState: chartPresentation91State,
        chartAreaFillLine,
        chartAxisAdvanced,
        chartAxisDisplayUnitCatalog,
        stableMediaLifecycle,
        mediaTargetIsolation: browserMediaCloneOnWrite,
        mediaOrphanCount,
      };
    },
    {
      moduleUrl: 'http://127.0.0.1:4173/packages/pptx/dist/browser.js',
      chartPresentationFixtureBase64,
      chartPresentationProbeBase64,
      imageIdentityEffects5ProbeBase64,
      base64: 'UEsDBAoAAAAIAOMg/FxMagnj0QAAAP0BAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK1RvU7DQAx+lejWqnHpwICaLsBKGXgB6+I0J+7HOrtVeXuctEiACixMlv39St68vDFJc0oxS+dGVb4DED9SQmkLUzZkKDWh2lr3wOhfcU+wXq1uwZeslHWpk4fbbh5owEPU5vFkZwkld65SFNfcn4lTVueQOQaPajgcc/8tZXlJaE05c2QMLAsjOLiaMCE/B1x0uyPVGnpqnrHqEyZjAbMCVxLTzdz2d6crVcswBE998YdkkvazWYpf1jZhyIs/yki0o5zHzX+3mV0/GsD89e07UEsDBAoAAAAAAOMg/FwAAAAAAAAAAAAAAAAGAAAAX3JlbHMvUEsDBAoAAAAIAOMg/Fwvm14oigAAAPUAAAALAAAAX3JlbHMvLnJlbHONzz0OwjAMBeCrVDlAXRgYUJKJpSvqBaLU+RFNYiVGgtsTMRXEwOjnp8+yvOJmOJbcQqQ2PNKWmxKBmc4AzQZMpo2FMPeNKzUZ7mP1QMbejEc4TtMJ6t4QWu7NYV6VqPN6EMPyJPzHLs5Fi5di7wkz/zjx1eiyqR5ZCSIGqth6+G6PXRagJXx8qV9QSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAQAAABwcHQvUEsDBAoAAAAIAOMg/FzLe24cTgAAAHEAAAAUAAAAcHB0L3ByZXNlbnRhdGlvbi54bWyzKbAqKEotTs0rSSzJzM9TqMjNySu2KrBVKlCCsotslYqU7GwKrIpzUjxTfIpL4GyFzBRbJSNTMyWFIisQs8gzxVBJ385GH1mtPqoFdgBQSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAoAAABwcHQvX3JlbHMvUEsDBAoAAAAIAOMg/Fw2SaGViAAAAOkAAAAfAAAAcHB0L19yZWxzL3ByZXNlbnRhdGlvbi54bWwucmVsc43PPQoCMRAF4KssOcDOroWFJKlsthUvEJLJD+aPTAS9vUEsVrCwfPPgGx6/YFQ9lEw+VJoeKWYSzPdeTwCkPSZFc6mYR2NLS6qP2BxUpW/KIRyW5QhtbzDJ9+a0GcHaZlY2XZ8V/7GLtUHjueh7wtx/vACKweAAVXPYBXvHz3Wdh8ZAcvhaJl9QSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAsAAABwcHQvc2xpZGVzL1BLAwQKAAAACADjIPxc5NE7A5MAAAD3AAAAFQAAAHBwdC9zbGlkZXMvc2xpZGUxLnhtbE2PUQrDIAyGryK5QGCPoj70AKPQXkCmYwXbhug6e/tNnWwvX0L+Pz+JIhmDE3kNW5SkgeDbWw0WjCJ5m4IrNdLM3reucDsmGrk6rsfIYnEaLiA2u3oN85KCB2y+5qKHSCd9tNQ17CL+p6U87O40ykoq4IJkBt5f0bO4Lzk92Sssw0KupBrSV7HdiL+jsf+B9V/zBlBLAQIUAAoAAAAIAOMg/FxMagnj0QAAAP0BAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAhQACgAAAAAA4yD8XAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAQAAAAAgEAAF9yZWxzL1BLAQIUAAoAAAAIAOMg/Fwvm14oigAAAPUAAAALAAAAAAAAAAAAAAAAACYBAABfcmVscy8ucmVsc1BLAQIUAAoAAAAAAOMg/FwAAAAAAAAAAAAAAAAEAAAAAAAAAAAAEAAAANkBAABwcHQvUEsBAhQACgAAAAgA4yD8XMt7bhxOAAAAcQAAABQAAAAAAAAAAAAAAAAA+wEAAHBwdC9wcmVzZW50YXRpb24ueG1sUEsBAhQACgAAAAAA4yD8XAAAAAAAAAAAAAAAAAoAAAAAAAAAAAAQAAAAewIAAHBwdC9fcmVscy9QSwECFAAKAAAACADjIPxcNkmhlYgAAADpAAAAHwAAAAAAAAAAAAAAAACjAgAAcHB0L19yZWxzL3ByZXNlbnRhdGlvbi54bWwucmVsc1BLAQIUAAoAAAAAAOMg/FwAAAAAAAAAAAAAAAALAAAAAAAAAAAAEAAAAGgDAABwcHQvc2xpZGVzL1BLAQIUAAoAAAAIAOMg/Fzk0TsDkwAAAPcAAAAVAAAAAAAAAAAAAAAAAJEDAABwcHQvc2xpZGVzL3NsaWRlMS54bWxQSwUGAAAAAAkACQAjAgAAVwQAAAAA',
    },
  );
  result.imageIdentityEffects5EvidenceFileName =
    'browser-image-identity-effects-5.pptx';
  const imageIdentityEffects5EvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_IMAGE_IDENTITY_EFFECTS_OUT
      ?? globalThis.__pptxBrowserImageIdentityEffectsOutput
    : globalThis.__pptxBrowserImageIdentityEffectsOutput;
  if (typeof imageIdentityEffects5EvidenceOutput === 'string') {
    const imageIdentityEffects5EvidenceBytes = await page.evaluate(async () => {
      const blob = globalThis.__pptxImageIdentityEffects5EvidenceBlob;
      if (!(blob instanceof Blob)) {
        throw new Error('Missing image identity/effects evidence Blob');
      }
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });
    await writeFile(
      imageIdentityEffects5EvidenceOutput,
      Uint8Array.from(imageIdentityEffects5EvidenceBytes),
    );
  }
  result.chartPresentation91EvidenceFileName = 'browser-chart-presentation-91.pptx';
  result.chartPresentationEvidenceFileName = result.chartPresentation91EvidenceFileName;
  const chartPresentation91EvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_CHART_PRESENTATION_OUT
      ?? globalThis.__pptxBrowserChartPresentationOutput
    : globalThis.__pptxBrowserChartPresentationOutput;
  if (typeof chartPresentation91EvidenceOutput === 'string') {
    const chartPresentation91EvidenceBytes = await page.evaluate(async () => {
      const blob = globalThis.__pptxChartPresentation91EvidenceBlob;
      if (!(blob instanceof Blob)) throw new Error('Missing chart-presentation evidence Blob');
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });
    await writeFile(
      chartPresentation91EvidenceOutput,
      Uint8Array.from(chartPresentation91EvidenceBytes),
    );
  }
  const textBoxFitEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTextBoxFitEvidenceBlob;
    if (!(blob instanceof Blob)) throw new Error('Missing text-box fit evidence Blob');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-text-box-fit.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const textBoxFitEvidenceDownload = await textBoxFitEvidenceDownloadPromise;
  result.textBoxFitEvidenceFileName = textBoxFitEvidenceDownload.suggestedFilename();
  const textBoxFitEvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_TEXT_BOX_FIT_OUT ??
      globalThis.__pptxTextBoxFitEvidenceOutput
    : globalThis.__pptxTextBoxFitEvidenceOutput;
  if (typeof textBoxFitEvidenceOutput === 'string') {
    await textBoxFitEvidenceDownload.saveAs(textBoxFitEvidenceOutput);
  }
  const textParagraphLayoutEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTextParagraphLayoutEvidenceBlob;
    if (!(blob instanceof Blob)) {
      throw new Error('Missing text paragraph/layout evidence Blob');
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-text-paragraph-layout.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const textParagraphLayoutEvidenceDownload =
    await textParagraphLayoutEvidenceDownloadPromise;
  result.textParagraphLayoutEvidenceFileName =
    textParagraphLayoutEvidenceDownload.suggestedFilename();
  const textParagraphLayoutEvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_TEXT_PARAGRAPH_LAYOUT_OUT ??
      globalThis.__pptxTextParagraphLayoutEvidenceOutput
    : globalThis.__pptxTextParagraphLayoutEvidenceOutput;
  if (typeof textParagraphLayoutEvidenceOutput === 'string') {
    await textParagraphLayoutEvidenceDownload.saveAs(textParagraphLayoutEvidenceOutput);
  }
  const richTextEffectsEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxRichTextEffectsEvidenceBlob;
    if (!(blob instanceof Blob)) {
      throw new Error('Missing rich-text effects evidence Blob');
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-rich-text-effects.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const richTextEffectsEvidenceDownload =
    await richTextEffectsEvidenceDownloadPromise;
  result.richTextEffectsEvidenceFileName =
    richTextEffectsEvidenceDownload.suggestedFilename();
  const richTextEffectsEvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_RICH_TEXT_EFFECTS_OUT ??
      globalThis.__pptxRichTextEffectsEvidenceOutput
    : globalThis.__pptxRichTextEffectsEvidenceOutput;
  if (typeof richTextEffectsEvidenceOutput === 'string') {
    await richTextEffectsEvidenceDownload.saveAs(richTextEffectsEvidenceOutput);
  }
  const customPathEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxCustomPathEvidenceBlob;
    if (!(blob instanceof Blob)) throw new Error('Missing custom-path evidence Blob');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-custom-paths.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const customPathEvidenceDownload = await customPathEvidenceDownloadPromise;
  result.customPathEvidenceFileName = customPathEvidenceDownload.suggestedFilename();
  const customPathEvidenceOutput = typeof process !== 'undefined'
    ? process.env.PPTX_BROWSER_CUSTOM_PATH_OUT
    : globalThis.__pptxCustomPathEvidenceOutput;
  if (typeof customPathEvidenceOutput === 'string') {
    await customPathEvidenceDownload.saveAs(customPathEvidenceOutput);
  }
  const tableCellMergesEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTableCellMergesEvidenceBlob;
    if (!(blob instanceof Blob)) throw new Error('Missing table-cell merge evidence Blob');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-table-cell-merges.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const tableCellMergesEvidenceDownload = await tableCellMergesEvidenceDownloadPromise;
  result.tableCellMergesEvidenceFileName =
    tableCellMergesEvidenceDownload.suggestedFilename();
  if (typeof process !== 'undefined' && process.env.PPTX_BROWSER_TABLE_CELL_MERGES_OUT) {
    await tableCellMergesEvidenceDownload.saveAs(
      process.env.PPTX_BROWSER_TABLE_CELL_MERGES_OUT,
    );
  }
  const tableStructureEditingEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTableStructureEditingEvidenceBlob;
    if (!(blob instanceof Blob)) throw new Error('Missing table structure editing evidence Blob');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-table-structure-editing.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const tableStructureEditingEvidenceDownload =
    await tableStructureEditingEvidenceDownloadPromise;
  result.tableStructureEditingEvidenceFileName =
    tableStructureEditingEvidenceDownload.suggestedFilename();
  if (typeof process !== 'undefined' && process.env.PPTX_BROWSER_TABLE_STRUCTURE_EDITING_OUT) {
    await tableStructureEditingEvidenceDownload.saveAs(
      process.env.PPTX_BROWSER_TABLE_STRUCTURE_EDITING_OUT,
    );
  }
  const tableAutoPageEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTableAutoPageEvidenceBlob;
    if (!(blob instanceof Blob)) throw new Error('Missing table auto-page evidence Blob');
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-table-auto-page.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const tableAutoPageEvidenceDownload = await tableAutoPageEvidenceDownloadPromise;
  result.tableAutoPageEvidenceFileName =
    tableAutoPageEvidenceDownload.suggestedFilename();
  if (typeof process !== 'undefined' && process.env.PPTX_BROWSER_TABLE_AUTO_PAGE_OUT) {
    await tableAutoPageEvidenceDownload.saveAs(
      process.env.PPTX_BROWSER_TABLE_AUTO_PAGE_OUT,
    );
  }
  const tableContentMeasurementEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTableContentMeasurementEvidenceBlob;
    if (!(blob instanceof Blob)) {
      throw new Error('Missing table content measurement evidence Blob');
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-table-content-measurement.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const tableContentMeasurementEvidenceDownload =
    await tableContentMeasurementEvidenceDownloadPromise;
  result.tableContentMeasurementEvidenceFileName =
    tableContentMeasurementEvidenceDownload.suggestedFilename();
  if (typeof process !== 'undefined' &&
      process.env.PPTX_BROWSER_TABLE_CONTENT_MEASUREMENT_OUT) {
    await tableContentMeasurementEvidenceDownload.saveAs(
      process.env.PPTX_BROWSER_TABLE_CONTENT_MEASUREMENT_OUT,
    );
  }
  const tableToSlidesEvidenceDownloadPromise = page.waitForEvent('download');
  await page.evaluate(() => {
    const blob = globalThis.__pptxTableToSlidesEvidenceBlob;
    if (!(blob instanceof Blob)) {
      throw new Error('Missing HTML table slide evidence Blob');
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'browser-table-to-slides.pptx';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  });
  const tableToSlidesEvidenceDownload = await tableToSlidesEvidenceDownloadPromise;
  result.tableToSlidesEvidenceFileName =
    tableToSlidesEvidenceDownload.suggestedFilename();
  if (typeof process !== 'undefined' &&
      process.env.PPTX_BROWSER_TABLE_TO_SLIDES_OUT) {
    await tableToSlidesEvidenceDownload.saveAs(
      process.env.PPTX_BROWSER_TABLE_TO_SLIDES_OUT,
    );
  }
  const downloadPromise = page.waitForEvent('download');
  await page.evaluate(
    async ({ moduleUrl, base64 }) => {
      const binary = atob(base64);
      const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
      const api = await import(moduleUrl);
      const document = await api.PptxDocument.open(bytes);
      await document.download('browser-smoke.pptx');
    },
    {
      moduleUrl: 'http://127.0.0.1:4173/packages/pptx/dist/browser.js',
      base64: 'UEsDBAoAAAAIAOMg/FxMagnj0QAAAP0BAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK1RvU7DQAx+lejWqnHpwICaLsBKGXgB6+I0J+7HOrtVeXuctEiACixMlv39St68vDFJc0oxS+dGVb4DED9SQmkLUzZkKDWh2lr3wOhfcU+wXq1uwZeslHWpk4fbbh5owEPU5vFkZwkld65SFNfcn4lTVueQOQaPajgcc/8tZXlJaE05c2QMLAsjOLiaMCE/B1x0uyPVGnpqnrHqEyZjAbMCVxLTzdz2d6crVcswBE998YdkkvazWYpf1jZhyIs/yki0o5zHzX+3mV0/GsD89e07UEsDBAoAAAAAAOMg/FwAAAAAAAAAAAAAAAAGAAAAX3JlbHMvUEsDBAoAAAAIAOMg/Fwvm14oigAAAPUAAAALAAAAX3JlbHMvLnJlbHONzz0OwjAMBeCrVDlAXRgYUJKJpSvqBaLU+RFNYiVGgtsTMRXEwOjnp8+yvOJmOJbcQqQ2PNKWmxKBmc4AzQZMpo2FMPeNKzUZ7mP1QMbejEc4TtMJ6t4QWu7NYV6VqPN6EMPyJPzHLs5Fi5di7wkz/zjx1eiyqR5ZCSIGqth6+G6PXRagJXx8qV9QSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAQAAABwcHQvUEsDBAoAAAAIAOMg/FzLe24cTgAAAHEAAAAUAAAAcHB0L3ByZXNlbnRhdGlvbi54bWyzKbAqKEotTs0rSSzJzM9TqMjNySu2KrBVKlCCsotslYqU7GwKrIpzUjxTfIpL4GyFzBRbJSNTMyWFIisQs8gzxVBJ385GH1mtPqoFdgBQSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAoAAABwcHQvX3JlbHMvUEsDBAoAAAAIAOMg/Fw2SaGViAAAAOkAAAAfAAAAcHB0L19yZWxzL3ByZXNlbnRhdGlvbi54bWwucmVsc43PPQoCMRAF4KssOcDOroWFJKlsthUvEJLJD+aPTAS9vUEsVrCwfPPgGx6/YFQ9lEw+VJoeKWYSzPdeTwCkPSZFc6mYR2NLS6qP2BxUpW/KIRyW5QhtbzDJ9+a0GcHaZlY2XZ8V/7GLtUHjueh7wtx/vACKweAAVXPYBXvHz3Wdh8ZAcvhaJl9QSwMECgAAAAAA4yD8XAAAAAAAAAAAAAAAAAsAAABwcHQvc2xpZGVzL1BLAwQKAAAACADjIPxc5NE7A5MAAAD3AAAAFQAAAHBwdC9zbGlkZXMvc2xpZGUxLnhtbE2PUQrDIAyGryK5QGCPoj70AKPQXkCmYwXbhug6e/tNnWwvX0L+Pz+JIhmDE3kNW5SkgeDbWw0WjCJ5m4IrNdLM3reucDsmGrk6rsfIYnEaLiA2u3oN85KCB2y+5qKHSCd9tNQ17CL+p6U87O40ykoq4IJkBt5f0bO4Lzk92Sssw0KupBrSV7HdiL+jsf+B9V/zBlBLAQIUAAoAAAAIAOMg/FxMagnj0QAAAP0BAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAhQACgAAAAAA4yD8XAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAQAAAAAgEAAF9yZWxzL1BLAQIUAAoAAAAIAOMg/Fwvm14oigAAAPUAAAALAAAAAAAAAAAAAAAAACYBAABfcmVscy8ucmVsc1BLAQIUAAoAAAAAAOMg/FwAAAAAAAAAAAAAAAAEAAAAAAAAAAAAEAAAANkBAABwcHQvUEsBAhQACgAAAAgA4yD8XMt7bhxOAAAAcQAAABQAAAAAAAAAAAAAAAAA+wEAAHBwdC9wcmVzZW50YXRpb24ueG1sUEsBAhQACgAAAAAA4yD8XAAAAAAAAAAAAAAAAAoAAAAAAAAAAAAQAAAAewIAAHBwdC9fcmVscy9QSwECFAAKAAAACADjIPxcNkmhlYgAAADpAAAAHwAAAAAAAAAAAAAAAACjAgAAcHB0L19yZWxzL3ByZXNlbnRhdGlvbi54bWwucmVsc1BLAQIUAAoAAAAAAOMg/FwAAAAAAAAAAAAAAAALAAAAAAAAAAAAEAAAAGgDAABwcHQvc2xpZGVzL1BLAQIUAAoAAAAIAOMg/Fzk0TsDkwAAAPcAAAAVAAAAAAAAAAAAAAAAAJEDAABwcHQvc2xpZGVzL3NsaWRlMS54bWxQSwUGAAAAAAkACQAjAgAAVwQAAAAA',
    },
  );
  result.downloadFileName = (await downloadPromise).suggestedFilename();
  const compressionDownloadBytes = Uint8Array.from(await page.evaluate(async (moduleUrl) => {
    const api = await import(moduleUrl);
    const document = api.PptxDocument.create();
    document.addSlide().addText('Chrome compression download');
    document.opcPackage.setPart(
      '/custom/chrome-download-compression.bin',
      new Uint8Array(65_536).fill(0x41),
      'application/octet-stream',
    );
    const blob = await document.writeBlob({ compression: true });
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  }, 'http://127.0.0.1:4173/packages/pptx/dist/browser.js'));
  const compressionDownloadMethods = (() => {
    const view = new DataView(
      compressionDownloadBytes.buffer,
      compressionDownloadBytes.byteOffset,
      compressionDownloadBytes.byteLength,
    );
    let eocd = compressionDownloadBytes.byteLength - 22;
    while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
    if (eocd < 0) throw new Error('Compression download EOCD not found');
    const entries = view.getUint16(eocd + 10, true);
    let offset = view.getUint32(eocd + 16, true);
    const methods = [];
    for (let index = 0; index < entries; index += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) {
        throw new Error('Compression download central directory entry not found');
      }
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const name = String.fromCharCode(
        ...compressionDownloadBytes.subarray(offset + 46, offset + 46 + nameLength),
      );
      if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return [...new Set(methods)];
  })();
  const compressionDownloadReopenTitle = await page.evaluate(
    async ({ moduleUrl, bytes }) => {
      const api = await import(moduleUrl);
      const reopened = await api.PptxDocument.open(Uint8Array.from(bytes));
      const shape = reopened.slides[0].shapes[0];
      return shape instanceof api.ShapeModel ? shape.text : undefined;
    },
    {
      moduleUrl: 'http://127.0.0.1:4173/packages/pptx/dist/browser.js',
      bytes: [...compressionDownloadBytes],
    },
  );
  result.compressionPolicyState = {
    ...result.compressionPolicyState,
    downloadFileName: 'compression-policy.pptx',
    downloadMethods: compressionDownloadMethods,
    downloadReopenTitle: compressionDownloadReopenTitle,
  };
  result.compressionPolicy = result.compressionPolicy &&
    result.downloadFileName === 'browser-smoke.pptx' &&
    JSON.stringify(compressionDownloadMethods) === JSON.stringify([8]) &&
    compressionDownloadReopenTitle === 'Chrome compression download';
  result.errorCounts = {
    console: consoleErrors.length,
    page: pageErrors.length,
    network: networkErrors.length,
  };
  page.off('console', onConsole);
  page.off('pageerror', onPageError);
  page.off('requestfailed', onRequestFailed);
  page.off('response', onResponse);
  const expected = {
    textRunScalarFamily: true,
    textRunScalarFamilyState: {
      catalog: true,
      owners: true,
      initial: true,
      edited: true,
      rollback: true,
      duplicate: true,
      reopened: true,
      sixFormats: true,
      xml: true,
      diagnostics: true,
    },
    presentationVersion: true,
    presentationVersionState: {
      constant: '0.1.0',
      created: '0.1.0',
      reopened: '0.1.0',
    },
    presentationLayouts: true,
    presentationLayoutState: {
      standard: { name: 'screen4x3', width: 9_144_000, height: 6_858_000 },
      custom: { name: 'custom', width: 10_698_480, height: 7_589_520 },
      edited: { name: 'custom', width: 12_192_000, height: 6_858_000 },
      reopened: { name: 'custom', width: 12_192_000, height: 6_858_000 },
    },
    horizontalAlignments: true,
    horizontalAlignmentState: {
      values: ['left', 'center', 'right', 'justify'],
      reopened: ['left', 'center', 'right', 'justify'],
      frozen: true,
    },
    verticalAlignments: true,
    verticalAlignmentState: {
      values: ['top', 'middle', 'bottom'],
      textReopened: ['top', 'middle', 'bottom'],
      tableReopened: ['top', 'middle', 'bottom'],
      frozen: true,
    },
    underlineFamily: true,
    underlineFamilyState: {
      catalog: true,
      trueAsSingle: true,
      immediate: true,
      initialOoxml: true,
      snapshotDetached: true,
      invalidRollback: true,
      editedOoxml: true,
      reopenedOwners: true,
      duplicateIsolation: true,
      diagnostics: true,
    },
    tabStops: true,
    tabStopsState: {
      snapshotDetached: true,
      immediate: true,
      initialOoxml: true,
      invalidRollback: true,
      editedOoxml: true,
      reopenedText: true,
      reopenedPlaceholder: true,
      reopenedTable: true,
      diagnostics: true,
    },
    bulletNumbering: true,
    bulletNumberingState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      catalog: [
        'alphaLcParenBoth',
        'alphaLcParenR',
        'alphaLcPeriod',
        'alphaUcParenBoth',
        'alphaUcParenR',
        'alphaUcPeriod',
        'arabicParenBoth',
        'arabicParenR',
        'arabicPeriod',
        'arabicPlain',
        'romanLcParenBoth',
        'romanLcParenR',
        'romanLcPeriod',
        'romanUcParenBoth',
        'romanUcParenR',
        'romanUcPeriod',
      ],
      immediateCatalog: [
        'alphaLcParenBoth',
        'alphaLcParenR',
        'alphaLcPeriod',
        'alphaUcParenBoth',
        'alphaUcParenR',
        'alphaUcPeriod',
        'arabicParenBoth',
        'arabicParenR',
        'arabicPeriod',
        'arabicPlain',
        'romanLcParenBoth',
        'romanLcParenR',
        'romanLcPeriod',
        'romanUcParenBoth',
        'romanUcParenR',
        'romanUcPeriod',
      ],
      reopenedCatalog: [
        'alphaLcParenBoth',
        'alphaLcParenR',
        'alphaLcPeriod',
        'alphaUcParenBoth',
        'alphaUcParenR',
        'alphaUcPeriod',
        'arabicParenBoth',
        'arabicParenR',
        'arabicPeriod',
        'arabicPlain',
        'romanLcParenBoth',
        'romanLcParenR',
        'romanLcPeriod',
        'romanUcParenBoth',
        'romanUcParenR',
        'romanUcPeriod',
      ],
      initial: {
        text: [
          { kind: 'bullet', character: '◆', indent: 18 },
          { kind: 'number', style: 'romanUcPeriod', startAt: 7, indent: 22 },
          { kind: 'bullet', character: '•', indent: 27 },
        ],
        placeholder: { kind: 'bullet', character: '→', indent: 19 },
        table: [
          { kind: 'bullet', character: '→', indent: 17 },
          { kind: 'number', style: 'arabicPlain', startAt: 6, indent: 19 },
        ],
      },
      edited: {
        text: [
          { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
          null,
          { kind: 'bullet', character: '•', indent: 27 },
        ],
        placeholder: { kind: 'number', style: 'alphaLcParenR', startAt: 5, indent: 20 },
        table: [
          { kind: 'bullet', character: '★', indent: 21 },
          null,
        ],
      },
      reopened: {
        text: [
          { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
          null,
          { kind: 'bullet', character: '•', indent: 27 },
        ],
        placeholder: { kind: 'number', style: 'alphaLcParenR', startAt: 5, indent: 20 },
        table: [
          { kind: 'bullet', character: '★', indent: 21 },
          null,
        ],
      },
      exactOoxml: {
        catalog: true,
        initialText: true,
        initialPlaceholder: true,
        initialTable: true,
        editedText: true,
        editedPlaceholder: true,
        editedTable: true,
        reopened: true,
      },
      diagnostics: 0,
    },
    presetShapes: true,
    presetShapeState: {
      count: 178,
      unique: true,
      frozen: true,
      canonicalBoundary: true,
      created: true,
      serialized: true,
      reopened: true,
      validationErrors: 0,
    },
    browserCustomPaths: true,
    browserCustomPathState: {
      live: true,
      detached: true,
      frozen: true,
      serialized: true,
      invalidRollback: true,
      reopened: true,
      validationErrors: 0,
    },
    shapeLines: true,
    shapeLineState: {
      immediate: ['solid', 'dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot'],
      serialized: ['solid', 'dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot'],
      reopened: ['solid', 'dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot'],
      preservedArrows: { begin: 'triangle', end: 'arrow' },
      edited: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
        width: 3,
        dash: 'sysDash',
      },
      reopenedEdited: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
        width: 3,
        dash: 'sysDash',
      },
      validationErrors: 0,
    },
    shapeArrows: true,
    shapeArrowState: {
      immediateBegins: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      immediateEnds: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      serializedBegins: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      serializedEnds: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      reopenedBegins: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      reopenedEnds: ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'],
      preservedLine: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
        width: 3,
        dash: 'sysDash',
      },
      reopenedEdited: { begin: 'diamond' },
      validationErrors: 0,
    },
    tableVerticalAlignment: true,
    tableVerticalAlignmentState: {
      uniform: 'middle',
      readIsolation: true,
      noOp: true,
      mixed: null,
      overwritten: 'bottom',
      overwrittenCells: ['bottom', 'bottom', 'bottom', 'bottom'],
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: 'top',
      reopenedCells: ['top', 'top', 'top', 'top'],
      invalidError: {
        name: 'TypeError',
        message: 'Table vertical alignment must be top, middle, or bottom',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    textDirectionFamily: true,
    textDirectionFamilyState: {
      table: true,
      catalog: true,
      initialOoxml: true,
      invalidRollback: true,
      editedOoxml: true,
      reopened: true,
      duplicateIsolation: true,
      diagnostics: true,
    },
    textBoxFitFamily: true,
    textBoxFitFamilyState: {
      initial: true,
      initialOoxml: true,
      invalidRollback: true,
      editedOoxml: true,
      reopened: true,
      duplicateIsolation: true,
      mime: true,
      diagnostics: true,
    },
    textParagraphLayoutFamily: true,
    textParagraphLayoutFamilyState: {
      immediate: true,
      initialOoxml: true,
      catalog: true,
      catalogOoxml: true,
      invalidRollback: true,
      edited: true,
      editedOoxml: true,
      reopened: true,
      duplicateIsolation: true,
      mime: true,
      diagnostics: true,
    },
    richTextEffectsFamily: true,
    richTextEffectsFamilyState: {
      immediate: true,
      initialOoxml: true,
      catalog: true,
      catalogOoxml: true,
      invalidRollback: true,
      edited: true,
      editedOoxml: true,
      reopened: true,
      duplicateIsolation: true,
      mime: true,
      diagnostics: true,
    },
    tableTextDirection: true,
    tableTextDirectionState: {
      uniform: 'vert270',
      readIsolation: true,
      noOp: true,
      mixed: null,
      overwritten: 'wordArtVert',
      overwrittenCells: ['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert'],
      horizontal: 'horz',
      horizontalCells: ['horz', 'horz', 'horz', 'horz'],
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: 'vert',
      reopenedCells: ['vert', 'vert', 'vert', 'vert'],
      invalidError: {
        name: 'TypeError',
        message: 'Table text direction must be horz, vert, vert270, or wordArtVert',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableHorizontalAlignment: true,
    tableHorizontalAlignmentState: {
      uniform: 'center',
      readIsolation: true,
      noOp: true,
      mixed: null,
      overwritten: 'justify',
      overwrittenCells: ['justify', 'justify', 'justify', 'justify'],
      explicitLeft: 'left',
      explicitLeftCells: ['left', 'left', 'left', 'left'],
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: 'right',
      reopenedCells: ['right', 'right', 'right', 'right'],
      invalidError: {
        name: 'TypeError',
        message: 'Table horizontal alignment must be left, center, right, or justify',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableMargins: true,
    tableMarginsState: {
      uniform: { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      readIsolation: true,
      noOp: true,
      mixed: null,
      overwritten: { top: 6, right: 6, bottom: 6, left: 6 },
      overwrittenCells: Array(4).fill({ top: 6, right: 6, bottom: 6, left: 6 }),
      partial: { top: 2, left: 4 },
      partialCells: Array(4).fill({ top: 2, left: 4 }),
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: { top: 1, right: 2, bottom: 3, left: 4 },
      reopenedCells: Array(4).fill({ top: 1, right: 2, bottom: 3, left: 4 }),
      invalidError: {
        name: 'TypeError',
        message: 'Table margins must be a number, four-value tuple, or margin object',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableCellHyperlinks: true,
    tableCellHyperlinkEditing: true,
    tableCellHyperlinkState: {
      immediate: [
        { url: 'https://shared.example?a=1&b=2', tooltip: 'Visit & learn' },
        { url: 'https://second.example', tooltip: '' },
        { url: 'https://third.example' },
        { slide: 2 },
        { slide: 1, tooltip: '' },
        null,
      ],
      inputDetached: true,
      snapshotsFrozen: true,
      independentRelationships: true,
      shared: [
        { url: 'https://shared.example?a=1&b=2', tooltip: 'Visit & learn' },
        { url: 'https://shared.example?a=1&b=2', tooltip: '' },
        { url: 'https://shared.example?a=1&b=2' },
        { slide: 2 },
        { slide: 1, tooltip: '' },
        null,
      ],
      noOp: true,
      tooltipIdReuse: true,
      cloneOnWrite: true,
      uniqueIdReuse: true,
      addClear: true,
      sharedGc: true,
      textEditPreserved: {
        url: 'https://edited.example?a=1&b=2',
        tooltip: 'Edited',
      },
      afterEdit: [
        { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
        null,
        null,
        { slide: 3, tooltip: '' },
        { slide: 1, tooltip: '' },
        null,
      ],
      movedInternal: { slide: 1, tooltip: '' },
      movedSelf: { slide: 2, tooltip: '' },
      restoredInternal: { slide: 3, tooltip: '' },
      duplicateSelf: { slide: 4, tooltip: '' },
      targetDeletion: { source: null, duplicate: null },
      finalRelationshipOwnership: {
        clickCount: 2,
        relationshipCount: 2,
        uniqueClickIds: 2,
        idsMatch: true,
        underlineCount: 6,
      },
      reopened: [
        { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
        null,
        null,
        null,
        { slide: 1, tooltip: '' },
        null,
      ],
      reopenedDuplicate: [
        { url: 'https://edited.example?a=1&b=2', tooltip: 'Edited' },
        null,
        null,
        null,
        { slide: 3, tooltip: '' },
        null,
      ],
      invalidError: {
        name: 'RangeError',
        message: 'Table cell 0,0 hyperlink slide 99 is out of range',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableBorders: true,
    tableBordersState: {
      uniform: {
        top: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'dash',
        },
        right: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'dash',
        },
        bottom: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'dash',
        },
        left: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'dash',
        },
      },
      readIsolation: true,
      noOp: true,
      mixed: null,
      partial: {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: 'D9EAF7' },
          width: 2,
        },
        bottom: { kind: 'none' },
      },
      partialCells: Array(4).fill({
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: 'D9EAF7' },
          width: 2,
        },
        bottom: { kind: 'none' },
      }),
      none: {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      },
      noneCells: Array(4).fill({
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      }),
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      },
      reopenedCells: Array(4).fill({
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      }),
      invalidError: {
        name: 'TypeError',
        message: 'Table borders must be an object',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableFill: true,
    tableFillState: {
      uniform: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      readIsolation: true,
      noOp: true,
      mixed: null,
      none: { kind: 'none' },
      noneCells: Array(4).fill({ kind: 'none' }),
      solid: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      },
      solidCells: Array(4).fill({
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      }),
      cleared: null,
      clearedCells: [null, null, null, null],
      reopened: { kind: 'none' },
      reopenedCells: Array(4).fill({ kind: 'none' }),
      invalidError: {
        name: 'TypeError',
        message: 'Table fill must be an object',
      },
      failureIsolation: true,
      validationErrors: 0,
    },
    tableTextDefaults: true,
    tableTextDefaultsState: {
      created: true,
      plainEdit: true,
      richReplacement: true,
      reopened: true,
      validationErrors: 0,
    },
    tableCellMerges: true,
    tableCellMergesState: {
      created: true,
      read: true,
      snapshotsFrozen: true,
      unmerged: true,
      edited: true,
      remerged: true,
      reopened: true,
      validationErrors: 0,
    },
    tableStructureEditing: true,
    tableStructureEditingState: {
      created: true,
      rowsInserted: true,
      columnsInserted: true,
      newCellEdited: true,
      rowsDeleted: true,
      columnsDeleted: true,
      dimensions: true,
      merge: true,
      survivor: true,
      relationships: true,
      reopened: true,
      validationErrors: 0,
    },
    tableAutoPage: true,
    tableAutoPageState: {
      created: true,
      edited: true,
      moved: true,
      deleted: true,
      relationships: true,
      reopened: true,
      validationErrors: 0,
    },
    tableContentMeasurement: true,
    tableContentMeasurementState: {
      created: true,
      automatic: true,
      minimum: true,
      fragment: true,
      placeholder: true,
      fixed: true,
      relationships: true,
      edited: true,
      moved: true,
      deleted: true,
      reopened: true,
      validationErrors: 0,
    },
    tableToSlides: true,
    tableToSlidesState: {
      created: true,
      styles: true,
      widths: true,
      headers: true,
      layout: true,
      additions: true,
      relationships: true,
      edited: true,
      reopened: true,
      validationErrors: 0,
    },
    schemeColors: true,
    schemeColorState: {
      entries: [
        ['text1', 'tx1'],
        ['text2', 'tx2'],
        ['background1', 'bg1'],
        ['background2', 'bg2'],
        ['accent1', 'accent1'],
        ['accent2', 'accent2'],
        ['accent3', 'accent3'],
        ['accent4', 'accent4'],
        ['accent5', 'accent5'],
        ['accent6', 'accent6'],
      ],
      frozen: true,
      mutationIsolation: true,
      textColor: { kind: 'scheme', value: 'tx1' },
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      validationErrors: 0,
    },
    outputTypes: true,
    outputTypeState: {
      values: [
        'arraybuffer',
        'base64',
        'binarystring',
        'blob',
        'nodebuffer',
        'uint8array',
      ],
      frozen: true,
      mutationIsolation: true,
    },
    writeOutputTypes: true,
    writeOutputTypeState: {
      defaultKind: 'uint8array',
      emptyKind: 'uint8array',
      arraybufferKind: 'arraybuffer',
      base64Kind: 'string',
      binarystringKind: 'string',
      blobKind: 'blob',
      blobType: 'application/zip',
      uint8arrayKind: 'uint8array',
      byteEquality: true,
      reopenTitles: Array(5).fill('Browser output types 你好'),
      writeBlobType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      nodebufferError: {
        name: 'Error',
        message: 'nodebuffer is not supported by this platform',
      },
      failureIsolation: true,
      mutationIsolation: true,
    },
    nodeReadableStream: true,
    nodeReadableStreamState: {
      error: {
        name: 'Error',
        message: 'PptxDocument.stream() is only supported in Node.js',
      },
      failureIsolation: true,
      laterWriteTitle: 'Browser node stream boundary',
    },
    compressionPolicy: true,
    compressionPolicyState: {
      defaultEqualsFalse: true,
      storeMethods: [0],
      deflateMethods: [8],
      storeBytes: result.compressionPolicyState.storeBytes,
      deflateBytes: result.compressionPolicyState.deflateBytes,
      deflateSmaller: true,
      blobEquality: true,
      reopenTitle: 'Chrome compression policy',
      invalidError: {
        name: 'TypeError',
        message: 'PptxDocument output compression must be a boolean',
      },
      invalidEarly: true,
      laterWriteTitle: 'Chrome compression policy',
      failureIsolation: true,
      downloadFileName: 'compression-policy.pptx',
      downloadMethods: [8],
      downloadReopenTitle: 'Chrome compression download',
    },
    format: 'pptx',
    title: 'Browser updated',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    transition: 'function',
    smartArt: 'function',
    blobInputTitle: 'Browser fixture',
    slideNumbers: {
      firstSlideNumber: -2,
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      slideCount: 2,
      values: [
        {
          width: 800000,
          height: 300000,
          align: 'center',
          rtl: true,
          valign: 'middle',
          margin: { top: 1, right: 2, bottom: 3, left: 4 },
          italic: true,
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        {
          width: 800000,
          height: 300000,
          align: 'center',
          rtl: true,
          valign: 'middle',
          margin: { top: 1, right: 2, bottom: 3, left: 4 },
          italic: true,
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
      ],
      caches: ['-2', '-1'],
      layoutX: 200,
      masterX: 300,
      layoutCache: '‹#›',
      masterCache: '‹#›',
      masterEnabled: true,
      ownerCounts: [1, 1, 1, 1],
      diagnostics: [],
    },
    masterLayouts: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      liveWrapperIdentity: true,
      layoutNames: ['DEFAULT', 'BROWSER-MASTER-LAYOUT'],
      backgroundKinds: { master: 'solid', layout: 'image' },
      marginBeforeWrite: { top: 91440, right: 182880, bottom: 274320, left: 365760 },
      placeholderTypes: ['title', 'body', 'pic', 'chart', 'tbl', 'media'],
      layoutPlaceholders: [
        { name: 'browser_title', placeholder: { type: 'title', index: 101 } },
        { name: 'browser_body', placeholder: { type: 'body', index: 102 } },
        { name: 'browser_picture', placeholder: { type: 'pic', index: 103 } },
        { name: 'browser_chart', placeholder: { type: 'chart', index: 104 } },
        { name: 'browser_table', placeholder: { type: 'tbl', index: 105 } },
        { name: 'browser_media', placeholder: { type: 'media', index: 106 } },
      ],
      slidePlaceholders: [
        { name: 'browser_title', kind: 'text', placeholder: { type: 'title', index: 101 } },
        { name: 'browser_body', kind: 'text', placeholder: { type: 'body', index: 102 } },
        { name: 'browser_picture', kind: 'image', placeholder: { type: 'pic', index: 103 } },
        { name: 'browser_chart', kind: 'chart', placeholder: { type: 'chart', index: 104 } },
        { name: 'browser_table', kind: 'table', placeholder: { type: 'tbl', index: 105 } },
        { name: 'browser_media', kind: 'audio', placeholder: { type: 'media', index: 106 } },
      ],
      placeholderGeometryPreserved: true,
      selectedTargets: [
        '/ppt/slideLayouts/slideLayout2.xml',
        '/ppt/slideLayouts/slideLayout2.xml',
      ],
      reopenedMargin: null,
      payloadHashes: {
        background: '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
        image: '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
        media: '9f64a747e1b97f131fabb6b447296c9b6f0201e79fb3c5356e6c77e89b6a806a',
      },
      chartDefinition: [{
        type: 'bar',
        axis: 'primary',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      }],
      validationErrors: 0,
    },
    slideDefaultColor: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      live: [
        { kind: 'scheme', value: 'accent1' },
        { kind: 'scheme', value: 'accent1' },
      ],
      duplicateIdentity: true,
      materialized: [
        [[
          { color: { kind: 'scheme', value: 'accent1' } },
          { color: { kind: 'srgb', value: '00AA00' } },
          { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
        ]],
        [
          [
            { color: { kind: 'scheme', value: 'accent1' } },
            { color: { kind: 'srgb', value: '00AA00' } },
            { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
          ],
          [{ color: { kind: 'scheme', value: 'accent1' } }],
        ],
      ],
      reopened: [null, null],
      validationErrors: 0,
    },
    shapeFills: true,
    shapeFillState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        { kind: 'none' },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
      ],
      detached: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'AB12CD' },
        transparency: 25,
      },
      replacedNone: { kind: 'none' },
      cleared: null,
      edited: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
      },
      preserved: { kind: 'none' },
      serialized: { none: true, edited: true, explicitZero: true },
      partIsolation: true,
      relationshipIsolation: true,
      reopened: [
        { kind: 'none' },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
      ],
      validationErrors: 0,
    },
    textShapeFills: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
        { kind: 'none' },
      ],
      detached: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'AB12CD' },
        transparency: 25,
      },
      reopened: [
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
        },
        { kind: 'none' },
      ],
      layout: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 100,
      },
      validationErrors: 0,
    },
    textShapeLines: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
          width: 2.5,
          dash: 'dashDot',
        },
        {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
          width: 1,
          dash: 'solid',
        },
        { kind: 'none' },
      ],
      detached: {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      },
      reopened: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
          width: 2.5,
          dash: 'dashDot',
        },
        {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 0,
          width: 1,
          dash: 'solid',
        },
        { kind: 'none' },
      ],
      layout: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 100,
        width: 0,
        dash: 'sysDot',
      },
      validationErrors: 0,
    },
    textShapeArrows: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        { begin: 'triangle', end: 'arrow' },
        { end: 'diamond' },
        { begin: 'arrow' },
      ],
      detached: { begin: 'triangle', end: 'arrow' },
      reopened: [
        { begin: 'triangle', end: 'arrow' },
        { end: 'diamond' },
        { begin: 'arrow' },
      ],
      layout: { begin: 'none', end: 'stealth' },
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 2,
        dash: 'dashDot',
      },
      validationErrors: 0,
    },
    textShapeShadows: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        {
          kind: 'outer',
          color: { kind: 'scheme', value: 'accent4' },
          opacity: 0.4,
          blur: 2,
          angle: 45,
          distance: 3,
          rotateWithShape: true,
        },
        {
          kind: 'inner',
          color: { kind: 'srgb', value: '667788' },
          opacity: 0,
          blur: 0,
          angle: 0,
          distance: 0,
        },
        {
          kind: 'outer',
          color: { kind: 'srgb', value: '000000' },
          opacity: 0.75,
          blur: 8,
          angle: 270,
          distance: 4,
          rotateWithShape: false,
        },
      ],
      detached: {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent4' },
        opacity: 0.4,
        blur: 2,
        angle: 45,
        distance: 3,
        rotateWithShape: true,
      },
      reopened: [
        {
          kind: 'outer',
          color: { kind: 'scheme', value: 'accent4' },
          opacity: 0.4,
          blur: 2,
          angle: 45,
          distance: 3,
          rotateWithShape: true,
        },
        {
          kind: 'inner',
          color: { kind: 'srgb', value: '667788' },
          opacity: 0,
          blur: 0,
          angle: 0,
          distance: 0,
        },
        {
          kind: 'outer',
          color: { kind: 'srgb', value: '000000' },
          opacity: 0.75,
          blur: 8,
          angle: 270,
          distance: 4,
          rotateWithShape: false,
        },
      ],
      layout: {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent1' },
        opacity: 0.75,
        blur: 8,
        angle: 270,
        distance: 4,
        rotateWithShape: true,
      },
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 2,
        dash: 'dashDot',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
      validationErrors: 0,
    },
    textShapePresetGeometry: true,
    textShapePresetGeometryState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: ['rect', 'ellipse', 'line', 'roundRect', 'foldedCorner'],
      edited: 'star5',
      reopened: ['rect', 'star5', 'line', 'roundRect'],
      layout: 'foldedCorner',
      texts: [
        'Browser populated geometry',
        'Browser default geometry',
        'Browser ellipse geometry',
        'Browser rich line',
      ],
      validationErrors: 0,
    },
    textShapeRectRadius: true,
    textShapeRectRadiusState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        [],
        [{ name: 'adj', value: 0 }],
        [{ name: 'adj', value: 50_000 }],
        [{ name: 'adj', value: 25_000 }],
        [{ name: 'adj', value: 25_000 }],
        [{ name: 'adj', value: 12_500 }],
      ],
      resizeStable: [{ name: 'adj', value: 50_000 }],
      edited: [{ name: 'adj', value: 12_500 }],
      cleared: [],
      reopened: [
        [],
        [{ name: 'adj', value: 0 }],
        [{ name: 'adj', value: 50_000 }],
        [{ name: 'adj', value: 12_500 }],
        [{ name: 'adj', value: 25_000 }],
      ],
      layout: [{ name: 'adj', value: 12_500 }],
      validationErrors: 0,
    },
    textShapeIsTextBox: true,
    textShapeIsTextBoxState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [false, true, false, true, false, true, false, true],
      materialized: [false, true],
      noOp: true,
      edited: [true, false],
      duplicate: [true, false],
      declarative: [true, false, true, false, true],
      reopened: [true, false, false, true],
      layout: [false, false, true],
      master: true,
      reopenedDeclarative: [true, false, true],
      validationErrors: 0,
    },
    richTextBreakLine: true,
    richTextBreakLineState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      direct: [
        ['Browser direct first'],
        [],
        ['Browser direct soft', 'Browser direct trailing'],
      ],
      materialized: [
        [],
      ],
      owners: [
        [
          ['Browser layout first'],
          [],
          ['Browser layout soft', 'Browser layout trailing'],
        ],
        [
          ['Browser master first'],
          [],
          ['Browser master soft', 'Browser master trailing'],
        ],
        [
          ['Browser layout prompt first'],
          [],
          ['Browser layout prompt soft', 'Browser layout prompt trailing'],
        ],
        [
          ['Browser master prompt first'],
          [],
          ['Browser master prompt soft', 'Browser master prompt trailing'],
        ],
      ],
      populated: [
        ['Browser populated first'],
        [],
        ['Browser populated soft', 'Browser populated trailing'],
      ],
      edited: [
        ['Browser edited first'],
        [],
        ['Browser edited soft', 'Browser edited trailing'],
      ],
      duplicate: [
        [
          ['Browser direct first'],
          [],
          ['Browser direct soft', 'Browser direct trailing'],
        ],
        [['Browser duplicate only']],
      ],
      declarative: [
        [
          ['Browser declarative first'],
          [],
          ['Browser declarative soft', 'Browser declarative trailing'],
        ],
        [
          ['Browser declarative prompt first'],
          [],
          ['Browser declarative prompt soft', 'Browser declarative prompt trailing'],
        ],
        [
          ['Browser declarative populated first'],
          [],
          [
            'Browser declarative populated soft',
            'Browser declarative populated trailing',
          ],
        ],
      ],
      reopened: [
        [
          ['Browser direct first'],
          [],
          ['Browser direct soft', 'Browser direct trailing'],
        ],
        [['Browser duplicate only']],
        [
          ['Browser layout first'],
          [],
          ['Browser layout soft', 'Browser layout trailing'],
        ],
        [
          ['Browser master first'],
          [],
          ['Browser master soft', 'Browser master trailing'],
        ],
        [
          ['Browser declarative first'],
          [],
          ['Browser declarative soft', 'Browser declarative trailing'],
        ],
      ],
      properties: true,
      canonical: true,
      softBreak: true,
      relationshipTarget: true,
      xml: {
        paragraphCount: 3,
        emptyParagraph: true,
        softBreak: true,
        privateMarker: false,
      },
      validationErrors: 0,
      validationWarnings: 0,
    },
    textShapeHyperlinks: true,
    textShapeHyperlinkState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        {
          url: 'https://browser-text.example/path?a=1&b=2',
          tooltip: 'Browser & text',
        },
        { slide: 2, tooltip: '' },
        { slide: 1 },
      ],
      detached: {
        url: 'https://browser-text.example/path?a=1&b=2',
        tooltip: 'Browser & text',
      },
      reopened: [
        {
          url: 'https://browser-text.example/path?a=1&b=2',
          tooltip: 'Browser & text',
        },
        { slide: 2, tooltip: '' },
        { slide: 1 },
      ],
      layout: { url: 'https://layout-browser.example', tooltip: '' },
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'DDEEFF' },
      },
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 2,
        dash: 'dashDot',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
      shadowKind: 'outer',
      clickCounts: [3, 3],
      sharedIds: [true, true],
      internalActions: true,
      validationErrors: 0,
    },
    richTextRunHyperlinks: true,
    richTextRunHyperlinkState: {
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      immediate: [
        { url: 'https://browser-outer-run.example', tooltip: 'Outer' },
        { url: 'https://browser-run.example', tooltip: 'One' },
        { url: 'https://browser-run.example', tooltip: 'Two' },
        undefined,
        { slide: 2, tooltip: '' },
        { slide: 1 },
      ],
      relationshipIndependence: true,
      afterEdit: [
        { url: 'https://browser-outer-run.example', tooltip: 'Outer' },
        { url: 'https://browser-run-edited.example', tooltip: '' },
        undefined,
        undefined,
        undefined,
        { slide: 1 },
      ],
      afterMove: [{ slide: 2 }, { slide: 3 }],
      reopened: [
        [
          { url: 'https://browser-outer-run.example', tooltip: 'Outer' },
          { url: 'https://browser-run-edited.example', tooltip: '' },
          undefined,
          undefined,
          undefined,
          { slide: 1 },
        ],
        [
          { url: 'https://browser-outer-run.example', tooltip: 'Outer' },
          { url: 'https://browser-run-edited.example', tooltip: '' },
          undefined,
          undefined,
          undefined,
          { slide: 2 },
        ],
      ],
      clickCounts: [4, 4],
      independentIds: [3, 3],
      internalActions: true,
      validationErrors: 0,
    },
    svgCreatedLive: true,
    svgState: [
      {
        name: 'Browser Blob SVG',
        fallbackType: 'image/png',
        svgType: 'image/svg+xml',
        pngSignature: [137, 80, 78, 71, 13, 10, 26, 10],
        transform: {
          x: 914400,
          y: 914400,
          width: 3657600,
          height: 2743200,
          rotation: 900000,
          flipHorizontal: true,
          flipVertical: false,
        },
        sourceRectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
        internalTargets: 2,
      },
      {
        name: 'Browser data SVG',
        fallbackType: 'image/png',
        svgType: 'image/svg+xml',
        pngSignature: [137, 80, 78, 71, 13, 10, 26, 10],
        transform: {
          x: 5029200,
          y: 914400,
          width: 3657600,
          height: 2743200,
          rotation: 0,
          flipHorizontal: false,
          flipVertical: true,
        },
        sourceRectangle: { left: 0, top: -16.667, right: 0, bottom: -16.667 },
        internalTargets: 2,
      },
    ],
    imageSourceSizingTransformState: {
      live: true,
      relativeRaster: {
        contentType: 'image/png',
        transform: {
          x: 914400,
          y: 4114800,
          width: 3657600,
          height: 2743200,
          rotation: -600000,
          flipHorizontal: true,
          flipVertical: true,
        },
        sourceRectangle: { left: 25, top: 25, right: 25, bottom: 25 },
        internalTargets: 1,
      },
    },
    backgroundMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    slideBackgroundKinds: ['image', 'image', 'solid', 'linear-gradient'],
    backgroundPayloadHashes: [
      '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
      '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
    ],
    backgroundRelationshipCounts: [1, 1, 0, 0],
    backgroundValidationErrors: 0,
    mediaMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    mediaNames: ['Browser MP3 narration edited', 'Browser Blob video', 'Browser online video'],
    mediaElementCounts: { audio: 1, video: 2 },
    mediaState: [
      {
        kind: 'audio',
        mediaType: 'audio/mpeg',
        mediaExtension: '.mp3',
        posterType: 'image/png',
        posterExtension: '.png',
        roles: [true, true, true],
        posterSignature: [137, 80, 78, 71],
      },
      {
        kind: 'video',
        mediaType: 'video/mp4',
        mediaExtension: '.mp4',
        posterType: 'image/jpeg',
        posterExtension: '.jpg',
        roles: [true, true, true],
        posterSignature: [255, 216, 255, 224],
      },
      {
        kind: 'video',
        externalUrl: 'https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0&controls=0',
        mediaPartUri: null,
        posterType: 'image/png',
        posterExtension: '.png',
        transform: {
          x: 5486400,
          y: 3657600,
          width: 3657600,
          height: 2057400,
          rotation: 0,
          flipHorizontal: false,
          flipVertical: false,
        },
        roles: [true, true],
        posterSignature: [137, 80, 78, 71],
      },
    ],
    browserOnlineVideoState: {
      externalUrl: 'https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0&controls=0',
      mediaPartUri: null,
      name: 'Browser online video',
      transform: {
        x: 5486400,
        y: 3657600,
        width: 3657600,
        height: 2057400,
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      posterType: 'image/png',
      externalRelationship: true,
      posterRelationship: true,
      escapedQuery: true,
    },
    mediaValidationErrors: 0,
    mediaTimingElementCounts: [
      { timing: 1, media: 2, commands: 3, playback: 2 },
      { timing: 1, media: 1, commands: 1, playback: 1 },
    ],
    mediaTimingUniqueIdCount: [17, 7],
    settingsAfterReopen: [
      [
        { play: 'click', loop: false, hideWhenStopped: false, volume: 1 },
        {},
        { play: 'click', loop: false, hideWhenStopped: false, volume: 1 },
      ],
      [
        { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25 },
      ],
    ],
    timingDiagnostics: [],
    nativeMediaTiming: true,
    nativeCharts: true,
    imageIdentityEffects5: true,
    imageIdentityEffects5State: {
      noOp: true,
      rollback: true,
      sourceIsolation: true,
      sharedMedia: true,
      relationships: true,
      media: true,
      exactOoxml: true,
      diagnostics: true,
    },
    chartPresentation91: true,
    chartPresentation91State: {
      fixtureShape: true,
      readback: true,
      triStateOverrides: true,
      unrelatedEditPreserved: true,
      explicitSafeEdit: true,
      customScatterPreserved: true,
      packageLifecycle: true,
    },
    chartPresentation: true,
    chartPresentationState: {
      fixtureShape: true,
      readback: true,
      triStateOverrides: true,
      unrelatedEditPreserved: true,
      explicitSafeEdit: true,
      customScatterPreserved: true,
      packageLifecycle: true,
    },
    chartAreaFillLine: true,
    chartAxisAdvanced: true,
    chartAxisDisplayUnitCatalog: true,
    stableMediaLifecycle: true,
    mediaTargetIsolation: true,
    mediaOrphanCount: 0,
    imageIdentityEffects5EvidenceFileName:
      'browser-image-identity-effects-5.pptx',
    chartPresentation91EvidenceFileName: 'browser-chart-presentation-91.pptx',
    chartPresentationEvidenceFileName: 'browser-chart-presentation-91.pptx',
    textBoxFitEvidenceFileName: 'browser-text-box-fit.pptx',
    textParagraphLayoutEvidenceFileName: 'browser-text-paragraph-layout.pptx',
    richTextEffectsEvidenceFileName: 'browser-rich-text-effects.pptx',
    customPathEvidenceFileName: 'browser-custom-paths.pptx',
    tableCellMergesEvidenceFileName: 'browser-table-cell-merges.pptx',
    tableStructureEditingEvidenceFileName: 'browser-table-structure-editing.pptx',
    tableAutoPageEvidenceFileName: 'browser-table-auto-page.pptx',
    tableContentMeasurementEvidenceFileName:
      'browser-table-content-measurement.pptx',
    tableToSlidesEvidenceFileName: 'browser-table-to-slides.pptx',
    downloadFileName: 'browser-smoke.pptx',
    errorCounts: { console: 0, page: 0, network: 0 },
  };
  if (JSON.stringify(result) !== JSON.stringify(expected)) {
    throw new Error(`Browser smoke mismatch: ${JSON.stringify({
      result,
      consoleErrors,
      pageErrors,
      networkErrors,
    })}`);
  }
  return result;
}
