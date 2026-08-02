async (page) => {
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
    async ({ moduleUrl, base64 }) => {
      const binary = atob(base64);
      const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
      const api = await import(moduleUrl);
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
            options: { name: 'browser_title', type: 'title', index: 101 },
          },
          {
            kind: 'placeholder', text: 'Body prompt',
            options: { name: 'browser_body', type: 'body', index: 102 },
          },
          {
            kind: 'placeholder', text: 'Picture prompt',
            options: { name: 'browser_picture', type: 'pic', index: 103 },
          },
          {
            kind: 'placeholder', text: 'Chart prompt',
            options: { name: 'browser_chart', type: 'chart', index: 104 },
          },
          {
            kind: 'placeholder', text: 'Table prompt',
            options: { name: 'browser_table', type: 'tbl', index: 105 },
          },
          {
            kind: 'placeholder', text: 'Media prompt',
            options: { name: 'browser_media', type: 'media', index: 106 },
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
        width: api.inches(3),
        height: api.inches(2),
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
          internalTargets: targets.filter((target) => reopenedSlide.relationships.some(
            ({ type, targetMode, resolvedTarget }) => type.endsWith('/image')
              && targetMode === 'Internal' && resolvedTarget === target,
          )).length,
        };
      });
      const mediaOutput = await mediaDocument.writeBlob({ compatibility: 'powerpoint-2010' });
      const reopenedMediaDocument = await api.PptxDocument.open(mediaOutput);
      await reopenedMediaDocument.write({ mode: 'permissive', compatibility: 'powerpoint-2010' });
      const reopenedMediaSlide = reopenedMediaDocument.slides[0];
      const reopenedMedia = reopenedMediaDocument.media(0);
      const reopenedDuplicateMedia = reopenedMediaDocument.media(1);
      const mediaXml = new TextDecoder().decode(
        reopenedMediaDocument.opcPackage.requirePart(reopenedMediaSlide.partUri).bytes,
      );
      const mediaState = reopenedMedia.map((model) => {
        const mediaPart = reopenedMediaDocument.opcPackage.requirePart(model.mediaPartUri);
        const posterPart = reopenedMediaDocument.opcPackage.requirePart(model.posterPartUri);
        const mediaRelationships = reopenedMediaSlide.relationships.filter(
          ({ resolvedTarget }) => resolvedTarget === model.mediaPartUri,
        );
        const posterRelationships = reopenedMediaSlide.relationships.filter(
          ({ resolvedTarget }) => resolvedTarget === model.posterPartUri,
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
        const configured = reopenedMediaDocument.media(slideIndex)
          .filter((model) => Object.keys(model.settings).length > 0)
          .map(({ shapeId }) => shapeId);
        return {
          timing: (source.match(/<p:timing>/g) ?? []).length,
          media: (source.match(/<p:cMediaNode\b/g) ?? []).length,
          commands: (source.match(/<p:cmd\b/g) ?? []).length,
          playback: (source.match(/<px:playback\b/g) ?? []).length,
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
        && timingSummaries.every((summary) => summary.timing === 1
          && summary.media === 1 && summary.commands === 1 && summary.playback === 1
          && summary.ids === summary.uniqueIds && summary.isolatedTargets)
        && timingDiagnostics.length === 0
        && settingsAfterReopen[0][0].play === 'click'
        && Object.keys(settingsAfterReopen[0][1]).length === 0
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
        chartModels.push(await slide.addChart(type, series, {
          name: `Browser ${type} chart`,
          x: api.inches(0.5),
          y: api.inches(0.5),
          width: api.inches(9),
          height: api.inches(6.5),
        }));
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
      const duplicateChartSlide = chartDocument.duplicateSlide(chartDocument.slides.length - 1);
      const duplicateChart = duplicateChartSlide.shapes.find(
        (shape) => shape instanceof api.ChartModel,
      );
      const duplicateChartPartUri = duplicateChart.chartPartUri;
      const comboChartPartUri = combo.chartPartUri;
      duplicateChart.remove();
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
      const nativeCharts = reopenedCharts.length === 10
        && api.CHART_TYPES.every((type) => reopenedChartTypes.has(type))
        && chartWorkbookResults.every(Boolean)
        && chartIdsUnique
        && chartOrphanCount === 0
        && !chartDocument.opcPackage.hasPart(duplicateChartPartUri)
        && chartDocument.opcPackage.hasPart(comboChartPartUri)
        && reopenedChartDocument.diagnostics.filter(({ code }) => code.startsWith('CHART_')).length === 0;
      return {
        format: reopened.format,
        title: reopened.slides[0].title.text,
        mime: output.type,
        transition: typeof api.transitions.TransitionCodec,
        smartArt: typeof api.smartArt.SmartArtDiagramCodec,
        blobInputTitle: fromBlob.slides[0].title.text,
        slideNumbers: slideNumberState,
        masterLayouts: masterLayoutState,
        slideDefaultColor: slideDefaultColorState,
        textShapeFills: textShapeFillState,
        textShapeLines: textShapeLineState,
        textShapeArrows: textShapeArrowState,
        textShapeShadows: textShapeShadowState,
        textShapeHyperlinks,
        textShapeHyperlinkState,
        richTextRunHyperlinks,
        richTextRunHyperlinkState,
        svgCreatedLive: svgDocument.slides[0].shapes.includes(blobSvg)
          && svgDocument.slides[0].shapes.includes(dataSvg),
        svgState,
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
        stableMediaLifecycle,
        mediaTargetIsolation: browserMediaCloneOnWrite,
        mediaOrphanCount,
      };
    },
    {
      moduleUrl: 'http://127.0.0.1:4173/packages/pptx/dist/browser.js',
      base64: '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',
    },
  );
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
      base64: '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',
    },
  );
  result.downloadFileName = (await downloadPromise).suggestedFilename();
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
        internalTargets: 2,
      },
      {
        name: 'Browser data SVG',
        fallbackType: 'image/png',
        svgType: 'image/svg+xml',
        pngSignature: [137, 80, 78, 71, 13, 10, 26, 10],
        internalTargets: 2,
      },
    ],
    backgroundMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    slideBackgroundKinds: ['image', 'image', 'solid', 'linear-gradient'],
    backgroundPayloadHashes: [
      '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
      '431ced6916a2a21a156e38701afe55bbd7f88969fbbfc56d7fe099d47f265460',
    ],
    backgroundRelationshipCounts: [1, 1, 0, 0],
    backgroundValidationErrors: 0,
    mediaMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    mediaNames: ['Browser MP3 narration edited', 'Browser Blob video'],
    mediaElementCounts: { audio: 1, video: 1 },
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
    ],
    mediaValidationErrors: 0,
    mediaTimingElementCounts: [
      { timing: 1, media: 1, commands: 1, playback: 1 },
      { timing: 1, media: 1, commands: 1, playback: 1 },
    ],
    mediaTimingUniqueIdCount: [7, 7],
    settingsAfterReopen: [
      [
        { play: 'click', loop: false, hideWhenStopped: false, volume: 1 },
        {},
      ],
      [
        { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25 },
      ],
    ],
    timingDiagnostics: [],
    nativeMediaTiming: true,
    nativeCharts: true,
    stableMediaLifecycle: true,
    mediaTargetIsolation: true,
    mediaOrphanCount: 0,
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
