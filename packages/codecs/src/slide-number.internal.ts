import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import type {
  SlideNumber,
  SlideNumberColor,
  SlideNumberMargins,
  SlideNumberOptions,
  SlideNumberOwnerKind,
  SlideNumberTextStyle,
} from './slide-number.js';
import type { CodecDiagnostic } from './registry.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_COORDINATE = 27_273_042_316_900;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const MAX_UINT32 = 4_294_967_295;
const EMU_PER_POINT = 12_700;
const FIELD_ID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
const INVALID = Symbol('invalid slide number value');
const OPTION_KEYS = new Set([
  'x',
  'y',
  'width',
  'height',
  'align',
  'rtl',
  'valign',
  'margin',
  'style',
]);
const STYLE_KEYS = new Set([
  'fontFamily',
  'fontSize',
  'lang',
  'bold',
  'italic',
  'color',
  'transparency',
]);
const COLOR_KEYS = new Set(['kind', 'value']);
const MARGIN_KEYS = new Set(['top', 'right', 'bottom', 'left']);
const SCHEME_COLORS = new Set([
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'bg1',
  'bg2',
  'dk1',
  'dk2',
  'folHlink',
  'hlink',
  'lt1',
  'lt2',
  'phClr',
  'tx1',
  'tx2',
]);
const ROOT_NAMES: Readonly<Record<SlideNumberOwnerKind, string>> = {
  slide: 'sld',
  layout: 'sldLayout',
  master: 'sldMaster',
};
const ALIGN_FROM_OOXML = new Map([
  ['l', 'left' as const],
  ['ctr', 'center' as const],
  ['r', 'right' as const],
  ['just', 'justify' as const],
]);
const VALIGN_FROM_OOXML = new Map([
  ['t', 'top' as const],
  ['ctr', 'middle' as const],
  ['b', 'bottom' as const],
]);

interface PartialSlideNumberStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

interface SupportedSlideNumberTarget {
  readonly shape: XmlElement;
  readonly offset: XmlElement;
  readonly extent: XmlElement;
  readonly bodyProperties: XmlElement;
  readonly paragraph: XmlElement;
  readonly paragraphProperties?: XmlElement;
  readonly field: XmlElement;
  readonly fieldProperties?: XmlElement;
  readonly text: XmlElement;
  readonly defaultProperties?: XmlElement;
}

export function normalizeSlideNumberOptions(value: unknown): Readonly<SlideNumber> {
  const input = readDataObject(value, OPTION_KEYS, 'Slide number options');
  const margin = normalizeMargins(input.margin);
  const style = normalizeStyle(input.style);
  return deepFreeze({
    x: normalizeCoordinate(input.x, 0, 'Slide number x'),
    y: normalizeCoordinate(input.y, 0, 'Slide number y'),
    width: normalizeExtent(input.width, 800_000, 'Slide number width'),
    height: normalizeExtent(input.height, 300_000, 'Slide number height'),
    align: normalizeEnum(
      input.align,
      ['left', 'center', 'right', 'justify'] as const,
      'left',
      'Slide number align',
    ),
    rtl: normalizeBoolean(input.rtl, false, 'Slide number rtl'),
    ...(input.valign === undefined
      ? {}
      : {
          valign: normalizeEnum(
            input.valign,
            ['top', 'middle', 'bottom'] as const,
            undefined,
            'Slide number valign',
          ),
        }),
    ...(margin === undefined ? {} : { margin }),
    style,
  });
}

export function readSlideNumber(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
): Readonly<SlideNumber> | undefined {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(ownerPartUri).bytes);
  const root = uniqueRoot(xml, ROOT_NAMES[ownerKind], PRESENTATION_NAMESPACE);
  if (!root) return undefined;
  if (ownerKind === 'master' && !masterSlideNumberEnabled(root)) return undefined;

  const commonSlide = uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE);
  const shapeTree = commonSlide
    ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
    : undefined;
  if (!shapeTree) return undefined;

  const candidates = directChildren(shapeTree).filter(
    (child) => isElement(child, 'sp', PRESENTATION_NAMESPACE)
      && containsDirectSlideNumberPlaceholder(child),
  );
  if (candidates.length !== 1) return undefined;
  const shape = candidates[0]!;
  if (!validUniqueShapeId(xml, root, shape)) return undefined;

  const shapeProperties = uniqueDirectChild(shape, 'spPr', PRESENTATION_NAMESPACE);
  const transform = shapeProperties
    ? uniqueDirectChild(shapeProperties, 'xfrm', DRAWING_NAMESPACE)
    : undefined;
  const offset = transform
    ? uniqueDirectChild(transform, 'off', DRAWING_NAMESPACE)
    : undefined;
  const extent = transform
    ? uniqueDirectChild(transform, 'ext', DRAWING_NAMESPACE)
    : undefined;
  if (!offset || !extent) return undefined;
  const x = readCoordinate(offset, 'x');
  const y = readCoordinate(offset, 'y');
  const width = readExtent(extent, 'cx');
  const height = readExtent(extent, 'cy');
  if (x === INVALID || y === INVALID || width === INVALID || height === INVALID) {
    return undefined;
  }

  const textBody = uniqueDirectChild(shape, 'txBody', PRESENTATION_NAMESPACE);
  const bodyProperties = textBody
    ? uniqueDirectChild(textBody, 'bodyPr', DRAWING_NAMESPACE)
    : undefined;
  if (!textBody || !bodyProperties) return undefined;
  const listStyles = directChildren(textBody).filter(
    (child) => isElement(child, 'lstStyle', DRAWING_NAMESPACE),
  );
  if (listStyles.length > 1) return undefined;
  const paragraphs = directChildren(textBody).filter(
    (child) => isElement(child, 'p', DRAWING_NAMESPACE),
  );
  if (paragraphs.length !== 1) return undefined;
  const paragraph = paragraphs[0]!;
  const paragraphProperties = optionalUniqueDirectChild(
    paragraph,
    'pPr',
    DRAWING_NAMESPACE,
  );
  if (paragraphProperties === INVALID) return undefined;
  const field = readUniqueSlideNumberField(paragraph);
  if (!field) return undefined;

  const align = paragraphProperties
    ? readMappedAttribute(paragraphProperties, 'algn', ALIGN_FROM_OOXML, 'left')
    : 'left';
  const rtl = paragraphProperties
    ? readBooleanAttribute(paragraphProperties, 'rtl', false)
    : false;
  const valign = readOptionalMappedAttribute(
    bodyProperties,
    'anchor',
    VALIGN_FROM_OOXML,
  );
  const margin = readMargins(bodyProperties);
  if (align === INVALID || rtl === INVALID || valign === INVALID || margin === INVALID) {
    return undefined;
  }

  const defaultStyle = readListStyle(listStyles[0]);
  const fieldProperties = optionalUniqueDirectChild(
    field,
    'rPr',
    DRAWING_NAMESPACE,
  );
  if (fieldProperties === INVALID) return undefined;
  const fieldStyle = readStyleElement(fieldProperties);
  if (defaultStyle === INVALID || fieldStyle === INVALID) return undefined;
  const style = mergeStyles(defaultStyle, fieldStyle);

  return deepFreeze({
    x,
    y,
    width,
    height,
    align,
    rtl,
    ...(valign === undefined ? {} : { valign }),
    ...(margin === undefined ? {} : { margin }),
    style,
  });
}

export function slideNumberDiagnostics(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
  expectedCachedText: string,
  compatibility: string,
): (CodecDiagnostic & { readonly compatibility: string })[] {
  assertCachedText(expectedCachedText);
  const xml = LosslessXmlDocument.parse(pkg.requirePart(ownerPartUri).bytes);
  const root = uniqueRoot(xml, ROOT_NAMES[ownerKind], PRESENTATION_NAMESPACE);
  const commonSlide = root
    ? uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE)
    : undefined;
  const shapeTree = commonSlide
    ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
    : undefined;
  const candidates = shapeTree ? directSlideNumberShapes(shapeTree) : [];
  if (!root || candidates.length !== 1) return [];

  const shape = candidates[0]!;
  const diagnostics: (CodecDiagnostic & { readonly compatibility: string })[] = [];
  const shapeId = diagnosticShapeId(shape);
  if (shapeId !== undefined) {
    const matchingIds = descendantsAndSelf(root)
      .filter((element) => isElement(element, 'cNvPr', PRESENTATION_NAMESPACE))
      .filter((element) => readUnsignedInteger(element, 'id', 1, MAX_UINT32) === shapeId);
    if (matchingIds.length > 1) {
      diagnostics.push({
        severity: 'error',
        code: 'SLIDE_NUMBER_SHAPE_ID_COLLISION',
        message: 'Slide-number shape id collides with another shape id',
        partUri: ownerPartUri,
        compatibility,
      });
    }
  }
  if (ownerKind === 'master' && masterSlideNumberExplicitlyDisabled(root)) {
    diagnostics.push({
      severity: 'warning',
      code: 'SLIDE_NUMBER_MASTER_DISABLED',
      message: 'The master slide-number placeholder is disabled by p:hf',
      partUri: ownerPartUri,
      compatibility,
    });
  }
  const target = supportedTarget(shape);
  if (target && textValue(target.text) !== expectedCachedText) {
    diagnostics.push({
      severity: 'warning',
      code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
      message: `Expected cached slide-number text ${expectedCachedText}`,
      partUri: ownerPartUri,
      compatibility,
    });
  }
  return diagnostics;
}

export function replaceSlideNumber(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
  value: SlideNumberOptions | undefined,
  cachedText: string,
): void {
  const normalized = value === undefined
    ? undefined
    : normalizeSlideNumberOptions(value);
  assertCachedText(cachedText);

  pkg.transaction(() => {
    const part = pkg.requirePart(ownerPartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const root = editableOwnerRoot(xml, ownerKind, ownerPartUri);
    const commonSlide = uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE);
    const shapeTree = commonSlide
      ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
      : undefined;
    if (!shapeTree) {
      throw new Error(`Slide-number owner has no unique editable shape tree: ${ownerPartUri}`);
    }
    const candidates = directSlideNumberShapes(shapeTree);

    if (normalized === undefined) {
      if (candidates.length === 0) return;
      if (candidates.length !== 1) {
        throw new Error(`Slide-number owner has ambiguous direct placeholders: ${ownerPartUri}`);
      }
      xml.removeElement(candidates[0]!);
      if (ownerKind === 'master') updateMasterFlag(xml, root, false);
      pkg.setPart(ownerPartUri, xml.serialize(), part.contentType);
      return;
    }

    if (candidates.length > 1) {
      throw new Error(`Slide-number owner has ambiguous direct placeholders: ${ownerPartUri}`);
    }
    const current = readSlideNumber(pkg, ownerPartUri, ownerKind);
    const candidate = candidates[0];
    const target = candidate && current
      ? supportedTarget(candidate)
      : undefined;
    if (target && current && slideNumbersEqual(current, normalized)) {
      const currentCache = textValue(target.text);
      const masterAlreadyEnabled = ownerKind !== 'master' || masterFlagValue(root) === true;
      if (currentCache === cachedText && masterAlreadyEnabled) return;
    }

    if (target) {
      patchSupportedTarget(xml, target, normalized, cachedText);
    } else if (candidate) {
      const identity = opaqueShapeIdentity(root, candidate);
      if (!identity) {
        throw new Error(`Slide-number placeholder identity is ambiguous: ${ownerPartUri}`);
      }
      xml.replaceElement(
        candidate,
        renderSlideNumberShape(normalized, identity.shapeId, identity.placeholderIndex, cachedText),
      );
    } else {
      const shapeId = allocateShapeId(root);
      const placeholderIndex = allocatePlaceholderIndex(root);
      const rendered = renderSlideNumberShape(
        normalized,
        shapeId,
        placeholderIndex,
        cachedText,
      );
      const extensions = directChildren(shapeTree).filter(
        (child) => isElement(child, 'extLst', PRESENTATION_NAMESPACE),
      );
      if (extensions[0]) xml.replace(extensions[0].start, extensions[0].start, rendered);
      else xml.appendChildXml(shapeTree, rendered);
    }
    if (ownerKind === 'master') updateMasterFlag(xml, root, true);
    pkg.setPart(ownerPartUri, xml.serialize(), part.contentType);
  });
}

export function replaceSlideNumberCachedText(
  pkg: OpcPackage,
  slidePartUri: string,
  cachedText: string,
): boolean {
  assertCachedText(cachedText);
  if (!readSlideNumber(pkg, slidePartUri, 'slide')) return false;
  return pkg.transaction(() => {
    const part = pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const root = uniqueRoot(xml, ROOT_NAMES.slide, PRESENTATION_NAMESPACE);
    const commonSlide = root
      ? uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE)
      : undefined;
    const shapeTree = commonSlide
      ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
      : undefined;
    const candidates = shapeTree ? directSlideNumberShapes(shapeTree) : [];
    const target = candidates.length === 1 ? supportedTarget(candidates[0]!) : undefined;
    if (!target || textValue(target.text) === cachedText) return false;
    xml.replaceElement(target.text, rewriteTextElement(xml, target.text, cachedText));
    pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    return true;
  });
}

function editableOwnerRoot(
  xml: LosslessXmlDocument,
  ownerKind: SlideNumberOwnerKind,
  ownerPartUri: string,
): XmlElement {
  const root = uniqueRoot(xml, ROOT_NAMES[ownerKind], PRESENTATION_NAMESPACE);
  if (!root) {
    throw new Error(`Slide-number owner root is missing or ambiguous: ${ownerPartUri}`);
  }
  return root;
}

function directSlideNumberShapes(shapeTree: XmlElement): XmlElement[] {
  return directChildren(shapeTree).filter(
    (child) => isElement(child, 'sp', PRESENTATION_NAMESPACE)
      && containsDirectSlideNumberPlaceholder(child),
  );
}

function diagnosticShapeId(shape: XmlElement): number | undefined {
  const nonVisual = uniqueDirectChild(shape, 'nvSpPr', PRESENTATION_NAMESPACE);
  const properties = nonVisual
    ? uniqueDirectChild(nonVisual, 'cNvPr', PRESENTATION_NAMESPACE)
    : undefined;
  if (!properties) return undefined;
  const value = readUnsignedInteger(properties, 'id', 1, MAX_UINT32);
  return value === INVALID ? undefined : value;
}

function supportedTarget(shape: XmlElement): SupportedSlideNumberTarget | undefined {
  const shapeProperties = uniqueDirectChild(shape, 'spPr', PRESENTATION_NAMESPACE);
  const transform = shapeProperties
    ? uniqueDirectChild(shapeProperties, 'xfrm', DRAWING_NAMESPACE)
    : undefined;
  const offset = transform
    ? uniqueDirectChild(transform, 'off', DRAWING_NAMESPACE)
    : undefined;
  const extent = transform
    ? uniqueDirectChild(transform, 'ext', DRAWING_NAMESPACE)
    : undefined;
  const textBody = uniqueDirectChild(shape, 'txBody', PRESENTATION_NAMESPACE);
  const bodyProperties = textBody
    ? uniqueDirectChild(textBody, 'bodyPr', DRAWING_NAMESPACE)
    : undefined;
  const paragraphs = textBody
    ? directChildren(textBody).filter((child) => isElement(child, 'p', DRAWING_NAMESPACE))
    : [];
  if (!offset || !extent || !textBody || !bodyProperties || paragraphs.length !== 1) {
    return undefined;
  }
  const paragraph = paragraphs[0]!;
  const paragraphProperties = optionalUniqueDirectChild(paragraph, 'pPr', DRAWING_NAMESPACE);
  if (paragraphProperties === INVALID) return undefined;
  const field = readUniqueSlideNumberField(paragraph);
  if (!field) return undefined;
  const fieldProperties = optionalUniqueDirectChild(field, 'rPr', DRAWING_NAMESPACE);
  if (fieldProperties === INVALID) return undefined;
  const texts = directChildren(field).filter((child) => isElement(child, 't', DRAWING_NAMESPACE));
  if (texts.length !== 1) return undefined;
  const defaultProperties = defaultRunProperties(textBody);
  return {
    shape,
    offset,
    extent,
    bodyProperties,
    paragraph,
    ...(paragraphProperties ? { paragraphProperties } : {}),
    field,
    ...(fieldProperties ? { fieldProperties } : {}),
    text: texts[0]!,
    ...(defaultProperties ? { defaultProperties } : {}),
  };
}

function defaultRunProperties(textBody: XmlElement): XmlElement | undefined {
  const listStyle = uniqueDirectChild(textBody, 'lstStyle', DRAWING_NAMESPACE);
  const level = listStyle
    ? uniqueDirectChild(listStyle, 'lvl1pPr', DRAWING_NAMESPACE)
    : undefined;
  return level ? uniqueDirectChild(level, 'defRPr', DRAWING_NAMESPACE) : undefined;
}

function patchSupportedTarget(
  xml: LosslessXmlDocument,
  target: SupportedSlideNumberTarget,
  value: SlideNumber,
  cachedText: string,
): void {
  xml.replaceElement(target.offset, rewriteElement(xml, target.offset, {
    x: String(value.x),
    y: String(value.y),
  }));
  xml.replaceElement(target.extent, rewriteElement(xml, target.extent, {
    cx: String(value.width),
    cy: String(value.height),
  }));
  xml.replaceElement(target.bodyProperties, rewriteElement(xml, target.bodyProperties, {
    anchor: value.valign === undefined
      ? undefined
      : { top: 't', middle: 'ctr', bottom: 'b' }[value.valign],
    lIns: marginRaw(value.margin?.left),
    tIns: marginRaw(value.margin?.top),
    rIns: marginRaw(value.margin?.right),
    bIns: marginRaw(value.margin?.bottom),
  }));
  const paragraphAttributes = {
    algn: { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[value.align],
    rtl: value.rtl ? '1' : '0',
  };
  if (target.paragraphProperties) {
    xml.replaceElement(
      target.paragraphProperties,
      rewriteElement(xml, target.paragraphProperties, paragraphAttributes),
    );
  } else {
    xml.replace(
      target.field.start,
      target.field.start,
      `<a:pPr xmlns:a="${DRAWING_NAMESPACE}" algn="${paragraphAttributes.algn}" rtl="${paragraphAttributes.rtl}"/>`,
    );
  }
  if (target.defaultProperties) {
    xml.replaceElement(target.defaultProperties, rewriteElement(
      xml,
      target.defaultProperties,
      { sz: undefined, lang: undefined, b: undefined, i: undefined },
      new Set(['solidFill', 'latin', 'ea', 'cs']),
      '',
    ));
  }
  const runAttributes = {
    lang: value.style.lang,
    b: value.style.bold ? '1' : '0',
    i: value.style.italic ? '1' : '0',
    sz: value.style.fontSize === undefined
      ? undefined
      : String(Math.round(value.style.fontSize * 100)),
  };
  const runChildren = renderStyleChildren(value.style);
  if (target.fieldProperties) {
    xml.replaceElement(target.fieldProperties, rewriteElement(
      xml,
      target.fieldProperties,
      runAttributes,
      new Set(['solidFill', 'latin', 'ea', 'cs']),
      runChildren,
    ));
  } else {
    xml.replace(
      target.text.start,
      target.text.start,
      renderRunProperties(value.style, true),
    );
  }
  xml.replaceElement(target.text, rewriteTextElement(xml, target.text, cachedText));
}

function rewriteElement(
  xml: LosslessXmlDocument,
  element: XmlElement,
  attributes: Readonly<Record<string, string | undefined>>,
  ownedChildren: ReadonlySet<string> = new Set(),
  newChildren = '',
): string {
  let startTag = xml.source.slice(element.start, element.startTagEnd);
  const removals = element.attributes
    .filter((attribute) => Object.hasOwn(attributes, attribute.name))
    .map((attribute) => ({
      start: attribute.start - element.start,
      end: attribute.end - element.start,
    }));
  for (const removal of removals.sort((left, right) => right.start - left.start)) {
    let start = removal.start;
    while (start > 0 && /\s/.test(startTag[start - 1]!)) start -= 1;
    startTag = startTag.slice(0, start) + startTag.slice(removal.end);
  }
  const renderedAttributes = Object.entries(attributes)
    .filter((entry): entry is [string, string] => entry[1] !== undefined)
    .map(([name, entryValue]) => ` ${name}="${escapeXmlAttribute(entryValue)}"`)
    .join('');
  const marker = element.selfClosing ? startTag.lastIndexOf('/>') : startTag.lastIndexOf('>');
  if (marker < 0) throw new Error(`Invalid XML start tag for ${element.name}`);
  startTag = startTag.slice(0, marker).replace(/\s+$/u, '')
    + renderedAttributes
    + startTag.slice(marker);

  const removedChildren = directChildren(element).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE
      && ownedChildren.has(child.localName),
  );
  if (element.selfClosing) {
    if (newChildren.length === 0) return startTag;
    return `${startTag.slice(0, startTag.lastIndexOf('/>'))}>${newChildren}</${element.name}>`;
  }
  const extension = directChildren(element).find(
    (child) => isElement(child, 'extLst', DRAWING_NAMESPACE),
  );
  const insertion = removedChildren[0]?.start ?? extension?.start ?? element.endTagStart;
  const content = rewriteContent(
    xml.source,
    element.startTagEnd,
    element.endTagStart,
    removedChildren,
    insertion,
    newChildren,
  );
  return startTag + content + xml.source.slice(element.endTagStart, element.end);
}

function rewriteContent(
  source: string,
  start: number,
  end: number,
  removals: readonly XmlElement[],
  insertion: number,
  inserted: string,
): string {
  let output = '';
  let cursor = start;
  let didInsert = false;
  for (const removal of [...removals].sort((left, right) => left.start - right.start)) {
    if (!didInsert && insertion >= cursor && insertion <= removal.start) {
      output += source.slice(cursor, insertion) + inserted + source.slice(insertion, removal.start);
      didInsert = true;
    } else {
      output += source.slice(cursor, removal.start);
    }
    cursor = removal.end;
  }
  if (!didInsert) {
    output += source.slice(cursor, insertion) + inserted;
    cursor = insertion;
  }
  return output + source.slice(cursor, end);
}

function rewriteTextElement(
  xml: LosslessXmlDocument,
  element: XmlElement,
  value: string,
): string {
  const startTag = xml.source.slice(element.start, element.startTagEnd);
  if (element.selfClosing) {
    const marker = startTag.lastIndexOf('/>');
    return `${startTag.slice(0, marker).replace(/\s+$/u, '')}>${escapeXmlText(value)}</${element.name}>`;
  }
  return startTag
    + escapeXmlText(value)
    + xml.source.slice(element.endTagStart, element.end);
}

function renderSlideNumberShape(
  value: SlideNumber,
  shapeId: number,
  placeholderIndex: number,
  cachedText: string,
): string {
  const bodyAttributes = [
    value.valign === undefined
      ? ''
      : ` anchor="${{ top: 't', middle: 'ctr', bottom: 'b' }[value.valign]}"`,
    renderMarginAttribute('lIns', value.margin?.left),
    renderMarginAttribute('tIns', value.margin?.top),
    renderMarginAttribute('rIns', value.margin?.right),
    renderMarginAttribute('bIns', value.margin?.bottom),
  ].join('');
  const align = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[value.align];
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">`
    + `<p:nvSpPr><p:cNvPr id="${shapeId}" name="Slide Number ${shapeId}"/>`
    + '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>'
    + `<p:ph type="sldNum" sz="quarter" idx="${placeholderIndex}"/></p:nvPr></p:nvSpPr>`
    + `<p:spPr><a:xfrm><a:off x="${value.x}" y="${value.y}"/>`
    + `<a:ext cx="${value.width}" cy="${value.height}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
    + `<p:txBody><a:bodyPr${bodyAttributes}/><a:lstStyle/><a:p>`
    + `<a:pPr algn="${align}" rtl="${value.rtl ? 1 : 0}"/>`
    + `<a:fld id="${FIELD_ID}" type="slidenum">${renderRunProperties(value.style, false)}`
    + `<a:t>${escapeXmlText(cachedText)}</a:t></a:fld>`
    + `<a:endParaRPr lang="${escapeXmlAttribute(value.style.lang)}"/></a:p></p:txBody></p:sp>`;
}

function renderRunProperties(style: SlideNumberTextStyle, includeNamespace: boolean): string {
  const attributes = [
    includeNamespace ? ` xmlns:a="${DRAWING_NAMESPACE}"` : '',
    ` lang="${escapeXmlAttribute(style.lang)}"`,
    ` b="${style.bold ? 1 : 0}"`,
    ` i="${style.italic ? 1 : 0}"`,
    style.fontSize === undefined ? '' : ` sz="${Math.round(style.fontSize * 100)}"`,
  ].join('');
  const children = renderStyleChildren(style);
  return children.length === 0
    ? `<a:rPr${attributes}/>`
    : `<a:rPr${attributes}>${children}</a:rPr>`;
}

function renderStyleChildren(style: SlideNumberTextStyle): string {
  const fill = style.color === undefined
    ? ''
    : renderColor(style.color, style.transparency);
  const family = style.fontFamily === undefined
    ? ''
    : `<a:latin typeface="${escapeXmlAttribute(style.fontFamily)}"/>`
      + `<a:ea typeface="${escapeXmlAttribute(style.fontFamily)}"/>`
      + `<a:cs typeface="${escapeXmlAttribute(style.fontFamily)}"/>`;
  return fill + family;
}

function renderColor(color: SlideNumberColor, transparency: number | undefined): string {
  const name = color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const alpha = transparency === undefined
    ? ''
    : `<a:alpha val="${Math.round((100 - transparency) * 1_000)}"/>`;
  const encoded = escapeXmlAttribute(color.value);
  return alpha.length === 0
    ? `<a:solidFill><a:${name} val="${encoded}"/></a:solidFill>`
    : `<a:solidFill><a:${name} val="${encoded}">${alpha}</a:${name}></a:solidFill>`;
}

function renderMarginAttribute(name: string, value: number | undefined): string {
  const raw = marginRaw(value);
  return raw === undefined ? '' : ` ${name}="${raw}"`;
}

function marginRaw(value: number | undefined): string | undefined {
  return value === undefined ? undefined : String(Math.round(value * EMU_PER_POINT));
}

function allocateShapeId(root: XmlElement): number {
  const used = new Set(descendantsAndSelf(root)
    .filter((element) => isElement(element, 'cNvPr', PRESENTATION_NAMESPACE))
    .map((element) => readUnsignedInteger(element, 'id', 1, MAX_UINT32))
    .filter((value): value is number => value !== INVALID));
  let candidate = 1;
  while (used.has(candidate)) candidate += 1;
  return candidate;
}

function allocatePlaceholderIndex(root: XmlElement): number {
  const used = placeholderIndexes(root);
  if (!used.has(MAX_UINT32)) return MAX_UINT32;
  let candidate = 0;
  while (used.has(candidate)) candidate += 1;
  return candidate;
}

function placeholderIndexes(root: XmlElement): Set<number> {
  return new Set(descendantsAndSelf(root)
    .filter((element) => isElement(element, 'ph', PRESENTATION_NAMESPACE))
    .map((element) => readUnsignedInteger(element, 'idx', 0, MAX_UINT32))
    .filter((value): value is number => value !== INVALID));
}

function opaqueShapeIdentity(
  root: XmlElement,
  shape: XmlElement,
): { readonly shapeId: number; readonly placeholderIndex: number } | undefined {
  const nonVisual = uniqueDirectChild(shape, 'nvSpPr', PRESENTATION_NAMESPACE);
  const properties = nonVisual
    ? uniqueDirectChild(nonVisual, 'cNvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const application = nonVisual
    ? uniqueDirectChild(nonVisual, 'nvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const placeholder = application
    ? uniqueDirectChild(application, 'ph', PRESENTATION_NAMESPACE)
    : undefined;
  if (!properties || !placeholder) return undefined;
  const shapeId = readUnsignedInteger(properties, 'id', 1, MAX_UINT32);
  const placeholderIndex = readUnsignedInteger(placeholder, 'idx', 0, MAX_UINT32);
  if (shapeId === INVALID || placeholderIndex === INVALID) return undefined;
  const matchingIds = descendantsAndSelf(root)
    .filter((element) => isElement(element, 'cNvPr', PRESENTATION_NAMESPACE))
    .filter((element) => readUnsignedInteger(element, 'id', 1, MAX_UINT32) === shapeId);
  const matchingIndexes = descendantsAndSelf(root)
    .filter((element) => isElement(element, 'ph', PRESENTATION_NAMESPACE))
    .filter((element) => readUnsignedInteger(element, 'idx', 0, MAX_UINT32) === placeholderIndex);
  return matchingIds.length === 1 && matchingIndexes.length === 1
    ? { shapeId, placeholderIndex }
    : undefined;
}

function updateMasterFlag(
  xml: LosslessXmlDocument,
  root: XmlElement,
  enabled: boolean,
): void {
  const headers = directChildren(root).filter(
    (child) => isElement(child, 'hf', PRESENTATION_NAMESPACE),
  );
  if (headers.length > 1) throw new Error('Slide master has ambiguous header-footer state');
  if (headers[0]) {
    xml.replaceElement(headers[0], rewriteElement(xml, headers[0], {
      sldNum: enabled ? '1' : '0',
    }));
    return;
  }
  if (!enabled) return;
  const rootSeparator = root.name.indexOf(':');
  const rootPrefix = rootSeparator < 0 ? '' : root.name.slice(0, rootSeparator);
  const headerName = rootPrefix.length === 0 ? 'hf' : `${rootPrefix}:hf`;
  const rendered = `<${headerName} sldNum="1"/>`;
  const follower = directChildren(root).find(
    (child) => isElement(child, 'txStyles', PRESENTATION_NAMESPACE)
      || isElement(child, 'extLst', PRESENTATION_NAMESPACE),
  );
  if (follower) {
    xml.replace(follower.start, follower.start, rendered);
    return;
  }
  const layoutList = directChildren(root).find(
    (child) => isElement(child, 'sldLayoutIdLst', PRESENTATION_NAMESPACE),
  );
  if (layoutList) xml.replace(layoutList.end, layoutList.end, rendered);
  else xml.appendChildXml(root, rendered);
}

function masterFlagValue(root: XmlElement): boolean | undefined {
  const headers = directChildren(root).filter(
    (child) => isElement(child, 'hf', PRESENTATION_NAMESPACE),
  );
  if (headers.length !== 1) return headers.length === 0 ? undefined : false;
  const value = directAttribute(headers[0]!, 'sldNum');
  if (value === undefined) return undefined;
  if (value === INVALID) return false;
  const parsed = parseBoolean(value);
  return parsed === INVALID ? false : parsed;
}

function masterSlideNumberExplicitlyDisabled(root: XmlElement): boolean {
  const headers = directChildren(root).filter(
    (child) => isElement(child, 'hf', PRESENTATION_NAMESPACE),
  );
  if (headers.length !== 1) return false;
  const value = directAttribute(headers[0]!, 'sldNum');
  return value !== undefined && value !== INVALID && parseBoolean(value) === false;
}

function slideNumbersEqual(left: SlideNumber, right: SlideNumber): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function textValue(element: XmlElement): string {
  return element.children
    .filter((child) => child.type === 'text')
    .map((child) => child.value)
    .join('');
}

function assertCachedText(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !isXmlSafe(value)) {
    throw new TypeError('Slide-number cached text must be an XML-safe string');
  }
}

function normalizeCoordinate(
  value: unknown,
  fallback: number,
  context: string,
): number {
  if (value === undefined) return fallback;
  const normalized = normalizeFiniteInteger(value, context);
  if (normalized < -MAX_COORDINATE || normalized > MAX_COORDINATE) {
    throw new RangeError(`${context} must fit the DrawingML coordinate range`);
  }
  return normalized;
}

function normalizeExtent(
  value: unknown,
  fallback: number,
  context: string,
): number {
  if (value === undefined) return fallback;
  const normalized = normalizeFiniteInteger(value, context);
  if (normalized <= 0 || normalized > MAX_COORDINATE) {
    throw new RangeError(`${context} must be positive and fit the DrawingML coordinate range`);
  }
  return normalized;
}

function normalizeFiniteInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  return Math.round(value);
}

function normalizeStyle(value: unknown): SlideNumberTextStyle {
  const input = value === undefined
    ? Object.create(null) as Record<string, unknown>
    : readDataObject(value, STYLE_KEYS, 'Slide number style');
  const fontFamily = input.fontFamily === undefined
    ? undefined
    : normalizeXmlString(input.fontFamily, 'Slide number font family');
  const fontSize = input.fontSize === undefined
    ? undefined
    : normalizeFontSize(input.fontSize);
  const lang = input.lang === undefined
    ? 'en-US'
    : normalizeXmlString(input.lang, 'Slide number language');
  const bold = normalizeBoolean(input.bold, false, 'Slide number bold');
  const italic = normalizeBoolean(input.italic, false, 'Slide number italic');
  let color = input.color === undefined ? undefined : normalizeColor(input.color);
  const transparency = input.transparency === undefined
    ? undefined
    : normalizeTransparency(input.transparency);
  if (transparency !== undefined && color === undefined) {
    color = { kind: 'scheme', value: 'tx1' };
  }
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    lang,
    bold,
    italic,
    ...(color === undefined ? {} : { color }),
    ...(transparency === undefined ? {} : { transparency }),
  };
}

function normalizeFontSize(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Slide number font size must be finite');
  }
  if (value < 1 || value > 4_000) {
    throw new RangeError('Slide number font size must be between 1 and 4000 points');
  }
  return Math.round(value * 100) / 100;
}

function normalizeTransparency(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Slide number transparency must be finite');
  }
  if (value < 0 || value > 100) {
    throw new RangeError('Slide number transparency must be between 0 and 100');
  }
  return Math.round(value * 1_000) / 1_000;
}

function normalizeColor(value: unknown): SlideNumberColor {
  const input = readDataObject(value, COLOR_KEYS, 'Slide number color');
  if (input.kind === 'srgb') {
    if (typeof input.value !== 'string' || !/^[\da-f]{6}$/i.test(input.value)) {
      throw new TypeError('Slide number sRGB color must contain six hex digits');
    }
    return { kind: 'srgb', value: input.value.toUpperCase() };
  }
  if (input.kind === 'scheme') {
    if (typeof input.value !== 'string' || !SCHEME_COLORS.has(input.value)) {
      throw new TypeError('Slide number scheme color is unsupported');
    }
    return { kind: 'scheme', value: input.value };
  }
  throw new TypeError('Slide number color kind must be srgb or scheme');
}

function normalizeMargins(value: unknown): SlideNumberMargins | undefined {
  if (value === undefined) return undefined;
  if (typeof value === 'number') {
    const side = normalizeMarginSide(value, 'Slide number margin');
    return { top: side, right: side, bottom: side, left: side };
  }
  if (Array.isArray(value)) return normalizeMarginTuple(value);
  const input = readDataObject(value, MARGIN_KEYS, 'Slide number margin');
  const margin: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    if (input[side] !== undefined) {
      margin[side] = normalizeMarginSide(input[side], `Slide number margin ${side}`);
    }
  }
  return Object.keys(margin).length === 0 ? undefined : margin;
}

function normalizeMarginTuple(value: unknown[]): SlideNumberMargins {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError('Slide number margin tuple must be an ordinary array');
  }
  const length = Object.getOwnPropertyDescriptor(value, 'length');
  if (!length || !Object.hasOwn(length, 'value') || length.value !== 4) {
    throw new RangeError('Slide number margin tuple must contain exactly four values');
  }
  const allowed = new Set(['0', '1', '2', '3', 'length']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`Slide number margin tuple contains unsupported property ${String(key)}`);
    }
  }
  const sides = Array.from({ length: 4 }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor) {
      throw new TypeError('Slide number margin tuple must not contain sparse values');
    }
    if (!Object.hasOwn(descriptor, 'value')) {
      throw new TypeError('Slide number margin tuple must contain only data properties');
    }
    return normalizeMarginSide(
      descriptor.value,
      `Slide number margin ${['top', 'right', 'bottom', 'left'][index]}`,
    );
  });
  return { top: sides[0]!, right: sides[1]!, bottom: sides[2]!, left: sides[3]! };
}

function normalizeMarginSide(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * EMU_PER_POINT);
  if (raw < MIN_INT32 || raw > MAX_INT32) {
    throw new RangeError(`${context} must fit the OOXML signed Int32 coordinate range`);
  }
  return raw / EMU_PER_POINT;
}

function normalizeXmlString(value: unknown, context: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`${context} must be a non-whitespace string`);
  }
  if (!isXmlSafe(value)) throw new TypeError(`${context} contains invalid XML characters`);
  return value;
}

function isXmlSafe(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if ((code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d)
      || code === 0xfffe || code === 0xffff) {
      return false;
    }
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function normalizeBoolean(
  value: unknown,
  fallback: boolean,
  context: string,
): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

function normalizeEnum<const T extends string>(
  value: unknown,
  supported: readonly T[],
  fallback: T | undefined,
  context: string,
): T {
  if (value === undefined && fallback !== undefined) return fallback;
  if (typeof value !== 'string' || !supported.includes(value as T)) {
    throw new TypeError(`${context} must be ${supported.join(', ')}`);
  }
  return value as T;
}

function readDataObject(
  value: unknown,
  supported: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !supported.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function uniqueRoot(
  xml: LosslessXmlDocument,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const roots = xml.roots.filter((root) => isElement(root, localName, namespace));
  return roots.length === 1 && xml.roots.length === 1 ? roots[0] : undefined;
}

function containsDirectSlideNumberPlaceholder(shape: XmlElement): boolean {
  return directChildren(shape)
    .filter((child) => isElement(child, 'nvSpPr', PRESENTATION_NAMESPACE))
    .some((nonVisual) => directChildren(nonVisual)
      .filter((child) => isElement(child, 'nvPr', PRESENTATION_NAMESPACE))
      .some((application) => directChildren(application)
        .filter((child) => isElement(child, 'ph', PRESENTATION_NAMESPACE))
        .some((placeholder) => placeholder.attributes.some(
          (attribute) => attribute.name === 'type' && attribute.value === 'sldNum',
        ))));
}

function validUniqueShapeId(
  xml: LosslessXmlDocument,
  root: XmlElement,
  shape: XmlElement,
): boolean {
  const nonVisual = uniqueDirectChild(shape, 'nvSpPr', PRESENTATION_NAMESPACE);
  const properties = nonVisual
    ? uniqueDirectChild(nonVisual, 'cNvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const application = nonVisual
    ? uniqueDirectChild(nonVisual, 'nvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const placeholder = application
    ? uniqueDirectChild(application, 'ph', PRESENTATION_NAMESPACE)
    : undefined;
  if (!properties || !placeholder || directAttribute(placeholder, 'type') !== 'sldNum') {
    return false;
  }
  const id = readUnsignedInteger(properties, 'id', 1, MAX_UINT32);
  if (id === INVALID) return false;
  return descendantsAndSelf(root)
    .filter((element) => isElement(element, 'cNvPr', PRESENTATION_NAMESPACE))
    .filter((element) => readUnsignedInteger(element, 'id', 1, MAX_UINT32) === id)
    .length === 1
    && xml.roots.length === 1;
}

function readUniqueSlideNumberField(paragraph: XmlElement): XmlElement | undefined {
  const drawingChildren = directChildren(paragraph).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  const fields = drawingChildren.filter((child) => child.localName === 'fld');
  const slideNumberFields = fields.filter(
    (field) => directAttribute(field, 'type') === 'slidenum',
  );
  if (fields.length !== 1 || slideNumberFields.length !== 1) return undefined;
  if (drawingChildren.some((child) => !['pPr', 'fld', 'endParaRPr'].includes(child.localName))) {
    return undefined;
  }
  const field = slideNumberFields[0]!;
  const fieldChildren = directChildren(field).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (fieldChildren.some((child) => !['rPr', 't'].includes(child.localName))) {
    return undefined;
  }
  const text = fieldChildren.filter((child) => child.localName === 't');
  return text.length === 1 && directChildren(text[0]!).length === 0
    ? field
    : undefined;
}

function readListStyle(
  listStyle: XmlElement | undefined,
): PartialSlideNumberStyle | typeof INVALID {
  if (!listStyle) return {};
  const levels = directChildren(listStyle).filter(
    (child) => isElement(child, 'lvl1pPr', DRAWING_NAMESPACE),
  );
  if (levels.length > 1) return INVALID;
  if (levels.length === 0) return {};
  const defaults = directChildren(levels[0]!).filter(
    (child) => isElement(child, 'defRPr', DRAWING_NAMESPACE),
  );
  if (defaults.length > 1) return INVALID;
  return readStyleElement(defaults[0]);
}

function readStyleElement(
  properties: XmlElement | undefined,
): PartialSlideNumberStyle | typeof INVALID {
  if (!properties) return {};
  const size = readOptionalInteger(properties, 'sz', 100, 400_000);
  const lang = readOptionalXmlString(properties, 'lang');
  const bold = readOptionalBooleanAttribute(properties, 'b');
  const italic = readOptionalBooleanAttribute(properties, 'i');
  const fontFamily = readFontFamily(properties);
  const fill = readStyleColor(properties);
  if (size === INVALID || lang === INVALID || bold === INVALID
    || italic === INVALID || fontFamily === INVALID || fill === INVALID) {
    return INVALID;
  }
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(size === undefined ? {} : { fontSize: size / 100 }),
    ...(lang === undefined ? {} : { lang }),
    ...(bold === undefined ? {} : { bold }),
    ...(italic === undefined ? {} : { italic }),
    ...(fill?.color === undefined ? {} : { color: fill.color }),
    ...(fill?.transparency === undefined ? {} : { transparency: fill.transparency }),
  };
}

function readFontFamily(properties: XmlElement): string | undefined | typeof INVALID {
  const families: string[] = [];
  for (const name of ['latin', 'ea', 'cs']) {
    const elements = directChildren(properties).filter(
      (child) => isElement(child, name, DRAWING_NAMESPACE),
    );
    if (elements.length > 1) return INVALID;
    if (elements[0]) {
      const typeface = directAttribute(elements[0], 'typeface');
      if (typeface === INVALID || typeface === undefined || !isValidXmlString(typeface)) {
        return INVALID;
      }
      families.push(typeface);
    }
  }
  const unique = [...new Set(families)];
  return unique.length > 1 ? INVALID : unique[0];
}

function readStyleColor(
  properties: XmlElement,
): { readonly color: SlideNumberColor; readonly transparency?: number }
  | undefined
  | typeof INVALID {
  const fillChoices = directChildren(properties).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE
      && ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(child.localName),
  );
  if (fillChoices.length === 0) return undefined;
  if (fillChoices.length !== 1 || fillChoices[0]?.localName !== 'solidFill') return INVALID;
  const colorElements = directChildren(fillChoices[0]).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (colorElements.length !== 1) return INVALID;
  const element = colorElements[0]!;
  const raw = directAttribute(element, 'val');
  let color: SlideNumberColor;
  if (element.localName === 'srgbClr') {
    if (typeof raw !== 'string' || !/^[\da-f]{6}$/i.test(raw)) return INVALID;
    color = { kind: 'srgb', value: raw.toUpperCase() };
  } else if (element.localName === 'schemeClr') {
    if (typeof raw !== 'string' || !SCHEME_COLORS.has(raw)) return INVALID;
    color = { kind: 'scheme', value: raw };
  } else {
    return INVALID;
  }
  const transforms = directChildren(element).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (transforms.some((child) => child.localName !== 'alpha')) return INVALID;
  if (transforms.length > 1) return INVALID;
  if (transforms[0] && directChildren(transforms[0]).length > 0) return INVALID;
  const alpha = transforms[0]
    ? readUnsignedInteger(transforms[0], 'val', 0, 100_000)
    : undefined;
  if (alpha === INVALID) return INVALID;
  return {
    color,
    ...(alpha === undefined ? {} : { transparency: 100 - alpha / 1_000 }),
  };
}

function mergeStyles(
  defaults: PartialSlideNumberStyle,
  direct: PartialSlideNumberStyle,
): SlideNumberTextStyle {
  const color = direct.color ?? defaults.color;
  const transparency = direct.color === undefined
    ? defaults.transparency
    : direct.transparency;
  return {
    ...(direct.fontFamily === undefined && defaults.fontFamily === undefined
      ? {}
      : { fontFamily: direct.fontFamily ?? defaults.fontFamily }),
    ...(direct.fontSize === undefined && defaults.fontSize === undefined
      ? {}
      : { fontSize: direct.fontSize ?? defaults.fontSize }),
    lang: direct.lang ?? defaults.lang ?? 'en-US',
    bold: direct.bold ?? defaults.bold ?? false,
    italic: direct.italic ?? defaults.italic ?? false,
    ...(color === undefined ? {} : { color }),
    ...(transparency === undefined ? {} : { transparency }),
  };
}

function readMargins(
  element: XmlElement,
): SlideNumberMargins | undefined | typeof INVALID {
  const mapping = [
    ['top', 'tIns'],
    ['right', 'rIns'],
    ['bottom', 'bIns'],
    ['left', 'lIns'],
  ] as const;
  const result: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side, attribute] of mapping) {
    const raw = readOptionalSignedInteger(element, attribute, MIN_INT32, MAX_INT32);
    if (raw === INVALID) return INVALID;
    if (raw !== undefined) result[side] = raw / EMU_PER_POINT;
  }
  return Object.keys(result).length === 0 ? undefined : result;
}

function readCoordinate(element: XmlElement, name: string): number | typeof INVALID {
  return readSignedInteger(element, name, -MAX_COORDINATE, MAX_COORDINATE);
}

function readExtent(element: XmlElement, name: string): number | typeof INVALID {
  return readUnsignedInteger(element, name, 1, MAX_COORDINATE);
}

function readSignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | typeof INVALID {
  const value = directAttribute(element, name);
  if (typeof value !== 'string' || !/^-?\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readUnsignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | typeof INVALID {
  const value = directAttribute(element, name);
  if (typeof value !== 'string' || !/^\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalSignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  if (value === undefined) return undefined;
  if (!/^-?\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  if (value === undefined) return undefined;
  if (!/^\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalXmlString(
  element: XmlElement,
  name: string,
): string | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined || isValidXmlString(value) ? value : INVALID;
}

function isValidXmlString(value: string): boolean {
  return /\S/u.test(value) && isXmlSafe(value);
}

function readMappedAttribute<const T>(
  element: XmlElement,
  name: string,
  mapping: ReadonlyMap<string, T>,
  fallback: T,
): T | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? fallback : mapping.get(value) ?? INVALID;
}

function readOptionalMappedAttribute<const T>(
  element: XmlElement,
  name: string,
  mapping: ReadonlyMap<string, T>,
): T | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? undefined : mapping.get(value) ?? INVALID;
}

function readBooleanAttribute(
  element: XmlElement,
  name: string,
  fallback: boolean,
): boolean | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? fallback : parseBoolean(value);
}

function readOptionalBooleanAttribute(
  element: XmlElement,
  name: string,
): boolean | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? undefined : parseBoolean(value);
}

function parseBoolean(value: string): boolean | typeof INVALID {
  if (['1', 'true', 'on'].includes(value)) return true;
  if (['0', 'false', 'off'].includes(value)) return false;
  return INVALID;
}

function masterSlideNumberEnabled(root: XmlElement): boolean {
  const headers = directChildren(root).filter(
    (child) => isElement(child, 'hf', PRESENTATION_NAMESPACE),
  );
  if (headers.length > 1) return false;
  if (headers.length === 0) return true;
  const value = directAttribute(headers[0]!, 'sldNum');
  if (value === INVALID) return false;
  if (value === undefined) return true;
  return parseBoolean(value) === true;
}

function directAttribute(
  element: XmlElement,
  name: string,
): string | undefined | typeof INVALID {
  const matches = element.attributes.filter((attribute) => attribute.name === name);
  if (matches.length > 1) return INVALID;
  return matches[0]?.value;
}

function uniqueDirectChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const matches = directChildren(parent).filter(
    (child) => isElement(child, localName, namespace),
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function optionalUniqueDirectChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined | typeof INVALID {
  const matches = directChildren(parent).filter(
    (child) => isElement(child, localName, namespace),
  );
  if (matches.length > 1) return INVALID;
  return matches[0];
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function descendantsAndSelf(element: XmlElement): XmlElement[] {
  const result = [element];
  for (const child of directChildren(element)) result.push(...descendantsAndSelf(child));
  return result;
}

function isElement(
  element: XmlElement,
  localName: string,
  namespace: string,
): boolean {
  return element.localName === localName && elementNamespaceUri(element) === namespace;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  let current: XmlElement | undefined = element;
  while (current) {
    const declarations = current.attributes.filter(
      (attribute) => attribute.name === declarationName,
    );
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) deepFreeze(child);
    Object.freeze(value);
  }
  return value;
}
