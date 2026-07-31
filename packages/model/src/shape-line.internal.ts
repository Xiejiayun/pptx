import {
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { ShapeLine } from './preset-shape.js';
import { readSimpleFillChoice } from './simple-fill.internal.js';
import {
  type NormalizedSimpleLine,
  readSimpleLine,
  renderSimpleLine,
  SIMPLE_LINE_DASH_CHOICE_NAMES,
  SIMPLE_LINE_FILL_CHOICE_NAMES,
  simpleLinesEqual,
} from './simple-line.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const EMU_PER_POINT = 12_700;
const MAX_LINE_WIDTH_EMUS = 20_116_800;
const FILL_CHOICES = new Set<string>(SIMPLE_LINE_FILL_CHOICE_NAMES);
const DASH_CHOICES = new Set<string>(SIMPLE_LINE_DASH_CHOICE_NAMES);
const PRESET_DASHES = new Set([
  'solid',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
]);
const GEOMETRY_CHOICES = new Set(['prstGeom', 'custGeom']);
const LATER_PROPERTY_CHOICES = new Set([
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
]);

interface ExistingLineState {
  readonly line: XmlElement;
  readonly width: XmlAttribute | undefined;
  readonly fill: XmlElement | undefined;
  readonly dash: XmlElement | undefined;
  readonly prefix: string;
}

interface LocalEdit {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export function readShapeLine(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): ShapeLine | undefined {
  const properties = resolveShapeProperties(shape);
  if (!properties) return undefined;
  const candidates = directChildren(properties).filter(({ localName }) => localName === 'ln');
  if (candidates.length !== 1) return undefined;
  const line = candidates[0]!;
  if (namespaceUri(line) !== DRAWING_NAMESPACE) return undefined;
  const prefix = renderPrefix(line);
  return readSimpleLine(line, prefix);
}

export function replaceShapeLine(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  line: NormalizedSimpleLine | undefined,
  partUri: string,
): boolean {
  const properties = resolveShapeProperties(shape);
  if (!properties) {
    throw new ModelParseError(
      'Shape must contain exactly one direct shape properties element',
      partUri,
    );
  }

  const candidates = directChildren(properties).filter(({ localName }) => localName === 'ln');
  if (candidates.length > 1) {
    throw new ModelParseError('Shape contains multiple direct line elements', partUri);
  }
  const candidate = candidates[0];
  if (candidate && namespaceUri(candidate) !== DRAWING_NAMESPACE) {
    throw new ModelParseError('Shape line uses an unsafe namespace', partUri);
  }
  if (!candidate) {
    if (line === undefined) return false;
    const anchor = resolveInsertionAnchor(properties, partUri);
    const encoded = renderLineForParent(line, anchor, properties);
    xml.replace(anchor.end, anchor.end, encoded);
    return true;
  }

  const state = analyzeExistingLine(candidate, partUri);
  const prefix = state.prefix === '' ? '' : `${state.prefix}:`;
  const current = readSimpleLine(candidate, prefix);
  if (line !== undefined && simpleLinesEqual(current, line)) return false;
  if (line === undefined && !state.width && !state.fill && !state.dash) return false;

  const original = xml.original(candidate);
  const updated = patchExistingLine(original, state, line);
  if (updated === original) return false;
  xml.replaceElement(candidate, updated);
  return true;
}

function analyzeExistingLine(
  line: XmlElement,
  partUri: string,
): ExistingLineState {
  const widths = line.attributes.filter(({ name }) => name === 'w');
  if (widths.length > 1) {
    throw new ModelParseError('Shape line contains repeated width attributes', partUri);
  }
  if (widths[0]) {
    const width = widths[0].value;
    const emus = /^\d+$/.test(width) ? Number(width) : Number.NaN;
    if (!Number.isSafeInteger(emus) || emus < 0 || emus > MAX_LINE_WIDTH_EMUS) {
      throw new ModelParseError('Shape line contains an invalid width', partUri);
    }
  }

  const children = directChildren(line);
  const fills = children.filter(({ localName }) => FILL_CHOICES.has(localName));
  if (fills.length > 1) {
    throw new ModelParseError('Shape line contains multiple fill choices', partUri);
  }
  if (fills[0] && namespaceUri(fills[0]) !== DRAWING_NAMESPACE) {
    throw new ModelParseError('Shape line fill uses an unsafe namespace', partUri);
  }
  if (
    fills[0]
    && (fills[0].localName === 'noFill' || fills[0].localName === 'solidFill')
    && (
      !subtreeUsesDrawingNamespace(fills[0])
      || readSimpleFillChoice(fills[0], renderPrefix(fills[0])) === undefined
    )
  ) {
    throw new ModelParseError('Shape line contains an invalid simple fill', partUri);
  }

  const dashes = children.filter(({ localName }) => DASH_CHOICES.has(localName));
  if (dashes.length > 1) {
    throw new ModelParseError('Shape line contains multiple dash choices', partUri);
  }
  if (dashes[0] && namespaceUri(dashes[0]) !== DRAWING_NAMESPACE) {
    throw new ModelParseError('Shape line dash uses an unsafe namespace', partUri);
  }
  if (dashes[0]?.localName === 'prstDash') {
    const attributes = dashes[0].attributes.filter(
      ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
    );
    if (
      directChildren(dashes[0]).length !== 0
      || attributes.length !== 1
      || attributes[0]?.name !== 'val'
      || !PRESET_DASHES.has(attributes[0].value)
    ) {
      throw new ModelParseError('Shape line contains an invalid preset dash', partUri);
    }
  }

  return {
    line,
    width: widths[0],
    fill: fills[0],
    dash: dashes[0],
    prefix: lexicalPrefix(line.name),
  };
}

function patchExistingLine(
  source: string,
  state: ExistingLineState,
  target: NormalizedSimpleLine | undefined,
): string {
  const offset = state.line.start;
  const edits: LocalEdit[] = [];
  const prefix = state.prefix === '' ? '' : `${state.prefix}:`;
  const renderedTarget = target === undefined
    ? undefined
    : renderSimpleLine(target, prefix);
  const targetFill = target?.kind === 'line'
    ? renderedTarget!.slice(0, renderedTarget!.indexOf(`<${prefix}prstDash`))
    : renderedTarget;
  const targetDash = target?.kind === 'line'
    ? `<${prefix}prstDash val="${target.dash}"/>`
    : undefined;

  if (target?.kind === 'line') {
    const width = String(Math.round(target.width * EMU_PER_POINT));
    if (state.width) {
      if (state.width.value !== width) {
        edits.push({
          start: state.width.valueStart - offset,
          end: state.width.valueEnd - offset,
          replacement: width,
        });
      }
    } else if (!state.line.selfClosing) {
      const insertion = state.line.startTagEnd - offset - 1;
      edits.push({ start: insertion, end: insertion, replacement: ` w="${width}"` });
    }
  } else if (state.width) {
    edits.push(removeAttributeEdit(source, state.width, offset));
  }

  if (state.line.selfClosing) {
    if (target === undefined) return applyLocalEdits(source, edits);
    const marker = source.lastIndexOf('/>');
    const width = target.kind === 'line' && !state.width
      ? ` w="${Math.round(target.width * EMU_PER_POINT)}"`
      : '';
    edits.push({
      start: marker,
      end: source.length,
      replacement: `${width}>${targetFill ?? ''}${targetDash ?? ''}</${state.line.name}>`,
    });
    return applyLocalEdits(source, edits);
  }

  if (state.fill) {
    edits.push({
      start: state.fill.start - offset,
      end: state.fill.end - offset,
      replacement: targetFill ?? '',
    });
  } else if (targetFill !== undefined) {
    const insertion = state.dash?.start
      ?? directChildren(state.line)[0]?.start
      ?? state.line.endTagStart;
    const combined = state.dash || targetDash === undefined
      ? targetFill
      : targetFill + targetDash;
    edits.push({
      start: insertion - offset,
      end: insertion - offset,
      replacement: combined,
    });
  }

  if (state.dash) {
    edits.push({
      start: state.dash.start - offset,
      end: state.dash.end - offset,
      replacement: targetDash ?? '',
    });
  } else if (targetDash !== undefined && state.fill) {
    edits.push({
      start: state.fill.end - offset,
      end: state.fill.end - offset,
      replacement: targetDash,
    });
  }

  return applyLocalEdits(source, edits);
}

function removeAttributeEdit(
  source: string,
  attribute: XmlAttribute,
  offset: number,
): LocalEdit {
  let start = attribute.start - offset;
  while (start > 0 && /[\t ]/.test(source[start - 1] ?? '')) start -= 1;
  return {
    start,
    end: attribute.end - offset,
    replacement: '',
  };
}

function applyLocalEdits(source: string, edits: readonly LocalEdit[]): string {
  let output = source;
  const ordered = [...edits].sort(
    (left, right) => right.start - left.start || right.end - left.end,
  );
  let previousStart = source.length;
  for (const edit of ordered) {
    if (
      edit.start < 0
      || edit.end < edit.start
      || edit.end > source.length
      || edit.end > previousStart
    ) {
      throw new Error('Overlapping local shape line edits');
    }
    output = output.slice(0, edit.start) + edit.replacement + output.slice(edit.end);
    previousStart = edit.start;
  }
  return output;
}

function resolveInsertionAnchor(
  properties: XmlElement,
  partUri: string,
): XmlElement {
  const children = directChildren(properties);
  const geometries = children.filter(({ localName }) => GEOMETRY_CHOICES.has(localName));
  if (
    geometries.length !== 1
    || namespaceUri(geometries[0]!) !== DRAWING_NAMESPACE
  ) {
    throw new ModelParseError(
      'Shape must contain exactly one direct geometry before inserting a line',
      partUri,
    );
  }
  const fills = children.filter(({ localName }) => FILL_CHOICES.has(localName));
  if (fills.length > 1 || (fills[0] && namespaceUri(fills[0]) !== DRAWING_NAMESPACE)) {
    throw new ModelParseError('Shape contains unsafe direct fill state', partUri);
  }

  const geometry = geometries[0]!;
  const fill = fills[0];
  const geometryIndex = children.indexOf(geometry);
  const anchor = fill ?? geometry;
  const anchorIndex = children.indexOf(anchor);
  if (fill && anchorIndex <= geometryIndex) {
    throw new ModelParseError('Shape fill is not in a safe line insertion position', partUri);
  }
  const transforms = children.filter(
    ({ localName }) => localName === 'xfrm',
  );
  if (transforms.length > 1 || transforms.some((child) => children.indexOf(child) >= geometryIndex)) {
    throw new ModelParseError('Shape transform is not in a safe line insertion position', partUri);
  }
  if (children.some(
    (child, index) => LATER_PROPERTY_CHOICES.has(child.localName)
      && namespaceUri(child) === DRAWING_NAMESPACE
      && index < anchorIndex,
  )) {
    throw new ModelParseError('Shape properties have an unsafe line insertion order', partUri);
  }
  return anchor;
}

function renderLineForParent(
  line: NormalizedSimpleLine,
  prefixSource: XmlElement,
  parent: XmlElement,
): string {
  const prefix = lexicalPrefix(prefixSource.name);
  const qualifiedPrefix = prefix === '' ? '' : `${prefix}:`;
  const namespaceDeclaration = namespaceUriForPrefix(parent, prefix) === DRAWING_NAMESPACE
    ? ''
    : prefix === ''
      ? ` xmlns="${DRAWING_NAMESPACE}"`
      : ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`;
  if (line.kind === 'none') {
    return `<${qualifiedPrefix}ln${namespaceDeclaration}>` +
      `${renderSimpleLine(line, qualifiedPrefix)}</${qualifiedPrefix}ln>`;
  }
  return `<${qualifiedPrefix}ln${namespaceDeclaration} w="${Math.round(line.width * EMU_PER_POINT)}">` +
    `${renderSimpleLine(line, qualifiedPrefix)}</${qualifiedPrefix}ln>`;
}

function resolveShapeProperties(shape: XmlElement): XmlElement | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const candidates = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (candidates.length !== 1) return undefined;
  const properties = candidates[0]!;
  return namespaceUri(properties) === PRESENTATION_NAMESPACE
    ? properties
    : undefined;
}

function renderPrefix(element: XmlElement): string {
  const prefix = lexicalPrefix(element.name);
  return prefix === '' ? '' : `${prefix}:`;
}

function subtreeUsesDrawingNamespace(element: XmlElement): boolean {
  if (namespaceUri(element) !== DRAWING_NAMESPACE) return false;
  return directChildren(element).every(subtreeUsesDrawingNamespace);
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
