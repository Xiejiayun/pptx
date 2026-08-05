import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readSimpleShadow,
  renderSimpleShadow,
  shapeShadowsEqual,
  type NormalizedShapeShadow,
} from './simple-shadow.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const EMU_PER_POINT = 12_700;
const ANGLE_SCALE = 60_000;
const OPACITY_SCALE = 100_000;
const PROPERTY_STAGES = new Map<string, number>([
  ['xfrm', 0],
  ['prstGeom', 1],
  ['custGeom', 1],
  ['noFill', 2],
  ['solidFill', 2],
  ['gradFill', 2],
  ['blipFill', 2],
  ['pattFill', 2],
  ['grpFill', 2],
  ['ln', 3],
  ['effectLst', 4],
  ['effectDag', 4],
  ['scene3d', 5],
  ['sp3d', 6],
  ['extLst', 7],
]);
const EFFECT_STAGES = new Map<string, number>([
  ['blur', 0],
  ['fillOverlay', 1],
  ['glow', 2],
  ['innerShdw', 3],
  ['outerShdw', 4],
  ['prstShdw', 5],
  ['reflection', 6],
  ['softEdge', 7],
]);

interface LocalEdit {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

interface ExistingShadowState {
  readonly element: XmlElement;
  readonly snapshot: NormalizedShapeShadow;
  readonly prefix: string;
  readonly attributes: ReadonlyMap<string, XmlAttribute>;
  readonly color: XmlElement;
  readonly colorValue: XmlAttribute;
  readonly alphaValue: XmlAttribute | undefined;
}

interface ShapeShadowOwnerState {
  readonly properties: XmlElement;
  readonly propertyChildren: readonly XmlElement[];
  readonly effectList: XmlElement | undefined;
  readonly effectChildren: readonly XmlElement[];
  readonly shadow: ExistingShadowState | undefined;
}

export function readShapeShadow(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeShadow | undefined {
  return inspectOwner(shape)?.shadow?.snapshot;
}

export function replaceShapeShadow(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  shadow: NormalizedShapeShadow | undefined,
  partUri: string,
): boolean {
  const state = inspectOwner(shape);
  if (!state) {
    throw new ModelParseError('Shape shadow state is not safely editable', partUri);
  }
  if (shapeShadowsEqual(state.shadow?.snapshot, shadow)) return false;

  if (!state.effectList) {
    if (!shadow) return false;
    let anchor: XmlElement | undefined;
    for (const child of state.propertyChildren) {
      if ((PROPERTY_STAGES.get(child.localName) ?? Number.MAX_SAFE_INTEGER) < 4) {
        anchor = child;
      }
    }
    if (!anchor) {
      throw new ModelParseError('Shape has no safe shadow insertion anchor', partUri);
    }
    xml.replace(
      anchor.end,
      anchor.end,
      renderEffectListForParent(shadow, anchor, state.properties),
    );
    return true;
  }

  if (!state.shadow) {
    if (!shadow) return false;
    insertIntoEffectList(xml, state.effectList, state.effectChildren, shadow);
    return true;
  }

  if (!shadow) {
    xml.removeElement(state.shadow.element);
    return true;
  }

  if (state.shadow.snapshot.kind !== shadow.kind) {
    xml.replaceElement(
      state.shadow.element,
      renderShadowForParent(shadow, state.shadow.prefix, state.effectList),
    );
    return true;
  }

  const original = xml.original(state.shadow.element);
  const updated = patchSameKindShadow(original, state.shadow, shadow);
  if (updated === original) return false;
  xml.replaceElement(state.shadow.element, updated);
  return true;
}

function inspectOwner(shape: XmlElement): ShapeShadowOwnerState | undefined {
  if (
    (shape.localName !== 'sp' && shape.localName !== 'pic')
    || namespaceUri(shape) !== PRESENTATION_NAMESPACE
  ) {
    return undefined;
  }
  const propertyCandidates = directChildren(shape).filter(
    ({ localName }) => localName === 'spPr',
  );
  if (propertyCandidates.length !== 1) return undefined;
  const properties = propertyCandidates[0]!;
  if (
    namespaceUri(properties) !== PRESENTATION_NAMESPACE
    || hasNonWhitespaceText(properties)
  ) return undefined;

  const propertyChildren = directChildren(properties);
  let previousStage = -1;
  const stageCounts = new Map<number, number>();
  let effectList: XmlElement | undefined;
  for (const child of propertyChildren) {
    const stage = PROPERTY_STAGES.get(child.localName);
    if (
      stage === undefined
      || namespaceUri(child) !== DRAWING_NAMESPACE
      || stage < previousStage
    ) return undefined;
    previousStage = stage;
    const count = (stageCounts.get(stage) ?? 0) + 1;
    stageCounts.set(stage, count);
    if (count > 1) return undefined;
    if (child.localName === 'effectDag') return undefined;
    if (child.localName === 'effectLst') effectList = child;
  }
  if (stageCounts.get(1) !== 1) return undefined;

  if (!effectList) {
    return {
      properties,
      propertyChildren,
      effectList: undefined,
      effectChildren: [],
      shadow: undefined,
    };
  }
  if (
    nonNamespaceAttributes(effectList).length !== 0
    || hasNonWhitespaceText(effectList)
  ) return undefined;

  const effectChildren = directChildren(effectList);
  let previousEffectStage = -1;
  const seen = new Set<string>();
  let shadow: ExistingShadowState | undefined;
  for (const child of effectChildren) {
    const stage = EFFECT_STAGES.get(child.localName);
    if (
      stage === undefined
      || namespaceUri(child) !== DRAWING_NAMESPACE
      || stage < previousEffectStage
      || seen.has(child.localName)
    ) return undefined;
    previousEffectStage = stage;
    seen.add(child.localName);
    if (child.localName === 'innerShdw' || child.localName === 'outerShdw') {
      if (shadow) return undefined;
      shadow = inspectShadow(child);
      if (!shadow) return undefined;
    }
  }

  return { properties, propertyChildren, effectList, effectChildren, shadow };
}

function inspectShadow(element: XmlElement): ExistingShadowState | undefined {
  if (hasNonWhitespaceText(element)) return undefined;
  const prefix = lexicalPrefix(element.name);
  const snapshot = readSimpleShadow(element, qualifiedPrefix(prefix));
  if (!snapshot) return undefined;
  const children = directChildren(element);
  const color = children[0];
  if (!color || hasNonWhitespaceText(color)) return undefined;
  const colorValue = nonNamespaceAttributes(color).find(({ name }) => name === 'val');
  if (!colorValue) return undefined;
  const alpha = directChildren(color)[0];
  if (alpha && hasNonWhitespaceText(alpha)) return undefined;
  const alphaValue = alpha
    ? nonNamespaceAttributes(alpha).find(({ name }) => name === 'val')
    : undefined;
  return {
    element,
    snapshot,
    prefix,
    attributes: new Map(
      nonNamespaceAttributes(element).map((attribute) => [attribute.name, attribute]),
    ),
    color,
    colorValue,
    alphaValue,
  };
}

function insertIntoEffectList(
  xml: LosslessXmlDocument,
  effectList: XmlElement,
  children: readonly XmlElement[],
  shadow: NormalizedShapeShadow,
): void {
  const prefix = lexicalPrefix(effectList.name);
  const encoded = renderShadowForParent(shadow, prefix, effectList);
  if (effectList.selfClosing) {
    const source = xml.original(effectList);
    const marker = source.lastIndexOf('/>');
    xml.replaceElement(
      effectList,
      source.slice(0, marker) + `>${encoded}</${effectList.name}>`,
    );
    return;
  }
  const targetStage = shadow.kind === 'inner' ? 3 : 4;
  const following = children.find(
    (child) => (EFFECT_STAGES.get(child.localName) ?? Number.MAX_SAFE_INTEGER) > targetStage,
  );
  const position = following?.start ?? effectList.endTagStart;
  xml.replace(position, position, encoded);
}

function patchSameKindShadow(
  source: string,
  state: ExistingShadowState,
  target: NormalizedShapeShadow,
): string {
  const offset = state.element.start;
  const edits: LocalEdit[] = [];
  const missingAttributes: string[] = [];
  const current = state.snapshot;

  if (
    current.kind === 'outer'
    && target.kind === 'outer'
    && current.rotateWithShape !== target.rotateWithShape
  ) {
    patchAttribute(
      edits,
      state.attributes.get('rotWithShape'),
      target.rotateWithShape ? '1' : '0',
      offset,
      missingAttributes,
      'rotWithShape',
    );
  }
  patchNumericAttribute(
    edits,
    state,
    'blurRad',
    current.blur,
    target.blur,
    Math.round(target.blur * EMU_PER_POINT),
    offset,
    missingAttributes,
  );
  patchNumericAttribute(
    edits,
    state,
    'dist',
    current.distance,
    target.distance,
    Math.round(target.distance * EMU_PER_POINT),
    offset,
    missingAttributes,
  );
  patchNumericAttribute(
    edits,
    state,
    'dir',
    current.angle,
    target.angle,
    Math.round(target.angle * ANGLE_SCALE),
    offset,
    missingAttributes,
  );
  if (missingAttributes.length > 0) {
    const position = state.element.startTagEnd - offset - 1;
    edits.push({
      start: position,
      end: position,
      replacement: missingAttributes.join(''),
    });
  }

  if (current.color.kind !== target.color.kind) {
    edits.push({
      start: state.color.start - offset,
      end: state.color.end - offset,
      replacement: renderColorForParent(target, state.prefix, state.element),
    });
  } else {
    if (current.color.value !== target.color.value) {
      edits.push({
        start: state.colorValue.valueStart - offset,
        end: state.colorValue.valueEnd - offset,
        replacement: escapeXmlAttribute(target.color.value),
      });
    }
    if (current.opacity !== target.opacity) {
      const opacity = String(Math.round(target.opacity * OPACITY_SCALE));
      if (state.alphaValue) {
        edits.push({
          start: state.alphaValue.valueStart - offset,
          end: state.alphaValue.valueEnd - offset,
          replacement: opacity,
        });
      } else {
        insertAlphaEdit(edits, state, opacity, offset, source);
      }
    }
  }

  return applyLocalEdits(source, edits);
}

function patchNumericAttribute(
  edits: LocalEdit[],
  state: ExistingShadowState,
  name: 'blurRad' | 'dist' | 'dir',
  current: number,
  target: number,
  rawTarget: number,
  offset: number,
  missingAttributes: string[],
): void {
  if (current === target) return;
  patchAttribute(
    edits,
    state.attributes.get(name),
    String(rawTarget),
    offset,
    missingAttributes,
    name,
  );
}

function patchAttribute(
  edits: LocalEdit[],
  attribute: XmlAttribute | undefined,
  value: string,
  offset: number,
  missingAttributes: string[],
  name: string,
): void {
  if (attribute) {
    edits.push({
      start: attribute.valueStart - offset,
      end: attribute.valueEnd - offset,
      replacement: value,
    });
  } else {
    missingAttributes.push(` ${name}="${value}"`);
  }
}

function insertAlphaEdit(
  edits: LocalEdit[],
  state: ExistingShadowState,
  opacity: string,
  offset: number,
  source: string,
): void {
  const prefix = qualifiedPrefix(state.prefix);
  const alpha = `<${prefix}alpha val="${opacity}"/>`;
  if (state.color.selfClosing) {
    const marker = state.color.end - offset - 2;
    if (source.slice(marker, marker + 2) !== '/>') {
      throw new Error('Invalid self-closing shape shadow color span');
    }
    edits.push({
      start: marker,
      end: state.color.end - offset,
      replacement: `>${alpha}</${state.color.name}>`,
    });
    return;
  }
  edits.push({
    start: state.color.endTagStart - offset,
    end: state.color.endTagStart - offset,
    replacement: alpha,
  });
}

function renderEffectListForParent(
  shadow: NormalizedShapeShadow,
  prefixSource: XmlElement,
  parent: XmlElement,
): string {
  const prefix = lexicalPrefix(prefixSource.name);
  const qualified = qualifiedPrefix(prefix);
  const declaration = namespaceDeclaration(parent, prefix);
  return `<${qualified}effectLst${declaration}>` +
    `${renderSimpleShadow(shadow, qualified)}</${qualified}effectLst>`;
}

function renderShadowForParent(
  shadow: NormalizedShapeShadow,
  prefix: string,
  parent: XmlElement,
): string {
  const qualified = qualifiedPrefix(prefix);
  const rendered = renderSimpleShadow(shadow, qualified);
  const declaration = namespaceDeclaration(parent, prefix);
  if (declaration === '') return rendered;
  const tag = shadow.kind === 'outer' ? 'outerShdw' : 'innerShdw';
  return rendered.replace(`<${qualified}${tag} `, `<${qualified}${tag}${declaration} `);
}

function renderColorForParent(
  shadow: NormalizedShapeShadow,
  prefix: string,
  parent: XmlElement,
): string {
  const qualified = qualifiedPrefix(prefix);
  const tag = shadow.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const declaration = namespaceDeclaration(parent, prefix);
  return `<${qualified}${tag}${declaration} val="${escapeXmlAttribute(shadow.color.value)}">` +
    `<${qualified}alpha val="${Math.round(shadow.opacity * OPACITY_SCALE)}"/>` +
    `</${qualified}${tag}>`;
}

function namespaceDeclaration(parent: XmlElement, prefix: string): string {
  if (namespaceUriForPrefix(parent, prefix) === DRAWING_NAMESPACE) return '';
  return prefix === ''
    ? ` xmlns="${DRAWING_NAMESPACE}"`
    : ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`;
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
      throw new Error('Overlapping local shape shadow edits');
    }
    output = output.slice(0, edit.start) + edit.replacement + output.slice(edit.end);
    previousStart = edit.start;
  }
  return output;
}

function hasNonWhitespaceText(element: XmlElement): boolean {
  return element.children.some(
    (child) => child.type === 'text' && /\S/u.test(child.value),
  );
}

function qualifiedPrefix(prefix: string): string {
  return prefix === '' ? '' : `${prefix}:`;
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

function nonNamespaceAttributes(element: XmlElement): XmlAttribute[] {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
