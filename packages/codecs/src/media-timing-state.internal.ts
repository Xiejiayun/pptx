import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';

const MAX_TIMING_ID = 0xFFFF_FFFF;

export interface NativeMediaTimingOwnership {
  readonly version: 1;
  readonly mediaTnId: number;
  readonly playTnId: number;
  readonly pauseTnId?: number;
}

export interface NativeMediaTimingSettings {
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
}

export type NativeMediaTimingStatus =
  | 'absent'
  | 'recognized-imported'
  | 'owned-healthy'
  | 'owned-stale'
  | 'unsupported'
  | 'ambiguous';

export interface NativeMediaTimingState {
  readonly status: NativeMediaTimingStatus;
  readonly settings?: Readonly<NativeMediaTimingSettings>;
  readonly ownership?: Readonly<NativeMediaTimingOwnership>;
  readonly reason?: string;
}

type MediaKind = 'audio' | 'video';

interface TimingContext {
  readonly slide: XmlElement;
  readonly timing: XmlElement;
  readonly rootChildren: XmlElement;
}

interface ParsedMediaNode {
  readonly id: number;
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
}

interface ParsedCommand {
  readonly id: number;
  readonly play?: 'click' | 'auto';
}

interface ParseFailure {
  readonly status: 'unsupported' | 'ambiguous';
  readonly reason: string;
}

export function readNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: MediaKind,
  recorded?: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingState> {
  const context = readTimingContext(xml);
  if (!context) {
    return recorded
      ? state('owned-stale', 'Owned native media timing graph is missing')
      : state('absent');
  }
  if ('status' in context) return state(context.status, context.reason);

  const ids = readTimingIds(xml, context.timing);
  if ('status' in ids) return state(ids.status, ids.reason);

  const candidates = directChildren(context.rootChildren)
    .filter(({ localName }) => localName === 'audio' || localName === 'video')
    .filter((candidate) => hasDescendantShapeTarget(xml, candidate, shapeId));
  if (candidates.length === 0) {
    return recorded
      ? state('owned-stale', 'Owned native media timing graph is missing')
      : state('absent');
  }
  if (candidates.length > 1) {
    return state('ambiguous', `Multiple native media timing graphs target shape ${shapeId}`);
  }

  if (!hasUniqueShape(xml, shapeId)) {
    return state('unsupported', `Native media timing targets missing or repeated shape ${shapeId}`);
  }

  const mediaElement = candidates[0]!;
  if (mediaElement.localName !== kind) {
    return state(
      'unsupported',
      `Native ${mediaElement.localName} timing does not match ${kind} shape ${shapeId}`,
    );
  }
  const media = parseMediaNode(xml, mediaElement, shapeId, kind);
  if ('status' in media) return state(media.status, media.reason);

  const playCommands = commandCandidates(xml, context.rootChildren, 'playFrom(0.0)', shapeId);
  if (playCommands.length === 0) {
    return state('unsupported', `Native media timing for shape ${shapeId} has no play command`);
  }
  if (playCommands.length > 1) {
    return state('ambiguous', `Multiple native media play commands target shape ${shapeId}`);
  }
  const play = parsePlayCommand(xml, playCommands[0]!, shapeId);
  if ('status' in play) return state(play.status, play.reason);

  const pauseCommands = commandCandidates(xml, context.rootChildren, 'togglePause', shapeId);
  if (pauseCommands.length > 1) {
    return state('ambiguous', `Multiple native media pause commands target shape ${shapeId}`);
  }
  const pause = pauseCommands[0]
    ? parsePauseCommand(xml, pauseCommands[0], shapeId)
    : undefined;
  if (pause && 'status' in pause) return state(pause.status, pause.reason);

  const settings = Object.freeze({
    play: play.play!,
    loop: media.loop,
    hideWhenStopped: media.hideWhenStopped,
    volume: media.volume,
  });
  const ownership = Object.freeze({
    version: 1 as const,
    mediaTnId: media.id,
    playTnId: play.id,
    ...(pause ? { pauseTnId: pause.id } : {}),
  });
  if (!recorded) {
    return state('recognized-imported', undefined, settings, ownership);
  }
  if (!ownershipEqual(recorded, ownership)) {
    return state(
      'owned-stale',
      `Recorded native media timing ids do not match shape ${shapeId}`,
      settings,
      ownership,
    );
  }
  return state('owned-healthy', undefined, settings, ownership);
}

export function allocateNativeTimingIds(
  xml: LosslessXmlDocument,
  count: number,
): readonly number[] {
  if (!Number.isSafeInteger(count) || count < 0) {
    throw new RangeError('Timing ID count must be a non-negative integer');
  }
  const slide = requireSlideRoot(xml);
  const timings = directChildren(slide, 'timing');
  if (timings.length > 1) throw new Error('Slide contains repeated timing trees');
  const ids = timings[0] ? readTimingIds(xml, timings[0]) : [];
  if ('status' in ids) throw new Error(ids.reason);
  const maximum = ids.length === 0 ? 0 : Math.max(...ids);
  if (maximum + count > MAX_TIMING_ID) {
    throw new RangeError('Native timing ID space is exhausted');
  }
  return Object.freeze(Array.from({ length: count }, (_, index) => maximum + index + 1));
}

function readTimingContext(
  xml: LosslessXmlDocument,
): TimingContext | ParseFailure | undefined {
  const slides = xml.roots.filter(({ localName }) => localName === 'sld');
  if (slides.length !== 1) {
    return {
      status: slides.length > 1 ? 'ambiguous' : 'unsupported',
      reason: 'Slide XML must contain exactly one slide root',
    };
  }
  const slide = slides[0]!;
  const timings = directChildren(slide, 'timing');
  if (timings.length === 0) return undefined;
  if (timings.length > 1) {
    return { status: 'ambiguous', reason: 'Slide contains repeated timing trees' };
  }
  const timing = timings[0]!;
  const lists = directChildren(timing, 'tnLst');
  if (lists.length !== 1) {
    return failureForCount(lists.length, 'native timing node lists');
  }
  const roots = directChildren(lists[0]!, 'par');
  if (roots.length !== 1 || directChildren(lists[0]!).length !== 1) {
    return failureForCount(roots.length, 'native timing roots');
  }
  const rootNodes = directChildren(roots[0]!, 'cTn');
  if (rootNodes.length !== 1 || directChildren(roots[0]!).length !== 1) {
    return failureForCount(rootNodes.length, 'native timing root nodes');
  }
  const root = rootNodes[0]!;
  if (singleAttribute(xml, root, 'nodeType')?.value !== 'tmRoot') {
    return { status: 'unsupported', reason: 'Native timing root is not tmRoot' };
  }
  const childLists = directChildren(root, 'childTnLst');
  if (childLists.length !== 1 || directChildren(root).length !== 1) {
    return failureForCount(childLists.length, 'native timing root child lists');
  }
  return { slide, timing, rootChildren: childLists[0]! };
}

function parseMediaNode(
  xml: LosslessXmlDocument,
  mediaElement: XmlElement,
  shapeId: number,
  kind: MediaKind,
): ParsedMediaNode | ParseFailure {
  const mediaNodes = directChildren(mediaElement, 'cMediaNode');
  if (mediaNodes.length !== 1 || directChildren(mediaElement).length !== 1) {
    return failureForCount(mediaNodes.length, 'common media nodes');
  }
  const mediaNode = mediaNodes[0]!;
  const timeNodes = directChildren(mediaNode, 'cTn');
  const targets = directChildren(mediaNode, 'tgtEl');
  if (timeNodes.length !== 1 || targets.length !== 1 || directChildren(mediaNode).length !== 2) {
    return { status: 'unsupported', reason: 'Common media node structure is not canonical' };
  }
  const targetFailure = requireShapeTarget(xml, targets[0]!, shapeId);
  if (targetFailure) return targetFailure;

  const numSld = attributes(mediaNode, 'numSld');
  if (numSld.length > 0) {
    return { status: 'unsupported', reason: 'Cross-slide media timing is not supported' };
  }
  const volume = percentageAttribute(xml, mediaNode, 'vol', 100_000);
  if (typeof volume !== 'number') return volume;
  const mute = booleanAttribute(xml, mediaNode, 'mute', false);
  if (typeof mute !== 'boolean') return mute;
  const shown = booleanAttribute(xml, mediaNode, 'showWhenStopped', true);
  if (typeof shown !== 'boolean') return shown;

  const timeNode = timeNodes[0]!;
  const id = timingId(xml, timeNode);
  if (typeof id !== 'number') return id;
  if (attributes(timeNode, 'repeatDur').length > 0) {
    return { status: 'unsupported', reason: 'Media repeat duration is not supported' };
  }
  const repeatAttributes = attributes(timeNode, 'repeatCount');
  if (repeatAttributes.length > 1) {
    return { status: 'ambiguous', reason: 'Media time node has repeated repeat count' };
  }
  const repeat = repeatAttributes[0]?.value;
  if (repeat !== undefined && repeat !== 'indefinite') {
    return { status: 'unsupported', reason: 'Finite media repeat count is not supported' };
  }
  const start = requireSingleCondition(xml, timeNode, 'stCondLst', {
    delay: 'indefinite',
  });
  if (start) return start;
  const end = kind === 'audio'
    ? requireAudioEndCondition(xml, timeNode)
    : rejectDirectChild(timeNode, 'endCondLst', 'Video media end condition is not supported');
  if (end) return end;

  const allowed = new Set(kind === 'audio'
    ? ['stCondLst', 'endCondLst']
    : ['stCondLst']);
  if (directChildren(timeNode).some(({ localName }) => !allowed.has(localName))) {
    return { status: 'unsupported', reason: 'Media time node contains unsupported children' };
  }
  return {
    id,
    loop: repeat === 'indefinite',
    hideWhenStopped: !shown,
    volume: mute ? 0 : volume / 100_000,
  };
}

function parsePlayCommand(
  xml: LosslessXmlDocument,
  command: XmlElement,
  shapeId: number,
): ParsedCommand | ParseFailure {
  const parsed = parseMediaCommand(xml, command, shapeId, 'playFrom(0.0)');
  if ('status' in parsed) return parsed;
  const effect = ancestorWithAttribute(xml, command, 'cTn', 'presetClass', 'mediacall');
  const main = ancestorWithAttribute(xml, effect, 'cTn', 'nodeType', 'mainSeq');
  if (!effect || !main) {
    return { status: 'unsupported', reason: 'Media play command is outside the main sequence' };
  }
  const lists = directChildren(main, 'childTnLst');
  if (lists.length !== 1) {
    return failureForCount(lists.length, 'main-sequence child lists');
  }
  const containers = directChildren(lists[0]!).filter((child) => contains(child, effect));
  if (containers.length !== 1) {
    return failureForCount(containers.length, 'media play sequence containers');
  }
  const topNodes = directChildren(containers[0]!, 'cTn');
  if (topNodes.length !== 1) {
    return failureForCount(topNodes.length, 'media play sequence nodes');
  }
  const conditionLists = directChildren(topNodes[0]!, 'stCondLst');
  if (conditionLists.length !== 1) {
    return failureForCount(conditionLists.length, 'media play start-condition lists');
  }
  const conditions = directChildren(conditionLists[0]!, 'cond');
  if (conditions.length !== 1) {
    return failureForCount(conditions.length, 'media play start conditions');
  }
  const delay = singleAttribute(xml, conditions[0]!, 'delay');
  if (!delay || (delay.value !== '0' && delay.value !== 'indefinite')) {
    return { status: 'unsupported', reason: 'Media play start condition is not supported' };
  }
  return { id: parsed.id, play: delay.value === '0' ? 'auto' : 'click' };
}

function parsePauseCommand(
  xml: LosslessXmlDocument,
  command: XmlElement,
  shapeId: number,
): ParsedCommand | ParseFailure {
  const parsed = parseMediaCommand(xml, command, shapeId, 'togglePause');
  if ('status' in parsed) return parsed;
  const sequence = ancestorWithAttribute(xml, command, 'cTn', 'nodeType', 'interactiveSeq');
  if (!sequence) {
    return { status: 'unsupported', reason: 'Media pause command is outside an interactive sequence' };
  }
  const conditionLists = directChildren(sequence, 'stCondLst');
  if (conditionLists.length !== 1) {
    return failureForCount(conditionLists.length, 'interactive start-condition lists');
  }
  const conditions = directChildren(conditionLists[0]!, 'cond');
  if (conditions.length !== 1 || directChildren(conditionLists[0]!).length !== 1) {
    return failureForCount(conditions.length, 'interactive start conditions');
  }
  const condition = conditions[0]!;
  if (
    singleAttribute(xml, condition, 'evt')?.value !== 'onClick'
    || singleAttribute(xml, condition, 'delay')?.value !== '0'
  ) {
    return { status: 'unsupported', reason: 'Media pause trigger is not a shape click' };
  }
  const targets = directChildren(condition, 'tgtEl');
  if (targets.length !== 1 || directChildren(condition).length !== 1) {
    return failureForCount(targets.length, 'interactive trigger targets');
  }
  const targetFailure = requireShapeTarget(xml, targets[0]!, shapeId);
  return targetFailure ?? parsed;
}

function parseMediaCommand(
  xml: LosslessXmlDocument,
  command: XmlElement,
  shapeId: number,
  name: 'playFrom(0.0)' | 'togglePause',
): ParsedCommand | ParseFailure {
  if (
    singleAttribute(xml, command, 'type')?.value !== 'call'
    || singleAttribute(xml, command, 'cmd')?.value !== name
  ) {
    return { status: 'unsupported', reason: `Media ${name} command attributes are invalid` };
  }
  const behaviors = directChildren(command, 'cBhvr');
  if (behaviors.length !== 1 || directChildren(command).length !== 1) {
    return failureForCount(behaviors.length, `media ${name} behaviors`);
  }
  const behavior = behaviors[0]!;
  const behaviorTimes = directChildren(behavior, 'cTn');
  const behaviorTargets = directChildren(behavior, 'tgtEl');
  if (
    behaviorTimes.length !== 1
    || behaviorTargets.length !== 1
    || directChildren(behavior).length !== 2
  ) {
    return { status: 'unsupported', reason: `Media ${name} behavior is not canonical` };
  }
  const behaviorId = timingId(xml, behaviorTimes[0]!);
  if (typeof behaviorId !== 'number') return behaviorId;
  if (directChildren(behaviorTimes[0]!).length > 0) {
    return { status: 'unsupported', reason: `Media ${name} behavior time node has children` };
  }
  const targetFailure = requireShapeTarget(xml, behaviorTargets[0]!, shapeId);
  if (targetFailure) return targetFailure;
  const effect = ancestorWithAttribute(xml, command, 'cTn', 'presetClass', 'mediacall');
  if (!effect || singleAttribute(xml, effect, 'nodeType')?.value !== 'clickEffect') {
    return { status: 'unsupported', reason: `Media ${name} effect node is not canonical` };
  }
  const id = timingId(xml, effect);
  if (typeof id !== 'number') return id;
  const start = requireSingleCondition(xml, effect, 'stCondLst', { delay: '0' });
  if (start) return start;
  const childLists = directChildren(effect, 'childTnLst');
  if (
    childLists.length !== 1
    || directChildren(childLists[0]!).length !== 1
    || command.parent !== childLists[0]
    || directChildren(effect).some(
      ({ localName }) => localName !== 'stCondLst' && localName !== 'childTnLst',
    )
  ) {
    return { status: 'unsupported', reason: `Media ${name} effect contains unsupported children` };
  }
  return { id };
}

function readTimingIds(
  xml: LosslessXmlDocument,
  timing: XmlElement,
): readonly number[] | ParseFailure {
  const result: number[] = [];
  for (const node of xml.descendants(timing, 'cTn')) {
    const id = timingId(xml, node);
    if (typeof id !== 'number') return id;
    result.push(id);
  }
  if (new Set(result).size !== result.length) {
    return { status: 'ambiguous', reason: 'Timing tree contains duplicate time-node IDs' };
  }
  return result;
}

function timingId(
  xml: LosslessXmlDocument,
  node: XmlElement,
): number | ParseFailure {
  const ids = attributes(node, 'id');
  if (ids.length !== 1 || !/^[1-9]\d*$/.test(ids[0]!.value)) {
    return { status: 'unsupported', reason: 'Timing tree contains an invalid time-node ID' };
  }
  const value = Number(ids[0]!.value);
  if (!Number.isSafeInteger(value) || value > MAX_TIMING_ID) {
    return { status: 'unsupported', reason: 'Timing tree contains an invalid time-node ID' };
  }
  return value;
}

function percentageAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  fallback: number,
): number | ParseFailure {
  const candidates = attributes(element, name);
  if (candidates.length === 0) return fallback;
  if (candidates.length > 1 || !/^\d+$/.test(candidates[0]!.value)) {
    return { status: 'ambiguous', reason: `Media ${name} attribute is invalid or repeated` };
  }
  const value = Number(candidates[0]!.value);
  return Number.isSafeInteger(value) && value >= 0 && value <= 100_000
    ? value
    : { status: 'unsupported', reason: `Media ${name} attribute is outside 0..100000` };
}

function booleanAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  fallback: boolean,
): boolean | ParseFailure {
  const candidates = attributes(element, name);
  if (candidates.length === 0) return fallback;
  if (candidates.length > 1) {
    return { status: 'ambiguous', reason: `Media ${name} attribute is repeated` };
  }
  const value = candidates[0]!.value;
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return { status: 'unsupported', reason: `Media ${name} attribute is invalid` };
}

function requireAudioEndCondition(
  xml: LosslessXmlDocument,
  timeNode: XmlElement,
): ParseFailure | undefined {
  const lists = directChildren(timeNode, 'endCondLst');
  if (lists.length !== 1) return failureForCount(lists.length, 'audio end-condition lists');
  const conditions = directChildren(lists[0]!, 'cond');
  if (conditions.length !== 1) return failureForCount(conditions.length, 'audio end conditions');
  const condition = conditions[0]!;
  if (
    singleAttribute(xml, condition, 'evt')?.value !== 'onStopAudio'
    || singleAttribute(xml, condition, 'delay')?.value !== '0'
  ) {
    return { status: 'unsupported', reason: 'Audio end condition is not onStopAudio' };
  }
  const targets = directChildren(condition, 'tgtEl');
  if (
    targets.length !== 1
    || directChildren(condition).length !== 1
    || directChildren(targets[0]!).length !== 1
    || directChildren(targets[0]!, 'sldTgt').length !== 1
  ) {
    return { status: 'unsupported', reason: 'Audio end condition has an invalid slide target' };
  }
  return undefined;
}

function requireSingleCondition(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  listName: string,
  expected: Readonly<Record<string, string>>,
): ParseFailure | undefined {
  const lists = directChildren(parent, listName);
  if (lists.length !== 1) return failureForCount(lists.length, `${listName} elements`);
  const conditions = directChildren(lists[0]!, 'cond');
  if (conditions.length !== 1 || directChildren(lists[0]!).length !== 1) {
    return failureForCount(conditions.length, `${listName} conditions`);
  }
  const condition = conditions[0]!;
  if (directChildren(condition).length > 0) {
    return { status: 'unsupported', reason: `${listName} condition contains unsupported children` };
  }
  for (const [name, value] of Object.entries(expected)) {
    if (singleAttribute(xml, condition, name)?.value !== value) {
      return { status: 'unsupported', reason: `${listName} condition has unsupported ${name}` };
    }
  }
  return undefined;
}

function requireShapeTarget(
  xml: LosslessXmlDocument,
  target: XmlElement,
  shapeId: number,
): ParseFailure | undefined {
  const children = directChildren(target);
  const shapes = directChildren(target, 'spTgt');
  if (children.length !== 1 || shapes.length !== 1) {
    return { status: 'unsupported', reason: 'Media timing target is not one shape target' };
  }
  const spid = attributes(shapes[0]!, 'spid');
  if (spid.length !== 1 || spid[0]!.value !== String(shapeId)) {
    return { status: 'unsupported', reason: `Media timing target does not match shape ${shapeId}` };
  }
  if (directChildren(shapes[0]!).length > 0) {
    return { status: 'unsupported', reason: 'Media shape target contains unsupported children' };
  }
  return undefined;
}

function commandCandidates(
  xml: LosslessXmlDocument,
  rootChildren: XmlElement,
  name: string,
  shapeId: number,
): readonly XmlElement[] {
  return xml.descendants(rootChildren, 'cmd').filter((command) =>
    singleAttribute(xml, command, 'cmd')?.value === name
    && hasDescendantShapeTarget(xml, command, shapeId));
}

function hasDescendantShapeTarget(
  xml: LosslessXmlDocument,
  element: XmlElement,
  shapeId: number,
): boolean {
  return xml.descendants(element, 'spTgt').some(
    (target) => singleAttribute(xml, target, 'spid')?.value === String(shapeId),
  );
}

function hasUniqueShape(xml: LosslessXmlDocument, shapeId: number): boolean {
  return xml.elements('cNvPr').filter(
    (properties) => singleAttribute(xml, properties, 'id')?.value === String(shapeId),
  ).length === 1;
}

function ancestorWithAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement | undefined,
  localName: string,
  attributeName: string,
  value: string,
): XmlElement | undefined {
  let current = element?.parent;
  while (current) {
    if (
      current.localName === localName
      && singleAttribute(xml, current, attributeName)?.value === value
    ) return current;
    current = current.parent;
  }
  return undefined;
}

function rejectDirectChild(
  parent: XmlElement,
  localName: string,
  reason: string,
): ParseFailure | undefined {
  return directChildren(parent, localName).length > 0
    ? { status: 'unsupported', reason }
    : undefined;
}

function requireSlideRoot(xml: LosslessXmlDocument): XmlElement {
  const slides = xml.roots.filter(({ localName }) => localName === 'sld');
  if (slides.length !== 1) throw new Error('Slide XML must contain exactly one slide root');
  return slides[0]!;
}

function directChildren(parent: XmlElement, localName?: string): readonly XmlElement[] {
  return parent.children.filter(
    (child): child is XmlElement => child.type === 'element'
      && (localName === undefined || child.localName === localName),
  );
}

function attributes(element: XmlElement, localName: string): readonly XmlAttribute[] {
  return element.attributes.filter((attribute) => attribute.localName === localName);
}

function singleAttribute(
  _xml: LosslessXmlDocument,
  element: XmlElement,
  localName: string,
): XmlAttribute | undefined {
  const candidates = attributes(element, localName);
  return candidates.length === 1 ? candidates[0] : undefined;
}

function contains(parent: XmlElement, child: XmlElement): boolean {
  return parent.start <= child.start && parent.end >= child.end;
}

function failureForCount(count: number, label: string): ParseFailure {
  return {
    status: count > 1 ? 'ambiguous' : 'unsupported',
    reason: `Expected exactly one ${label}`,
  };
}

function ownershipEqual(
  left: Readonly<NativeMediaTimingOwnership>,
  right: Readonly<NativeMediaTimingOwnership>,
): boolean {
  return left.version === right.version
    && left.mediaTnId === right.mediaTnId
    && left.playTnId === right.playTnId
    && left.pauseTnId === right.pauseTnId;
}

function state(
  status: NativeMediaTimingStatus,
  reason?: string,
  settings?: Readonly<NativeMediaTimingSettings>,
  ownership?: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingState> {
  return Object.freeze({
    status,
    ...(settings ? { settings } : {}),
    ...(ownership ? { ownership } : {}),
    ...(reason ? { reason } : {}),
  });
}
