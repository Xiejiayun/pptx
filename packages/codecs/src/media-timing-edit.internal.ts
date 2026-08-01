import type { LosslessXmlDocument, XmlElement } from '@pptx/lossless-xml';
import type { NormalizedMediaPlaybackSettings } from './media-edit.internal.js';
import {
  allocateNativeTimingIds,
  readNativeMediaTiming,
  type NativeMediaTimingOwnership,
  type NativeMediaTimingSettings,
} from './media-timing-state.internal.js';

type MediaKind = 'audio' | 'video';

export interface NativeMediaTimingSyncResult {
  readonly changed: boolean;
  readonly ownership: Readonly<NativeMediaTimingOwnership>;
}

interface TimingContext {
  readonly timing: XmlElement;
  readonly rootChildren: XmlElement;
}

interface LocatedGraph extends TimingContext {
  readonly mediaBranch: XmlElement;
  readonly playBranch: XmlElement;
  readonly pauseBranch?: XmlElement;
}

export function syncNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: MediaKind,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
  recorded?: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingSyncResult> {
  requireShapeId(shapeId);
  const current = readNativeMediaTiming(xml, shapeId, kind, recorded);
  if (current.status === 'unsupported' || current.status === 'ambiguous') {
    throw timingError(current.status, current.reason);
  }
  if (current.settings && current.ownership) {
    if (settingsEqual(current.settings, settings)) {
      return syncResult(false, current.ownership);
    }
    return replaceRecognizedGraph(xml, shapeId, kind, settings, current.ownership);
  }
  return createGraph(xml, shapeId, kind, settings);
}

export function clearNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: MediaKind,
  recorded?: Readonly<NativeMediaTimingOwnership>,
): boolean {
  requireShapeId(shapeId);
  const current = readNativeMediaTiming(xml, shapeId, kind, recorded);
  if (current.status === 'unsupported' || current.status === 'ambiguous') {
    throw timingError(current.status, current.reason);
  }
  if (!current.settings || !current.ownership) return false;

  const graph = locateGraph(xml, current.ownership);
  if (isExactOwnedTiming(xml, graph.timing, kind, shapeId, current.settings)) {
    xml.removeElement(graph.timing);
    return true;
  }
  xml.removeElement(graph.mediaBranch);
  xml.removeElement(graph.playBranch);
  if (graph.pauseBranch) xml.removeElement(graph.pauseBranch);
  return true;
}

function createGraph(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: MediaKind,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
): Readonly<NativeMediaTimingSyncResult> {
  const context = timingContext(xml);
  if (!context) {
    const ids = allocateNativeTimingIds(xml, kind === 'video' ? 12 : 7);
    const timing = renderTimingTree(kind, shapeId, settings, ids);
    const slide = requireSlide(xml);
    insertTiming(xml, slide, timing);
    return syncResult(true, ownership(ids[6]!, ids[4]!, kind === 'video' ? ids[10] : undefined));
  }

  const main = findMainSequence(xml, context.rootChildren);
  if (main) {
    const ids = allocateNativeTimingIds(xml, kind === 'video' ? 10 : 5);
    const playIds = ids.slice(0, 4);
    const mediaId = ids[4]!;
    const pauseIds = ids.slice(5, 10);
    xml.appendChildXml(main.list, renderPlayBranch(shapeId, settings.play, playIds));
    xml.appendChildXml(
      context.rootChildren,
      renderMediaNode(kind, shapeId, mediaId, settings)
        + (kind === 'video' ? renderPauseBranch(shapeId, pauseIds) : ''),
    );
    return syncResult(
      true,
      ownership(mediaId, playIds[2]!, kind === 'video' ? pauseIds[3] : undefined),
    );
  }

  const ids = allocateNativeTimingIds(xml, kind === 'video' ? 11 : 6);
  const mainId = ids[0]!;
  const playIds = ids.slice(1, 5);
  const mediaId = ids[5]!;
  const pauseIds = ids.slice(6, 11);
  xml.appendChildXml(
    context.rootChildren,
    renderMainSequence(shapeId, settings.play, mainId, playIds)
      + renderMediaNode(kind, shapeId, mediaId, settings)
      + (kind === 'video' ? renderPauseBranch(shapeId, pauseIds) : ''),
  );
  return syncResult(
    true,
    ownership(mediaId, playIds[2]!, kind === 'video' ? pauseIds[3] : undefined),
  );
}

function replaceRecognizedGraph(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: MediaKind,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
  current: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingSyncResult> {
  const graph = locateGraph(xml, current);
  const ids = allocateNativeTimingIds(xml, kind === 'video' ? 10 : 5);
  const playIds = ids.slice(0, 4);
  const mediaId = ids[4]!;
  const pauseIds = ids.slice(5, 10);
  xml.replaceElement(graph.playBranch, renderPlayBranch(shapeId, settings.play, playIds));
  xml.replaceElement(graph.mediaBranch, renderMediaNode(kind, shapeId, mediaId, settings));
  if (kind === 'video') {
    const pause = renderPauseBranch(shapeId, pauseIds);
    if (graph.pauseBranch) xml.replaceElement(graph.pauseBranch, pause);
    else xml.appendChildXml(graph.rootChildren, pause);
  } else if (graph.pauseBranch) {
    xml.removeElement(graph.pauseBranch);
  }
  return syncResult(
    true,
    ownership(mediaId, playIds[2]!, kind === 'video' ? pauseIds[3] : undefined),
  );
}

function timingContext(xml: LosslessXmlDocument): TimingContext | undefined {
  const slide = requireSlide(xml);
  const timings = directChildren(slide, 'timing');
  if (timings.length === 0) return undefined;
  if (timings.length !== 1) throw new Error('Native media timing is ambiguous');
  const timing = timings[0]!;
  const lists = directChildren(timing, 'tnLst');
  const roots = lists.length === 1 ? directChildren(lists[0]!, 'par') : [];
  const rootNodes = roots.length === 1 ? directChildren(roots[0]!, 'cTn') : [];
  const childLists = rootNodes.length === 1 ? directChildren(rootNodes[0]!, 'childTnLst') : [];
  if (lists.length !== 1 || roots.length !== 1 || rootNodes.length !== 1 || childLists.length !== 1) {
    throw new Error('Native media timing root is not canonical');
  }
  return { timing, rootChildren: childLists[0]! };
}

function locateGraph(
  xml: LosslessXmlDocument,
  ids: Readonly<NativeMediaTimingOwnership>,
): LocatedGraph {
  const context = timingContext(xml);
  if (!context) throw new Error('Owned native media timing graph is missing');
  const mediaTime = timingNode(xml, context.timing, ids.mediaTnId);
  const mediaNode = mediaTime.parent;
  const mediaBranch = mediaNode?.parent;
  if (
    mediaNode?.localName !== 'cMediaNode'
    || (mediaBranch?.localName !== 'audio' && mediaBranch?.localName !== 'video')
    || mediaBranch.parent !== context.rootChildren
  ) throw new Error('Owned native media timing state branch is missing');

  const playEffect = timingNode(xml, context.timing, ids.playTnId);
  const main = ancestorWithAttribute(playEffect, 'cTn', 'nodeType', 'mainSeq');
  const playList = main ? directChildren(main, 'childTnLst')[0] : undefined;
  const playBranch = playList
    ? directElementChildren(playList).find((candidate) => contains(candidate, playEffect))
    : undefined;
  if (!playList || !playBranch) throw new Error('Owned native media play branch is missing');

  let pauseBranch: XmlElement | undefined;
  if (ids.pauseTnId !== undefined) {
    const pauseEffect = timingNode(xml, context.timing, ids.pauseTnId);
    pauseBranch = directElementChildren(context.rootChildren)
      .find((candidate) => contains(candidate, pauseEffect));
    if (!pauseBranch) throw new Error('Owned native media pause branch is missing');
  }
  return { ...context, mediaBranch, playBranch, ...(pauseBranch ? { pauseBranch } : {}) };
}

function findMainSequence(
  xml: LosslessXmlDocument,
  rootChildren: XmlElement,
): { readonly list: XmlElement } | undefined {
  const nodes = xml.descendants(rootChildren, 'cTn').filter(
    (node) => singleAttributeValue(node, 'nodeType') === 'mainSeq',
  );
  if (nodes.length > 1) throw new Error('Native timing contains multiple main sequences');
  if (nodes.length === 0) return undefined;
  const node = nodes[0]!;
  const container = node.parent;
  const lists = directChildren(node, 'childTnLst');
  if (!container || container.parent !== rootChildren || lists.length !== 1) {
    throw new Error('Native timing main sequence is not canonical');
  }
  return { list: lists[0]! };
}

function insertTiming(xml: LosslessXmlDocument, slide: XmlElement, timing: string): void {
  const extensions = directChildren(slide, 'extLst');
  if (extensions.length > 1) throw new Error('Slide contains repeated extension lists');
  xml.replace(extensions[0]?.start ?? slide.endTagStart, extensions[0]?.start ?? slide.endTagStart, timing);
}

function renderTimingTree(
  kind: MediaKind,
  shapeId: number,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
  ids: readonly number[],
): string {
  if (ids.length !== (kind === 'video' ? 12 : 7)) {
    throw new Error('Native media timing ID allocation is incomplete');
  }
  return '<p:timing><p:tnLst><p:par>'
    + `<p:cTn id="${ids[0]}" dur="indefinite" restart="never" nodeType="tmRoot">`
    + '<p:childTnLst>'
    + renderMainSequence(shapeId, settings.play, ids[1]!, ids.slice(2, 6))
    + renderMediaNode(kind, shapeId, ids[6]!, settings)
    + (kind === 'video' ? renderPauseBranch(shapeId, ids.slice(7, 12)) : '')
    + '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
}

function renderMainSequence(
  shapeId: number,
  play: 'click' | 'auto',
  mainId: number,
  playIds: readonly number[],
): string {
  return '<p:seq concurrent="1" nextAc="seek">'
    + `<p:cTn id="${mainId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>`
    + renderPlayBranch(shapeId, play, playIds)
    + '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0">'
    + '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst>'
    + '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
    + '</p:cond></p:nextCondLst></p:seq>';
}

function renderPlayBranch(
  shapeId: number,
  play: 'click' | 'auto',
  ids: readonly number[],
): string {
  if (ids.length !== 4) throw new Error('Native media play ID allocation is incomplete');
  return `<p:par><p:cTn id="${ids[0]}" fill="hold"><p:stCondLst>`
    + `<p:cond delay="${play === 'click' ? 'indefinite' : '0'}"/></p:stCondLst>`
    + `<p:childTnLst><p:par><p:cTn id="${ids[1]}" fill="hold"><p:stCondLst>`
    + '<p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>'
    + `<p:cTn id="${ids[2]}" presetID="1" presetClass="mediacall" `
    + 'presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst>'
    + '<p:cond delay="0"/></p:stCondLst><p:childTnLst>'
    + '<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>'
    + `<p:cTn id="${ids[3]}" dur="3000" fill="hold"/>`
    + `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
    + '</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst>'
    + '</p:cTn></p:par></p:childTnLst></p:cTn></p:par>';
}

function renderMediaNode(
  kind: MediaKind,
  shapeId: number,
  id: number,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
): string {
  const repeat = settings.loop ? ' repeatCount="indefinite"' : '';
  const shown = settings.hideWhenStopped ? 0 : 1;
  const end = kind === 'audio'
    ? '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl>'
      + '<p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>'
    : '';
  return `<p:${kind}><p:cMediaNode vol="${canonicalVolume(settings.volume)}" `
    + `showWhenStopped="${shown}"><p:cTn id="${id}" fill="hold" display="0"${repeat}>`
    + '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' + end
    + `</p:cTn><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
    + `</p:cMediaNode></p:${kind}>`;
}

function renderPauseBranch(shapeId: number, ids: readonly number[]): string {
  if (ids.length !== 5) throw new Error('Native media pause ID allocation is incomplete');
  return '<p:seq concurrent="1" nextAc="seek">'
    + `<p:cTn id="${ids[0]}" restart="whenNotActive" fill="hold" `
    + 'evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst>'
    + `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${shapeId}"/>`
    + '</p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0">'
    + `<p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="${ids[1]}" fill="hold">`
    + '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>'
    + `<p:cTn id="${ids[2]}" fill="hold"><p:stCondLst><p:cond delay="0"/>`
    + `</p:stCondLst><p:childTnLst><p:par><p:cTn id="${ids[3]}" presetID="2" `
    + 'presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect">'
    + '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
    + '<p:cmd type="call" cmd="togglePause"><p:cBhvr>'
    + `<p:cTn id="${ids[4]}" dur="1" fill="hold"/>`
    + `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
    + '</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst>'
    + '</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
    + `<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl>`
    + `<p:spTgt spid="${shapeId}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
}

function isExactOwnedTiming(
  xml: LosslessXmlDocument,
  timing: XmlElement,
  kind: MediaKind,
  shapeId: number,
  settings: Readonly<NativeMediaTimingSettings>,
): boolean {
  const ids = xml.descendants(timing, 'cTn').map((node) => Number(singleAttributeValue(node, 'id')));
  return ids.length === (kind === 'video' ? 12 : 7)
    && ids.every(Number.isSafeInteger)
    && xml.original(timing) === renderTimingTree(kind, shapeId, settings, ids);
}

function timingNode(
  xml: LosslessXmlDocument,
  timing: XmlElement,
  id: number,
): XmlElement {
  const nodes = xml.descendants(timing, 'cTn').filter(
    (node) => singleAttributeValue(node, 'id') === String(id),
  );
  if (nodes.length !== 1) throw new Error(`Native timing node ${id} is missing or repeated`);
  return nodes[0]!;
}

function requireSlide(xml: LosslessXmlDocument): XmlElement {
  const slides = xml.roots.filter(({ localName }) => localName === 'sld');
  if (slides.length !== 1) throw new Error('Slide XML must contain exactly one slide root');
  return slides[0]!;
}

function requireShapeId(shapeId: number): void {
  if (!Number.isSafeInteger(shapeId) || shapeId <= 0 || shapeId > 0xFFFF_FFFF) {
    throw new RangeError('Media shape ID must be a positive unsigned integer');
  }
}

function directChildren(parent: XmlElement, localName: string): XmlElement[] {
  return parent.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function directElementChildren(parent: XmlElement): XmlElement[] {
  return parent.children.filter((child): child is XmlElement => child.type === 'element');
}

function singleAttributeValue(element: XmlElement, localName: string): string | undefined {
  const values = element.attributes.filter((attribute) => attribute.localName === localName);
  return values.length === 1 ? values[0]!.value : undefined;
}

function ancestorWithAttribute(
  element: XmlElement,
  localName: string,
  attributeName: string,
  value: string,
): XmlElement | undefined {
  let current = element.parent;
  while (current) {
    if (
      current.localName === localName
      && singleAttributeValue(current, attributeName) === value
    ) return current;
    current = current.parent;
  }
  return undefined;
}

function contains(parent: XmlElement, child: XmlElement): boolean {
  return parent.start <= child.start && parent.end >= child.end;
}

function settingsEqual(
  left: Readonly<NativeMediaTimingSettings>,
  right: Readonly<NormalizedMediaPlaybackSettings>,
): boolean {
  return left.play === right.play
    && left.loop === right.loop
    && left.hideWhenStopped === right.hideWhenStopped
    && left.volume === canonicalVolume(right.volume) / 100_000;
}

function canonicalVolume(value: number): number {
  return Math.round(value * 100_000);
}

function ownership(
  mediaTnId: number,
  playTnId: number,
  pauseTnId?: number,
): Readonly<NativeMediaTimingOwnership> {
  return Object.freeze({
    version: 1,
    mediaTnId,
    playTnId,
    ...(pauseTnId === undefined ? {} : { pauseTnId }),
  });
}

function syncResult(
  changed: boolean,
  nativeOwnership: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingSyncResult> {
  return Object.freeze({ changed, ownership: nativeOwnership });
}

function timingError(status: string, reason: string | undefined): Error {
  return new Error(`Native media timing is ${status}${reason ? `: ${reason}` : ''}`);
}
