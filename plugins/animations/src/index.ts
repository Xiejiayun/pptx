import {
  allocateNativeTimingIds,
  MediaCodec,
  type CodecDiagnostic,
  type CodecRegistry,
} from '@pptx/codecs';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';

export type TimingNodeKind = 'parallel' | 'sequence' | 'effect' | 'motion' | 'media' | 'unknown';
export type AnimationEffect = 'appear' | 'fade' | 'wipe' | 'fly' | string;

export interface TimingNode {
  readonly id?: number;
  readonly kind: TimingNodeKind;
  readonly durationMs?: number | 'indefinite';
  readonly delayMs?: number;
  readonly repeat?: number | 'indefinite';
  readonly targetShapeId?: number;
  readonly textRange?: { readonly start: number; readonly end: number };
  readonly effect?: string;
  readonly motionPath?: string;
  readonly children: readonly TimingNode[];
}

export interface AnimationSpec {
  readonly effect: AnimationEffect;
  readonly targetShapeId: number;
  readonly durationMs?: number;
  readonly delayMs?: number;
  readonly trigger?: 'after-previous' | 'with-previous' | 'on-click';
  readonly repeat?: number | 'indefinite';
  readonly textRange?: { readonly start: number; readonly end: number };
  readonly motionPath?: string;
}

export interface PluginHost {
  readonly opcPackage: OpcPackage;
  readonly codecRegistry: CodecRegistry;
}

export class AnimationTimingCodec {
  readonly id = 'plugin.animations-timing';
  readonly priority = 200;
  readonly ownership = { elements: ['p:timing'] } as const;

  constructor(readonly pkg: OpcPackage) {}

  tree(slidePartUri: string): TimingNode | undefined {
    const xml = this.parse(slidePartUri);
    const timing = xml.elements('timing')[0];
    if (!timing) return undefined;
    const root = timing.children.find((child): child is XmlElement => child.type === 'element');
    return root ? decodeNode(xml, root) : { kind: 'parallel', children: [] };
  }

  add(slidePartUri: string, spec: AnimationSpec): number {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    if (!shapeIds(xml).has(spec.targetShapeId)) throw new Error(`Animation target shape ${spec.targetShapeId} does not exist`);
    const timing = ensureTimingRoot(xml);
    const id = nextTimingId(timing.document);
    const encoded = encodeAnimation(spec, id);
    timing.document.appendChildXml(timing.list, encoded);
    this.pkg.setPart(slidePartUri, timing.document.serialize(), part.contentType);
    return id;
  }

  remove(slidePartUri: string, timingId: number): boolean {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const cTn = xml
      .elements('cTn')
      .find((element) => Number(xml.attribute(element, 'id')?.value ?? -1) === timingId);
    if (!cTn) return false;
    const container = ancestor(cTn, 'par') ?? ancestor(cTn, 'seq') ?? cTn;
    xml.removeElement(container);
    this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    return true;
  }

  retargetShape(slidePartUri: string, oldShapeId: number, newShapeId: number): number {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    if (!shapeIds(xml).has(newShapeId)) throw new Error(`Replacement shape ${newShapeId} does not exist`);
    let count = 0;
    for (const target of xml.elements('spTgt')) {
      const attribute = xml.attribute(target, 'spid');
      if (Number(attribute?.value) === oldShapeId && attribute) {
        xml.replaceAttribute(attribute, String(newShapeId));
        count += 1;
      }
    }
    if (count > 0) this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    return count;
  }

  validate(slidePartUri: string): CodecDiagnostic[] {
    const xml = this.parse(slidePartUri);
    const ids = shapeIds(xml);
    return xml.elements('spTgt').flatMap((target) => {
      const id = Number(xml.attribute(target, 'spid')?.value ?? -1);
      return ids.has(id)
        ? []
        : [{
            severity: 'error' as const,
            code: 'TIMING_DANGLING_SHAPE_TARGET',
            message: `Timing node targets missing shape ${id}`,
            partUri: slidePartUri,
          }];
    });
  }

  materializeMediaPlayback(slidePartUri: string): number {
    return new MediaCodec(this.pkg).materializePlayback(slidePartUri);
  }

  private parse(partUri: string): LosslessXmlDocument {
    return LosslessXmlDocument.parse(this.pkg.requirePart(partUri).bytes);
  }
}

export function installAnimationPlugin(host: PluginHost): AnimationTimingCodec {
  const codec = new AnimationTimingCodec(host.opcPackage);
  host.codecRegistry.register(codec);
  for (const part of host.opcPackage.parts) {
    if (part.contentType === 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml') {
      codec.materializeMediaPlayback(part.uri);
    }
  }
  return codec;
}

function ensureTimingRoot(xml: LosslessXmlDocument): { document: LosslessXmlDocument; list: XmlElement } {
  let timing = xml.elements('timing')[0];
  let list = timing ? xml.descendants(timing, 'childTnLst')[0] : undefined;
  if (timing && list) return { document: xml, list };
  const root = xml.elements('sld')[0];
  if (!root) throw new Error('Slide root is missing');
  if (!timing) {
    const skeleton = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
    const insertionTarget = root.children.find(
      (child): child is XmlElement => child.type === 'element' && child.localName === 'extLst',
    );
    const insertion = insertionTarget?.start ?? root.endTagStart;
    xml.replace(insertion, insertion, skeleton);
    const interim = LosslessXmlDocument.parse(xml.serialize());
    list = interim.elements('childTnLst')[0]!;
    return { document: interim, list };
  }
  xml.appendChildXml(timing, '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst></p:childTnLst></p:cTn></p:par></p:tnLst>');
  const interim = LosslessXmlDocument.parse(xml.serialize());
  return { document: interim, list: interim.elements('childTnLst')[0]! };
}

function encodeAnimation(spec: AnimationSpec, id: number): string {
  const duration = Math.max(1, Math.round(spec.durationMs ?? 500));
  const delay = Math.max(0, Math.round(spec.delayMs ?? 0));
  const repeat = spec.repeat === 'indefinite' ? ' repeatCount="indefinite"' : spec.repeat ? ` repeatCount="${spec.repeat * 1000}"` : '';
  const trigger = spec.trigger === 'on-click' ? ' evt="onClick"' : '';
  const range = spec.textRange ? `<p:txEl><p:charRg st="${spec.textRange.start}" end="${spec.textRange.end}"/></p:txEl>` : '';
  const target = `<p:tgtEl><p:spTgt spid="${spec.targetShapeId}">${range}</p:spTgt></p:tgtEl>`;
  const behavior = `<p:cBhvr><p:cTn id="${id + 1}" dur="${duration}"${repeat}/>${target}</p:cBhvr>`;
  const effect = spec.motionPath
    ? `<p:animMotion path="${escapeAttribute(spec.motionPath)}">${behavior}</p:animMotion>`
    : `<p:animEffect transition="in" filter="${escapeAttribute(effectFilter(spec.effect))}">${behavior}</p:animEffect>`;
  return `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="${delay}"${trigger}/></p:stCondLst><p:childTnLst>${effect}</p:childTnLst></p:cTn></p:par>`;
}

function decodeNode(xml: LosslessXmlDocument, element: XmlElement): TimingNode {
  const cTn = element.localName === 'cTn' ? element : xml.descendants(element, 'cTn')[0];
  const target = xml.descendants(element, 'spTgt')[0];
  const range = target ? xml.descendants(target, 'charRg')[0] : undefined;
  const effect = xml.descendants(element, 'animEffect')[0];
  const motion = xml.descendants(element, 'animMotion')[0];
  const durationValue = cTn ? xml.attribute(cTn, 'dur')?.value : undefined;
  const repeatValue = cTn ? xml.attribute(cTn, 'repeatCount')?.value : undefined;
  const delay = xml.descendants(element, 'cond')[0];
  const kind: TimingNodeKind =
    element.localName === 'par'
      ? 'parallel'
      : element.localName === 'seq'
        ? 'sequence'
        : motion
          ? 'motion'
          : effect
            ? 'effect'
            : xml.descendants(element, 'cMediaNode').length > 0
              ? 'media'
              : 'unknown';
  const childList = cTn ? xml.descendants(cTn, 'childTnLst')[0] : undefined;
  const children = childList
    ? childList.children.filter((child): child is XmlElement => child.type === 'element').map((child) => decodeNode(xml, child))
    : [];
  return {
    kind,
    ...(cTn && xml.attribute(cTn, 'id') ? { id: Number(xml.attribute(cTn, 'id')!.value) } : {}),
    ...(durationValue ? { durationMs: durationValue === 'indefinite' ? 'indefinite' : Number(durationValue) } : {}),
    ...(delay ? { delayMs: Number(xml.attribute(delay, 'delay')?.value ?? 0) } : {}),
    ...(repeatValue ? { repeat: repeatValue === 'indefinite' ? 'indefinite' : Number(repeatValue) / 1000 } : {}),
    ...(target ? { targetShapeId: Number(xml.attribute(target, 'spid')?.value ?? -1) } : {}),
    ...(range
      ? { textRange: { start: Number(xml.attribute(range, 'st')?.value ?? 0), end: Number(xml.attribute(range, 'end')?.value ?? 0) } }
      : {}),
    ...(effect ? { effect: xml.attribute(effect, 'filter')?.value ?? 'appear' } : {}),
    ...(motion ? { motionPath: xml.attribute(motion, 'path')?.value ?? '' } : {}),
    children,
  };
}

function nextTimingId(xml: LosslessXmlDocument): number {
  return allocateNativeTimingIds(xml, 2)[0]!;
}

function shapeIds(xml: LosslessXmlDocument): Set<number> {
  return new Set(xml.elements('cNvPr').map((element) => Number(xml.attribute(element, 'id')?.value ?? -1)));
}

function ancestor(element: XmlElement, localName: string): XmlElement | undefined {
  let current = element.parent;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parent;
  }
  return undefined;
}

function effectFilter(effect: string): string {
  return { appear: 'appear', fade: 'fade', wipe: 'wipe(right)', fly: 'fly(left)' }[effect] ?? effect;
}

function escapeAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
