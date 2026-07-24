import type { CodecRegistry, CodecDiagnostic } from '@pptx/codecs';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';

export type TransitionEffect =
  | 'none'
  | 'cut'
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'random'
  | 'morph'
  | string;

export interface SlideTransition {
  readonly effect: TransitionEffect;
  readonly speed?: 'slow' | 'med' | 'fast';
  readonly durationMs?: number;
  readonly advanceOnClick?: boolean;
  readonly advanceAfterMs?: number;
  readonly direction?: string;
  readonly soundRelationshipId?: string;
}

export interface PluginHost {
  readonly opcPackage: OpcPackage;
  readonly codecRegistry: CodecRegistry;
}

export class TransitionCodec {
  readonly id = 'plugin.transitions';
  readonly priority = 200;
  readonly ownership = { elements: ['p:transition'] } as const;

  constructor(readonly pkg: OpcPackage) {}

  get(slidePartUri: string): SlideTransition | undefined {
    const xml = this.parse(slidePartUri);
    const transition = xml.elements('transition')[0];
    if (!transition) return undefined;
    const effect = transition.children.find(
      (child): child is XmlElement =>
        child.type === 'element' && !['sndAc', 'extLst'].includes(child.localName),
    );
    const sound = xml.descendants(transition, 'snd')[0];
    const direction = effect
      ? xml.attribute(effect, 'dir')?.value ?? xml.attribute(effect, 'orient')?.value
      : undefined;
    const speed = normalizeSpeed(xml.attribute(transition, 'spd')?.value);
    const durationMs = numberValue(xml.attribute(transition, 'p14:dur')?.value ?? xml.attribute(transition, 'dur')?.value);
    const advanceAfterMs = numberValue(xml.attribute(transition, 'advTm')?.value);
    const soundRelationshipId = sound ? xml.attribute(sound, 'r:embed')?.value : undefined;
    const morph = xml.descendants(transition).find(({ localName }) => localName === 'morph');
    return {
      effect: morph ? 'morph' : effect?.localName ?? 'none',
      ...(speed ? { speed } : {}),
      ...(durationMs !== undefined ? { durationMs } : {}),
      advanceOnClick: xml.attribute(transition, 'advClick')?.value !== '0',
      ...(advanceAfterMs !== undefined ? { advanceAfterMs } : {}),
      ...(direction ? { direction } : {}),
      ...(soundRelationshipId ? { soundRelationshipId } : {}),
    };
  }

  set(slidePartUri: string, value: SlideTransition): void {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const root = xml.elements('sld')[0];
    if (!root) throw new Error(`Slide ${slidePartUri} has no root`);
    let transition = xml.elements('transition')[0];
    if (!transition) {
      const insertionTarget = root.children.find(
        (child): child is XmlElement => child.type === 'element' && ['timing', 'extLst'].includes(child.localName),
      );
      const encoded = encodeTransition(value);
      xml.replace(insertionTarget?.start ?? root.endTagStart, insertionTarget?.start ?? root.endTagStart, encoded);
      this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
      return;
    }
    if (value.speed) setAttribute(xml, transition, 'spd', value.speed);
    if (value.durationMs !== undefined) setAttribute(xml, transition, 'dur', Math.max(0, Math.round(value.durationMs)));
    if (value.advanceOnClick !== undefined) setAttribute(xml, transition, 'advClick', value.advanceOnClick ? 1 : 0);
    if (value.advanceAfterMs !== undefined) setAttribute(xml, transition, 'advTm', Math.max(0, Math.round(value.advanceAfterMs)));
    const currentEffect = transition.children.find(
      (child): child is XmlElement => child.type === 'element' && !['sndAc', 'extLst'].includes(child.localName),
    );
    const encodedEffect = encodeEffect(value);
    if (currentEffect) xml.replaceElement(currentEffect, encodedEffect);
    else if (encodedEffect) xml.replace(transition.startTagEnd, transition.startTagEnd, encodedEffect);
    this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
  }

  clear(slidePartUri: string): void {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const transition = xml.elements('transition')[0];
    if (!transition) return;
    xml.removeElement(transition);
    this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
  }

  setSound(slidePartUri: string, mediaPartUri: string, loop = false): string {
    this.pkg.requirePart(mediaPartUri);
    if (!this.get(slidePartUri)) this.set(slidePartUri, { effect: 'none' });
    if (this.get(slidePartUri)?.soundRelationshipId) this.clearSound(slidePartUri);
    const relationship = this.pkg.addRelationship(slidePartUri, {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
      target: relativeTarget(slidePartUri, mediaPartUri),
    });
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const transition = xml.elements('transition')[0]!;
    const current = xml.descendants(transition, 'sndAc')[0];
    const sound = `<p:sndAc><p:stSnd loop="${loop ? 1 : 0}"><p:snd r:embed="${relationship.id}" name="${escapeAttribute(
      mediaPartUri.split('/').at(-1) ?? 'transition-sound',
    )}"/></p:stSnd></p:sndAc>`;
    if (current) xml.replaceElement(current, sound);
    else xml.appendChildXml(transition, sound);
    this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    return relationship.id;
  }

  clearSound(slidePartUri: string): boolean {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const transition = xml.elements('transition')[0];
    const soundAction = transition ? xml.descendants(transition, 'sndAc')[0] : undefined;
    if (!soundAction) return false;
    const sound = xml.descendants(soundAction, 'snd')[0];
    const relationshipId = sound ? xml.attribute(sound, 'r:embed')?.value : undefined;
    xml.removeElement(soundAction);
    this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    if (relationshipId) this.pkg.removeRelationship(slidePartUri, relationshipId);
    return true;
  }

  diagnostics(value: SlideTransition, slidePartUri: string, profile: string): CodecDiagnostic[] {
    if (value.effect === 'morph' && profile === 'powerpoint-2010') {
      return [{
        severity: 'warning',
        code: 'TRANSITION_MORPH_PRESERVED_ONLY',
        message: 'Morph is preserved but is not native in PowerPoint 2010',
        partUri: slidePartUri,
      }];
    }
    return [];
  }

  private parse(partUri: string): LosslessXmlDocument {
    return LosslessXmlDocument.parse(this.pkg.requirePart(partUri).bytes);
  }
}

export function installTransitionPlugin(host: PluginHost): TransitionCodec {
  const codec = new TransitionCodec(host.opcPackage);
  host.codecRegistry.register(codec);
  return codec;
}

function encodeTransition(value: SlideTransition): string {
  const attributes = [
    value.speed ? ` spd="${value.speed}"` : '',
    value.durationMs !== undefined ? ` dur="${Math.max(0, Math.round(value.durationMs))}"` : '',
    value.advanceOnClick !== undefined ? ` advClick="${value.advanceOnClick ? 1 : 0}"` : '',
    value.advanceAfterMs !== undefined ? ` advTm="${Math.max(0, Math.round(value.advanceAfterMs))}"` : '',
  ].join('');
  return `<p:transition${attributes}>${encodeEffect(value)}</p:transition>`;
}

function encodeEffect(value: SlideTransition): string {
  if (value.effect === 'none') return '';
  if (value.effect === 'morph') {
    throw new Error('Morph creation requires a preserved PowerPoint extension node');
  }
  const direction = value.direction ? ` dir="${escapeAttribute(value.direction)}"` : '';
  return `<p:${safeName(value.effect)}${direction}/>`;
}

function setAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  value: string | number,
): void {
  const attribute = xml.attribute(element, name);
  if (attribute) xml.replaceAttribute(attribute, String(value));
  else xml.replace(element.startTagEnd - (element.selfClosing ? 2 : 1), element.startTagEnd - (element.selfClosing ? 2 : 1), ` ${name}="${value}"`);
}

function normalizeSpeed(value: string | undefined): SlideTransition['speed'] {
  return value === 'slow' || value === 'fast' ? value : 'med';
}

function numberValue(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function safeName(value: string): string {
  if (!/^[A-Za-z_][\w.-]*$/.test(value)) throw new Error(`Invalid transition effect ${value}`);
  return value;
}

function escapeAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function relativeTarget(sourcePartUri: string, targetPartUri: string): string {
  const sourceDirectory = sourcePartUri.slice(0, sourcePartUri.lastIndexOf('/'));
  const targetSegments = targetPartUri.split('/').filter(Boolean);
  const sourceSegments = sourceDirectory.split('/').filter(Boolean);
  while (targetSegments.length > 0 && sourceSegments.length > 0 && targetSegments[0] === sourceSegments[0]) {
    targetSegments.shift();
    sourceSegments.shift();
  }
  return `${'../'.repeat(sourceSegments.length)}${targetSegments.join('/')}`;
}
