import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  joinPartUri,
  OpcPackage,
  partUriBasename,
  partUriDirname,
  relativeRelationshipTarget,
  relationshipPartUri,
  type Relationship,
} from '@pptx/opc';
import {
  readThemeFonts,
  replaceThemeFonts,
} from './theme-fonts.internal.js';
import type { ThemeFontSnapshot, ThemeFontUpdate } from './theme-fonts.js';
import {
  readSlideNumber,
  replaceSlideNumber,
} from './slide-number.internal.js';
import type {
  SlideNumber,
  SlideNumberOptions,
} from './slide-number.js';

export type { ThemeFontSnapshot, ThemeFontUpdate } from './theme-fonts.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CONTENT = 'application/vnd.openxmlformats-officedocument.presentationml.';
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

export interface PlaceholderModel {
  readonly shapeId: number;
  readonly type: string;
  readonly index: number;
  readonly text: string;
}

export interface ThemeColorModel {
  readonly name: string;
  readonly kind: string;
  readonly value: string;
  readonly lastColor?: string;
}

export class ThemeModel {
  constructor(private readonly codec: MasterLayoutThemeCodec, readonly partUri: string) {}

  get colors(): readonly ThemeColorModel[] {
    const xml = this.codec.parse(this.partUri);
    const scheme = xml.elements('clrScheme')[0];
    if (!scheme) return [];
    return scheme.children
      .filter((child): child is XmlElement => child.type === 'element')
      .map((entry) => {
        const color = entry.children.find((child): child is XmlElement => child.type === 'element');
        const lastColor = color ? xml.attribute(color, 'lastClr')?.value : undefined;
        return {
          name: entry.localName,
          kind: color?.localName ?? 'unknown',
          value: color ? xml.attribute(color, 'val')?.value ?? '' : '',
          ...(lastColor ? { lastColor } : {}),
        };
      });
  }

  get fonts(): ThemeFontSnapshot | undefined {
    return readThemeFonts(this.codec.parse(this.partUri));
  }

  setFonts(value: ThemeFontUpdate): void {
    this.codec.pkg.transaction(() => {
      const xml = this.codec.parse(this.partUri);
      replaceThemeFonts(xml, value);
      if (xml.changed) this.codec.save(this.partUri, xml);
    });
  }

  setColor(name: string, value: string): void {
    const xml = this.codec.parse(this.partUri);
    const scheme = xml.elements('clrScheme')[0];
    const entry = scheme?.children.find(
      (child): child is XmlElement => child.type === 'element' && child.localName === name,
    );
    const color = entry?.children.find((child): child is XmlElement => child.type === 'element');
    if (!color) throw new Error(`Theme color ${name} was not found`);
    const attribute = xml.attribute(color, 'val');
    if (!attribute) throw new Error(`Theme color ${name} has no value`);
    xml.replaceAttribute(attribute, value.replace(/^#/, '').toUpperCase());
    this.codec.save(this.partUri, xml);
  }
}

export class LayoutModel {
  constructor(private readonly codec: MasterLayoutThemeCodec, readonly partUri: string) {}

  get name(): string {
    const xml = this.codec.parse(this.partUri);
    const layout = xml.elements('sldLayout')[0];
    const commonSlideData = xml.elements('cSld')[0];
    return (
      (layout ? xml.attribute(layout, 'name')?.value : undefined) ??
      (commonSlideData ? xml.attribute(commonSlideData, 'name')?.value : undefined) ??
      partUriBasename(this.partUri)
    );
  }

  get masterPartUri(): string | undefined {
    return this.codec.relationship(this.partUri, 'slideMaster')?.resolvedTarget;
  }

  get placeholders(): readonly PlaceholderModel[] {
    return this.codec.placeholders(this.partUri);
  }

  get slideNumber(): Readonly<SlideNumber> | undefined {
    return readSlideNumber(this.codec.pkg, this.partUri, 'layout');
  }

  set slideNumber(value: SlideNumberOptions | undefined) {
    replaceSlideNumber(this.codec.pkg, this.partUri, 'layout', value, '‹#›');
  }
}

export class MasterModel {
  constructor(private readonly codec: MasterLayoutThemeCodec, readonly partUri: string) {}

  get layouts(): readonly LayoutModel[] {
    return this.codec.pkg
      .relationships(this.partUri)
      .filter(({ type, resolvedTarget }) => type.endsWith('/slideLayout') && resolvedTarget)
      .map(({ resolvedTarget }) => this.codec.modelForLayout(resolvedTarget!));
  }

  get theme(): ThemeModel | undefined {
    const target = this.codec.relationship(this.partUri, 'theme')?.resolvedTarget;
    return target ? this.codec.modelForTheme(target) : undefined;
  }

  get placeholders(): readonly PlaceholderModel[] {
    return this.codec.placeholders(this.partUri);
  }

  get slideNumber(): Readonly<SlideNumber> | undefined {
    return readSlideNumber(this.codec.pkg, this.partUri, 'master');
  }

  set slideNumber(value: SlideNumberOptions | undefined) {
    replaceSlideNumber(this.codec.pkg, this.partUri, 'master', value, '‹#›');
  }
}

export interface MaterializedPlaceholder {
  readonly type: string;
  readonly index: number;
  readonly slide?: { partUri: string; xml: string };
  readonly layout?: { partUri: string; xml: string };
  readonly master?: { partUri: string; xml: string };
}

export class MasterLayoutThemeCodec {
  readonly id = 'builtin.master-layout-theme';
  readonly priority = 100;
  readonly ownership = {
    relationshipTypes: [`${REL}slideLayout`, `${REL}slideMaster`, `${REL}theme`],
    contentTypes: [`${CONTENT}slideMaster+xml`, `${CONTENT}slideLayout+xml`, THEME_CONTENT_TYPE],
  } as const;
  readonly #masterModels = new Map<string, MasterModel>();
  readonly #layoutModels = new Map<string, LayoutModel>();
  readonly #themeModels = new Map<string, ThemeModel>();

  constructor(readonly pkg: OpcPackage, readonly presentationPartUri = '/ppt/presentation.xml') {}

  get masters(): readonly MasterModel[] {
    return this.pkg
      .relationships(this.presentationPartUri)
      .filter(({ type, resolvedTarget }) => type.endsWith('/slideMaster') && resolvedTarget)
      .map(({ resolvedTarget }) => this.modelForMaster(resolvedTarget!));
  }

  get layouts(): readonly LayoutModel[] {
    return this.masters.flatMap(({ layouts }) => layouts);
  }

  get presentationTheme(): ThemeModel | undefined {
    const relationships = this.pkg
      .relationships(this.presentationPartUri)
      .filter(({ type }) => type === `${REL}theme`);
    if (relationships.length !== 1) return undefined;
    const relationship = relationships[0];
    if (
      !relationship
      || relationship.targetMode === 'External'
      || !relationship.resolvedTarget
    ) {
      return undefined;
    }
    const part = this.pkg.getPart(relationship.resolvedTarget);
    return part?.contentType === THEME_CONTENT_TYPE
      ? this.modelForTheme(relationship.resolvedTarget)
      : undefined;
  }

  get themes(): readonly ThemeModel[] {
    const uris = new Set(
      this.pkg.parts
        .filter(({ contentType }) => contentType === THEME_CONTENT_TYPE)
        .map(({ uri }) => uri),
    );
    return [...uris].map((uri) => this.modelForTheme(uri));
  }

  /** @internal */
  modelForMaster(partUri: string): MasterModel {
    const existing = this.#masterModels.get(partUri);
    if (existing) return existing;
    const created = new MasterModel(this, partUri);
    this.#masterModels.set(partUri, created);
    return created;
  }

  /** @internal */
  modelForLayout(partUri: string): LayoutModel {
    const existing = this.#layoutModels.get(partUri);
    if (existing) return existing;
    const created = new LayoutModel(this, partUri);
    this.#layoutModels.set(partUri, created);
    return created;
  }

  /** @internal */
  modelForTheme(partUri: string): ThemeModel {
    const existing = this.#themeModels.get(partUri);
    if (existing) return existing;
    const created = new ThemeModel(this, partUri);
    this.#themeModels.set(partUri, created);
    return created;
  }

  createTheme(xml: string): ThemeModel {
    return this.pkg.transaction(() => {
      LosslessXmlDocument.parse(xml);
      const uri = this.pkg.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'theme'),
        'theme',
        '.xml',
      );
      this.pkg.setPart(uri, xml, THEME_CONTENT_TYPE);
      return this.modelForTheme(uri);
    });
  }

  copyTheme(themePartUri: string): ThemeModel {
    return this.pkg.transaction(() => {
      const part = this.pkg.requirePart(themePartUri);
      return this.createTheme(new TextDecoder().decode(part.bytes));
    });
  }

  deleteTheme(themePartUri: string): void {
    this.pkg.transaction(() => {
      const incoming = this.pkg.graph.find(({ uri }) => uri === themePartUri)?.incoming ?? [];
      if (incoming.length > 0) throw new Error(`Theme ${themePartUri} is still referenced by ${incoming.length} part(s)`);
      this.pkg.deletePart(themePartUri);
    });
  }

  relinkMasterTheme(masterPartUri: string, themePartUri: string): void {
    this.pkg.transaction(() => {
      const relationship = this.relationship(masterPartUri, 'theme');
      const target = relativeTarget(masterPartUri, themePartUri);
      if (relationship) this.pkg.updateRelationship(masterPartUri, relationship.id, { target, targetMode: 'Internal' });
      else this.pkg.addRelationship(masterPartUri, { type: `${REL}theme`, target });
    });
  }

  createMaster(xml: string, themePartUri: string): MasterModel {
    return this.pkg.transaction(() => {
      LosslessXmlDocument.parse(xml);
      const uri = this.pkg.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'slideMasters'),
        'slideMaster',
        '.xml',
      );
      this.pkg.setPart(uri, xml, `${CONTENT}slideMaster+xml`);
      this.pkg.addRelationship(uri, { type: `${REL}theme`, target: relativeTarget(uri, themePartUri) });
      this.attachMaster(uri);
      return this.modelForMaster(uri);
    });
  }

  copyMaster(masterPartUri: string): MasterModel {
    return this.pkg.transaction(() => {
      const source = this.pkg.requirePart(masterPartUri);
      const uri = this.pkg.allocatePartUri(partUriDirname(masterPartUri), 'slideMaster', '.xml');
      this.pkg.setPart(uri, source.bytes, source.contentType);
      for (const relationship of this.pkg.relationships(masterPartUri)) {
        if (relationship.type.endsWith('/slideLayout') && relationship.resolvedTarget) {
          const layoutUri = this.copyLayoutPart(relationship.resolvedTarget, uri);
          this.pkg.addRelationship(uri, {
            id: relationship.id,
            type: relationship.type,
            target: relativeTarget(uri, layoutUri),
          });
        } else {
          this.pkg.addRelationship(uri, {
            id: relationship.id,
            type: relationship.type,
            target: relationship.target,
            targetMode: relationship.targetMode,
          });
        }
      }
      this.attachMaster(uri);
      return this.modelForMaster(uri);
    });
  }

  deleteMaster(masterPartUri: string, replacementMasterPartUri?: string): void {
    this.pkg.transaction(() => {
      const master = this.modelForMaster(masterPartUri);
      const replacementLayout = replacementMasterPartUri
        ? this.modelForMaster(replacementMasterPartUri).layouts[0]?.partUri
        : undefined;
      for (const layout of master.layouts) this.deleteLayout(layout.partUri, replacementLayout);
      const relationship = this.pkg
        .relationships(this.presentationPartUri)
        .find(({ resolvedTarget, type }) => type.endsWith('/slideMaster') && resolvedTarget === masterPartUri);
      if (relationship) {
        const xml = this.parse(this.presentationPartUri);
        const element = xml
          .elements('sldMasterId')
          .find((candidate) => xml.attribute(candidate, 'r:id')?.value === relationship.id);
        if (element) xml.removeElement(element);
        this.save(this.presentationPartUri, xml);
        this.pkg.removeRelationship(this.presentationPartUri, relationship.id);
      }
      this.pkg.deletePart(masterPartUri);
    });
  }

  createLayout(masterPartUri: string, xml: string): LayoutModel {
    return this.pkg.transaction(() => {
      LosslessXmlDocument.parse(xml);
      const uri = this.pkg.allocatePartUri(
        joinPartUri(partUriDirname(masterPartUri), '../slideLayouts'),
        'slideLayout',
        '.xml',
      );
      this.pkg.setPart(uri, xml, `${CONTENT}slideLayout+xml`);
      this.pkg.addRelationship(uri, { type: `${REL}slideMaster`, target: relativeTarget(uri, masterPartUri) });
      this.attachLayout(masterPartUri, uri);
      return this.modelForLayout(uri);
    });
  }

  copyLayout(layoutPartUri: string, masterPartUri?: string): LayoutModel {
    return this.pkg.transaction(() => {
      const sourceMaster = this.relationship(layoutPartUri, 'slideMaster')?.resolvedTarget;
      const targetMaster = masterPartUri ?? sourceMaster;
      if (!targetMaster) throw new Error(`Layout ${layoutPartUri} has no master`);
      const uri = this.copyLayoutPart(layoutPartUri, targetMaster);
      this.attachLayout(targetMaster, uri);
      return this.modelForLayout(uri);
    });
  }

  deleteLayout(layoutPartUri: string, replacementLayoutPartUri?: string): void {
    this.pkg.transaction(() => {
      const incomingSlides = (this.pkg.graph.find(({ uri }) => uri === layoutPartUri)?.incoming ?? []).filter(
        ({ sourceUri, relationship }) =>
          relationship.type.endsWith('/slideLayout') &&
          this.pkg.getPart(sourceUri)?.contentType === `${CONTENT}slide+xml`,
      );
      if (incomingSlides.length > 0 && !replacementLayoutPartUri) {
        throw new Error(`Layout ${layoutPartUri} is still used by ${incomingSlides.length} slide(s)`);
      }
      for (const { sourceUri } of incomingSlides) this.relinkSlideLayout(sourceUri, replacementLayoutPartUri!);
      const masterPartUri = this.relationship(layoutPartUri, 'slideMaster')?.resolvedTarget;
      if (masterPartUri) {
        const relationship = this.pkg
          .relationships(masterPartUri)
          .find(({ resolvedTarget, type }) => type.endsWith('/slideLayout') && resolvedTarget === layoutPartUri);
        if (relationship) {
          const xml = this.parse(masterPartUri);
          const element = xml
            .elements('sldLayoutId')
            .find((candidate) => xml.attribute(candidate, 'r:id')?.value === relationship.id);
          if (element) xml.removeElement(element);
          this.save(masterPartUri, xml);
          this.pkg.removeRelationship(masterPartUri, relationship.id);
        }
      }
      this.pkg.deletePart(layoutPartUri);
    });
  }

  relinkSlideLayout(slidePartUri: string, layoutPartUri: string): void {
    this.pkg.transaction(() => {
      const relationship = this.relationship(slidePartUri, 'slideLayout');
      const target = relativeTarget(slidePartUri, layoutPartUri);
      if (relationship) this.pkg.updateRelationship(slidePartUri, relationship.id, { target, targetMode: 'Internal' });
      else this.pkg.addRelationship(slidePartUri, { type: `${REL}slideLayout`, target });
    });
  }

  materializeInheritedStyle(slidePartUri: string, shapeId: number): MaterializedPlaceholder | undefined {
    const slideXml = this.parse(slidePartUri);
    const slideShape = findShape(slideXml, shapeId);
    const placeholder = slideShape ? slideXml.descendants(slideShape, 'ph')[0] : undefined;
    if (!slideShape || !placeholder) return undefined;
    const type = slideXml.attribute(placeholder, 'type')?.value ?? 'body';
    const index = Number(slideXml.attribute(placeholder, 'idx')?.value ?? 0);
    const result: MaterializedPlaceholder = {
      type,
      index,
      slide: { partUri: slidePartUri, xml: slideXml.original(slideShape) },
    };
    const layoutPartUri = this.relationship(slidePartUri, 'slideLayout')?.resolvedTarget;
    if (!layoutPartUri) return result;
    const layoutXml = this.parse(layoutPartUri);
    const layoutShape = findPlaceholder(layoutXml, type, index);
    const masterPartUri = this.relationship(layoutPartUri, 'slideMaster')?.resolvedTarget;
    const masterXml = masterPartUri ? this.parse(masterPartUri) : undefined;
    const masterShape = masterXml ? findPlaceholder(masterXml, type, index) : undefined;
    return {
      ...result,
      ...(layoutShape ? { layout: { partUri: layoutPartUri, xml: layoutXml.original(layoutShape) } } : {}),
      ...(masterShape && masterPartUri && masterXml
        ? { master: { partUri: masterPartUri, xml: masterXml.original(masterShape) } }
        : {}),
    };
  }

  placeholders(partUri: string): PlaceholderModel[] {
    const xml = this.parse(partUri);
    return xml.elements('sp').flatMap((shape) => {
      const placeholder = xml.descendants(shape, 'ph')[0];
      const properties = xml.descendants(shape, 'cNvPr')[0];
      if (!placeholder || !properties) return [];
      return [{
        shapeId: Number(xml.attribute(properties, 'id')?.value ?? 0),
        type: xml.attribute(placeholder, 'type')?.value ?? 'body',
        index: Number(xml.attribute(placeholder, 'idx')?.value ?? 0),
        text: xml.descendants(shape, 't').map((text) => xml.text(text)).join(''),
      }];
    });
  }

  relationship(partUri: string, suffix: string): Relationship | undefined {
    return this.pkg.relationships(partUri).find(({ type }) => type.endsWith(`/${suffix}`));
  }

  parse(partUri: string): LosslessXmlDocument {
    return LosslessXmlDocument.parse(this.pkg.requirePart(partUri).bytes);
  }

  save(partUri: string, xml: LosslessXmlDocument): void {
    const part = this.pkg.requirePart(partUri);
    this.pkg.setPart(partUri, xml.serialize(), part.contentType);
  }

  private attachMaster(masterPartUri: string): void {
    const relationship = this.pkg.addRelationship(this.presentationPartUri, {
      type: `${REL}slideMaster`,
      target: relativeTarget(this.presentationPartUri, masterPartUri),
    });
    const xml = this.parse(this.presentationPartUri);
    const root = xml.elements('presentation')[0];
    if (!root) throw new Error('Presentation XML is invalid');
    const list = xml.elements('sldMasterIdLst')[0];
    const id = Math.max(2_147_483_647, ...xml.elements('sldMasterId').map((item) => Number(xml.attribute(item, 'id')?.value ?? 0))) + 1;
    const entry = `<p:sldMasterId id="${id}" r:id="${relationship.id}"/>`;
    if (list) xml.appendChildXml(list, entry);
    else xml.replace(root.startTagEnd, root.startTagEnd, `<p:sldMasterIdLst>${entry}</p:sldMasterIdLst>`);
    this.save(this.presentationPartUri, xml);
  }

  private attachLayout(masterPartUri: string, layoutPartUri: string): void {
    const relationship = this.pkg.addRelationship(masterPartUri, {
      type: `${REL}slideLayout`,
      target: relativeTarget(masterPartUri, layoutPartUri),
    });
    const xml = this.parse(masterPartUri);
    const root = xml.elements('sldMaster')[0];
    if (!root) throw new Error(`Master ${masterPartUri} XML is invalid`);
    const list = xml.elements('sldLayoutIdLst')[0];
    const id = Math.max(0, ...xml.elements('sldLayoutId').map((item) => Number(xml.attribute(item, 'id')?.value ?? 0))) + 1;
    const entry = `<p:sldLayoutId id="${id}" r:id="${relationship.id}"/>`;
    if (list) xml.appendChildXml(list, entry);
    else xml.replace(root.startTagEnd, root.startTagEnd, `<p:sldLayoutIdLst>${entry}</p:sldLayoutIdLst>`);
    this.save(masterPartUri, xml);
  }

  private copyLayoutPart(layoutPartUri: string, targetMasterPartUri: string): string {
    const source = this.pkg.requirePart(layoutPartUri);
    const uri = this.pkg.allocatePartUri(partUriDirname(layoutPartUri), 'slideLayout', '.xml');
    this.pkg.setPart(uri, source.bytes, source.contentType);
    for (const relationship of this.pkg.relationships(layoutPartUri)) {
      this.pkg.addRelationship(uri, {
        id: relationship.id,
        type: relationship.type,
        target: relationship.type.endsWith('/slideMaster')
          ? relativeTarget(uri, targetMasterPartUri)
          : relationship.target,
        targetMode: relationship.targetMode,
      });
    }
    return uri;
  }
}

function relativeTarget(sourcePartUri: string, targetPartUri: string): string {
  return relativeRelationshipTarget(sourcePartUri, targetPartUri);
}

function findShape(xml: LosslessXmlDocument, shapeId: number): XmlElement | undefined {
  return xml.elements('sp').find((shape) => {
    const properties = xml.descendants(shape, 'cNvPr')[0];
    return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
  });
}

function findPlaceholder(xml: LosslessXmlDocument, type: string, index: number): XmlElement | undefined {
  return xml.elements('sp').find((shape) => {
    const placeholder = xml.descendants(shape, 'ph')[0];
    return (
      placeholder &&
      (xml.attribute(placeholder, 'type')?.value ?? 'body') === type &&
      Number(xml.attribute(placeholder, 'idx')?.value ?? 0) === index
    );
  });
}
