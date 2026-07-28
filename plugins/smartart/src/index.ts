import type { CodecDiagnostic, CodecRegistry } from '@pptx/codecs';
import { escapeXmlAttribute, escapeXmlText, LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';

export interface SmartArtNode {
  readonly id: string;
  readonly type: string;
  readonly text: string;
}

export interface SmartArtModel {
  readonly slidePartUri: string;
  readonly dataPartUri: string;
  readonly layoutPartUri?: string;
  readonly quickStylePartUri?: string;
  readonly colorsPartUri?: string;
  readonly fallbackDrawingPartUri?: string;
  readonly nodes: readonly SmartArtNode[];
  readonly needsLayout: boolean;
}

export interface PluginHost {
  readonly opcPackage: OpcPackage;
  readonly codecRegistry: CodecRegistry;
}

export class SmartArtDiagramCodec {
  readonly id = 'plugin.smartart-diagram';
  readonly priority = 200;
  readonly ownership = {
    elements: ['dgm:relIds', 'dgm:dataModel'],
    relationshipTypes: [
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors',
      'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
    ],
  } as const;
  readonly #mutatedDataParts = new Set<string>();

  constructor(readonly pkg: OpcPackage) {}

  list(slidePartUri: string): readonly SmartArtModel[] {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(slidePartUri).bytes);
    const relationships = this.pkg.relationships(slidePartUri);
    return xml.elements('relIds').flatMap((relIds) => {
      const resolve = (name: string): string | undefined => {
        const id = xml.attribute(relIds, name)?.value;
        return relationships.find((relationship) => relationship.id === id)?.resolvedTarget;
      };
      const dataPartUri = resolve('r:dm');
      if (!dataPartUri) return [];
      const dataXml = LosslessXmlDocument.parse(this.pkg.requirePart(dataPartUri).bytes);
      const fallbackDrawingPartUri = resolve('r:drawing');
      return [{
        slidePartUri,
        dataPartUri,
        ...(resolve('r:lo') ? { layoutPartUri: resolve('r:lo')! } : {}),
        ...(resolve('r:qs') ? { quickStylePartUri: resolve('r:qs')! } : {}),
        ...(resolve('r:cs') ? { colorsPartUri: resolve('r:cs')! } : {}),
        ...(fallbackDrawingPartUri ? { fallbackDrawingPartUri } : {}),
        nodes: decodeNodes(dataXml),
        needsLayout:
          Boolean(fallbackDrawingPartUri) &&
          (this.#mutatedDataParts.has(dataPartUri) || dataXml.elements('needsLayout').length > 0),
      }];
    });
  }

  replaceText(dataPartUri: string, nodeId: string, text: string): void {
    const { xml, node } = this.resolveNode(dataPartUri, nodeId);
    const textElement = xml.descendants(node, 'a:t')[0];
    if (!textElement) throw new Error(`SmartArt node ${nodeId} has no text body`);
    xml.replaceText(textElement, text);
    this.save(dataPartUri, xml);
  }

  addNode(dataPartUri: string, text: string, parentId?: string): SmartArtNode {
    const part = this.pkg.requirePart(dataPartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const list = xml.elements('ptLst')[0];
    if (!list) throw new Error(`SmartArt data ${dataPartUri} has no point list`);
    const id = `{${randomUuid().toUpperCase()}}`;
    xml.appendChildXml(
      list,
      `<dgm:pt modelId="${escapeXmlAttribute(id)}" type="node"><dgm:prSet/><dgm:t><a:p><a:r><a:t>${escapeXmlText(
        text,
      )}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
    );
    if (parentId) {
      const connections = xml.elements('cxnLst')[0];
      if (!connections) throw new Error(`SmartArt data ${dataPartUri} has no connection list`);
      const connectionId = `{${randomUuid().toUpperCase()}}`;
      xml.appendChildXml(
        connections,
        `<dgm:cxn modelId="${escapeXmlAttribute(connectionId)}" type="parOf" srcId="${escapeXmlAttribute(
          parentId,
        )}" destId="${escapeXmlAttribute(id)}" srcOrd="0" destOrd="0"/>`,
      );
    }
    this.save(dataPartUri, xml);
    return { id, type: 'node', text };
  }

  deleteNode(dataPartUri: string, nodeId: string): void {
    const { xml, node } = this.resolveNode(dataPartUri, nodeId);
    xml.removeElement(node);
    for (const connection of xml.elements('cxn')) {
      if (
        xml.attribute(connection, 'srcId')?.value === nodeId ||
        xml.attribute(connection, 'destId')?.value === nodeId
      ) {
        xml.removeElement(connection);
      }
    }
    this.save(dataPartUri, xml);
  }

  diagnostics(model: SmartArtModel): CodecDiagnostic[] {
    if (model.needsLayout) {
      return [{
        severity: 'warning',
        code: 'SMARTART_RELAYOUT_REQUIRED',
        message: 'SmartArt data changed; PowerPoint should regenerate or realign the fallback drawing',
        partUri: model.dataPartUri,
      }];
    }
    return [];
  }

  private resolveNode(dataPartUri: string, nodeId: string): { xml: LosslessXmlDocument; node: XmlElement } {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(dataPartUri).bytes);
    const node = xml.elements('pt').find((point) => xml.attribute(point, 'modelId')?.value === nodeId);
    if (!node) throw new Error(`SmartArt node ${nodeId} was not found`);
    return { xml, node };
  }

  private save(dataPartUri: string, xml: LosslessXmlDocument): void {
    const part = this.pkg.requirePart(dataPartUri);
    if (xml.elements('needsLayout').length === 0) {
      const root = xml.elements('dataModel')[0];
      if (!root) throw new Error(`SmartArt data ${dataPartUri} has no dataModel root`);
      const extensions = xml.elements('extLst')[0];
      const marker = '<a:ext uri="{50505458-534D-4152-5441-52544C41594F}"><px:needsLayout xmlns:px="urn:pptx-ooxml:smartart"/></a:ext>';
      if (extensions) xml.appendChildXml(extensions, marker);
      else xml.appendChildXml(root, `<dgm:extLst>${marker}</dgm:extLst>`);
    }
    this.pkg.setPart(dataPartUri, xml.serialize(), part.contentType);
    this.#mutatedDataParts.add(dataPartUri);
  }
}

function randomUuid(): string {
  if (!globalThis.crypto?.randomUUID) throw new Error('SmartArt node creation requires crypto.randomUUID()');
  return globalThis.crypto.randomUUID();
}

export function installSmartArtPlugin(host: PluginHost): SmartArtDiagramCodec {
  const codec = new SmartArtDiagramCodec(host.opcPackage);
  host.codecRegistry.register(codec);
  return codec;
}

function decodeNodes(xml: LosslessXmlDocument): SmartArtNode[] {
  return xml.elements('pt').map((point) => ({
    id: xml.attribute(point, 'modelId')?.value ?? '',
    type: xml.attribute(point, 'type')?.value ?? 'node',
    text: xml.descendants(point, 'a:t').map((element) => xml.text(element)).join(''),
  }));
}
