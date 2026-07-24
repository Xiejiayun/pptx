import { describe, expect, it } from 'vitest';
import { CodecRegistry } from '@pptx/codecs';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage } from '@pptx/opc';
import { createMinimalPptx } from '@pptx/testkit';
import { installSmartArtPlugin } from './index.js';

describe('SmartArtDiagramCodec', () => {
  it('reads the part set and edits nodes while preserving style parts', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const slidePart = pkg.requirePart('/ppt/slides/slide1.xml');
    const slideXml = LosslessXmlDocument.parse(slidePart.bytes);
    const tree = slideXml.elements('spTree')[0]!;
    slideXml.appendChildXml(tree, '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="SmartArt"/></p:nvGraphicFramePr><a:graphic><a:graphicData><dgm:relIds xmlns:dgm="dgm" xmlns:r="r" r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4" r:drawing="rId5"/></a:graphicData></a:graphic></p:graphicFrame>');
    pkg.setPart('/ppt/slides/slide1.xml', slideXml.serialize(), slidePart.contentType);
    pkg.setPart('/ppt/diagrams/data1.xml', '<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst><dgm:pt modelId="root" type="node"><dgm:t><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t></dgm:pt><x:unknown xmlns:x="urn:test"/></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>', 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml');
    pkg.setPart('/ppt/diagrams/layout1.xml', '<dgm:layoutDef xmlns:dgm="dgm"/>', 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml');
    pkg.setPart('/ppt/diagrams/quickStyle1.xml', '<dgm:styleDef xmlns:dgm="dgm"/>', 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml');
    pkg.setPart('/ppt/diagrams/colors1.xml', '<dgm:colorsDef xmlns:dgm="dgm"/>', 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml');
    pkg.setPart('/ppt/diagrams/drawing1.xml', '<dsp:drawing xmlns:dsp="dsp"/>', 'application/vnd.ms-office.drawingml.diagramDrawing+xml');
    const relationships = [
      ['rId1', 'diagramData', '../diagrams/data1.xml'],
      ['rId2', 'diagramLayout', '../diagrams/layout1.xml'],
      ['rId3', 'diagramQuickStyle', '../diagrams/quickStyle1.xml'],
      ['rId4', 'diagramColors', '../diagrams/colors1.xml'],
      ['rId5', 'diagramDrawing', '../diagrams/drawing1.xml'],
    ] as const;
    for (const [id, type, target] of relationships) {
      pkg.addRelationship('/ppt/slides/slide1.xml', {
        id,
        type: type === 'diagramDrawing'
          ? 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
          : `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}`,
        target,
      });
    }
    const codec = installSmartArtPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    const initial = codec.list('/ppt/slides/slide1.xml')[0]!;
    expect(initial.nodes[0]?.text).toBe('Root');
    codec.replaceText(initial.dataPartUri, 'root', 'Updated root');
    const child = codec.addNode(initial.dataPartUri, 'Child', 'root');
    expect(codec.list('/ppt/slides/slide1.xml')[0]?.nodes).toHaveLength(2);
    expect(codec.diagnostics(codec.list('/ppt/slides/slide1.xml')[0]!)[0]?.code).toBe('SMARTART_RELAYOUT_REQUIRED');
    codec.deleteNode(initial.dataPartUri, child.id);
    const data = new TextDecoder().decode(pkg.requirePart(initial.dataPartUri).bytes);
    expect(data).toContain('<x:unknown xmlns:x="urn:test"/>');
    expect(data).not.toContain(child.id);
    expect(pkg.hasPart('/ppt/diagrams/quickStyle1.xml')).toBe(true);
    const reopenedCodec = new (await import('./index.js')).SmartArtDiagramCodec(
      await OpcPackage.open(await pkg.write()),
    );
    expect(reopenedCodec.list('/ppt/slides/slide1.xml')[0]?.needsLayout).toBe(true);
  });
});
