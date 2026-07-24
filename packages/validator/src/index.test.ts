import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { validatePackage } from './index.js';

async function invalidFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="1 bad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/missing.xml"/><Relationship Id="1 bad" Type="https://example.test/external" Target="https://example.test/a" TargetMode="External"/></Relationships>');
  return zip.generateAsync({ type: 'uint8array' });
}

describe('validatePackage', () => {
  it('reports duplicate ids, invalid ids, dangling targets, and external resources', async () => {
    const diagnostics = validatePackage(await OpcPackage.open(await invalidFixture()));
    const codes = diagnostics.map(({ code }) => code);
    expect(codes).toContain('OPC_DUPLICATE_RELATIONSHIP_ID');
    expect(codes).toContain('OPC_INVALID_RELATIONSHIP_ID');
    expect(codes).toContain('OPC_DANGLING_RELATIONSHIP');
    expect(codes).toContain('OPC_EXTERNAL_RELATIONSHIP');
  });
});
