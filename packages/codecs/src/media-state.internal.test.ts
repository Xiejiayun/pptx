import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage } from '@pptx/opc';
import { readMediaState } from './media-state.internal.js';
import { syncNativeMediaTiming } from './media-timing-edit.internal.js';

const SLIDE = '/ppt/slides/slide1.xml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';

describe('readMediaState', () => {
  it('reads canonical embedded and external state without mutation', () => {
    const pkg = OpcPackage.create();
    const source = slideXml(
      mediaPicture(2, 'audioFile', 'rId1', 'rId2', 'rId3', 'Narration', 'Overview'),
      mediaPicture(3, 'videoFile', 'rId4', undefined, 'rId5', 'Remote video'),
    );
    pkg.setPart(SLIDE, source, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    pkg.setPart('/ppt/media/audio1.mp3', Uint8Array.of(1), 'audio/mpeg');
    pkg.setPart('/ppt/media/poster1.png', Uint8Array.of(2), 'image/png');
    pkg.setPart('/ppt/media/poster2.jpg', Uint8Array.of(3), 'image/jpeg');
    addRelationship(pkg, 'rId1', `${REL}audio`, '../media/audio1.mp3');
    addRelationship(pkg, 'rId2', MEDIA_REL, '../media/audio1.mp3');
    addRelationship(pkg, 'rId3', `${REL}image`, '../media/poster1.png');
    addRelationship(pkg, 'rId4', `${REL}video`, 'https://example.com/video.mp4', 'External');
    addRelationship(pkg, 'rId5', `${REL}image`, '../media/poster2.jpg');
    const before = snapshot(pkg);
    const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE).bytes);
    const pictures = xml.elements('pic');

    expect(readMediaState(pkg, SLIDE, xml, pictures[0]!)).toEqual({
      kind: 'audio',
      shapeId: 2,
      slidePartUri: SLIDE,
      name: 'Narration',
      altText: 'Overview',
      mediaPartUri: '/ppt/media/audio1.mp3',
      posterPartUri: '/ppt/media/poster1.png',
      settings: {
        play: 'auto',
        loop: true,
        hideWhenStopped: false,
        volume: 0.5,
      },
    });
    expect(readMediaState(pkg, SLIDE, xml, pictures[1]!)).toMatchObject({
      kind: 'video',
      shapeId: 3,
      externalUrl: 'https://example.com/video.mp4',
      posterPartUri: '/ppt/media/poster2.jpg',
    });
    expect(snapshot(pkg)).toEqual(before);
  });

  it('prefers relationship, then MIME, then marker for PptxGenJS legacy audio', () => {
    const pkg = OpcPackage.create();
    const source = slideXml(
      mediaPicture(2, 'videoFile', 'rId1', 'rId2', 'rId3', 'Wrong marker'),
      mediaPicture(3, 'videoFile', 'rId4', 'rId4', 'rId3', 'MIME fallback'),
      mediaPicture(4, 'audioFile', 'rId5', undefined, 'rId3', 'Marker fallback'),
      '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Ordinary"/><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr/></p:pic>',
    );
    pkg.setPart(SLIDE, source, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    pkg.setPart('/ppt/media/audio1.mp3', Uint8Array.of(1), 'audio/mp3');
    pkg.setPart('/ppt/media/opaque.bin', Uint8Array.of(2), 'application/octet-stream');
    pkg.setPart('/ppt/media/poster1.png', Uint8Array.of(3), 'image/png');
    addRelationship(pkg, 'rId1', `${REL}audio`, '../media/audio1.mp3');
    addRelationship(pkg, 'rId2', MEDIA_REL, '../media/audio1.mp3');
    addRelationship(pkg, 'rId3', `${REL}image`, '../media/poster1.png');
    addRelationship(pkg, 'rId4', MEDIA_REL, '../media/audio1.mp3');
    addRelationship(pkg, 'rId5', MEDIA_REL, '../media/opaque.bin');
    const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE).bytes);
    const pictures = xml.elements('pic');

    expect(readMediaState(pkg, SLIDE, xml, pictures[0]!)?.kind).toBe('audio');
    expect(readMediaState(pkg, SLIDE, xml, pictures[1]!)?.kind).toBe('audio');
    expect(readMediaState(pkg, SLIDE, xml, pictures[2]!)?.kind).toBe('audio');
    expect(readMediaState(pkg, SLIDE, xml, pictures[3]!)).toBeUndefined();
  });

  it('returns frozen detached optional state and rejects invalid picture identity', () => {
    const pkg = OpcPackage.create();
    pkg.setPart(
      SLIDE,
      slideXml(
        mediaPicture(2, 'audioFile', 'rId1', undefined, undefined, 'Audio'),
        '<p:pic><p:nvPicPr><p:cNvPr id="bad" name="Bad"/><p:nvPr>' +
          '<a:audioFile r:link="rId1"/></p:nvPr></p:nvPicPr><p:spPr/></p:pic>',
      ),
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    );
    addRelationship(pkg, 'rId1', `${REL}audio`, 'https://example.com/audio.mp3', 'External');
    const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE).bytes);
    const first = readMediaState(pkg, SLIDE, xml, xml.elements('pic')[0]!)!;

    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.settings)).toBe(true);
    expect(first.settings).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.5,
    });
    expect(readMediaState(pkg, SLIDE, xml, xml.elements('pic')[1]!)).toBeUndefined();
  });

  it('projects one recognized native graph only when no valid private preference exists', () => {
    const pkg = OpcPackage.create();
    const source = slideXml(
      mediaPicture(2, 'audioFile', 'rId1', undefined, undefined, 'Native only', undefined, false),
      mediaPicture(3, 'videoFile', 'rId2', undefined, undefined, 'Private wins'),
      mediaPicture(4, 'audioFile', 'rId3', undefined, undefined, 'Unsupported', undefined, false),
    );
    const xml = LosslessXmlDocument.parse(source);
    syncNativeMediaTiming(xml, 2, 'audio', {
      play: 'auto',
      loop: false,
      hideWhenStopped: true,
      volume: 0.25,
    });
    const withFirst = LosslessXmlDocument.parse(xml.serialize());
    syncNativeMediaTiming(withFirst, 3, 'video', {
      play: 'click',
      loop: false,
      hideWhenStopped: false,
      volume: 1,
    });
    const withSecond = LosslessXmlDocument.parse(withFirst.serialize());
    syncNativeMediaTiming(withSecond, 4, 'audio', {
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.5,
    });
    pkg.setPart(
      SLIDE,
      withSecond.serialize().replace('repeatCount="indefinite"', 'repeatCount="2000"'),
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    );
    addRelationship(pkg, 'rId1', `${REL}audio`, 'https://example.com/native.mp3', 'External');
    addRelationship(pkg, 'rId2', `${REL}video`, 'https://example.com/private.mp4', 'External');
    addRelationship(pkg, 'rId3', `${REL}audio`, 'https://example.com/unsupported.mp3', 'External');
    const before = snapshot(pkg);
    const parsed = LosslessXmlDocument.parse(pkg.requirePart(SLIDE).bytes);
    const pictures = parsed.elements('pic');

    expect(readMediaState(pkg, SLIDE, parsed, pictures[0]!)?.settings).toEqual({
      play: 'auto',
      loop: false,
      hideWhenStopped: true,
      volume: 0.25,
    });
    expect(readMediaState(pkg, SLIDE, parsed, pictures[1]!)?.settings).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.5,
    });
    expect(readMediaState(pkg, SLIDE, parsed, pictures[2]!)?.settings).toEqual({});
    expect(snapshot(pkg)).toEqual(before);
  });
});

function slideXml(...pictures: string[]): string {
  return '<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>' +
    pictures.join('') + '</p:spTree></p:cSld></p:sld>';
}

function mediaPicture(
  id: number,
  marker: 'audioFile' | 'videoFile',
  kindId: string,
  mediaId: string | undefined,
  posterId: string | undefined,
  name: string,
  altText?: string,
  includePlayback = true,
): string {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"${
    altText === undefined ? '' : ` descr="${altText}"`
  }/><p:nvPr><a:${marker} r:link="${kindId}"/><p:extLst>${
    mediaId
      ? `<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="p14" r:embed="${mediaId}"/></p:ext>`
      : ''
  }${includePlayback
    ? '<p:ext uri="{C13D3E4A-5148-4B6D-A7E7-505054582D4F}"><px:playback xmlns:px="urn:pptx-ooxml:media" play="auto" loop="1" hideWhenStopped="0" volume="50000"/></p:ext>'
    : ''}</p:extLst></p:nvPr></p:nvPicPr><p:blipFill>${
    posterId ? `<a:blip r:embed="${posterId}"/>` : ''
  }</p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:pic>`;
}

function addRelationship(
  pkg: OpcPackage,
  id: string,
  type: string,
  target: string,
  targetMode: 'Internal' | 'External' = 'Internal',
): void {
  pkg.addRelationship(SLIDE, { id, type, target, targetMode });
}

function snapshot(pkg: OpcPackage): unknown {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
      relationships,
    })),
    graph: pkg.graph,
    mutations: [...pkg.mutations],
  };
}
