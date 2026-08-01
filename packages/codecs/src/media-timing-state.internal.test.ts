import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import {
  allocateNativeTimingIds,
  readNativeMediaTiming,
  type NativeMediaTimingOwnership,
} from './media-timing-state.internal.js';

describe('readNativeMediaTiming', () => {
  it('reads exact PowerPoint audio and video graphs without mutation', () => {
    const audioXml = LosslessXmlDocument.parse(powerPointAudioSlide());
    const videoXml = LosslessXmlDocument.parse(powerPointVideoSlide());
    const audioBefore = audioXml.serialize();
    const videoBefore = videoXml.serialize();

    const audio = readNativeMediaTiming(audioXml, 4, 'audio');
    const video = readNativeMediaTiming(videoXml, 2, 'video');

    expect(audio).toEqual({
      status: 'recognized-imported',
      settings: {
        play: 'click',
        loop: false,
        hideWhenStopped: false,
        volume: 0.8,
      },
      ownership: { version: 1, mediaTnId: 7, playTnId: 5 },
    });
    expect(video).toEqual({
      status: 'recognized-imported',
      settings: {
        play: 'click',
        loop: false,
        hideWhenStopped: false,
        volume: 0.8,
      },
      ownership: { version: 1, mediaTnId: 7, playTnId: 5, pauseTnId: 11 },
    });
    expect(Object.isFrozen(audio)).toBe(true);
    expect(Object.isFrozen(audio.settings)).toBe(true);
    expect(Object.isFrozen(audio.ownership)).toBe(true);
    expect(audioXml.serialize()).toBe(audioBefore);
    expect(videoXml.serialize()).toBe(videoBefore);
  });

  it('decodes automatic start, indefinite loop, mute, and stopped visibility', () => {
    const xml = LosslessXmlDocument.parse(powerPointVideoSlide({
      play: 'auto',
      repeatCount: 'indefinite',
      mute: '1',
      showWhenStopped: '0',
      volume: '25000',
    }));

    expect(readNativeMediaTiming(xml, 2, 'video')).toMatchObject({
      status: 'recognized-imported',
      settings: {
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0,
      },
    });
  });

  it('distinguishes healthy, stale-id, and missing owned graphs', () => {
    const xml = LosslessXmlDocument.parse(powerPointVideoSlide());
    const ownership: NativeMediaTimingOwnership = {
      version: 1,
      mediaTnId: 7,
      playTnId: 5,
      pauseTnId: 11,
    };

    expect(readNativeMediaTiming(xml, 2, 'video', ownership)).toMatchObject({
      status: 'owned-healthy',
      ownership,
    });
    expect(readNativeMediaTiming(xml, 2, 'video', {
      ...ownership,
      playTnId: 99,
    })).toMatchObject({
      status: 'owned-stale',
      ownership,
    });
    expect(readNativeMediaTiming(
      LosslessXmlDocument.parse(slideXml(mediaPicture(2, 'video'))),
      2,
      'video',
      ownership,
    )).toMatchObject({
      status: 'owned-stale',
      reason: expect.stringMatching(/missing/i),
    });
  });

  it('returns absent for slides without a media timing graph', () => {
    const empty = LosslessXmlDocument.parse(slideXml(mediaPicture(2, 'video')));
    const animated = LosslessXmlDocument.parse(slideXml(
      mediaPicture(2, 'video'),
      timingRoot(
        '<p:par><p:cTn id="2" fill="hold"><p:childTnLst>' +
          '<p:animEffect transition="in" filter="fade"><p:cBhvr>' +
          '<p:cTn id="3" dur="500"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>' +
          '</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>',
      ),
    ));

    expect(readNativeMediaTiming(empty, 2, 'video')).toEqual({ status: 'absent' });
    expect(readNativeMediaTiming(animated, 2, 'video')).toEqual({ status: 'absent' });
  });

  it.each([
    ['kind mismatch', powerPointAudioSlide({ shapeId: 2 }), 2, 'video'],
    ['cross-slide playback', powerPointVideoSlide({ numSld: '2' }), 2, 'video'],
    ['finite repeat', powerPointVideoSlide({ repeatCount: '2000' }), 2, 'video'],
    ['invalid volume', powerPointVideoSlide({ volume: '100001' }), 2, 'video'],
    ['missing play command', powerPointVideoSlide({ omitPlayCommand: true }), 2, 'video'],
    [
      'missing play start condition',
      powerPointVideoSlide().replace(
        '<p:cTn id="3" fill="hold"><p:stCondLst>' +
          '<p:cond delay="indefinite"/></p:stCondLst>',
        '<p:cTn id="3" fill="hold">',
      ),
      2,
      'video',
    ],
    ['repeated target', powerPointVideoSlide({ repeatedTarget: true }), 2, 'video'],
    ['sound target', powerPointVideoSlide({ soundTarget: true }), 2, 'video'],
    ['dangling target', powerPointVideoSlide({ shapeId: 99, omitPicture: true }), 99, 'video'],
  ] as const)('classifies unsupported %s without mutation', (_label, source, shapeId, kind) => {
    const xml = LosslessXmlDocument.parse(source);
    const before = xml.serialize();

    expect(readNativeMediaTiming(xml, shapeId, kind)).toMatchObject({
      status: 'unsupported',
      reason: expect.any(String),
    });
    expect(xml.serialize()).toBe(before);
  });

  it('classifies repeated candidates, play commands, and timing ids as ambiguous', () => {
    const repeatedMedia = LosslessXmlDocument.parse(powerPointVideoSlide({ repeatedMedia: true }));
    const repeatedPlay = LosslessXmlDocument.parse(powerPointVideoSlide({ repeatedPlay: true }));
    const repeatedId = LosslessXmlDocument.parse(
      powerPointVideoSlide().replace('<p:cTn id="12"', '<p:cTn id="11"'),
    );

    for (const xml of [repeatedMedia, repeatedPlay, repeatedId]) {
      expect(readNativeMediaTiming(xml, 2, 'video')).toMatchObject({
        status: 'ambiguous',
        reason: expect.any(String),
      });
    }
  });

  it('classifies repeated timing roots and direct lists without throwing', () => {
    const source = powerPointVideoSlide();
    const repeatedTiming = LosslessXmlDocument.parse(
      source.replace('</p:sld>', timingRoot('') + '</p:sld>'),
    );
    const repeatedNodeList = LosslessXmlDocument.parse(
      source.replace('</p:timing>', '<p:tnLst/></p:timing>'),
    );
    const repeatedChildList = LosslessXmlDocument.parse(
      source.replace(
        '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>',
        '</p:childTnLst><p:childTnLst/></p:cTn></p:par></p:tnLst></p:timing>',
      ),
    );

    for (const xml of [repeatedTiming, repeatedNodeList, repeatedChildList]) {
      expect(() => readNativeMediaTiming(xml, 2, 'video')).not.toThrow();
      expect(readNativeMediaTiming(xml, 2, 'video')).toMatchObject({
        status: 'ambiguous',
        reason: expect.any(String),
      });
    }
  });
});

describe('allocateNativeTimingIds', () => {
  it('allocates frozen deterministic ids above an empty or sparse tree', () => {
    const empty = LosslessXmlDocument.parse(slideXml(mediaPicture(2, 'video')));
    const sparse = LosslessXmlDocument.parse(slideXml(
      mediaPicture(2, 'video'),
      timingRoot('<p:par><p:cTn id="9"/><p:cTn id="42"/></p:par>'),
    ));

    const first = allocateNativeTimingIds(empty, 3);
    const next = allocateNativeTimingIds(sparse, 2);
    expect(first).toEqual([1, 2, 3]);
    expect(next).toEqual([43, 44]);
    expect(Object.isFrozen(first)).toBe(true);
    expect(empty.changed).toBe(false);
    expect(sparse.changed).toBe(false);
  });

  it.each([-1, 0.5, Number.NaN, Number.POSITIVE_INFINITY])(
    'rejects invalid allocation count %s',
    (count) => {
      const xml = LosslessXmlDocument.parse(slideXml());
      expect(() => allocateNativeTimingIds(xml, count)).toThrow(/count|integer/i);
      expect(xml.changed).toBe(false);
    },
  );

  it.each([
    ['zero', '<p:cTn id="0"/>'],
    ['negative', '<p:cTn id="-1"/>'],
    ['decimal', '<p:cTn id="1.5"/>'],
    ['overflow', '<p:cTn id="4294967296"/>'],
    ['repeated attribute', '<p:cTn id="2" id="3"/>'],
    ['duplicate', '<p:cTn id="2"/><p:cTn id="2"/>'],
  ])('rejects %s timing ids', (_label, nodes) => {
    const xml = LosslessXmlDocument.parse(slideXml(undefined, timingRoot(nodes)));
    expect(() => allocateNativeTimingIds(xml, 1)).toThrow(/timing|id|duplicate/i);
    expect(xml.changed).toBe(false);
  });

  it('rejects unsigned 32-bit exhaustion', () => {
    const xml = LosslessXmlDocument.parse(slideXml(
      undefined,
      timingRoot('<p:cTn id="4294967295"/>'),
    ));

    expect(() => allocateNativeTimingIds(xml, 1)).toThrow(/exhausted/i);
    expect(xml.changed).toBe(false);
  });
});

interface PowerPointFixtureOptions {
  readonly shapeId?: number;
  readonly play?: 'click' | 'auto';
  readonly repeatCount?: string;
  readonly mute?: string;
  readonly showWhenStopped?: string;
  readonly volume?: string;
  readonly numSld?: string;
  readonly omitPicture?: boolean;
  readonly omitPlayCommand?: boolean;
  readonly repeatedTarget?: boolean;
  readonly soundTarget?: boolean;
  readonly repeatedMedia?: boolean;
  readonly repeatedPlay?: boolean;
}

function powerPointAudioSlide(options: PowerPointFixtureOptions = {}): string {
  const shapeId = options.shapeId ?? 4;
  const children = [
    ...(options.omitPlayCommand ? [] : [mainSequence(shapeId, options.play ?? 'click')]),
    mediaNode('audio', shapeId, 7, options),
    ...(options.repeatedMedia ? [mediaNode('audio', shapeId, 13, options)] : []),
  ];
  if (options.repeatedPlay) children.push(mainSequence(shapeId, options.play ?? 'click', 13));
  return slideXml(
    options.omitPicture ? undefined : mediaPicture(shapeId, 'audio'),
    timingRoot(children.join('')),
  );
}

function powerPointVideoSlide(options: PowerPointFixtureOptions = {}): string {
  const shapeId = options.shapeId ?? 2;
  const children = [
    ...(options.omitPlayCommand ? [] : [mainSequence(shapeId, options.play ?? 'click')]),
    mediaNode('video', shapeId, 7, options),
    ...(options.repeatedMedia ? [mediaNode('video', shapeId, 13, options)] : []),
    interactivePause(shapeId),
  ];
  if (options.repeatedPlay) children.push(mainSequence(shapeId, options.play ?? 'click', 13));
  return slideXml(
    options.omitPicture ? undefined : mediaPicture(shapeId, 'video'),
    timingRoot(children.join('')),
  );
}

function slideXml(picture?: string, timing = ''): string {
  return '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
    (picture ?? '') + '</p:spTree></p:cSld>' + timing + '</p:sld>';
}

function mediaPicture(shapeId: number, kind: 'audio' | 'video'): string {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="Media"/>` +
    `<p:nvPr><a:${kind}File/></p:nvPr></p:nvPicPr><p:spPr/></p:pic>`;
}

function timingRoot(children: string): string {
  return '<p:timing><p:tnLst><p:par>' +
    '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">' +
    `<p:childTnLst>${children}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

function mainSequence(
  shapeId: number,
  play: 'click' | 'auto',
  firstId = 2,
): string {
  const effectId = firstId + 3;
  const behaviorId = firstId + 4;
  return '<p:seq concurrent="1" nextAc="seek">' +
    `<p:cTn id="${firstId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    `<p:par><p:cTn id="${firstId + 1}" fill="hold"><p:stCondLst>` +
    `<p:cond delay="${play === 'click' ? 'indefinite' : '0'}"/></p:stCondLst>` +
    `<p:childTnLst><p:par><p:cTn id="${firstId + 2}" fill="hold"><p:stCondLst>` +
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>' +
    `<p:cTn id="${effectId}" presetID="1" presetClass="mediacall" ` +
    'presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst>' +
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    '<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>' +
    `<p:cTn id="${behaviorId}" dur="3000" fill="hold"/>` +
    `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>` +
    '</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst>' +
    '</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>' +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>' +
    '</p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst>' +
    '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>' +
    '</p:tgtEl></p:cond></p:nextCondLst></p:seq>';
}

function mediaNode(
  kind: 'audio' | 'video',
  shapeId: number,
  id: number,
  options: PowerPointFixtureOptions,
): string {
  const attributes = [
    `vol="${options.volume ?? '80000'}"`,
    ...(options.mute === undefined ? [] : [`mute="${options.mute}"`]),
    ...(options.showWhenStopped === undefined
      ? []
      : [`showWhenStopped="${options.showWhenStopped}"`]),
    ...(options.numSld === undefined ? [] : [`numSld="${options.numSld}"`]),
  ].join(' ');
  const repeat = options.repeatCount === undefined
    ? ''
    : ` repeatCount="${options.repeatCount}"`;
  const end = kind === 'audio'
    ? '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl>' +
      '<p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>'
    : '';
  const target = options.soundTarget
    ? `<p:tgtEl><p:sndTgt/><p:spTgt spid="${shapeId}"/></p:tgtEl>`
    : `<p:tgtEl><p:spTgt spid="${shapeId}"/>${
      options.repeatedTarget ? `<p:spTgt spid="${shapeId}"/>` : ''
    }</p:tgtEl>`;
  return `<p:${kind}><p:cMediaNode ${attributes}>` +
    `<p:cTn id="${id}" fill="hold" display="0"${repeat}>` +
    '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' + end +
    `</p:cTn>${target}</p:cMediaNode></p:${kind}>`;
}

function interactivePause(shapeId: number): string {
  return '<p:seq concurrent="1" nextAc="seek"><p:cTn id="8" restart="whenNotActive" ' +
    'fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst>' +
    `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${shapeId}"/>` +
    '</p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0">' +
    '<p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="9" fill="hold">' +
    '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>' +
    '<p:cTn id="10" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
    '<p:childTnLst><p:par><p:cTn id="11" presetID="2" presetClass="mediacall" ' +
    'presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst>' +
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    '<p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="12" dur="1" fill="hold"/>' +
    `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>` +
    '</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst>' +
    '</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>' +
    `<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl>` +
    `<p:spTgt spid="${shapeId}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
}
