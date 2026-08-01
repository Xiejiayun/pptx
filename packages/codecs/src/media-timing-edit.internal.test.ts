import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import type { NormalizedMediaPlaybackSettings } from './media-edit.internal.js';
import {
  clearNativeMediaTiming,
  syncNativeMediaTiming,
} from './media-timing-edit.internal.js';
import { readNativeMediaTiming } from './media-timing-state.internal.js';

const DEFAULTS = Object.freeze({
  play: 'click',
  loop: false,
  hideWhenStopped: false,
  volume: 1,
}) satisfies Readonly<NormalizedMediaPlaybackSettings>;

const CUSTOM = Object.freeze({
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.25,
}) satisfies Readonly<NormalizedMediaPlaybackSettings>;

describe('syncNativeMediaTiming', () => {
  it.each([
    ['audio', 7],
    ['video', 12],
  ] as const)('creates exact canonical %s timing in slide schema order', (kind, lastId) => {
    const xml = LosslessXmlDocument.parse(slideXml(kind, '<p:transition/><p:extLst keep="1"/>'));

    const result = syncNativeMediaTiming(xml, 2, kind, CUSTOM);
    const rendered = xml.serialize();

    expect(result.changed).toBe(true);
    expect(result.ownership).toMatchObject({ version: 1, mediaTnId: 7, playTnId: 5 });
    expect(result.ownership.pauseTnId).toBe(kind === 'video' ? 11 : undefined);
    expect(rendered.indexOf('<p:transition/>')).toBeLessThan(rendered.indexOf('<p:timing>'));
    expect(rendered.indexOf('<p:timing>')).toBeLessThan(rendered.indexOf('<p:extLst keep="1"/>'));
    expect(rendered).toContain('repeatCount="indefinite"');
    expect(rendered).toContain('showWhenStopped="0"');
    expect(rendered).toContain('vol="25000"');
    expect(rendered).toContain(`id="${lastId}"`);
    expect(readNativeMediaTiming(
      LosslessXmlDocument.parse(rendered),
      2,
      kind,
      result.ownership,
    )).toMatchObject({ status: 'owned-healthy', settings: CUSTOM });
  });

  it('appends to a valid imported root without changing ordinary animation bytes', () => {
    const ordinary = '<p:par><p:cTn id="9" fill="hold"><p:childTnLst>'
      + '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="42" dur="500"/>'
      + '<p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:animEffect>'
      + '</p:childTnLst></p:cTn></p:par>';
    const xml = LosslessXmlDocument.parse(slideXml('video', timingRoot(ordinary)));

    const result = syncNativeMediaTiming(xml, 2, 'video', DEFAULTS);
    const rendered = xml.serialize();

    expect(result.changed).toBe(true);
    expect(rendered).toContain(ordinary);
    expect(result.ownership.mediaTnId).toBeGreaterThan(42);
    expect(readNativeMediaTiming(
      LosslessXmlDocument.parse(rendered),
      2,
      'video',
      result.ownership,
    )).toMatchObject({ status: 'owned-healthy', settings: DEFAULTS });
  });

  it('adopts healthy imports, repairs stale ownership, and no-ops verified state', () => {
    const created = LosslessXmlDocument.parse(slideXml('video'));
    const initial = syncNativeMediaTiming(created, 2, 'video', DEFAULTS);
    const source = created.serialize();

    const imported = LosslessXmlDocument.parse(source);
    expect(syncNativeMediaTiming(imported, 2, 'video', DEFAULTS)).toEqual({
      changed: false,
      ownership: initial.ownership,
    });
    expect(imported.serialize()).toBe(source);

    const stale = LosslessXmlDocument.parse(source);
    expect(syncNativeMediaTiming(stale, 2, 'video', DEFAULTS, {
      ...initial.ownership,
      playTnId: 99,
    })).toEqual({ changed: false, ownership: initial.ownership });
    expect(stale.serialize()).toBe(source);

    const healthy = LosslessXmlDocument.parse(source);
    expect(syncNativeMediaTiming(healthy, 2, 'video', DEFAULTS, initial.ownership)).toEqual({
      changed: false,
      ownership: initial.ownership,
    });
    expect(healthy.serialize()).toBe(source);
  });

  it('replaces only recognized media branches when settings change', () => {
    const initialXml = LosslessXmlDocument.parse(slideXml('audio'));
    const initial = syncNativeMediaTiming(initialXml, 2, 'audio', DEFAULTS);
    const source = initialXml.serialize();
    const xml = LosslessXmlDocument.parse(source);

    const updated = syncNativeMediaTiming(xml, 2, 'audio', CUSTOM, initial.ownership);
    const rendered = xml.serialize();

    expect(updated.changed).toBe(true);
    expect(updated.ownership).not.toEqual(initial.ownership);
    expect(readNativeMediaTiming(
      LosslessXmlDocument.parse(rendered),
      2,
      'audio',
      updated.ownership,
    )).toMatchObject({ status: 'owned-healthy', settings: CUSTOM });
  });

  it.each([
    ['unsupported', (source: string) => source.replace('repeatCount="indefinite"', 'repeatCount="2000"')],
    ['ambiguous', (source: string) => source.replace('</p:childTnLst></p:cTn></p:par></p:tnLst>',
      '<p:video><p:cMediaNode><p:cTn id="13"><p:stCondLst><p:cond delay="indefinite"/>'
      + '</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>'
      + '</p:cMediaNode></p:video></p:childTnLst></p:cTn></p:par></p:tnLst>')],
  ] as const)('rejects %s target timing without mutation', (_label, mutate) => {
    const created = LosslessXmlDocument.parse(slideXml('video'));
    syncNativeMediaTiming(created, 2, 'video', CUSTOM);
    const xml = LosslessXmlDocument.parse(mutate(created.serialize()));
    const before = xml.serialize();

    expect(() => syncNativeMediaTiming(xml, 2, 'video', DEFAULTS)).toThrow(/timing|media/i);
    expect(xml.serialize()).toBe(before);
    expect(xml.changed).toBe(false);
  });
});

describe('clearNativeMediaTiming', () => {
  it('clears a complete imported graph but preserves its timing containers', () => {
    const created = LosslessXmlDocument.parse(slideXml('video'));
    const result = syncNativeMediaTiming(created, 2, 'video', DEFAULTS);
    const importedSource = created.serialize().replace('<p:timing>', '<p:timing imported="1">');
    const imported = LosslessXmlDocument.parse(importedSource);

    expect(clearNativeMediaTiming(imported, 2, 'video', result.ownership)).toBe(true);
    const rendered = imported.serialize();
    expect(rendered).toContain('<p:timing imported="1">');
    expect(readNativeMediaTiming(LosslessXmlDocument.parse(rendered), 2, 'video')).toEqual({
      status: 'absent',
    });
  });

  it('removes the last exact library-created timing root', () => {
    const created = LosslessXmlDocument.parse(slideXml('audio'));
    const result = syncNativeMediaTiming(created, 2, 'audio', DEFAULTS);
    const xml = LosslessXmlDocument.parse(created.serialize());

    expect(clearNativeMediaTiming(xml, 2, 'audio', result.ownership)).toBe(true);
    expect(xml.serialize()).not.toContain('<p:timing>');
  });

  it('preserves imported empty containers and rejects unsafe clears without mutation', () => {
    const emptySource = slideXml('video', timingRoot(''));
    const empty = LosslessXmlDocument.parse(emptySource);
    expect(clearNativeMediaTiming(empty, 2, 'video', {
      version: 1,
      mediaTnId: 7,
      playTnId: 5,
      pauseTnId: 11,
    })).toBe(false);
    expect(empty.serialize()).toBe(emptySource);

    const created = LosslessXmlDocument.parse(slideXml('video'));
    syncNativeMediaTiming(created, 2, 'video', CUSTOM);
    const unsafe = LosslessXmlDocument.parse(
      created.serialize().replace('repeatCount="indefinite"', 'repeatCount="2000"'),
    );
    const before = unsafe.serialize();
    expect(() => clearNativeMediaTiming(unsafe, 2, 'video')).toThrow(/timing|media/i);
    expect(unsafe.serialize()).toBe(before);
    expect(unsafe.changed).toBe(false);
  });
});

function slideXml(kind: 'audio' | 'video', tail = ''): string {
  return '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>'
    + `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Media"/><p:nvPr><a:${kind}File/>`
    + '</p:nvPr></p:nvPicPr><p:spPr/></p:pic>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Other"/></p:nvSpPr></p:sp>'
    + '</p:spTree></p:cSld>' + tail + '</p:sld>';
}

function timingRoot(children: string): string {
  return '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
    + `nodeType="tmRoot"><p:childTnLst>${children}</p:childTnLst>`
    + '</p:cTn></p:par></p:tnLst></p:timing>';
}
