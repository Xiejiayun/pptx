import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, LosslessXmlError } from '@pptx/lossless-xml';
import {
  assignPresentationSlideToSection,
  copyPresentationSlideSection,
  createPresentationSectionId,
  deletePresentationSection,
  insertPresentationSection,
  movePresentationSection,
  normalizeAddPresentationSectionOptions,
  normalizeAddPresentationSlideOptions,
  normalizePresentationSectionId,
  normalizePresentationSectionIndex,
  normalizePresentationSectionTitle,
  readPresentationSections,
  removePresentationSlideFromSections,
  renamePresentationSection,
  sortPresentationSectionSlides,
} from './presentation-sections.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const SECTION_NAMESPACE =
  'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTION_EXTENSION_URI =
  '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const A = '{00000000-0000-0000-0000-000000000001}';
const B = '{00000000-0000-0000-0000-000000000002}';
const C = '{00000000-0000-0000-0000-000000000003}';
const SLIDES = new Set([256, 257, 258, 259]);

function presentationXml(extension = '', prefix = 'p'): string {
  const q = prefix ? `${prefix}:` : '';
  const namespace = prefix
    ? `xmlns:${prefix}="${PRESENTATION_NAMESPACE}"`
    : `xmlns="${PRESENTATION_NAMESPACE}"`;
  return `<?xml version="1.0"?><${q}presentation ${namespace}><${q}sldIdLst><${q}sldId id="256"/><${q}sldId id="257"/><${q}sldId id="258"/><${q}sldId id="259"/></${q}sldIdLst>${extension}</${q}presentation>`;
}

function section(
  id: string,
  title: string,
  slideIds: readonly number[] = [],
  prefix = 'p14',
): string {
  const q = prefix ? `${prefix}:` : '';
  return `<${q}section name="${title}" id="${id}"><${q}sldIdLst>${slideIds
    .map((slideId) => `<${q}sldId id="${slideId}"/>`)
    .join('')}</${q}sldIdLst></${q}section>`;
}

function extension(
  sections: string,
  options: {
    readonly presentationPrefix?: string;
    readonly sectionPrefix?: string;
    readonly before?: string;
    readonly after?: string;
    readonly listAttributes?: string;
  } = {},
): string {
  const presentationPrefix = options.presentationPrefix ?? 'p';
  const sectionPrefix = options.sectionPrefix ?? 'p14';
  const p = presentationPrefix ? `${presentationPrefix}:` : '';
  const s = sectionPrefix ? `${sectionPrefix}:` : '';
  const declaration = sectionPrefix
    ? `xmlns:${sectionPrefix}="${SECTION_NAMESPACE}"`
    : `xmlns="${SECTION_NAMESPACE}"`;
  return `<${p}extLst>${options.before ?? ''}<${p}ext uri="${SECTION_EXTENSION_URI}"><${s}sectionLst ${declaration}${options.listAttributes ?? ''}>${sections}</${s}sectionLst></${p}ext>${options.after ?? ''}</${p}extLst>`;
}

function read(source: string, slideIds = SLIDES) {
  const xml = LosslessXmlDocument.parse(source);
  const snapshot = readPresentationSections(xml, slideIds);
  expect(xml.changed).toBe(false);
  return snapshot;
}

function expectRejectedWithoutPatch(
  source: string,
  operation: (xml: LosslessXmlDocument) => unknown,
): void {
  const xml = LosslessXmlDocument.parse(source);
  expect(() => operation(xml)).toThrow();
  expect(xml.changed).toBe(false);
  expect(xml.serialize()).toBe(source);
}

describe('presentation section codec', () => {
  it('reads detached canonical, empty, loose, duplicate-title, and alternate-prefix states', () => {
    const source = presentationXml(extension([
      section(A, 'Intro &amp; Start', [256]),
      section(B, 'Data', [258, 259]),
      section(C, 'Data'),
    ].join('')));
    const first = read(source)!;
    expect(first).toEqual([
      { id: A, title: 'Intro & Start', slideIds: [256] },
      { id: B, title: 'Data', slideIds: [258, 259] },
      { id: C, title: 'Data', slideIds: [] },
    ]);
    expect(read(presentationXml())).toEqual([]);

    const mutable = first as unknown as { title: string; slideIds: number[] }[];
    mutable[0]!.title = 'detached';
    mutable[0]!.slideIds.push(257);
    expect(read(source)?.[0]).toEqual({ id: A, title: 'Intro & Start', slideIds: [256] });

    const alternateSections = [
      section(A.toLowerCase(), 'Alternate', [256], 's'),
    ].join('');
    const alternate = presentationXml(
      extension(alternateSections, {
        presentationPrefix: 'q',
        sectionPrefix: 's',
        before: '<q:ext uri="{FOREIGN}"><x:sectionLst xmlns:x="urn:keep"><x:keep/></x:sectionLst></q:ext>',
        after: '<q:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></q:ext>',
      }),
      'q',
    );
    expect(read(alternate)).toEqual([
      { id: A.toLowerCase(), title: 'Alternate', slideIds: [256] },
    ]);
  });

  it('does not guess ambiguous, malformed, wrong-namespace, or dangling section state', () => {
    const validSection = section(A, 'Intro', [256]);
    const validExtension = extension(validSection);
    const cases = [
      `${presentationXml()}${presentationXml()}`,
      presentationXml(`<p:extLst/><p:extLst/>`),
      presentationXml(`<p:extLst><p:ext uri="${SECTION_EXTENSION_URI}"/><p:ext uri="${SECTION_EXTENSION_URI}"/></p:extLst>`),
      presentationXml(`<p:extLst><p:ext uri="{WRONG}"><p14:sectionLst xmlns:p14="${SECTION_NAMESPACE}"/></p:ext></p:extLst>`),
      presentationXml(`<p:extLst><p:ext uri="${SECTION_EXTENSION_URI}"><x:sectionLst xmlns:x="urn:wrong"/></p:ext></p:extLst>`),
      presentationXml(validExtension.replace('<p14:sectionLst ', '<p14:sectionLst/><p14:sectionLst ')),
      presentationXml(validExtension.replace(`id="${A}"`, 'id="bad"')),
      presentationXml(validExtension.replace(`id="${A}"`, `id="${A}" id="${B}"`)),
      presentationXml(validExtension.replace('<p14:sldIdLst>', '<p14:sldIdLst/><p14:sldIdLst>')),
      presentationXml(validExtension.replace('<p14:sldId id="256"/>', '<p14:sldId id="999"/>')),
      presentationXml(validExtension.replace('<p14:sldId id="256"/>', '<p14:sldId id="1.5"/>')),
      presentationXml(extension(`${validSection}${section(B, 'Other', [256])}`)),
      presentationXml(validExtension.replace('<p14:sldId id="256"/>', '<x:sldId xmlns:x="urn:wrong" id="256"/>')),
      presentationXml(validExtension.replace('name="Intro"', 'name="Intro" name="Again"')),
      presentationXml(validExtension.replace('<p14:sldId id="256"/>', '<p14:sldId id="256"><x:child xmlns:x="urn:x"/></p14:sldId>')),
    ];
    for (const source of cases) expect(read(source)).toBeUndefined();
    expect(() => LosslessXmlDocument.parse('<p:presentation>')).toThrow(LosslessXmlError);
  });

  it('creates, inserts, renames, and moves sections while preserving foreign extension state', () => {
    const absent = LosslessXmlDocument.parse(presentationXml());
    expect(insertPresentationSection(absent, SLIDES, 'Intro & <Start>', 0, A)).toEqual({
      id: A,
      title: 'Intro & <Start>',
      slideIds: [],
    });
    expect(absent.serialize()).toContain(
      `<p14:section name="Intro &amp; &lt;Start&gt;" id="${A}"><p14:sldIdLst/></p14:section>`,
    );
    expect(read(absent.serialize())?.[0]?.title).toBe('Intro & <Start>');

    const source = presentationXml(extension(
      `${section(A, 'A', [256])}<x:keep xmlns:x="urn:keep" value="FOREIGN"/>${section(C, 'C', [258])}`,
      {
        after: '<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>',
      },
    ));
    const xml = LosslessXmlDocument.parse(source);
    insertPresentationSection(xml, SLIDES, 'B', 1, B);
    expect(read(xml.serialize())?.map(({ title }) => title)).toEqual(['A', 'B', 'C']);
    expect(xml.serialize()).toContain('<x:keep xmlns:x="urn:keep" value="FOREIGN"/>');
    expect(xml.serialize()).toContain('<p15:sldGuideLst');

    const renamed = LosslessXmlDocument.parse(xml.serialize());
    expect(renamePresentationSection(renamed, SLIDES, B, 'B & <Two> "Q"')).toBe(true);
    expect(renamed.serialize()).toContain('name="B &amp; &lt;Two&gt; &quot;Q&quot;"');
    const moved = LosslessXmlDocument.parse(renamed.serialize());
    expect(movePresentationSection(moved, SLIDES, C, 0)).toBe(true);
    expect(read(moved.serialize())?.map(({ title }) => title)).toEqual([
      'C',
      'A',
      'B & <Two> "Q"',
    ]);
  });

  it('assigns, removes, copies, sorts, and deletes memberships with exact no-op behavior', () => {
    const source = presentationXml(extension([
      section(A, 'A', [258, 256]),
      section(B, 'B', [257]),
      section(C, 'C'),
    ].join('')));
    const assigned = LosslessXmlDocument.parse(source);
    expect(assignPresentationSlideToSection(assigned, SLIDES, 256, B)).toBe(true);
    expect(read(assigned.serialize())?.map(({ slideIds }) => slideIds)).toEqual([
      [258],
      [257, 256],
      [],
    ]);

    const copied = LosslessXmlDocument.parse(assigned.serialize());
    expect(copyPresentationSlideSection(copied, SLIDES, 257, 259)).toBe(true);
    expect(read(copied.serialize())?.[1]?.slideIds).toEqual([257, 256, 259]);
    const removed = LosslessXmlDocument.parse(copied.serialize());
    expect(removePresentationSlideFromSections(removed, SLIDES, 256)).toBe(true);
    expect(read(removed.serialize())?.[1]?.slideIds).toEqual([257, 259]);
    const sorted = LosslessXmlDocument.parse(removed.serialize());
    expect(sortPresentationSectionSlides(sorted, SLIDES, [259, 258, 257, 256])).toBe(true);
    expect(read(sorted.serialize())?.map(({ slideIds }) => slideIds)).toEqual([
      [258],
      [259, 257],
      [],
    ]);
    const deleted = LosslessXmlDocument.parse(sorted.serialize());
    expect(deletePresentationSection(deleted, SLIDES, A)).toBe(true);
    expect(read(deleted.serialize())?.map(({ id }) => id)).toEqual([B, C]);

    const same = LosslessXmlDocument.parse(source);
    expect(renamePresentationSection(same, SLIDES, A, 'A')).toBe(false);
    expect(movePresentationSection(same, SLIDES, A, 0)).toBe(false);
    expect(assignPresentationSlideToSection(same, SLIDES, 256, A)).toBe(false);
    expect(removePresentationSlideFromSections(same, SLIDES, 259)).toBe(false);
    expect(same.changed).toBe(false);
    expect(same.serialize()).toBe(source);
  });

  it('removes only clean empty section containers and preserves foreign siblings', () => {
    const clean = LosslessXmlDocument.parse(presentationXml(extension(section(A, 'A'))));
    expect(deletePresentationSection(clean, SLIDES, A)).toBe(true);
    expect(clean.serialize()).not.toContain('sectionLst');
    expect(clean.serialize()).not.toContain('extLst');
    expect(read(clean.serialize())).toEqual([]);

    const foreign = LosslessXmlDocument.parse(presentationXml(extension(
      `${section(A, 'A')}<x:keep xmlns:x="urn:keep"/>`,
      { after: '<p:ext uri="{FOREIGN}"/>' },
    )));
    expect(deletePresentationSection(foreign, SLIDES, A)).toBe(true);
    expect(foreign.serialize()).toContain('<p14:sectionLst');
    expect(foreign.serialize()).toContain('<x:keep xmlns:x="urn:keep"/>');
    expect(foreign.serialize()).toContain('<p:ext uri="{FOREIGN}"/>');
    expect(read(foreign.serialize())).toEqual([]);
  });

  it('normalizes strict data objects, titles, IDs, indices, and secure IDs', () => {
    const frozen = Object.freeze({ title: 'Frozen', order: 0 });
    expect(normalizeAddPresentationSectionOptions(frozen)).toEqual(frozen);
    const nullPrototype = Object.create(null) as { title?: string };
    Object.defineProperty(nullPrototype, 'title', { value: 'Null', enumerable: true });
    expect(normalizeAddPresentationSectionOptions(nullPrototype)).toEqual({ title: 'Null' });
    expect(normalizeAddPresentationSlideOptions(undefined)).toEqual({});
    expect(normalizeAddPresentationSlideOptions({ sectionTitle: 'Intro' })).toEqual({
      sectionTitle: 'Intro',
    });
    expect(normalizePresentationSectionTitle(' A ')).toBe(' A ');
    expect(normalizePresentationSectionId(A.toLowerCase())).toBe(A.toLowerCase());
    expect(normalizePresentationSectionIndex(1, 'index', 2, true)).toBe(1);
    expect(createPresentationSectionId()).toMatch(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/);

    const accessor = Object.create(null) as Record<string, unknown>;
    let calls = 0;
    Object.defineProperty(accessor, 'title', {
      enumerable: true,
      get() {
        calls += 1;
        return 'Never';
      },
    });
    const custom = Object.create({ inherited: true }) as { title?: string };
    custom.title = 'Custom';
    const invalidOptions = [
      null,
      [],
      'Intro',
      {},
      { title: '' },
      { title: '   ' },
      { title: 'bad\u0000title' },
      { title: 1 },
      { title: 'Intro', order: 1.5 },
      { title: 'Intro', extra: true },
      { title: 'Intro', [Symbol('extra')]: true },
      accessor,
      custom,
    ];
    for (const value of invalidOptions) {
      expect(() => normalizeAddPresentationSectionOptions(value)).toThrow(TypeError);
    }
    expect(calls).toBe(0);
    for (const value of ['', '   ', 'bad\u0000title', null, 1]) {
      expect(() => normalizePresentationSectionTitle(value)).toThrow(TypeError);
    }
    for (const value of ['bad', A.slice(1), null, 1]) {
      expect(() => normalizePresentationSectionId(value)).toThrow(TypeError);
    }
    for (const value of [-1, 3, 1.5, Number.NaN, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => normalizePresentationSectionIndex(value, 'index', 2, true)).toThrow();
    }
  });

  it('rejects unsafe mutations before adding any patch', () => {
    const source = presentationXml(extension(`${section(A, 'A', [256])}${section(B, 'B', [256])}`));
    expectRejectedWithoutPatch(source, (xml) =>
      renamePresentationSection(xml, SLIDES, A, 'Changed'));
    const valid = presentationXml(extension(section(A, 'A', [256])));
    expectRejectedWithoutPatch(valid, (xml) =>
      insertPresentationSection(xml, SLIDES, 'B', 2, B));
    expectRejectedWithoutPatch(valid, (xml) =>
      assignPresentationSlideToSection(xml, SLIDES, 999, A));
    expectRejectedWithoutPatch(valid, (xml) =>
      assignPresentationSlideToSection(xml, SLIDES, 257, B));
    expectRejectedWithoutPatch(valid, (xml) =>
      sortPresentationSectionSlides(xml, SLIDES, [256, 257]));
  });
});
