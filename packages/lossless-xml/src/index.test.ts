import { describe, expect, it } from 'vitest';
import { canonicalizeXml, LosslessXmlDocument, LosslessXmlError } from './index.js';

describe('LosslessXmlDocument', () => {
  it('returns the exact source when unchanged', () => {
    const source = '<?xml version="1.0"?><x:root b="2" a="1">\n  <!--keep--><x:t>A&amp;B</x:t><u:opaque q="x"/>\n</x:root>';
    expect(LosslessXmlDocument.parse(source).serialize()).toBe(source);
  });

  it('patches only the requested source span', () => {
    const source = '<p:sld><a:t>Old</a:t><x:unknown keep="yes"/></p:sld>';
    const document = LosslessXmlDocument.parse(source);
    document.replaceText(document.elements('t')[0]!, 'New & safe');
    expect(document.serialize()).toBe('<p:sld><a:t>New &amp; safe</a:t><x:unknown keep="yes"/></p:sld>');
  });

  it('rejects DTDs and overlapping patches', () => {
    expect(() => LosslessXmlDocument.parse('<!DOCTYPE x><x/>')).toThrow(LosslessXmlError);
    const document = LosslessXmlDocument.parse('<x><a>one</a><b>two</b></x>');
    document.replace(3, 13, '<a>changed</a>');
    expect(() => document.replace(6, 9, 'x')).toThrow(/Overlapping/);
  });

  it('supports structural patches and deterministic canonical output', () => {
    const document = LosslessXmlDocument.parse('<x b="2" a="1">\n<a>one</a>\n</x>');
    const root = document.elements('x')[0]!;
    document.removeElement(document.elements('a')[0]!);
    document.appendChildXml(root, '<b>two</b>');
    expect(document.serialize()).toBe('<x b="2" a="1">\n\n<b>two</b></x>');
    expect(canonicalizeXml('<x b="2" a="1">\n<a>one</a>\n</x>')).toBe('<x a="1" b="2"><a>one</a></x>');
  });

  it('expands self-closing elements when appending children', () => {
    const document = LosslessXmlDocument.parse('<x><empty a="1" /></x>');
    document.appendChildXml(document.elements('empty')[0]!, '<child/>');
    expect(document.serialize()).toBe('<x><empty a="1"><child/></empty></x>');
  });
});
