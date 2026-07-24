export interface XmlAttribute {
  readonly name: string;
  readonly localName: string;
  readonly value: string;
  readonly start: number;
  readonly end: number;
  readonly valueStart: number;
  readonly valueEnd: number;
}

export interface XmlText {
  readonly type: 'text';
  readonly start: number;
  readonly end: number;
  readonly value: string;
  readonly parent: XmlElement;
}

export interface XmlElement {
  readonly type: 'element';
  readonly name: string;
  readonly localName: string;
  readonly start: number;
  readonly startTagEnd: number;
  endTagStart: number;
  end: number;
  readonly selfClosing: boolean;
  readonly attributes: readonly XmlAttribute[];
  readonly children: XmlNode[];
  readonly parent?: XmlElement;
}

export type XmlNode = XmlElement | XmlText;

export class LosslessXmlError extends Error {
  constructor(message: string, readonly offset?: number) {
    super(offset === undefined ? message : `${message} at offset ${offset}`);
    this.name = 'LosslessXmlError';
  }
}

interface XmlPatch {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export class LosslessXmlDocument {
  readonly roots: readonly XmlElement[];
  readonly declaration: string | undefined;
  readonly #patches: XmlPatch[] = [];

  private constructor(readonly source: string, roots: XmlElement[], declaration?: string) {
    this.roots = roots;
    this.declaration = declaration;
  }

  static parse(source: string | Uint8Array): LosslessXmlDocument {
    const text = typeof source === 'string' ? source : new TextDecoder().decode(source);
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) {
      throw new LosslessXmlError('DTD and entity declarations are not allowed');
    }

    const roots: XmlElement[] = [];
    const stack: XmlElement[] = [];
    let declaration: string | undefined;
    let cursor = 0;

    while (cursor < text.length) {
      const open = text.indexOf('<', cursor);
      if (open < 0) {
        appendText(text, cursor, text.length, stack.at(-1));
        break;
      }
      appendText(text, cursor, open, stack.at(-1));

      if (text.startsWith('<!--', open)) {
        cursor = requireTerminator(text, open, '-->') + 3;
        continue;
      }
      if (text.startsWith('<![CDATA[', open)) {
        cursor = requireTerminator(text, open, ']]>') + 3;
        continue;
      }
      if (text.startsWith('<?', open)) {
        const end = requireTerminator(text, open, '?>') + 2;
        if (open === 0 && /^<\?xml\b/i.test(text.slice(open, end))) {
          declaration = text.slice(open, end);
        }
        cursor = end;
        continue;
      }
      if (text.startsWith('</', open)) {
        const close = findTagEnd(text, open);
        const name = text.slice(open + 2, close).trim();
        const element = stack.pop();
        if (!element || element.name !== name) {
          throw new LosslessXmlError(`Unexpected closing tag ${name}`, open);
        }
        element.endTagStart = open;
        element.end = close + 1;
        cursor = close + 1;
        continue;
      }
      if (text.startsWith('<!', open)) {
        throw new LosslessXmlError('Unsupported declaration', open);
      }

      const close = findTagEnd(text, open);
      const raw = text.slice(open + 1, close);
      const selfClosing = /\/\s*$/.test(raw);
      const body = selfClosing ? raw.replace(/\/\s*$/, '') : raw;
      const nameMatch = /^\s*([^\s/>]+)/.exec(body);
      if (!nameMatch?.[1]) throw new LosslessXmlError('Missing element name', open);
      const name = nameMatch[1];
      const parent = stack.at(-1);
      const element: XmlElement = {
        type: 'element',
        name,
        localName: localName(name),
        start: open,
        startTagEnd: close + 1,
        endTagStart: selfClosing ? close : -1,
        end: selfClosing ? close + 1 : -1,
        selfClosing,
        attributes: parseAttributes(text, open + 1 + (nameMatch.index ?? 0) + nameMatch[0].length, close),
        children: [],
        ...(parent ? { parent } : {}),
      };
      if (parent) parent.children.push(element);
      else roots.push(element);
      if (!selfClosing) stack.push(element);
      cursor = close + 1;
    }

    if (stack.length > 0) {
      throw new LosslessXmlError(`Unclosed element ${stack.at(-1)?.name}`);
    }
    return new LosslessXmlDocument(text, roots, declaration);
  }

  elements(localOrQualifiedName?: string): XmlElement[] {
    const result: XmlElement[] = [];
    const visit = (element: XmlElement): void => {
      if (!localOrQualifiedName || element.name === localOrQualifiedName || element.localName === localOrQualifiedName) {
        result.push(element);
      }
      for (const child of element.children) if (child.type === 'element') visit(child);
    };
    for (const root of this.roots) visit(root);
    return result;
  }

  attribute(element: XmlElement, localOrQualifiedName: string): XmlAttribute | undefined {
    return element.attributes.find(
      (attribute) => attribute.name === localOrQualifiedName || attribute.localName === localOrQualifiedName,
    );
  }

  descendants(element: XmlElement, localOrQualifiedName?: string): XmlElement[] {
    const result: XmlElement[] = [];
    const visit = (node: XmlElement): void => {
      for (const child of node.children) {
        if (child.type !== 'element') continue;
        if (!localOrQualifiedName || child.name === localOrQualifiedName || child.localName === localOrQualifiedName) {
          result.push(child);
        }
        visit(child);
      }
    };
    visit(element);
    return result;
  }

  text(element: XmlElement): string {
    let value = '';
    const visit = (node: XmlElement): void => {
      for (const child of node.children) {
        if (child.type === 'text') value += child.value;
        else visit(child);
      }
    };
    visit(element);
    return value;
  }

  replaceText(element: XmlElement, value: string): void {
    if (element.selfClosing || element.endTagStart < element.startTagEnd) {
      throw new LosslessXmlError(`Cannot replace content of self-closing element ${element.name}`, element.start);
    }
    this.replace(element.startTagEnd, element.endTagStart, escapeXmlText(value));
  }

  replaceAttribute(attribute: XmlAttribute, value: string): void {
    this.replace(attribute.valueStart, attribute.valueEnd, escapeXmlAttribute(value));
  }

  replace(start: number, end: number, replacement: string): void {
    if (start < 0 || end < start || end > this.source.length) {
      throw new LosslessXmlError('Invalid patch range', start);
    }
    for (const patch of this.#patches) {
      if (start < patch.end && end > patch.start) {
        throw new LosslessXmlError('Overlapping XML patches are not allowed', start);
      }
    }
    this.#patches.push({ start, end, replacement });
  }

  get changed(): boolean {
    return this.#patches.length > 0;
  }

  serialize(): string {
    if (!this.changed) return this.source;
    let output = this.source;
    for (const patch of [...this.#patches].sort((left, right) => right.start - left.start)) {
      output = output.slice(0, patch.start) + patch.replacement + output.slice(patch.end);
    }
    return output;
  }
}

export function escapeXmlText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replaceAll('"', '&quot;').replaceAll("'", '&apos;');
}

export function decodeXmlEntities(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity, body: string) => {
    if (body.startsWith('#x')) return String.fromCodePoint(Number.parseInt(body.slice(2), 16));
    if (body.startsWith('#')) return String.fromCodePoint(Number.parseInt(body.slice(1), 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[body.toLowerCase()] ?? entity;
  });
}

function appendText(source: string, start: number, end: number, parent: XmlElement | undefined): void {
  if (!parent || end <= start) return;
  parent.children.push({ type: 'text', start, end, value: decodeXmlEntities(source.slice(start, end)), parent });
}

function localName(name: string): string {
  return name.includes(':') ? name.slice(name.indexOf(':') + 1) : name;
}

function requireTerminator(source: string, start: number, terminator: string): number {
  const end = source.indexOf(terminator, start + 2);
  if (end < 0) throw new LosslessXmlError(`Missing ${terminator}`, start);
  return end;
}

function findTagEnd(source: string, start: number): number {
  let quote: '"' | "'" | undefined;
  for (let index = start + 1; index < source.length; index += 1) {
    const character = source[index];
    if (quote) {
      if (character === quote) quote = undefined;
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      return index;
    }
  }
  throw new LosslessXmlError('Unclosed tag', start);
}

function parseAttributes(source: string, start: number, end: number): XmlAttribute[] {
  const attributes: XmlAttribute[] = [];
  let cursor = start;
  while (cursor < end) {
    while (/\s/.test(source[cursor] ?? '')) cursor += 1;
    if (cursor >= end || source[cursor] === '/') break;
    const nameStart = cursor;
    while (cursor < end && !/[\s=/>]/.test(source[cursor] ?? '')) cursor += 1;
    const name = source.slice(nameStart, cursor);
    while (/\s/.test(source[cursor] ?? '')) cursor += 1;
    if (source[cursor] !== '=') throw new LosslessXmlError(`Missing '=' after attribute ${name}`, cursor);
    cursor += 1;
    while (/\s/.test(source[cursor] ?? '')) cursor += 1;
    const quote = source[cursor];
    if (quote !== '"' && quote !== "'") throw new LosslessXmlError(`Unquoted attribute ${name}`, cursor);
    cursor += 1;
    const valueStart = cursor;
    while (cursor < end && source[cursor] !== quote) cursor += 1;
    if (cursor >= end) throw new LosslessXmlError(`Unclosed attribute ${name}`, valueStart);
    const valueEnd = cursor;
    cursor += 1;
    attributes.push({
      name,
      localName: localName(name),
      value: decodeXmlEntities(source.slice(valueStart, valueEnd)),
      start: nameStart,
      end: cursor,
      valueStart,
      valueEnd,
    });
  }
  return attributes;
}
