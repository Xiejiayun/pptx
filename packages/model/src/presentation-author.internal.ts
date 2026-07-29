import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const AUTHOR_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'author',
  localName: 'creator',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationAuthor(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, AUTHOR_PROPERTY);
}

export function replacePresentationAuthor(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, AUTHOR_PROPERTY, normalizePresentationAuthor(value));
}

function normalizePresentationAuthor(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation author must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation author contains invalid XML characters');
  }
  return value;
}
