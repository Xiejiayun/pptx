import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const TITLE_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'title',
  localName: 'title',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationTitle(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, TITLE_PROPERTY);
}

export function replacePresentationTitle(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, TITLE_PROPERTY, normalizePresentationTitle(value));
}

function normalizePresentationTitle(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation title must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation title contains invalid XML characters');
  }
  return value;
}
