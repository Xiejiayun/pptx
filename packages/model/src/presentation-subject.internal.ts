import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const SUBJECT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'subject',
  localName: 'subject',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationSubject(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, SUBJECT_PROPERTY);
}

export function replacePresentationSubject(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, SUBJECT_PROPERTY, normalizePresentationSubject(value));
}

function normalizePresentationSubject(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation subject must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation subject contains invalid XML characters');
  }
  return value;
}
