import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const LAST_MODIFIED_BY_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'last modified by',
  localName: 'lastModifiedBy',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};

export function readPresentationLastModifiedBy(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, LAST_MODIFIED_BY_PROPERTY);
}

export function replacePresentationLastModifiedBy(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(
    pkg,
    LAST_MODIFIED_BY_PROPERTY,
    normalizePresentationLastModifiedBy(value),
  );
}

function normalizePresentationLastModifiedBy(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation lastModifiedBy must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation lastModifiedBy contains invalid XML characters');
  }
  return value;
}
