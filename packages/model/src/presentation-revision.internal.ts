import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const REVISION_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'revision',
  localName: 'revision',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};

export function readPresentationRevision(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, REVISION_PROPERTY);
  return value !== undefined && isPresentationRevision(value) ? value : undefined;
}

export function replacePresentationRevision(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, REVISION_PROPERTY, normalizePresentationRevision(value));
}

function normalizePresentationRevision(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !isPresentationRevision(value)) {
    throw new TypeError(
      'Presentation revision must be a non-empty ASCII whole-number string or undefined',
    );
  }
  return value;
}

function isPresentationRevision(value: string): boolean {
  return /^[0-9]+$/.test(value);
}
