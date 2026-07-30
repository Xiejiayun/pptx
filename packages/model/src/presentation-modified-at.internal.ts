import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';
import {
  isPresentationTimestamp,
  normalizePresentationTimestamp,
  W3CDTF_QUALIFIED_TYPE,
} from './presentation-timestamp.internal.js';

const MODIFIED_AT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'modified timestamp',
  localName: 'modified',
  namespace: 'http://purl.org/dc/terms/',
  preferredPrefix: 'dcterms',
  qualifiedType: W3CDTF_QUALIFIED_TYPE,
};

export function readPresentationModifiedAt(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, MODIFIED_AT_PROPERTY);
  return value !== undefined && isPresentationTimestamp(value) ? value : undefined;
}

export function replacePresentationModifiedAt(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(
    pkg,
    MODIFIED_AT_PROPERTY,
    normalizePresentationTimestamp(value, 'modifiedAt'),
  );
}
