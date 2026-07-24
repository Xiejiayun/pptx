import type { OpcPackage, Relationship } from '@pptx/opc';

export type DiagnosticSeverity = 'info' | 'warning' | 'error';
export type CompatibilityProfile =
  | 'powerpoint-2010'
  | 'powerpoint-current'
  | 'keynote-current'
  | 'libreoffice-current'
  | 'google-slides-import';

export interface Diagnostic {
  readonly severity: DiagnosticSeverity;
  readonly code: string;
  readonly message: string;
  readonly partUri?: string;
  readonly xmlPath?: string;
  readonly objectId?: string;
  readonly compatibility?: CompatibilityProfile;
  readonly suggestion?: string;
}

export class ValidationError extends Error {
  constructor(readonly diagnostics: readonly Diagnostic[]) {
    super(`Validation failed with ${diagnostics.filter(({ severity }) => severity === 'error').length} error(s)`);
    this.name = 'ValidationError';
  }
}

export function validatePackage(pkg: OpcPackage): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const officeDocumentRelationships = pkg.relationships('/').filter(({ type }) => type.endsWith('/officeDocument'));
  if (officeDocumentRelationships.length !== 1) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_OFFICE_DOCUMENT_CARDINALITY',
      message: `Expected exactly one root officeDocument relationship; found ${officeDocumentRelationships.length}`,
      partUri: '/_rels/.rels',
      suggestion: 'Add one root relationship targeting the presentation part.',
    });
  }
  const idsByRelationshipPart = new Map<string, Set<string>>();
  for (const part of pkg.parts.filter(({ uri }) => uri.endsWith('.rels'))) {
    const ids = idsByRelationshipPart.get(part.uri) ?? new Set<string>();
    idsByRelationshipPart.set(part.uri, ids);
    for (const relationship of part.relationships) {
      validateRelationship(pkg, part.uri, relationship, ids, diagnostics);
    }
  }
  return diagnostics;
}

function validateRelationship(
  pkg: OpcPackage,
  relationshipPartUri: string,
  relationship: Relationship,
  ids: Set<string>,
  diagnostics: Diagnostic[],
): void {
  if (ids.has(relationship.id)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_DUPLICATE_RELATIONSHIP_ID',
      message: `Duplicate relationship id ${relationship.id}`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Allocate a unique relationship id in the source part.',
    });
  }
  ids.add(relationship.id);
  if (!/^[A-Za-z_][\w.-]*$/.test(relationship.id)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_INVALID_RELATIONSHIP_ID',
      message: `Relationship id ${relationship.id} is not a valid XML ID`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Use an id such as rId1.',
    });
  }
  if (relationship.targetMode === 'Internal' && relationship.resolvedTarget && !pkg.hasPart(relationship.resolvedTarget)) {
    diagnostics.push({
      severity: 'error',
      code: 'OPC_DANGLING_RELATIONSHIP',
      message: `Relationship ${relationship.id} targets missing part ${relationship.resolvedTarget}`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Add the target part or remove the relationship.',
    });
  }
  if (relationship.targetMode === 'External') {
    diagnostics.push({
      severity: 'warning',
      code: 'OPC_EXTERNAL_RELATIONSHIP',
      message: `External relationship ${relationship.id} is not portable`,
      partUri: relationshipPartUri,
      objectId: relationship.id,
      suggestion: 'Embed the resource when the presentation must be self-contained.',
    });
  }
}

