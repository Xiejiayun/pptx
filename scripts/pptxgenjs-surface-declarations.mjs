import { createHash } from 'node:crypto';
import { readFile, stat } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';

const EXPECTED_PACKAGE_VERSION = '4.0.1';
const BUILTIN_TYPES = new Set([
  'Array',
  'ArrayBuffer',
  'Blob',
  'Boolean',
  'Buffer',
  'Date',
  'Error',
  'Function',
  'HTMLElement',
  'Map',
  'Number',
  'Object',
  'Promise',
  'Readonly',
  'ReadonlyArray',
  'Record',
  'Set',
  'String',
  'Uint8Array',
]);

function deepFreeze(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return value;
  seen.add(value);
  for (const nested of Object.values(value)) deepFreeze(nested, seen);
  return Object.freeze(value);
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function compareText(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function isInside(parent, child) {
  const childRelative = relative(parent, child);
  return childRelative === ''
    || (!childRelative.startsWith('..') && !isAbsolute(childRelative));
}

async function readJson(path, context) {
  let parsed;
  try {
    parsed = JSON.parse(await readFile(path, 'utf8'));
  } catch (error) {
    throw new Error(`${context} is not valid JSON: ${error.message}`, { cause: error });
  }
  if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new TypeError(`${context} must contain an object`);
  }
  return parsed;
}

async function findPackageRoot(entryPath) {
  let directory = dirname(entryPath);
  while (true) {
    const packageJsonPath = join(directory, 'package.json');
    try {
      const packageJson = await readJson(packageJsonPath, packageJsonPath);
      if (packageJson.name === 'pptxgenjs') {
        return { packageJson, packageJsonPath, root: directory };
      }
    } catch (error) {
      if (error?.code !== 'ENOENT' && error?.cause?.code !== 'ENOENT') throw error;
    }
    const parent = dirname(directory);
    if (parent === directory) break;
    directory = parent;
  }
  throw new Error(`unable to find the pptxgenjs package root from ${entryPath}`);
}

export async function resolvePptxGenJSPackage(adapterPackagePath) {
  if (typeof adapterPackagePath !== 'string' || adapterPackagePath.length === 0) {
    throw new TypeError('adapterPackagePath must be a non-empty string');
  }
  const adapterPath = resolve(adapterPackagePath);
  const packageRequire = createRequire(adapterPath);
  let entryPath;
  try {
    entryPath = packageRequire.resolve('pptxgenjs');
  } catch (error) {
    throw new Error(`unable to resolve pptxgenjs from ${adapterPath}: ${error.message}`, {
      cause: error,
    });
  }

  const { packageJson, packageJsonPath, root } = await findPackageRoot(entryPath);
  if (packageJson.version !== EXPECTED_PACKAGE_VERSION) {
    throw new Error(
      `expected pptxgenjs@${EXPECTED_PACKAGE_VERSION} but resolved ${String(packageJson.version)}`,
    );
  }
  if (!isInside(root, entryPath)) {
    throw new Error('resolved pptxgenjs entry is outside its package root');
  }

  const typesEntry = packageJson.types ?? 'types/index.d.ts';
  if (typeof typesEntry !== 'string' || typesEntry.length === 0) {
    throw new TypeError('pptxgenjs package types entry must be a non-empty string');
  }
  let declarationPath = resolve(root, typesEntry);
  if (!isInside(root, declarationPath)) {
    throw new Error('resolved pptxgenjs declaration is outside its package root');
  }
  const declarationStat = await stat(declarationPath);
  if (declarationStat.isDirectory()) declarationPath = join(declarationPath, 'index.d.ts');
  const declarationBytes = await readFile(declarationPath);

  return deepFreeze({
    root,
    packageJsonPath,
    declarationPath,
    entryPath,
    version: packageJson.version,
    declarationSha256: sha256(declarationBytes),
  });
}

function nodeName(typescript, node) {
  const name = node?.name;
  if (typescript.isIdentifier(name) || typescript.isPrivateIdentifier?.(name)) {
    return name.text;
  }
  if (
    typescript.isStringLiteral(name)
    || typescript.isNumericLiteral(name)
    || typescript.isNoSubstitutionTemplateLiteral?.(name)
  ) {
    return name.text;
  }
  return undefined;
}

function hasModifier(typescript, node, kind) {
  return node.modifiers?.some((modifier) => modifier.kind === kind) ?? false;
}

function isPublicMember(typescript, node) {
  return !hasModifier(typescript, node, typescript.SyntaxKind.PrivateKeyword)
    && !hasModifier(typescript, node, typescript.SyntaxKind.ProtectedKeyword);
}

function isDeprecated(typescript, node) {
  return typescript.getJSDocTags(node).some(({ tagName }) => tagName.text === 'deprecated');
}

function entityNameText(typescript, node) {
  if (typescript.isIdentifier(node)) return node.text;
  if (typescript.isQualifiedName(node)) {
    return `${entityNameText(typescript, node.left)}.${node.right.text}`;
  }
  if (typescript.isPropertyAccessExpression?.(node)) {
    return `${entityNameText(typescript, node.expression)}.${node.name.text}`;
  }
  return '';
}

function finalEntityName(text) {
  const parts = text.split('.');
  return parts.at(-1) ?? text;
}

export function extractPptxGenJSPublicSurface({
  sourceText,
  fileName = 'types/index.d.ts',
  typescript,
}) {
  if (typeof sourceText !== 'string') throw new TypeError('sourceText must be a string');
  if (typeof fileName !== 'string' || fileName.length === 0) {
    throw new TypeError('fileName must be a non-empty string');
  }
  if (typescript === null || typeof typescript !== 'object') {
    throw new TypeError('typescript compiler API is required');
  }

  const ts = typescript;
  const sourceFile = ts.createSourceFile(
    fileName,
    sourceText,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  if (sourceFile.parseDiagnostics.length > 0) {
    const details = sourceFile.parseDiagnostics
      .map(({ messageText }) => ts.flattenDiagnosticMessageText(messageText, ' '))
      .join('; ');
    throw new SyntaxError(`PptxGenJS declaration parse diagnostics: ${details}`);
  }

  const defaultExport = sourceFile.statements.find(
    (statement) => ts.isExportAssignment(statement) && !statement.isExportEquals,
  );
  const defaultClass = sourceFile.statements.find(
    (statement) => ts.isClassDeclaration(statement)
      && hasModifier(ts, statement, ts.SyntaxKind.DefaultKeyword),
  );
  const presentationName = defaultExport && ts.isIdentifier(defaultExport.expression)
    ? defaultExport.expression.text
    : defaultClass?.name?.text;
  if (!presentationName) {
    throw new Error('PptxGenJS declarations must contain an identifiable default export');
  }

  const presentationClass = sourceFile.statements.find(
    (statement) => ts.isClassDeclaration(statement) && statement.name?.text === presentationName,
  ) ?? defaultClass;
  if (!presentationClass) {
    throw new Error(`default export ${presentationName} does not resolve to a class`);
  }

  const namespaceStatements = [];
  for (const statement of sourceFile.statements) {
    if (!ts.isModuleDeclaration(statement) || statement.name.text !== presentationName) continue;
    let body = statement.body;
    while (body && ts.isModuleDeclaration(body)) body = body.body;
    if (body && ts.isModuleBlock(body)) namespaceStatements.push(...body.statements);
  }
  if (namespaceStatements.length === 0) {
    throw new Error(`default export ${presentationName} has no matching namespace`);
  }

  const declarationsByName = new Map();
  for (const statement of namespaceStatements) {
    const name = statement.name?.text;
    if (!name) continue;
    const declarations = declarationsByName.get(name) ?? [];
    declarations.push(statement);
    declarationsByName.set(name, declarations);
  }

  const slideClass = (declarationsByName.get('Slide') ?? []).find(ts.isClassDeclaration);
  if (!slideClass) throw new Error(`${presentationName}.Slide does not resolve to a class`);

  const printer = ts.createPrinter({
    newLine: ts.NewLineKind.LineFeed,
    removeComments: true,
  });
  const printNode = (node) => printer
    .printNode(ts.EmitHint.Unspecified, node, sourceFile)
    .replace(/\s+/gu, ' ')
    .trim();
  const printSignature = (node) => printNode(node).replace(/;$/u, '');
  const printType = (node) => (node ? printNode(node) : 'unknown');

  const atomsById = new Map();
  const inlineTypeTexts = new Map();
  const memberSlots = new Map();
  const queuedTypes = [];
  const enqueuedTypes = new Set();
  const processedTypes = new Set();
  const flattenedInterfaces = new Map();
  const auditableUnionKeywordKinds = new Set([
    ts.SyntaxKind.AnyKeyword,
    ts.SyntaxKind.BooleanKeyword,
    ts.SyntaxKind.NeverKeyword,
    ts.SyntaxKind.NumberKeyword,
    ts.SyntaxKind.ObjectKeyword,
    ts.SyntaxKind.StringKeyword,
    ts.SyntaxKind.SymbolKeyword,
    ts.SyntaxKind.UndefinedKeyword,
    ts.SyntaxKind.UnknownKeyword,
    ts.SyntaxKind.VoidKeyword,
  ]);

  const addAtom = (atom) => {
    const existing = atomsById.get(atom.id);
    if (!existing) {
      atomsById.set(atom.id, atom);
      return;
    }
    if (existing.kind === 'method' && atom.kind === 'method') {
      existing.signatures = [...new Set([...existing.signatures, ...atom.signatures])].sort();
      existing.deprecatedSignatures = [
        ...new Set([...existing.deprecatedSignatures, ...atom.deprecatedSignatures]),
      ].sort();
      existing.deprecated &&= atom.deprecated;
      return;
    }
    if (
      atom.id.startsWith('inline:')
      && existing.kind === 'property'
      && atom.kind === 'property'
      && existing.owner === atom.owner
      && existing.name === atom.name
      && existing.declaredIn === atom.declaredIn
    ) {
      const typeTexts = inlineTypeTexts.get(atom.id)
        ?? new Set(existing.typeText.length > 0 ? [existing.typeText] : []);
      if (atom.typeText.length > 0) typeTexts.add(atom.typeText);
      inlineTypeTexts.set(atom.id, typeTexts);
      existing.typeText = [...typeTexts].sort().join(' | ');
      existing.optional ||= atom.optional;
      existing.readonly ||= atom.readonly;
      existing.deprecated ||= atom.deprecated;
      return;
    }
    if (JSON.stringify(existing) !== JSON.stringify(atom)) {
      throw new Error(`conflicting atom ${atom.id}`);
    }
  };

  const claimMemberSlot = (ownerKind, owner, name, kind) => {
    const slot = `${ownerKind}:${owner}:${name}`;
    const existing = memberSlots.get(slot);
    if (existing && existing !== kind) {
      throw new Error(`conflicting atom member ${owner}.${name}`);
    }
    memberSlots.set(slot, kind);
  };

  const makeAtom = ({
    id,
    kind,
    owner,
    name,
    declaredIn = owner,
    node,
    typeNode,
    signatures = [],
  }) => {
    const deprecated = Boolean(node && isDeprecated(ts, node));
    return {
      id,
      kind,
      owner,
      name,
      declaredIn,
      optional: Boolean(node?.questionToken),
      readonly: Boolean(
        node && hasModifier(ts, node, ts.SyntaxKind.ReadonlyKeyword),
      ),
      deprecated,
      signatures: [...signatures].sort(),
      deprecatedSignatures: deprecated ? [...signatures].sort() : [],
      catalogKey: '',
      typeText: typeNode ? printType(typeNode) : '',
    };
  };

  const resolveLocalReference = (fullName) => {
    const name = finalEntityName(fullName);
    if (BUILTIN_TYPES.has(name) || name === presentationName || name === 'Slide') {
      return undefined;
    }
    if (declarationsByName.has(name)) return name;
    if (fullName.startsWith(`${presentationName}.`)) {
      throw new Error(`unresolved local type ${fullName}`);
    }
    return undefined;
  };

  const enqueueType = (fullName) => {
    const name = resolveLocalReference(fullName);
    if (!name || enqueuedTypes.has(name)) return;
    enqueuedTypes.add(name);
    queuedTypes.push(name);
  };

  const canonicalUnionToken = (node) => {
    if (ts.isLiteralTypeNode(node)) {
      if (
        ts.isStringLiteral(node.literal)
        || ts.isNumericLiteral(node.literal)
        || ts.isNoSubstitutionTemplateLiteral?.(node.literal)
      ) {
        return node.literal.text;
      }
      if (node.literal.kind === ts.SyntaxKind.TrueKeyword) return 'true';
      if (node.literal.kind === ts.SyntaxKind.FalseKeyword) return 'false';
    }
    const printed = printType(node);
    if (ts.isTemplateLiteralTypeNode(node) && printed.startsWith('`') && printed.endsWith('`')) {
      return printed.slice(1, -1);
    }
    return printed.replace(/\s+/gu, '');
  };

  const addUnionAtom = (
    owner,
    token,
    typeNode,
    node = typeNode,
    catalogKey = token,
  ) => {
    addAtom({
      ...makeAtom({
        id: `union:${owner}#${token}`,
        kind: 'union-member',
        owner,
        name: token,
        node,
        typeNode,
      }),
      catalogKey,
    });
  };

  const unionOwner = (owningAtomId, pathPrefix) => {
    const propertyMarker = '@property:';
    const markerIndex = owningAtomId.lastIndexOf(propertyMarker);
    const directPropertyName = markerIndex === -1
      ? undefined
      : owningAtomId.slice(markerIndex + propertyMarker.length);
    return pathPrefix && pathPrefix !== directPropertyName
      ? `${owningAtomId}@path:${pathPrefix}`
      : owningAtomId;
  };

  const isAuditableUnionAlternative = (node) => {
    if (
      ts.isLiteralTypeNode(node)
      || ts.isTemplateLiteralTypeNode(node)
      || ts.isTypeReferenceNode(node)
      || ts.isArrayTypeNode(node)
      || ts.isTupleTypeNode(node)
    ) {
      return true;
    }
    return auditableUnionKeywordKinds.has(node.kind);
  };

  let visitTypeNode;

  const emitInlineMembers = (typeLiteral, owningAtomId, pathPrefix, declaredIn) => {
    for (const member of typeLiteral.members) {
      if (!isPublicMember(ts, member)) continue;
      const name = nodeName(ts, member);
      if (!name) {
        if (member.type) visitTypeNode(member.type, owningAtomId, pathPrefix, declaredIn);
        continue;
      }
      const path = pathPrefix ? `${pathPrefix}.${name}` : name;
      if (ts.isPropertySignature(member)) {
        const id = `inline:${owningAtomId}@property:${path}`;
        addAtom(makeAtom({
          id,
          kind: 'property',
          owner: owningAtomId,
          name: path,
          declaredIn,
          node: member,
          typeNode: member.type,
        }));
        if (member.type) visitTypeNode(member.type, owningAtomId, path, declaredIn);
      } else if (ts.isMethodSignature(member)) {
        const id = `inline:${owningAtomId}@method:${path}`;
        addAtom(makeAtom({
          id,
          kind: 'method',
          owner: owningAtomId,
          name: path,
          declaredIn,
          node: member,
          signatures: [printSignature(member)],
        }));
        for (const parameter of member.parameters) {
          if (parameter.type) visitTypeNode(parameter.type, owningAtomId, path, declaredIn);
        }
        if (member.type) visitTypeNode(member.type, owningAtomId, path, declaredIn);
      } else {
        ts.forEachChild(member, (child) => {
          if (ts.isTypeNode(child)) visitTypeNode(child, owningAtomId, path, declaredIn);
        });
      }
    }
  };

  visitTypeNode = (node, owningAtomId, pathPrefix = '', declaredIn = '') => {
    if (!node) return;
    if (ts.isUnionTypeNode(node) && owningAtomId) {
      for (const member of node.types) {
        if (isAuditableUnionAlternative(member)) {
          addUnionAtom(
            unionOwner(owningAtomId, pathPrefix),
            canonicalUnionToken(member),
            member,
          );
        }
      }
    }
    if (ts.isTypeReferenceNode(node)) {
      enqueueType(entityNameText(ts, node.typeName));
    } else if (ts.isTypeQueryNode(node)) {
      enqueueType(entityNameText(ts, node.exprName));
    } else if (ts.isExpressionWithTypeArguments(node)) {
      enqueueType(entityNameText(ts, node.expression));
    } else if (ts.isTypeLiteralNode(node) && owningAtomId) {
      emitInlineMembers(node, owningAtomId, pathPrefix, declaredIn);
      return;
    }
    ts.forEachChild(node, (child) => {
      if (ts.isTypeNode(child) || ts.isExpressionWithTypeArguments(child)) {
        visitTypeNode(child, owningAtomId, pathPrefix, declaredIn);
      }
    });
  };

  const emitClassMembers = (classNode, kind) => {
    const owner = kind === 'presentation' ? presentationName : 'Slide';
    for (const member of classNode.members) {
      if (!isPublicMember(ts, member)) continue;
      const name = nodeName(ts, member);
      if (!name) continue;
      if (
        ts.isPropertyDeclaration(member)
        || ts.isGetAccessorDeclaration(member)
        || ts.isSetAccessorDeclaration(member)
      ) {
        claimMemberSlot(kind, owner, name, 'property');
        const id = kind === 'presentation'
          ? `class:${owner}@property:${name}`
          : `property:Slide#${name}`;
        addAtom(makeAtom({
          id,
          kind: 'property',
          owner,
          name,
          node: member,
          typeNode: member.type,
        }));
        if (member.type) visitTypeNode(member.type, id, name, owner);
      } else if (ts.isMethodDeclaration(member)) {
        claimMemberSlot(kind, owner, name, 'method');
        const id = kind === 'presentation'
          ? `class:${owner}#${name}`
          : `method:Slide#${name}`;
        addAtom(makeAtom({
          id,
          kind: 'method',
          owner,
          name,
          node: member,
          signatures: [printSignature(member)],
        }));
        for (const parameter of member.parameters) {
          if (parameter.type) visitTypeNode(parameter.type, id, parameter.name.getText(sourceFile), owner);
        }
        if (member.type) visitTypeNode(member.type, id, 'return', owner);
      }
    }
  };

  const flattenInterface = (name, visiting = new Set()) => {
    const cached = flattenedInterfaces.get(name);
    if (cached) return cached;
    if (visiting.has(name)) return [];
    const nextVisiting = new Set(visiting).add(name);
    const members = new Map();
    const declarations = (declarationsByName.get(name) ?? []).filter(ts.isInterfaceDeclaration);
    for (const declaration of declarations) {
      for (const clause of declaration.heritageClauses ?? []) {
        for (const heritageType of clause.types) {
          const fullName = entityNameText(ts, heritageType.expression);
          const baseName = resolveLocalReference(fullName);
          if (!baseName) continue;
          const baseDeclarations = declarationsByName.get(baseName) ?? [];
          if (!baseDeclarations.some(ts.isInterfaceDeclaration)) continue;
          for (const inherited of flattenInterface(baseName, nextVisiting)) {
            if (!members.has(inherited.name)) members.set(inherited.name, inherited);
          }
        }
      }
      for (const member of declaration.members) {
        if (!isPublicMember(ts, member)) continue;
        const memberName = nodeName(ts, member);
        if (!memberName) continue;
        members.set(memberName, { member, name: memberName, declaredIn: name });
      }
    }
    const result = [...members.values()];
    flattenedInterfaces.set(name, result);
    return result;
  };

  const processInterface = (name) => {
    for (const { member, name: memberName, declaredIn } of flattenInterface(name)) {
      if (ts.isPropertySignature(member)) {
        const id = `interface:${name}@property:${memberName}`;
        addAtom(makeAtom({
          id,
          kind: 'property',
          owner: name,
          name: memberName,
          declaredIn,
          node: member,
          typeNode: member.type,
        }));
        if (member.type) visitTypeNode(member.type, id, memberName, declaredIn);
      } else if (ts.isMethodSignature(member)) {
        const id = `interface:${name}#${memberName}`;
        addAtom(makeAtom({
          id,
          kind: 'method',
          owner: name,
          name: memberName,
          declaredIn,
          node: member,
          signatures: [printSignature(member)],
        }));
        for (const parameter of member.parameters) {
          if (parameter.type) visitTypeNode(parameter.type, id, parameter.name.getText(sourceFile), declaredIn);
        }
        if (member.type) visitTypeNode(member.type, id, 'return', declaredIn);
      }
    }
  };

  const processTypeAlias = (name, declaration) => {
    if (ts.isUnionTypeNode(declaration.type)) {
      for (const member of declaration.type.types) {
        addUnionAtom(name, canonicalUnionToken(member), member);
      }
    }
    if (ts.isTypeLiteralNode(declaration.type)) {
      emitInlineMembers(declaration.type, `type:${name}`, name, name);
    } else {
      visitTypeNode(declaration.type, undefined, '', name);
    }
  };

  const processEnum = (name, declaration) => {
    for (const member of declaration.members) {
      const memberName = nodeName(ts, member);
      if (!memberName) continue;
      let token = memberName;
      if (member.initializer) {
        if (ts.isStringLiteral(member.initializer) || ts.isNumericLiteral(member.initializer)) {
          token = member.initializer.text;
        } else {
          token = printNode(member.initializer).replace(/\s+/gu, '');
        }
      }
      addUnionAtom(name, token, member.initializer ?? member, member, memberName);
    }
  };

  emitClassMembers(presentationClass, 'presentation');
  emitClassMembers(slideClass, 'slide');

  while (queuedTypes.length > 0) {
    const name = queuedTypes.shift();
    if (processedTypes.has(name)) continue;
    processedTypes.add(name);
    const declarations = declarationsByName.get(name) ?? [];
    const interfaces = declarations.filter(ts.isInterfaceDeclaration);
    if (interfaces.length > 0) {
      processInterface(name);
      continue;
    }
    const alias = declarations.find(ts.isTypeAliasDeclaration);
    if (alias) {
      processTypeAlias(name, alias);
      continue;
    }
    const enumeration = declarations.find(ts.isEnumDeclaration);
    if (enumeration) {
      processEnum(name, enumeration);
      continue;
    }
  }

  const atoms = [...atomsById.values()]
    .map((atom) => ({
      ...atom,
      signatures: [...atom.signatures].sort(),
      deprecatedSignatures: [...atom.deprecatedSignatures].sort(),
    }))
    .sort((left, right) => compareText(left.id, right.id));
  return deepFreeze({
    schemaVersion: 1,
    atoms,
    roots: { presentation: presentationName, slide: 'Slide' },
    diagnostics: [],
  });
}
