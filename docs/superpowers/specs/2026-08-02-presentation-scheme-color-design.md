# Presentation `SchemeColor` Runtime Helper Design

## Status

Validated under the standing project direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. This item follows the completed presentation compression policy and does not change the overall approximately 97% parity estimate by itself.

## Goal

Publish a stable native runtime equivalent of PptxGenJS 4.0.1's public `SchemeColor` helper so callers can discover and reuse the ten upstream theme-color names and values from Node.js, browsers, TypeScript, and the packed root package without creating or mutating a presentation.

## Current context

PptxGenJS 4.0.1 exposes one prototype getter named `SchemeColor`. All instances return the same mutable enum-shaped object in this order:

| Key | Value |
| --- | --- |
| `text1` | `tx1` |
| `text2` | `tx2` |
| `background1` | `bg1` |
| `background2` | `bg2` |
| `accent1` | `accent1` |
| `accent2` | `accent2` |
| `accent3` | `accent3` |
| `accent4` | `accent4` |
| `accent5` | `accent5` |
| `accent6` | `accent6` |

The native library already reads, validates, writes, and reopens scheme colors throughout text, fills, lines, shadows, tables, charts, slide numbers, backgrounds, and default text color. Those focused APIs accept seventeen DrawingML tokens, including hyperlink, dark/light, and placeholder colors that PptxGenJS does not publish through this helper. The remaining gap is runtime discovery and a precise public helper type, not OOXML color serialization.

Existing native runtime helpers use immutable package-level exports and derived literal types. They do not reproduce PptxGenJS instance namespaces, mutable aliases, or implementation-specific property descriptors.

## Considered approaches

### 1. Frozen root mapping with a derived value union — selected

Add one model-owned frozen mapping whose keys and values exactly match the ten upstream entries. Derive the public `SchemeColor` union from its values and re-export the same object through model, SDK, and the aggregate root.

This preserves the meaningful `text1 -> tx1` and `background1 -> bg1` aliases, gives callers familiar property access, and corrects the upstream shared-mutation hazard. It also follows the established single-owner and root-re-export pattern.

### 2. Values-only tuple

A tuple would match `TEXT_ALIGNMENTS`, `TEXT_VERTICAL_ALIGNMENTS`, and `OUTPUT_TYPES`, but it would discard the four upstream friendly-name mappings. Callers would need to know that `text1` is represented by `tx1`, weakening helper parity.

### 3. `PptxDocument.SchemeColor` instance getter

This would most closely copy the upstream namespace shape, but it would add presentation-instance state for a stateless catalog and encourage parity with a shared mutable object. It is inconsistent with the native root-catalog architecture and is rejected.

## Public API

The model package owns this exact value:

```ts
export const SCHEME_COLORS = Object.freeze({
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const);

export type SchemeColor = (typeof SCHEME_COLORS)[keyof typeof SCHEME_COLORS];
```

`@pptx/model`, `@pptx/sdk`, and `@jiayunxie/pptx` expose the same object identity and type. The root usage is:

```ts
import { PptxDocument, SCHEME_COLORS, type SchemeColor } from '@jiayunxie/pptx';

const accent: SchemeColor = SCHEME_COLORS.accent1;
const text: SchemeColor = SCHEME_COLORS.text1;

const document = PptxDocument.create();
document.addSlide().addText('Theme colors', {
  color: { kind: 'scheme', value: accent },
});
```

The mapping is deeply immutable because every value is a string primitive and the sole object is frozen. Enumeration order is the declaration order shown above. No `PptxDocument.SchemeColor`, class static, enum, mutable copy, alias object, or lowercase/uppercase duplicate is added.

## Existing color type boundary

`RichTextColor` remains unchanged in this item. Its scheme `value` stays a string because native APIs intentionally support seven additional DrawingML scheme tokens beyond the ten-value PptxGenJS helper. Every `SchemeColor` value is already assignable to `RichTextColor.value`, so the helper works across existing color-taking APIs without a breaking type restriction.

This item does not consolidate the five internal validation sets. That would mix a discovery-only public surface with wider low-level DrawingML validation and create unrelated regression risk. Internal consolidation may be considered separately only if a later feature needs one canonical seventeen-token type.

## Conformance and failure behavior

PptxGenJS conformance compares public keys, values, enumeration order, shared instance identity, and getter shape. Native must match the ten key/value pairs and order. Native deliberately differs in these documented ways:

- the helper is a package-level root export rather than an instance getter;
- the mapping is frozen rather than shared and mutable;
- mutation attempts throw in strict ESM and cannot affect later reads, other documents, or package output.

There is no runtime parser or new error path. TypeScript rejects unknown keys and values through the inferred readonly object and `SchemeColor` union. Existing color normalizers remain responsible for validating API inputs.

## Test design

### Model and root tests

- Require the exact ten entries and enumeration order.
- Require ten unique values, `Object.isFrozen()`, and stable state after failed assignment, deletion, or extension.
- Require `SchemeColor` to accept all ten values and reject key names such as `text1`, raw sRGB, unsupported scheme values, and arbitrary strings.
- Require SDK and aggregate-root exports to share model object identity.
- Read the catalog before and after create/write/reopen and require an unchanged package mutation journal.
- Use `SCHEME_COLORS.text1` and `SCHEME_COLORS.accent1` in real text/fill state and require canonical `a:schemeClr` values after reopen.

### PptxGenJS evidence

- Inspect a real 4.0.1 instance through public APIs and require its ten keys and values in order.
- Confirm upstream instances share the same mutable object as evidence for the native immutable divergence.
- Generate legal text/fill output with `SchemeColor.text1` and `SchemeColor.accent1`, import it, and compare native semantic color snapshots.

### Package and browser evidence

- Extend generated declarations to require `SCHEME_COLORS` and `SchemeColor` from model through SDK/root.
- Extend the actual-tarball Node, TypeScript, browser-condition, CLI, and real-Chrome smoke with one stable `schemeColors: true` field plus detailed values, frozen/identity/mutation-isolation, use, write/reopen, and zero-error state.
- The browser bundle must not fetch package metadata or import Node-only modules for catalog discovery.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. core catalog, type, exports, conformance, and focused tests;
4. packed Node/type/browser/CLI and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/` and generated workspace tarballs remain outside commits.

## Non-goals

- Replacing existing seventeen-token DrawingML validation.
- Narrowing `RichTextColor` or other current color input types.
- Resolving theme colors to sRGB values.
- Adding theme mutation, color transforms, tint/shade, alpha, or inheritance APIs.
- Adding an instance getter, mutable enum object, compatibility alias, or default export.
- Changing OOXML bytes for existing presentations that do not use new public constants.

## Acceptance criteria

The item is complete when the ten upstream key/value pairs and order are available through one frozen root mapping; `SchemeColor` is derived from its values; model, SDK, and root share identity; real color creation and reopen work; upstream legal output imports semantically; mutation and package isolation hold; focused/full/performance/type/bundle/declaration gates pass; the actual tarball and real Chrome report `schemeColors: true` with zero browser errors; documentation records the immutable root-surface divergence; and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
