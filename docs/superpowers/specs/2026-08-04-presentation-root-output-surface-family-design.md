# PptxGenJS Presentation Root, Output, and Theme Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit still contains 35 unverified atoms for
the presentation root, presentation layout projection, theme fonts, and output
runtime. The native SDK already implements the legal presentation metadata,
theme, layout, write, file, and Node stream lifecycles. This family closes the
declaration inventory by reusing those implementations and adding one aggregate
upstream control for the remaining method-return and declaration defects.

This is the first family in a three-family lifecycle artifact batch. Focused
tests, review, commit, and push remain family-local. npm pack, persistent Chrome,
OOXML package inspection, TypeScript, and full Vitest run once at the batch
boundary after the slide/section and master/background families are closed.

## Batch Boundary

This family contains exactly 35 atoms:

- the seven root methods `addSection`, `addSlide`, `defineLayout`,
  `defineSlideMaster`, `stream`, `write`, and `writeFile`;
- the eight root properties `author`, `company`, `layout`, `revision`,
  `rtlMode`, `subject`, `theme`, and `title`;
- `PresLayout.name`, `PresLayout.width`, and `PresLayout.height`;
- `ThemeProps.headFontFace` and `ThemeProps.bodyFontFace`;
- the five direct properties owned by `WriteBaseProps`, `WriteFileProps`, and
  `WriteProps`;
- the `JSZIP_OUTPUT_TYPE` and `STREAM` branches of `WRITE_OUTPUT_TYPE`;
- the four declared return branches of `PptxGenJS.stream()`;
- the four declared return branches of `PptxGenJS.write()`.

It excludes `PptxGenJS.ChartType`, slide members, section option properties,
master/background declarations, and presentation metadata not declared by
PptxGenJS 4.0.1.

## Locked Classification

### Supported: 18

- `PptxGenJS.addSlide`;
- `PptxGenJS.author`, `company`, `revision`, `rtlMode`, `subject`, `theme`, and
  `title`;
- `PresLayout.width` and `PresLayout.height`;
- `ThemeProps.headFontFace` and `ThemeProps.bodyFontFace`;
- `WriteBaseProps.compression`;
- `WRITE_OUTPUT_TYPE.JSZIP_OUTPUT_TYPE`;
- the `ArrayBuffer`, `Blob`, `Uint8Array`, and `string` return branches of
  `PptxGenJS.write()`.

Native preserves every legal explicit value and representation. Stricter input
validation, transactional isolation, editable reopen, and different omitted
metadata defaults do not remove any legal explicit capability.

### Deliberate differences: 12

- `PptxGenJS.addSection`, `defineLayout`, `defineSlideMaster`, `stream`,
  `write`, and `writeFile`;
- `PptxGenJS.layout` and `PresLayout.name`;
- `WriteFileProps.compression`, `WriteFileProps.fileName`,
  `WriteProps.compression`, and `WriteProps.outputType`;
- `WRITE_OUTPUT_TYPE.STREAM`;
- the `Uint8Array` return branch of `PptxGenJS.stream()`.

PptxGenJS returns `undefined` from section, layout, and master definitions,
stores user-defined layout names, accepts `writeFile({ fileName })`, and returns
byte buffers from both `stream()` and the misleading `STREAM` output selector.
Native returns editable section/layout identities, normalizes custom
presentation layout names, accepts `writeFile(path, options)`, and makes
`stream()` a real Node readable. PptxGenJS also ignores compression for explicit
JSZip output types and for `writeFile`, while native applies the legal request
consistently. Native exposes the six JSZip write selectors through `write()` and
the `STREAM` intent through the separate real-stream API.

### Defect exclusions: 3

- the `ArrayBuffer`, `Blob`, and `string` return branches of
  `PptxGenJS.stream()`.

PptxGenJS 4.0.1 always resolves `stream()` to a Node `Buffer` in Node, which is
a `Uint8Array`; the other declared branches cannot be selected and never occur.
Native's real readable stream is a strict correction, not evidence that the
three broken upstream return declarations are delivered capabilities.

## Evidence Architecture

- Adapter: add one aggregate control titled `locks the presentation root,
  output, and theme declarations against PptxGenJS 4.0.1`. Reuse existing
  output-catalog, real-stream, compression, layout, RTL, metadata, theme,
  section, and master controls.
- Native: reuse `PptxDocument.write`, `stream`, `writeFile`,
  `defineSlideMaster`, presentation metadata/theme tests, and section/slide
  lifecycle tests.
- npm/browser: reuse the packed output representation, Node readable stream,
  layout, theme, metadata, and lifecycle states. The persistent Chrome batch
  keeps the Node-only stream error boundary and later-write isolation.
- OOXML: reuse all-format metadata/theme/layout/master tests and inspect the
  retained lifecycle artifact once at the three-family batch boundary.
- Matrix: generate the exact 35 entries from frozen ID/status groups and assert
  the exact family and global counts.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 35 atoms move from `unverified`: 18 supported, 14
  deliberate-difference, and three defect-excluded.
- Global totals become 606 supported, 254 deliberate-difference, 86
  deprecated-alias, 355 defect-excluded, and 473 unverified: 1,301 of 1,774
  atoms closed (73.34%).
- The aggregate control locks method returns, custom layout naming, file output,
  real-stream behavior, compression divergence, the misleading `STREAM` token,
  ignored explicit/file compression, and all three impossible `stream()` return
  branches.
- Focused adapter/audit tests and deterministic matrix generation pass before
  review, one capability-family commit, and push.
- The worktree contains no staged package/browser/PPTX artifacts and local and
  remote main remain at divergence `0 0`.
