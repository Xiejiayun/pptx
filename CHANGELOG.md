# Changelog

## Unreleased

- Added explicit detection and public profiles for `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, and `.potm` packages.
- Replaced host `node:path` semantics in the OPC graph, model, codecs, and transition plugin with portable OPC part URI operations.
- Added a browser conditional export with Blob/File/Web Stream input, Blob/download output, portable media hashing, and browser-safe SmartArt IDs.
- Added nested synchronous mutation transactions with rollback across parts, content types, relationships, ZIP state, and the mutation journal.
- Added stable object identity for slide, shape, master, layout, and theme models while keeping properties live against current OOXML.
- Added lifecycle-aware slide duplication and deletion: owned dependency subgraphs are cloned/collected while layout, image, media, external, and opaque targets remain shared or preserved.
- Added clone-on-write isolation for shared image payload and raw chart XML edits, including chart-owned workbook subgraphs and transaction rollback.
- Added native zero-input presentation creation for all six formats, four PptxGenJS-compatible slide sizes, and blank slides linked to a canonical master/layout/theme chain.
- Added transactional single-paragraph text-box creation with stable shape identity, XML-safe content, unique IDs, and editable transforms.
- Added custom OOXML-valid slide dimensions at presentation creation, matching PptxGenJS custom-layout output semantics.
- Added lossless slide-canvas size reading and editing without silently scaling shapes or changing the notes page.
- Added PptxGenJS-compatible plain-text paragraphs with normalized line endings, empty-line preservation, and transactional text overwrite.
- Added structured rich-text run creation, snapshots, and replacement with fonts, point sizes, bold/italic, sRGB/theme colors, and soft breaks.
- Added PptxGenJS-compatible left, center, right, and justified paragraph alignment for plain and structured text creation and editing.
- Added paragraph bullets and automatic numbering with Unicode characters, all 16 PptxGenJS numbering styles, start values, and point-based hanging indents.
- Added paragraph before/after spacing plus exact and multiple line spacing with creation defaults, per-paragraph overrides, and lossless OOXML editing.
- Added zero-based paragraph list levels 0–8 with PptxGenJS-compatible nested bullet and numbering indents.
- Added paragraph tab stops with left, center, right, and decimal alignment, creation defaults, per-paragraph overrides, and lossless OOXML editing.
- Added rich-text underline creation and editing with all 17 valid OOXML styles, explicit none, independent sRGB/theme colors, and strict lossless reads.
- Added PptxGenJS-compatible single and double rich-text strike creation, explicit inherited-style suppression, editing, and strict lossless reads.
- Added rich-text highlight creation and editing with normalized sRGB/theme colors, strict reads, and PptxGenJS output conformance.
- Added solid rich-text outline creation and editing with point widths, sRGB/theme colors, strict reads, and PptxGenJS output conformance.
- Added rich-text glow creation and editing with point radii, opacity, sRGB/theme colors, strict effect-list reads, and PptxGenJS output conformance.
- Added rich-text superscript, subscript, and custom percentage baseline creation and editing with explicit normal-baseline suppression and strict reads.
- Added point-based rich-text character spacing creation and editing with explicit zero, strict Int32 reads, and PptxGenJS output conformance.
- Added rich-text main-fill transparency creation and editing with 0–100 percentages, 0.001% quantization, strict direct-alpha reads, and PptxGenJS output conformance.
- Added PptxGenJS-compatible outer and run-level text languages with direct reads, strict validation, inheritance, and XML-safe serialization.
- Added PptxGenJS-compatible paragraph RTL creation defaults, per-paragraph true/false overrides, strict direct reads, editing, and clearing.
- Added PptxGenJS-compatible presentation RTL creation with strict direct root reads, editing, explicit false, and clearing.
- Added point-based non-list paragraph left margins with creation defaults, strict direct reads, editing, clearing, and bullet isolation.
- Added point-based paragraph right margins with creation defaults, strict direct reads, editing, clearing, and list coexistence.
- Added signed point-based ordinary paragraph first-line/hanging indents with creation defaults, strict direct reads, editing, clearing, and bullet isolation.
- Added point-based text-box margin creation and direct four-side editing with scalar, documented TRBL tuple, named-object, and PptxGenJS output support.
- Added PptxGenJS-compatible top, middle, and bottom text-box vertical alignment with direct lossless editing and clearing.
- Added PptxGenJS-compatible text-box wrapping with strict boolean creation plus direct lossless editing and clearing.
- Added all seven PptxGenJS text-box directions with strict creation plus direct lossless editing and clearing.
- Added PptxGenJS-compatible four-value table-cell text-direction snapshots and physical-cell direct editing with explicit horizontal and clear operations.
- Added strict table-cell text-fit snapshots and physical-cell direct editing for existing `noAutofit`, `normAutofit`, and `spAutoFit` choices.
- Added PptxGenJS-compatible table-cell top, middle, and bottom vertical-alignment snapshots plus physical-cell direct editing and clearing.
- Added PptxGenJS-compatible none, shrink, and resize text-box fit modes with direct lossless editing.

## 0.1.0 - 2026-07-25

- Added lossless OOXML source-span patching and OPC package graph.
- Added editable slide, shape, text, image, table, and chart semantic models.
- Added slide create/copy/delete/reorder operations.
- Added PptxGenJS `^4.0.1` public-output adapter.
- Added Master/Layout/Theme, Gradient/Transparency, and Audio/Video codecs.
- Added compatibility diagnostics, security budgets, CLI, testkit, fuzz and performance harnesses.
- Added optional Transition, Animation/Timing, Advanced Charts, and SmartArt plugins.
- Added the self-contained `@jiayunxie/pptx` package with a single API entry, namespaced optional plugins, and the `pptx-inspect` binary.
- Added npm Trusted Publishing through GitHub Actions with OIDC provenance and the `next` dist-tag.
