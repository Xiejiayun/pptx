# Changelog

## Unreleased

- Added explicit detection and public profiles for `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, and `.potm` packages.
- Replaced host `node:path` semantics in the OPC graph, model, codecs, and transition plugin with portable OPC part URI operations.
- Added a browser conditional export with Blob/File/Web Stream input, Blob/download output, portable media hashing, and browser-safe SmartArt IDs.
- Added nested synchronous mutation transactions with rollback across parts, content types, relationships, ZIP state, and the mutation journal.
- Added stable object identity for slide, shape, master, layout, and theme models while keeping properties live against current OOXML.
- Added lifecycle-aware slide duplication and deletion: owned dependency subgraphs are cloned/collected while layout, image, media, external, and opaque targets remain shared or preserved.

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
