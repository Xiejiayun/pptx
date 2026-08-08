---
name: ppt
description: Create, edit, read, inspect, validate, render, or troubleshoot PowerPoint presentations and PPTX files. Use for slide decks, presentations, templates, masters, layouts, speaker notes, charts, tables, media, or lossless PowerPoint editing with @jiayunxie/pptx.
---

# PowerPoint

Use `@jiayunxie/pptx` for native creation and lossless semantic editing. Read [pptx.md](pptx.md) completely before creating or editing a presentation.

## Route the request

- Create: use the 180-second fast path below and build editable slides from a query plus available local inputs.
- Edit: open with `PptxDocument.open()`, change only requested semantic objects, and save to a separate output path.
- Inspect: use `pptx-inspect` to list slides, inspect parts, validate OOXML, or diff two files.
- Troubleshoot: reproduce, inspect diagnostics and relationships, make the narrowest semantic edit, and verify mutation isolation.

## Default source policy

Resolve assets in this order: user-provided files; repository or template assets; native editable text, shapes, tables, charts, and diagrams; then simple local SVG or raster derivations from available inputs.

Do not browse, search, scrape, or download stock assets on the default path. Do not call image search or image generation by default. If no photograph is available, create an intentional native composition; never insert a placeholder or broken URL. External acquisition requires an explicit user request and is outside the 180-second fast path.

## 180-second create path

1. In 20 seconds, derive the audience, purpose, narrative arc, one declarative takeaway per slide, and a compact `ThemeSpec` from the query and available files.
2. In 25 seconds, express the selected layout families as one `DeckSpec`. Call `assertDeckSpec()` before creating the presentation; fix every diagnostic in the spec rather than discovering basic geometry and title-fit defects after rendering.
3. In 75 seconds, run one generator execution. Keep text, diagrams, tables, and charts editable and use semantic APIs rather than raw OOXML. Create straight lines from endpoints with `connectorTransform()`; never emit zero-width or zero-height transforms.
4. In 45 seconds, run `node scripts/ppt-fast-qa.mjs <deck.pptx> --out-dir <new-run-dir> --expected-slides <count> --max-warnings 0 --json`. It reopens, validates, renders, checks overflow, and builds a montage concurrently with the bundled runtime.
5. Inspect the full-size rendered slides and montage. Reserve 15 seconds for at most one targeted repair when inspection finds a concrete defect. Recheck affected content and run the QA command into another new directory. If no defect exists, deliver the first valid output without inventing a change.

## Editing and delivery contract

- Preserve all inputs and unknown package content. Never overwrite an input unless explicitly requested; always use a separate output path by default.
- Prefer semantic APIs for slides, text, shapes, images, SVG, tables, charts, media, notes, masters, layouts, and metadata.
- Add a `[Sources]` block to speaker notes for factual claims and supplied or local third-party assets. Native decorative shapes need no citation.
- Verify slide count, order, titles, notes, alt text, links, data, metadata, and expected objects after reopening.
- Render every slide and inspect readable images for clipping, overlap, crop, contrast, wrapping, spacing, hierarchy, and repetition. A successful write or zero-error package is not visual proof.
- Keep the query, `DeckSpec`, generator source, PPTX, QA JSON, rendered slides, and montage together for a timed acceptance run. Start the task-cold clock before the query is handed to the generating agent and stop only after the final visual verdict is recorded; a replay of an existing generator is not query-to-PPT evidence.
- For edits, diff against the untouched source and investigate unexpected changed parts before delivery.
- Treat a missed 180-second warm-path target as a performance defect; never conceal structural or visual failures to meet the clock.

## Boundaries

The package implements the audited PptxGenJS 4.0.1 public capability surface and adds semantic editing. Do not claim broad PowerPoint, Keynote, Google Slides, or LibreOffice certification without current independent evidence. Treat unsupported provider-specific media metadata, rich notes-page editing, comments, captions, transcoding, and other native-client extensions as explicit boundaries rather than silently approximating them.
