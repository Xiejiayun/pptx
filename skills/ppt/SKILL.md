---
name: ppt
description: Create, edit, read, inspect, validate, render, or troubleshoot PowerPoint presentations and PPTX files. Use for slide decks, presentations, templates, masters, layouts, speaker notes, charts, tables, media, or lossless PowerPoint editing with @jiayunxie/pptx.
---

# PowerPoint

Use `@jiayunxie/pptx` for native creation and lossless semantic editing. Read [pptx.md](pptx.md) completely before creating or editing a presentation.

## Route the request

- Create: start with `PptxDocument.create()`, define the narrative and visual system, then build editable slides.
- Edit: open the source with `PptxDocument.open()`, change only the requested semantic objects, and save to a different output path.
- Inspect: use `pptx-inspect` to list slides, inspect package parts, validate OOXML, or diff before and after files.
- Troubleshoot: reproduce the issue, inspect relationships and diagnostics, make the narrowest semantic edit, and verify mutation isolation.

## Required workflow

1. Confirm the audience, purpose, source material, output path, and intended slide format from the request and available files. Make reasonable assumptions when details are absent.
2. Preserve every source file. Never overwrite an input presentation unless the user explicitly requests it; write to a separate output path.
3. Outline one clear message per slide. Use a coherent story, varied compositions, and an intentional visual system rather than a repeated grid of cards.
4. Prefer semantic APIs for slides, text, shapes, images, tables, charts, media, notes, masters, and metadata. Do not edit raw OOXML when the public model can express the change.
5. Keep text and data editable. Use raster imagery for photographs and textures, not for ordinary text or charts.
6. Add a `[Sources]` block to speaker notes for factual claims and externally sourced or generated assets. Use one source per line with enough detail to identify it.
7. Write the PPTX, reopen it, and run content checks plus `pptx-inspect package validate` with the relevant compatibility profile.
8. Render every slide, inspect the montage and each slide at readable resolution, and check clipping, overlaps, contrast, crops, title wrapping, repetition, and visual hierarchy.
9. Make at least one concrete correction based on the first render, then regenerate, revalidate, and rerender the final deck.

## Quality bar

- Use a 16:9 layout unless the source, template, or request requires another format.
- Use large, legible typography and safe margins. As a practical floor, use about 50 pt for title slides, 35 pt for slide titles, 24 pt for subheads, and 16 pt for body copy.
- Favor diagrams, photographs, native charts, and spatial composition over dense prose.
- Ensure every slide has a dominant focal point and no accidental alignment or spacing inconsistencies.
- Verify facts, labels, numbers, source notes, slide order, and file metadata as well as visual appearance.

## Boundaries

The package implements the audited PptxGenJS 4.0.1 public capability surface and adds semantic editing. Do not claim broad PowerPoint, Keynote, Google Slides, or LibreOffice certification without current independent evidence. Treat unsupported provider-specific media metadata, rich notes-page editing, comments, captions, transcoding, and other native-client extensions as explicit boundaries rather than silently approximating them.
