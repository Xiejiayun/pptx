# Repository PPT skill and Amazon biodiversity deck design

## Goal

Add a reusable repository skill at `skills/ppt` that teaches an agent to create, inspect, and losslessly edit PPTX files with `@jiayunxie/pptx`. Prove the skill with a visually rich presentation about Amazon rainforest biodiversity.

## Structure

The skill contains exactly the two requested files:

- `skills/ppt/SKILL.md`: concise trigger metadata, routing, mandatory workflow, safety rules, and QA loop.
- `skills/ppt/pptx.md`: detailed package-specific creation and editing reference, representative TypeScript examples, supported capability families, deliberate boundaries, and troubleshooting guidance.

`SKILL.md` stays below 500 lines and directs the agent to read `pptx.md` whenever it must create or modify a deck. No copied Anthropic scripts, bundled binaries, assets, or unrelated documentation are added.

## Reference adaptation

Use the Anthropic PPTX skill as a workflow and quality reference, not as an API source. Preserve its strongest general practices: content planning before implementation, topic-specific visual language, varied layouts, strong hierarchy, deliberate spacing, content and visual QA, and at least one fix-and-verify pass.

Replace its PptxGenJS/raw-XML workflow with this repository's verified capabilities:

- Zero-input creation and all six OOXML presentation formats.
- Lossless open/edit/write behavior with unknown XML preservation.
- Native slides, layouts, masters, placeholders, shapes, rich text, tables, charts, images, SVG, media, backgrounds, metadata, notes, sections, slide numbers, and hyperlinks.
- Table auto-pagination and `tableToSlides` support.
- Node and browser output types, compression controls, and stream/file/blob output.
- `pptx-inspect` package inspection, validation, diff, exact-part reads, and narrow dry-run edits.
- PptxGenJS 4.0.1 surface parity where certified, with deliberate differences and defects documented rather than copied.

The reference must not claim unsupported client certification or recommend unsafe direct OOXML mutation as the default.

## Operational workflow

For creation, require the agent to plan the narrative and visual system, create a new `PptxDocument`, select a layout, add editable native content, write to a new output path, inspect and validate the package, render slides, visually review them, fix issues, and re-verify.

For editing, require the agent to preserve the source, open bytes through the SDK, use semantic live models, write a separate output file, run package diff/validation, and confirm that unrelated parts remain stable. Raw part reading is diagnostic only.

Examples use the public aggregate package `@jiayunxie/pptx@0.1.1` and compile against its exported declarations. Detailed examples focus on representative patterns rather than duplicating the complete API reference.

## Amazon biodiversity proof deck

Create a 16:9 English presentation from the supplied query. Use a lush nature-inspired system with deep forest green, canopy green, warm earth, mist, and restrained amber accents. Use strong photography or generated nature imagery, editable native text/shapes/charts, and varied compositions.

The narrative contains approximately eight slides:

1. Cinematic title and thesis.
2. Biodiversity-at-a-glance statistics and biome layers.
3. Mammals, including jaguar, Amazon river dolphin, and sloth.
4. Insects and other invertebrates, emphasizing ecological functions.
5. Plants and vertical forest structure.
6. The Amazon's role in global water, carbon, and climate systems.
7. Deforestation threats and cascading impacts.
8. Protection priorities and a memorable conclusion.

Use source notes or compact citations for factual claims. Avoid decorative clutter, repetitive card grids, text-only slides, and claims that cannot be supported. Keep body copy concise enough for live presentation.

## Validation

Validate the skill frontmatter and folder naming with the skill-creator validator. Check every code block or a consolidated equivalent against the installed package declarations.

For the deck:

- Inspect and validate the PPTX package with the repository CLI/validator.
- Extract or inspect slide content to confirm titles, ordering, and required topics.
- Render every slide to images and inspect for overlap, clipping, contrast, alignment, spacing, and visual repetition.
- Perform at least one concrete fix followed by a second render of affected slides.
- Keep the generated deck and any intentional source assets in a clear repository output location; exclude temporary render/build artifacts.

## Completion criteria

- Both requested skill files exist and contain no placeholders or unsupported capability claims.
- The skill validator and code/example checks pass.
- The Amazon deck opens as a valid PPTX, covers every requested topic, is visually rich, and passes the fix-and-verify loop.
- Skill, intentional deck artifact, and supporting source are reviewed, committed, and pushed without including `.pnpm-store/` or temporary artifacts.
