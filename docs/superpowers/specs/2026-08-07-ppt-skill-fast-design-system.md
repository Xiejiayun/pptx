# PPT Skill Fast Design System

## Objective

Upgrade the repository's `ppt` skill so a presentation request can move from a natural-language query to a structurally validated, visually checked PPTX in no more than three minutes on the warm path. The skill must carry its own theme, layout, typography, spacing, and validation guidance. It must not spend generation time searching the web for stock images or other presentation assets.

## Scope

Change only:

- `skills/ppt/SKILL.md`
- `skills/ppt/pptx.md`
- `scripts/ppt-skill.test.mjs`

Keep `skills/ppt` limited to `SKILL.md` and `pptx.md`. Continue using the repository's public `@jiayunxie/pptx` APIs and `pptx-inspect` validation workflow.

## Source policy

The default fast path is offline and deterministic.

Asset priority is:

1. user-provided files;
2. assets already present in the repository or selected template;
3. native editable PowerPoint text, shapes, tables, charts, and diagrams;
4. simple local SVG or raster assets that can be derived from already available inputs.

Do not browse, search, scrape, or download stock imagery during ordinary deck generation. Do not call image search or image generation by default. External acquisition is a separate user-requested workflow and is outside the three-minute fast path.

When no photographic asset is available, use a designed native composition rather than a placeholder image, broken URL, or last-minute search.

## Theme system

Before creating slides, define one compact `ThemeSpec` containing:

- `mode`: light, dark, or sandwich;
- `primary`, `secondary`, `accent`, `background`, `surface`, `text`, and `mutedText` colors;
- a dominant visual weight of roughly 60–70% for the primary/background family;
- one repeated visual motif tied to the content;
- `displayFont` and `bodyFont`;
- title, section, body, caption, margin, and gap sizes.

Choose colors from the subject rather than defaulting to generic blue. Preserve strong text/background contrast. Use dark title and conclusion slides with lighter content slides by default when the topic does not require an all-dark treatment.

The reference will include several starting palettes inspired by Anthropic's skill, but they are starting points rather than templates. The agent must adapt the palette to the query.

## Typography system

Use fonts that render predictably in both Office and LibreOffice QA.

Default pairs:

- Cambria + Arial;
- Cambria + Calibri;
- Arial + Arial;
- Calibri + Calibri.

Do not default to Aptos. If the user specifies another font, honor it and reserve about 10% extra width for substitution variance.

Minimum sizes:

- cover title: 50–64 pt;
- slide title: 36–44 pt;
- section heading: 20–24 pt;
- body: 14–18 pt, with 16 pt preferred;
- captions and labels: 10–12 pt.

Reduce copy or split the slide before shrinking body text below the minimum.

## Layout system

Use a 16:9 canvas by default with:

- at least 0.5 inch outer margins;
- 0.3 or 0.5 inch internal gaps chosen once per deck;
- consistent alignment anchors and zero-margin text boxes when exact edge alignment is required;
- one dominant focal point per slide;
- left-aligned body copy.

The skill will define eight reusable layout families:

1. cinematic cover;
2. section divider;
3. statement plus hero visual;
4. asymmetric two-column;
5. large statistic plus explanation;
6. comparison;
7. process or timeline;
8. native chart or table plus takeaway.

Do not use the same family on consecutive content slides unless the content requires continuity. Do not default to title-and-bullets, repeated card grids, title underlines, decorative bars, or accent stripes.

Every slide must contain a meaningful visual element: an available image, native chart, table, diagram, icon-like shape, or large typographic statistic.

## Content density

Write one declarative takeaway per slide. Keep body copy to the amount that fits at the defined font size without aggressive shrinking. Prefer one main idea, up to three supporting points, and one visual hierarchy. Convert relationships to native diagrams and quantitative claims to native charts where appropriate.

## Three-minute execution budget

The warm-path budget starts when the query is available and ends when the final PPTX has passed content, structural, and visual checks.

| Stage | Budget |
|---|---:|
| Narrative outline | 20 seconds |
| Theme and layout assignment | 25 seconds |
| Native PPTX generation | 75 seconds |
| Reopen, validate, render, and inspect in parallel | 45 seconds |
| One repair and targeted recheck buffer | 15 seconds |
| Total | 180 seconds |

Use one generator execution. Run reopen/content checks, `pptx-inspect` validation, rendering, and overflow checks concurrently where possible. Re-render only changed slides after a repair. Do not launch multi-round review or perform full rebuilds for each slide.

If the first output has no defect, deliver it without inventing a correction. If a defect exists, make at most one targeted repair within the fast path. Report a missed SLA as a performance defect; do not hide structural or visual failures to meet the clock.

## Validation contract

Before delivery:

1. reopen the output with `PptxDocument.open()`;
2. verify slide count, order, titles, notes, alt text, and expected semantic objects;
3. run `pptx-inspect package validate` with the applicable profile;
4. render all slides once and inspect a montage for overflow, overlap, crop, contrast, wrapping, spacing, and repetition;
5. run the available automated overflow check;
6. after a repair, repeat only the affected semantic and visual checks plus final package validation.

Source notes remain required for claims and for any supplied or local third-party assets. Native decorative shapes do not need asset citations.

## Test coverage

Extend `scripts/ppt-skill.test.mjs` to assert that the skill:

- defines the theme contract and palette roles;
- defines the eight layout families;
- defines safe font pairs and rejects Aptos as the default;
- includes the 180-second stage budget;
- forbids web or stock-asset search on the default path;
- allows only one defect-driven repair rather than requiring a correction every time;
- preserves the existing native create, reopen, edit, chart, notes, and slide-number runtime coverage.

## Success criteria

- The skill can select a coherent theme, layout sequence, and typography system from the query without external asset discovery.
- The default generation workflow is explicitly offline and bounded to 180 seconds.
- Existing public API examples remain accurate and executable.
- Skill validation and Node tests pass.
- The capability family completes review, commit, and push without staging unrelated `.pnpm-store/` or `artifacts/` content.
