---
name: ppt
description: Create, edit, read, inspect, validate, render, or troubleshoot PowerPoint presentations and PPTX files. Use for slide decks, presentations, templates, masters, layouts, speaker notes, charts, tables, media, or lossless PowerPoint editing with @jiayunxie/pptx.
---

# PowerPoint

Use `@jiayunxie/pptx` for native creation and lossless semantic editing. Standard from-zero creation uses the fast compiler below and does not require reading [pptx.md](pptx.md). Read `pptx.md` completely for edits, unsupported compiler layouts, or API-level troubleshooting.

## Route the request

- Create: use the 180-second fast path below and build editable slides from a query plus available local inputs.
- Edit: open with `PptxDocument.open()`, change only requested semantic objects, and save to a separate output path.
- Inspect: use `pptx-inspect` to list slides, inspect parts, validate OOXML, or diff two files.
- Troubleshoot: reproduce, inspect diagnostics and relationships, make the narrowest semantic edit, and verify mutation isolation.

## Default source policy

Resolve assets in this order: user-provided files; repository or template assets; native editable text, shapes, tables, charts, and diagrams; then simple local SVG or raster derivations from available inputs.

Do not browse, search, scrape, or download stock assets on the default path. Do not call image search or image generation by default. If no photograph is available, create an intentional native composition; never insert a placeholder or broken URL. External acquisition requires an explicit user request and is outside the 180-second fast path.

## 180-second create path

1. Before generating content, start a fresh task-cold acceptance run. The supervisor copies the query and starts a monotonic clock before the generating agent begins:

   ```sh
   node scripts/ppt-fast-accept.mjs begin \
     --run-dir <new-run-directory> \
     --query-file <query-source.txt> \
     --expected-slides <count> \
     --sla-ms 180000
   ```

   Treat this task-cold clock as authoritative; do not reset or pause it between content generation, build, visual review, and finalization.

2. In one reasoning pass and at most 45 seconds, derive the audience, purpose, narrative arc, one declarative takeaway per slide, and a compact `ThemeSpec`; then write only valid JSON to `<run-directory>/content.json`. Do not read `pptx.md`, browse the repository, search for assets, draft prose in another file, or inspect/reuse an older topic-specific content file, generator, or deck during a task-cold run. Do not ask another agent to rewrite or polish the JSON before the first build.
3. Run `node scripts/ppt-fast-accept.mjs build --run-dir <run-directory>`. It wraps `scripts/ppt-fast-create.mjs` and `scripts/ppt-fast-qa.mjs`, performs one compiler execution, writes `deck-spec.json` before serialization, creates `deck.pptx`, and runs reopen, OOXML validation, rendering, overflow checks, and montage creation concurrently in `qa-1/`.
4. Inspect every full-size PNG and the montage. Copy `review-template.json`, replace its pending values with the actual review, and preserve every supplied hash. Record exact slide coverage, query/content consistency, sources, and unresolved issues. If a concrete defect exists, make at most one targeted content/compiler repair and rerun `build` with `--force-repair`; discard the old review and use the new template for `qa-2/`.

   ```json
   {
     "verdict": "pass",
     "inspectedSlides": [1, 2, 3, 4, 5, 6, 7, 8, 9],
     "inspectedMontage": true,
     "querySatisfied": true,
     "contentConsistent": true,
     "sourcesPresent": true,
     "issues": [],
     "pptxSha256": "preserve from review-template.json",
     "reviewManifestSha256": "preserve from review-template.json",
     "qaResultSha256": "preserve from review-template.json",
     "montageSha256": "preserve from review-template.json",
     "renderedSlideSha256s": ["preserve every slide hash from review-template.json"]
   }
   ```
5. Run `node scripts/ppt-fast-accept.mjs finalize --run-dir <run-directory> --review-file <review.json>`. Deliver only when the atomic `final-verdict.json` says `verdict: "pass"`, `qualityPass: true`, `slaPass: true`, and `elapsedMs < 180000`.

The QA runner's 45-second value is a stage budget recorded as `qaBudgetPass`; it is diagnostic. The acceptance gate is the complete query-to-verdict 180-second SLA, while structural QA failures remain quality failures.

## Fast compiler contract

Write one JSON object with `title`, optional `author`, optional theme colors/fonts, and `slides`. Each slide requires `family`, `title`, optional `kicker`, a family-specific payload, and a `sources` URL array. Use 8–10 slides and at least five silhouettes.

For the predictable nine-slide cold path, use exactly `cover → roles → stats → spotlight → branches → bands → chart → process → actions`, adapting the narrative jobs to the query. This sequence is a performance skeleton, not topic-specific content. Keep typical Latin-script prose within these operational character targets, including spaces; DeckSpec preflight remains authoritative for actual glyph widths:

| Field | Maximum characters |
|---|---:|
| Any non-cover slide title / kicker | 40 / 36 |
| Cover title / subtitle | 30 / 90 |
| `bands` heading / body / detail | 13 / 64 / 34 |
| `spotlight.hero` heading / subheading / body | 13 / 20 / 90 |
| `spotlight.items` heading / body | 22 / 74 |
| `roles` heading / body / footer | 18 / 64 / 45 |
| `branches` heading / body / callout | 22 / 74 / 48 |
| `stats` value / unit / heading / body | 9 / 20 / 19 / 62 |
| `chart.name` / category / callout value / heading / body | 30 / 10 / 9 / 22 / 62 |
| `process` heading / body / footer | 20 / 65 / 75 |
| `actions` heading / body | 30 / 72 |

Prefer complete short sentences over filling the budget. Make `chart.callout.value` a number, percentage, or number plus short unit; never use prose there. Never compensate for excess copy by requesting smaller fonts. All theme colors must be six-digit hexadecimal values without `#`; choose a light `background`, a readable `surface`, and a dark `deep`. The compiler selects accessible text colors automatically.

- `cover`: `subtitle`.
- `bands`: four `rows` with `heading`, `body`, and optional `detail`.
- `spotlight`: one `hero` plus up to three `items`.
- `roles`, `branches`, `process`, or `actions`: three or four `items` with `heading` and `body`.
- `stats`: three `items` with `value`, `unit`, `heading`, and `body`.
- `chart`: `chart` with `name`, `categories`, and numeric `values`, plus a `callout` with `value`, `heading`, and `body`.

For a standalone compiler regression check outside a timed acceptance run:

```sh
node scripts/ppt-fast-create.mjs \
  --input <content.json> \
  --output <deck.pptx> \
  --deck-spec-out <deck-spec.json>
```

The timed `build` command wraps this compiler and QA; do not call them separately inside an acceptance run. Do not write a bespoke generator unless the request cannot be expressed by these families. That exception leaves the standard fast path and requires the full `pptx.md` reference.

## Editing and delivery contract

- Preserve all inputs and unknown package content. Never overwrite an input unless explicitly requested; always use a separate output path by default.
- Prefer semantic APIs for slides, text, shapes, images, SVG, tables, charts, media, notes, masters, layouts, and metadata.
- Add a `[Sources]` block to speaker notes for factual claims and supplied or local third-party assets. Native decorative shapes need no citation.
- Verify slide count, order, titles, notes, alt text, links, data, metadata, and expected objects after reopening.
- Render every slide and inspect readable images for clipping, overlap, crop, contrast, wrapping, spacing, hierarchy, and repetition. A successful write or zero-error package is not visual proof.
- Keep the query, content JSON, `DeckSpec`, PPTX, QA JSON, rendered slides, montage, visual review, hashes, and `final-verdict.json` in the acceptance directory. A replay of an existing generator is not query-to-PPT evidence.
- For edits, diff against the untouched source and investigate unexpected changed parts before delivery.
- Treat a missed 180-second warm-path target as a performance defect; never conceal structural or visual failures to meet the clock.

## Boundaries

The package implements the audited PptxGenJS 4.0.1 public capability surface and adds semantic editing. Do not claim broad PowerPoint, Keynote, Google Slides, or LibreOffice certification without current independent evidence. Treat unsupported provider-specific media metadata, rich notes-page editing, comments, captions, transcoding, and other native-client extensions as explicit boundaries rather than silently approximating them.
