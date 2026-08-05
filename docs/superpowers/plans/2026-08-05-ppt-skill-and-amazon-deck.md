# PPT Skill and Amazon Biodiversity Deck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a repository-local PPT skill for `@jiayunxie/pptx` and prove it by generating and visually validating an Amazon rainforest biodiversity presentation.

**Architecture:** Keep the skill entry point concise and place detailed package guidance in the requested sibling `pptx.md`. Build the proof deck from editable native SDK objects, use generated photographic imagery for visual richness, and validate package structure plus rendered slide appearance before delivery.

**Tech Stack:** Markdown Agent Skill, `@jiayunxie/pptx@0.1.1`, Node.js ES modules, OpenAI image generation, LibreOffice/Poppler rendering, `pptx-inspect`.

## Global Constraints

- Create exactly `skills/ppt/SKILL.md` and `skills/ppt/pptx.md` inside the skill folder.
- Use Anthropic's PPTX skill for workflow and QA inspiration, but document this repository's APIs and verified boundaries.
- Generate a 16:9 English deck covering mammals, insects, plants, global ecosystem roles, and deforestation threats.
- Use lush green nature imagery, editable native slide content, speaker-note sources, and at least one render/fix/re-render cycle.
- Do not commit `.pnpm-store/`, temporary renders, or transient build artifacts.

---

### Task 1: Initialize and author the repository PPT skill

**Files:**
- Create: `skills/ppt/SKILL.md`
- Create: `skills/ppt/pptx.md`

**Interfaces:**
- Consumes: `@jiayunxie/pptx` public aggregate API and `pptx-inspect` CLI.
- Produces: `$ppt` trigger metadata and a detailed package workflow reference.

- [ ] **Step 1: Initialize the skill skeleton**

Run the skill-creator initializer with `ppt` under `skills/`, then remove generated UI metadata so the requested folder contains only the two specified files.

- [ ] **Step 2: Write the concise entry point**

Define broad PPT/PPTX creation, editing, inspection, and validation triggers. Require reading `pptx.md`, preserving source files, using semantic SDK operations, and completing package plus visual QA.

- [ ] **Step 3: Write the package reference**

Document installation, zero-input creation, opening/editing, units, slides, text, shapes, images/SVG, tables, charts, media, masters/layouts/placeholders, notes/metadata/sections, output, validation, and deliberate boundaries. Use compilable representative TypeScript examples.

- [ ] **Step 4: Validate the skill**

Run:

```sh
python /Users/jeremy/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/ppt
node --test scripts/ppt-skill.test.mjs
```

Expected: frontmatter PASS and reference assertions PASS.

- [ ] **Step 5: Review, commit, and push**

```sh
git add skills/ppt scripts/ppt-skill.test.mjs
git commit -m "docs: add native PPT creation skill"
git push origin main
```

### Task 2: Generate the Amazon visual system and deck source

**Files:**
- Create: `examples/amazon-biodiversity/assets/*.png`
- Create: `examples/amazon-biodiversity/generate.mjs`
- Create: `examples/amazon-biodiversity/sources.txt`
- Create: `examples/amazon-biodiversity/amazon-biodiversity.pptx`

**Interfaces:**
- Consumes: the new `skills/ppt/pptx.md`, generated raster assets, and public aggregate package.
- Produces: an eight-slide editable native PPTX and reproducible JavaScript source.

- [ ] **Step 1: Generate visual assets**

Create distinct wide photographic assets for canopy, jaguar, river dolphin, leafcutter ants, rainforest flora, atmospheric river, and deforestation. Save selected project-bound images under `examples/amazon-biodiversity/assets/` and record prompts/sources.

- [ ] **Step 2: Implement the deck**

Create a 13.333 × 7.5 inch presentation with forest/earth/mist palette, strong type hierarchy, varied full-bleed and split compositions, native charts/shapes/text, and `[Sources]` blocks in speaker notes.

- [ ] **Step 3: Generate the PPTX**

Run:

```sh
node examples/amazon-biodiversity/generate.mjs
```

Expected: `amazon-biodiversity.pptx` is written with eight slides and no SDK error diagnostics.

### Task 3: Validate, render, fix, and deliver

**Files:**
- Modify: `examples/amazon-biodiversity/generate.mjs`
- Modify: `examples/amazon-biodiversity/amazon-biodiversity.pptx`

**Interfaces:**
- Consumes: generated deck.
- Produces: validated and visually reviewed final deck.

- [ ] **Step 1: Inspect package and content**

```sh
pptx-inspect --json package inspect examples/amazon-biodiversity/amazon-biodiversity.pptx
pptx-inspect --json package validate examples/amazon-biodiversity/amazon-biodiversity.pptx --profile powerpoint-2010
pptx-inspect --json slides list examples/amazon-biodiversity/amazon-biodiversity.pptx
```

Expected: eight ordered slides, zero validation errors, and all requested topics present.

- [ ] **Step 2: Render and inspect every slide**

Render to a temporary directory, create a montage, and inspect individual full-size slides for overlap, clipping, contrast, crop quality, inconsistent spacing, and repetitive composition.

- [ ] **Step 3: Apply at least one concrete visual correction**

Update the source based on the first QA pass, regenerate the deck, and re-render all affected slides. Do not declare success without this fix-and-verify cycle.

- [ ] **Step 4: Run final gates**

Repeat package validation, slide listing, overflow checks, and visual inspection. Confirm every generated asset is project-local and every temporary render remains untracked.

- [ ] **Step 5: Review, commit, and push**

```sh
git add examples/amazon-biodiversity docs/superpowers/plans/2026-08-05-ppt-skill-and-amazon-deck.md
git commit -m "examples: add Amazon biodiversity presentation"
git push origin main
```
