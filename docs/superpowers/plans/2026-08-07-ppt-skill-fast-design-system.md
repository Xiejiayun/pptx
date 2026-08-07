# PPT Skill Fast Design System Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the repository `ppt` skill choose a complete offline design system and produce a reopened, structurally validated, visually inspected PPTX within a 180-second warm path.

**Architecture:** Keep `SKILL.md` as the concise routing and execution contract, and keep all reusable design-system rules and public API examples in `pptx.md`. Extend the existing Node test so documentation requirements are machine-checked while retaining its real create/reopen/edit runtime probe.

**Tech Stack:** Markdown skill instructions, Node.js `node:test`, `@jiayunxie/pptx`, `pptx-inspect`, skill-creator validation.

## Global Constraints

- Keep `skills/ppt` limited to exactly `SKILL.md` and `pptx.md`.
- Use only public `@jiayunxie/pptx` APIs and the repository `pptx-inspect` workflow.
- Do not browse, search, scrape, download, or generate presentation assets on the default path.
- Use a 16:9 default canvas, at least 0.5-inch margins, and a fixed 0.3- or 0.5-inch internal gap.
- Do not default to Aptos; use the approved Cambria, Arial, and Calibri font pairs.
- Bound the warm path to 180 seconds and allow at most one targeted repair when an actual defect exists.
- Do not stage `.pnpm-store/` or `artifacts/`.

---

### Task 1: Lock the design-system contract in tests

**Files:**
- Modify: `scripts/ppt-skill.test.mjs`
- Test: `scripts/ppt-skill.test.mjs`

**Interfaces:**
- Consumes: Markdown text from `skills/ppt/SKILL.md` and `skills/ppt/pptx.md`.
- Produces: Assertions for `ThemeSpec`, semantic palette roles, typography, eight layouts, offline asset sourcing, 180-second timing, and defect-driven repair.

- [ ] **Step 1: Replace the mandatory-correction assertion and add the new contract assertions**

```js
assert.match(combined, /ThemeSpec/);
assert.match(combined, /primary.*secondary.*accent.*background.*surface.*text.*mutedText/s);
assert.match(combined, /Cambria \+ Arial/);
assert.match(combined, /Do not default to Aptos/);
assert.match(combined, /cinematic cover/);
assert.match(combined, /native chart or table plus takeaway/);
assert.match(combined, /180 seconds/);
assert.match(combined, /Do not browse, search, scrape, or download/);
assert.match(combined, /at most one targeted repair/);
assert.doesNotMatch(combined, /at least one concrete (?:correction|improvement)/i);
```

- [ ] **Step 2: Run the focused test and verify the new assertions fail**

Run: `node --test scripts/ppt-skill.test.mjs`

Expected: the Markdown contract test fails because the current skill lacks `ThemeSpec`, the complete layout catalog, the 180-second budget, and defect-driven repair wording; the runtime creation/edit test still passes.

- [ ] **Step 3: Preserve runtime coverage while removing Aptos from its example**

Change the runtime probe's display font and corresponding assertion from `Aptos Display` to `Cambria`. Leave create, chart replacement, notes, slide-number, reopen, edit, and second reopen assertions intact.

### Task 2: Implement the fast-path router

**Files:**
- Modify: `skills/ppt/SKILL.md`
- Test: `scripts/ppt-skill.test.mjs`

**Interfaces:**
- Consumes: The detailed `ThemeSpec`, layouts, asset policy, and API guidance in `pptx.md`.
- Produces: A concise, mandatory create/edit/inspect workflow with one-generator execution and bounded QA.

- [ ] **Step 1: Replace the current generic workflow with the approved fast path**

Document these ordered stages in `SKILL.md`: read `pptx.md`, resolve local inputs, outline in 20 seconds, assign theme/layouts in 25 seconds, generate once in 75 seconds, reopen/validate/render/inspect concurrently in 45 seconds, and use a 15-second repair buffer only for a concrete defect.

- [ ] **Step 2: Add the source policy and delivery contract**

State the four-level local asset priority, prohibit browsing/search/scraping/downloading and default image generation, require native compositions when imagery is unavailable, preserve separate output paths, require `[Sources]`, and deliver without an invented correction when the first output passes.

- [ ] **Step 3: Keep edit and inspection behavior explicit**

Retain semantic-only editing, mutation isolation, reopen checks, rendering, package validation, compatibility boundaries, and the requirement to preserve unknown package content.

### Task 3: Implement the reusable design system and API guidance

**Files:**
- Modify: `skills/ppt/pptx.md`
- Test: `scripts/ppt-skill.test.mjs`

**Interfaces:**
- Consumes: A natural-language query plus optional user/repository/template assets.
- Produces: A concrete `ThemeSpec`, slide-to-layout assignment, native editable deck code, and bounded validation workflow.

- [ ] **Step 1: Add a table of contents and source policy near the top**

Include direct links to source policy, `ThemeSpec`, typography, spacing and layout families, creation/editing APIs, fast execution, validation, and deliberate boundaries. Make the default offline/no-search behavior unambiguous.

- [ ] **Step 2: Define `ThemeSpec` and a valid package example**

Define `mode`, `primary`, `secondary`, `accent`, `background`, `surface`, `text`, `mutedText`, `motif`, `displayFont`, `bodyFont`, sizes, margin, and gap. Show a query-adapted example and map its values to valid `PptxDocument.create()`, rich-text, fill, line, and background fields rather than inventing a package-level theme API.

- [ ] **Step 3: Define typography, spacing, and composition rules**

List all four approved font pairs, the 50–64/36–44/20–24/14–18/10–12 point ranges, the 10% substitution allowance, 16:9 canvas, 0.5-inch margins, fixed gaps, zero-margin alignment boxes, left-aligned body text, one focal point, and 60–70% dominant color weight.

- [ ] **Step 4: Define all eight layout families and anti-patterns**

Describe cinematic cover, section divider, statement plus hero visual, asymmetric two-column, large statistic plus explanation, comparison, process or timeline, and native chart or table plus takeaway. Require layout variation and a meaningful visual on every slide; reject repeated title-and-bullets, repeated card grids, title underlines, decorative bars, and accent stripes.

- [ ] **Step 5: Replace the open-ended QA loop with the 180-second workflow**

Document the five exact timing stages, concurrent verification, one render of all slides, targeted rerender after a concrete defect, and at most one targeted repair. Keep content, OOXML, overflow, and visual checks, but remove the requirement to manufacture a correction.

- [ ] **Step 6: Retain accurate creation and editing examples**

Keep examples for `PptxDocument.create`, `PptxDocument.open`, rich text, shapes, images, SVG, tables, `tableToSlides`, charts, media, masters/layouts, notes, slide numbers, output types, inspection, validation, and deliberate boundaries. Replace Aptos defaults in examples with approved safe fonts.

### Task 4: Verify, review, and publish the capability family

**Files:**
- Modify: `docs/superpowers/plans/2026-08-07-ppt-skill-fast-design-system.md` only to mark completed checkboxes.

**Interfaces:**
- Consumes: Updated skill documentation and focused runtime test.
- Produces: Passing skill validation, passing Node tests, clean diff review, one capability-family commit, and pushed branch.

- [ ] **Step 1: Run focused tests**

Run: `node --test scripts/ppt-skill.test.mjs`

Expected: both the documentation contract and runtime create/reopen/edit tests pass.

- [ ] **Step 2: Run skill validation**

Run: `python /Users/jeremy/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/ppt`

Expected: `Skill is valid!`. If the base interpreter lacks PyYAML, install it into a temporary directory and set `PYTHONPATH` only for this command.

- [ ] **Step 3: Review the exact change set**

Run: `git diff --check` and `git diff -- skills/ppt/SKILL.md skills/ppt/pptx.md scripts/ppt-skill.test.mjs docs/superpowers/plans/2026-08-07-ppt-skill-fast-design-system.md`.

Expected: no whitespace errors, no placeholders, no mandatory correction language, no accidental unrelated changes, and no inaccurate public API examples.

- [ ] **Step 4: Commit and push the capability family**

```bash
git add skills/ppt/SKILL.md skills/ppt/pptx.md scripts/ppt-skill.test.mjs docs/superpowers/plans/2026-08-07-ppt-skill-fast-design-system.md
git commit -m "feat: add fast offline ppt design system"
git push origin main
```
