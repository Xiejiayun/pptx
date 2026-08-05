# Placeholder Text Style Surface Family Implementation Plan

> **For agentic workers:** Evidence agents write only to `/tmp`; the main
> agent owns repository edits, review, matrix generation, commit, and push.

**Goal:** Close four placeholder text-style atoms through canonical text owners
and one shared cross-runtime lifecycle probe.

**Architecture:** Reuse paragraph alignment, text margins, rich-run alpha, and
vertical alignment. Invoke the same lifecycle probe from workspace, packed npm,
and persistent Chrome. Agent A/B/C produce isolated `/tmp` evidence shards in
parallel; the main agent is the only repository writer and merges one reviewed
capability-family commit.

**Tech Stack:** TypeScript, Vitest, PptxGenJS 4.0.1, npm pack, Chrome, OOXML,
and `pptx-inspect`.

## Global Constraints

- Close exactly four atoms as three supported / one deliberate difference.
- Add no placeholder-only product aliases or codecs.
- Run focused once and npm/browser/full gates once for the family.
- Evidence agents do not modify repository files.
- Review, commit, and push the capability family as one unit.
- Reuse existing tests, documentation, package smoke checks, and retained batch
  artifacts before adding or rerunning evidence.

---

### Task 1: Aggregate upstream and native semantics

- [x] Lock all legal align, valign, transparency, and margin cases in one
  PptxGenJS 4.0.1 control.
- [x] Add the shared native lifecycle probe with six-format and exact OOXML gates.
- [x] Integrate the probe with packed npm and persistent browser batches.

**Parallel ownership:**

- Agent A: exact manifest IDs/status split, reusable evidence anchors, generated
  report equivalence, and focused-test review in `/tmp`.
- Agent B: PptxGenJS runtime/control behavior, caller-mutation boundary, and
  legal/invalid case comparison in `/tmp`.
- Agent C: native lifecycle, actual tarball, persistent Chrome, exact OOXML, and
  retained-artifact validation in `/tmp`.
- Main: repository diffs, evidence reconciliation, generated matrix, public
  documentation, final gates, commit, and push.

### Task 2: Backfill matrix and documentation

- [x] Register the exact four IDs and 3/1 split.
- [x] Regenerate the audit to 1,748 classified and 26 unverified.
- [x] Update `README.md`, `packages/pptx/README.md`,
  `docs/compatibility/pptxgenjs-baseline.md`, and
  `docs/implementation-progress.md` to the exact 1,748/26 checkpoint.

### Task 3: Verify and synchronize

- [x] Run the aggregate focused test once with local Vitest; require the exact
  family title to pass.
- [x] Run the audit library and generator twice; require zero diagnostics,
  zero stale entries, and byte-identical generated JSON/Markdown.
- [x] Run TypeScript/build, one actual `npm pack` smoke batch, one persistent
  Chrome batch, the retained OOXML/client checks, and one full test batch.
- [x] Require Agent A/B/C PASS reports plus a clean main-agent diff review.
- [x] Stage only family files, commit once, push once, fetch, and require local/
  remote divergence `0 0`.
- [x] Report the completed family, 26 remaining atoms, 98.53% matrix progress,
  exact gates, commit, and push before starting the next family.
