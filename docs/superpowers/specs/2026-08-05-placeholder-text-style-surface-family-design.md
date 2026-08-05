# Placeholder Text Style Surface Family Design

## Scope

Close four PptxGenJS 4.0.1 `PlaceholderProps` style atoms as one capability
family: `align`, `margin`, `transparency`, and `valign`. Runtime and native
evidence fix the split at three supported atoms and one deliberate difference.
No product codec change is required.

## Canonical owners

Native placeholder text uses the existing canonical text owners inherited by
`AddPlaceholderOptions`: paragraph alignment, strict text-box margins,
explicit rich-run transparency, and text-box vertical alignment. Prompt,
populated, edited, duplicated, and reopened owners retain these states without
introducing placeholder-only aliases.

PptxGenJS delays margin application until population and maps `[1,2,3,4]` to
semantic TRBL `4/2/3/1`. Native deliberately keeps documented point-based
TRBL and partial-side state. All legal horizontal/vertical alignment and alpha
states remain representable through canonical native owners.

## Batch evidence

One aggregate PptxGenJS control covers all legal catalogs, prompt/population
two-phase behavior, caller mutation, and the margin ordering difference. One
shared native lifecycle probe covers layout, master, slide, populated and
duplicate owners, exact no-op, invalid isolation, rollback, six presentation
formats, exact OOXML, write/reopen, and diagnostics. The same probe is invoked
once from the actual npm package and persistent browser batches.

The matrix must finish at 1,748/1,774 classified (98.53%), 26 unverified, and
zero unsupported, stale, or diagnostic entries.

## Parallel capability-family workflow

The remaining declaration atoms are scheduled as capability families rather
than 1,774 independent tasks. Each family owns one aggregate control, one
focused test run, one shared lifecycle probe, one matrix update, and one
review/commit/push boundary. Existing tests, documentation, package smoke
checks, and OOXML evidence are indexed and reused before new proof is added.

Four lines run concurrently with non-overlapping ownership:

- Agent A parses the matrix, identifies reusable canonical owners, and writes
  mapping/review evidence only under `/tmp`.
- Agent B runs the PptxGenJS 4.0.1 runtime/control comparison and writes its
  report only under `/tmp`.
- Agent C validates native behavior, the actual npm tarball, persistent Chrome,
  and OOXML, reusing the current batch artifacts where they remain applicable.
- The main agent alone edits repository files, reviews all evidence, regenerates
  the matrix, runs the batch gates, commits the family, and pushes it.

No two agents modify the same repository file. Expensive npm packing, browser,
full-suite, and client gates run once at a batch boundary instead of once per
atom. A failed evidence line blocks classification of only the affected family;
the main agent records the failure and keeps the matrix state unchanged until
the family passes review.
