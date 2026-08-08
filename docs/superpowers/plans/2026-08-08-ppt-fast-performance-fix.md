# PPT fast-path performance fix plan

**Goal:** make a fresh query-to-validated PPT run reproducibly finish in under 180 seconds without external asset acquisition.

## Capability family 1: deterministic layout preflight

- Add a browser-safe `DeckSpec` model and stable preflight diagnostics to `@pptx/sdk`.
- Reject non-finite or non-positive geometry, unsafe bounds, region collisions, title overflow, and zero-length connectors before PPTX serialization.
- Canonicalize connector endpoints to positive OOXML transforms so generators never emit zero-width or zero-height lines.
- Cover the Amazon clipped-title regression and boundary cases with focused tests.

## Capability family 2: reproducible fast QA

- Add a single command that reopens the deck, validates OOXML, renders slides, checks overflow, creates a montage, and writes stable JSON evidence.
- Resolve the bundled Python and rendering tools once, then run independent checks concurrently.
- Fail on stale artifacts, missing slides, validation errors, overflow, timeout, or incomplete evidence.

## Capability family 3: skill and benchmark contract

- Update `skills/ppt` to require query → DeckSpec → preflight → one generator run → parallel QA → at most one targeted repair.
- Keep `skills/ppt` limited to `SKILL.md` and `pptx.md`.
- Add a task-cold benchmark with a nonce and an atomic final verdict; include planning, generation, reopen, validation, rendering, visual inspection, and repair time.
- Validate with three fresh Amazon runs, each below 180 seconds, before claiming the SLA.

Each capability family is reviewed, tested, committed, and pushed independently.
