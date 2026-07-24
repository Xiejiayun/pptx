---
name: pptx-inspect
description: Inspect, validate, diff, and make narrow dry-run-first edits to local PPTX files with stable JSON. Use when Codex needs to discover slide titles or package parts, check OOXML relationships and compatibility diagnostics, verify mutation isolation, read an exact part, or change one slide title without overwriting the source file.
---

# PPTX Inspect

Verify the installed command, then check the offline runtime:

```sh
command -v pptx-inspect
pptx-inspect --json doctor
```

No authentication or network access is required.

## Read workflow

Start broad, then narrow to exact objects:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
```

Compare decompressed part hashes when verifying isolation:

```sh
pptx-inspect --json package diff before.pptx after.pptx
```

Use the raw escape hatch only for an exact read:

```sh
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

## Write workflow

Preview first. Write only to an explicit output path after the preview matches the request:

```sh
pptx-inspect --json slides set-title deck.pptx --slide 1 --text "Updated" --dry-run
pptx-inspect --json slides set-title deck.pptx --slide 1 --text "Updated" --out updated.pptx
```

Do not overwrite the input file. Do not delete parts or bypass validation; this CLI intentionally exposes no raw write command.
