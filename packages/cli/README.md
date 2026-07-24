# pptx-inspect

Offline CLI for inspecting, validating, diffing, and making narrow PPTX edits.

```sh
pnpm --filter @pptx/cli build
pnpm --filter @pptx/cli link --global

pptx-inspect --json doctor
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json slides set-title deck.pptx --slide 1 --text "Updated" --dry-run
pptx-inspect --json slides set-title deck.pptx --slide 1 --text "Updated" --out updated.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

## JSON contract

Success:

```json
{"ok":true,"command":"slides.list","data":[]}
```

Error:

```json
{"ok":false,"error":{"code":"CLI_ERROR","message":"..."}}
```

With `--json`, stdout contains JSON only. Human diagnostics go to stderr. The CLI is offline and requires no authentication. `slides set-title` is dry-run first and requires an explicit `--out` for writes. `part read` is the raw read-only escape hatch.

