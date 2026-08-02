# Cross-client testing

Public CI runs Node 20/22 on Linux, macOS, and Windows, deterministic fuzz tests, a 1,000-part performance smoke test, and LibreOffice headless open/export.

`office-conformance.yml` is a manually triggered private workflow:

- Windows runner labels: `self-hosted`, `windows`, `powerpoint`.
- macOS runner labels: `self-hosted`, `macos`, `keynote`.
- PowerPoint opens every corpus PPTX through COM and saves a copy; COM/open/save errors fail the job.
- Keynote opens every corpus PPTX and exports PDF; AppleScript errors fail the job.

Google Slides import/export uses a controlled test account and remains a release-candidate sampling job rather than a per-PR blocker. Test results must record input hash, client version/date, import diagnostics, exported file hash, and visual-review outcome.

## 2026-08-01 native media timing evidence

The installed `@jiayunxie/pptx@0.1.0` tarball created a nine-slide, 12-media gallery with playable MP3/M4A/WAV/OGG and MP4/MOV/WebM payloads plus PNG/JPEG/GIF posters. It covered click/auto, loop, show/hide when stopped, volume 0/0.25/0.5/1, two media on one slide, ordinary animation coexistence, legacy materialization, native-only adoption, duplication, editing, clearing, and deletion.

The source package strictly reopened with zero timing diagnostics and no orphan media targets. `pptx-inspect` doctor, package inspection, PowerPoint 2010 validation, exact slide-part read, and before/after diff passed. Validation reported 0 errors and only the expected PowerPoint 2010 portability warnings for OGG and WebM. All nine slides rendered at 180 DPI with zero overflow and passed individual poster/title inspection.

LibreOffice 26.8 opened and saved the gallery, retained all nine slides and text, but removed every media picture, poster, media relationship, and timing branch. The saved package still strictly reopened and validated with 0 errors / 0 warnings. This is a client degradation record, not a media round-trip pass.

Local PowerPoint 16.112 automation returned the same `-9074` while opening the source gallery, the LibreOffice-saved package, and a minimal independent control deck. Because the control failed identically, no package-specific PowerPoint defect or successful PowerPoint round trip is claimed from this run; the Windows COM corpus remains the decisive gate.

## 2026-08-02 native chart evidence

The installed `@jiayunxie/pptx@0.1.0` tarball created an 11-slide gallery covering all nine standard chart types, a bar/line primary-secondary combination, semantic data editing, type conversion, duplication, deletion, and reopen. The package contained ten chart parts and ten embedded XLSX workbooks with synchronized A1 formulas and caches, unique shape IDs, and zero orphan chart-owned parts.

`pptx-inspect` package inspection, exact chart-part read, slide listing, and PowerPoint 2010 validation passed with 0 errors / 0 warnings. The package strictly reopened with zero chart diagnostics. All 11 slides rendered at 180 DPI with zero overflow and passed slide-by-slide visual inspection.

LibreOffice 26.8 displayed all eight 2D types plus the combination chart. Its `bar3D` rendering showed only the title; an independent PptxGenJS 4.0.1 control deck behaved identically. After save/reopen, LibreOffice retained the group types and cached values of all ten chart objects but removed every embedded workbook and replaced A1 formulas with client placeholders. The library recognizes this as editable cache-only state: validation reports 0 errors and ten `CHART_WORKBOOK_MISSING` warnings, and the first semantic replacement regenerates a canonical synchronized XLSX for that chart.

Local PowerPoint 16.112 automation returned the same `-9074` for the native gallery and an independent PptxGenJS control deck. No successful PowerPoint round trip is claimed from this machine; the Windows COM corpus remains the decisive gate.

## 2026-08-02 named master/layout/placeholder evidence

The installed 57-file `@jiayunxie/pptx@0.1.0` tarball created a two-slide native gallery containing 32 parts, 29 relationships, two layouts, one master, three PNG parts, two chart parts, two embedded workbooks, and one audio part, with no orphan owned dependencies. It covered named layout selection, inherited ordinary objects, all six placeholder types and populated domains, master/layout backgrounds, slide numbers, image/chart/table/media payloads, whole-definition replacement, deletion, and slide retargeting. PowerPoint 2010 validation reported 0 errors / 0 warnings. The independent two-slide PptxGenJS 4.0.1 control contained 36 parts and 34 relationships.

Packed Node, TypeScript, CLI, and real-Chrome checks all reported `masterLayouts: true`. Chrome exactly restored stable wrapper identity, two layout names, solid master and image layout backgrounds, all six layout/slide placeholder names, identities, and populated kinds, both selected layout targets, background/image/media SHA-256 values, the bar-chart definition, and the transient-margin reopen boundary. Validation, console, page, and network error counts were all zero.

Native/control source files and both LibreOffice 26.8 saved files produced eight 2400×1350 pages at 180 DPI and were inspected individually. Full-bleed fixture backgrounds yielded an expected 0px minimum non-white margin. The fixtures deliberately use 1×1 black PNG background/image payloads, and the second native slide is deliberately retargeted to the blank default layout; black/blank pages are therefore fixture behavior rather than evidence of lost inheritance.

LibreOffice retained two slides, two layouts, and one master, but rewrote placeholder identities and slide-number caches, removed audio and both embedded chart workbooks, and generated chart style/color parts. The native saved package contained 28 parts and 25 relationships and validated with one placeholder-identity error plus three workbook/cache warnings. This is a client degradation record, not a complete round-trip pass.

Local PowerPoint 16.112 automation returned the same `-9074` for native and control inputs and produced no saved PPTX or PDF. No successful PowerPoint round trip is claimed from this machine; the Windows COM corpus remains the decisive gate.
