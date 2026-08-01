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
