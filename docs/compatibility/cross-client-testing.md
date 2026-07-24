# Cross-client testing

Public CI runs Node 20/22 on Linux, macOS, and Windows, deterministic fuzz tests, a 1,000-part performance smoke test, and LibreOffice headless open/export.

`office-conformance.yml` is a manually triggered private workflow:

- Windows runner labels: `self-hosted`, `windows`, `powerpoint`.
- macOS runner labels: `self-hosted`, `macos`, `keynote`.
- PowerPoint opens every corpus PPTX through COM and saves a copy; COM/open/save errors fail the job.
- Keynote opens every corpus PPTX and exports PDF; AppleScript errors fail the job.

Google Slides import/export uses a controlled test account and remains a release-candidate sampling job rather than a per-PR blocker. Test results must record input hash, client version/date, import diagnostics, exported file hash, and visual-review outcome.

