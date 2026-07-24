# Security model

- ZIP entries reject absolute paths, backslashes, and `..` traversal.
- Defaults limit entry count, single-part bytes, total uncompressed bytes, and compression ratio.
- XML rejects DTD and ENTITY declarations; it performs no filesystem or network entity resolution.
- External relationships are preserved but never fetched automatically.
- Media, macro, OLE, and ActiveX payloads are data only; the library never executes them.
- `AbortSignal` cancels file/stream ingestion.
- CLI JSON errors do not include file payloads unless the caller explicitly runs `part read`.

Report security issues privately to the repository maintainers. Do not attach confidential presentation files to public issues.

