# Transaction ZIP Entry-Date Rollback Design

## Problem

`OpcPackage.transaction()` snapshots parts, relationships, content types, and the mutation journal. On rollback, it rebuilds every JSZip entry from the saved part bytes. When the package has no configured fixed `entryDate`, JSZip assigns the rebuild time. A rollback that crosses a ZIP timestamp boundary therefore changes output bytes even though the semantic package and mutation journal are restored.

This violates the existing zero-observable-change rollback contract and makes package-snapshot tests time-dependent.

## Scope

Preserve ZIP entry dates across failed outer and nested synchronous transactions. Do not change normal creation, opening, mutation, compression, write, or fixed-`entryDate` behavior. Do not broaden the transaction API or alter package-part data.

## Options

1. Record each current ZIP entry date in the transaction savepoint and reuse it during rollback reconstruction. This is the selected approach because it restores the metadata that rollback currently loses and leaves successful writes untouched.
2. Force a global fixed date for every package. Rejected because it changes public output semantics beyond rollback.
3. Exclude ZIP bytes from rollback tests. Rejected because it hides a real observable mutation.

## Design

Extend the private package savepoint with a detached map of non-directory ZIP entry names to copied `Date` values. `#createSavepoint()` captures those values before the transaction begins. `#restoreSavepoint()` continues rebuilding from the saved canonical part records, but passes the corresponding saved date into `#writeZipFile()` for each restored entry.

`#writeZipFile()` gains an optional restore-only date parameter. An explicit saved date wins; otherwise the existing package-level fixed `#entryDate` remains in force; otherwise JSZip keeps its current default behavior. Dates are copied on both capture and restore so caller or JSZip mutation cannot alias savepoint state.

Missing saved metadata falls back to the existing behavior. This is defensive only; a saved part should normally have a matching ZIP entry.

## Verification

Add an OPC regression test that creates a package without fixed `entryDate`, captures output at one mocked time, advances across a ZIP timestamp boundary, performs a failing transaction, and requires byte-identical output after rollback. Keep existing transaction, nested-savepoint, compression, focused table-cell hyperlink, full Vitest, and performance gates green.

## Success Criteria

- A failed transaction restores part bytes, relationships, content types, journal, and ZIP entry dates.
- Cross-boundary rollback output is byte-identical.
- Successful mutations and fixed/default entry-date behavior are unchanged.
- The change remains private to `@pptx/opc` and adds no public API.
