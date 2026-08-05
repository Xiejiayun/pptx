# Data/Path Inheritance Surface Family Design

## Scope

Close exactly four residual PptxGenJS 4.0.1 declaration atoms as one
evidence-only capability family: `DataOrPathProps.data/path` and the two fields
inherited by `TextPropsOptions`. Runtime evidence fixes the split at two
deliberate differences and two defect exclusions; no native implementation
changes are required.

## Native source model

PptxGenJS image and media owners accept separate optional `data` and `path`
fields. Legal individual fields work, while simultaneous fields give `data`
precedence. Images additionally leak the losing path into `cNvPr@descr`.
Native retains one typed `ImageSource` or `MediaSource`, validates it before
package mutation, and supports data URIs, paths where applicable, bytes,
buffers, Blob/File, streams, and supported URLs without ambiguous precedence.

The `TextPropsOptions` fields are declaration inheritance leakage. PptxGenJS
ignores them for plain and rich text, produces byte-identical normalized slide
XML, and serializes neither value anywhere in the package. Native therefore
does not add inert text aliases.

## Evidence and verification

One aggregate adapter control locks the exact four-ID inventory, image and
media data-only/path-only/both behavior, the image-only description leak,
plain/rich text byte isolation, package-wide sentinel absence, and native
typed-source write/reopen state. Existing source-owner SDK, npm, browser, and
OOXML evidence is reused because this family changes no packed implementation.

The manifest records real source owners as deliberate differences with full
serialization/client evidence and text inheritance as defect exclusions with
control-only evidence. The audit must regenerate byte-identically twice and
finish at 1,744/1,774 classified (98.31%), 30 unverified, and zero unsupported,
stale, or diagnostic entries.
