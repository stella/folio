# stella-docx-kernel

`stella-docx-kernel` projects bounded DOCX packages and WordprocessingML into a
versioned, host-neutral document representation.

The crate supports native Rust consumers directly. Browser consumers can enable the
`wasm` feature; `@stll/docx-core/projection` provides the corresponding TypeScript
binding.

The parser treats package identifiers as document facts, not durable application
identities. ZIP and XML resource limits are enforced at the input boundary;
semantic scans also bound XML events and inline-context copies and report both as
deterministic work units.

`project_docx_with_review_facts` reuses the package-directory scan and selected
part decompression for the document projection, plus attributed revisions and
comments. Each review family is complete or explicitly unknown. Content and
location details use the same typed known/unknown contract; they remain unknown
until the canonical paragraph state machine can prove offsets after revision-view
normalization, rather than exposing a second parser's guessed coordinates.
Revision attribution covers every recognized revision element in the effective
main-document body, including text boxes and structural or formatting changes;
historical markup nested inside a change snapshot is not counted a second time.
