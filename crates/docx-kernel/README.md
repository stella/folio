# stella-docx-kernel

`stella-docx-kernel` projects bounded DOCX packages and WordprocessingML into a
versioned, host-neutral document representation.

The crate supports native Rust consumers directly. Browser consumers can enable the
`wasm` feature; `@stll/docx-core/projection` provides the corresponding TypeScript
binding.

The parser treats package identifiers as document facts, not durable application
identities. ZIP and XML resource limits are enforced at the input boundary.
