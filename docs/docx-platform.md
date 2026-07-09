# DOCX platform boundary

Folio is the reusable engine and toolkit for working with `.docx` files. It
provides shared document semantics to interactive editors, headless processes,
and agent tools. Host products such as stella add workspace and domain
workflows without maintaining a second DOCX implementation.

This document defines that ownership boundary and the contracts that keep the
surfaces interoperable. The package boundaries that make it possible are
described in [Folio seam architecture](./seam-architecture.md).

## Source of truth

DOCX remains the canonical artifact. HTML, Markdown, plain text, JSON, and
rendered pages are inputs or views of that artifact; they do not silently
replace its package structure or document semantics.

Folio owns reusable DOCX behavior:

- create a document from structured content or a template;
- open, validate, inspect, extract, render, edit, and save a document;
- preserve supported and unknown OOXML that an operation did not intend to
  change;
- compare versions and produce machine-readable changes or a redlined DOCX;
- apply direct edits, tracked changes, and comments with provenance;
- expose the same capabilities through framework adapters and headless APIs;
- provide portable host capabilities for browser and server runtimes.

Folio APIs should be deterministic and provider-neutral. Model selection,
prompting, and agent orchestration are concerns of the caller, not the DOCX
engine.

## Folio and host ownership

| Folio owns                                                               | A host such as stella owns                                           |
| ------------------------------------------------------------------------ | -------------------------------------------------------------------- |
| OOXML package, relationship, and part handling                           | Authentication, authorization, and workspace membership              |
| Parsing, validation, compatibility profiles, and round-trip preservation | Blob storage, retention, and resolving workspace version identifiers |
| Document creation, extraction, layout, rendering, editing, and save      | Product-specific file and matter workflows                           |
| Comments, tracked changes, version comparison, and redlined output       | Model selection, prompts, conversations, approvals, and tool policy  |
| Template and content-control primitives                                  | Organization template catalogs and legal playbooks                   |
| A versioned document-operation contract                                  | Product UI outside the reusable editor and its adapters              |
| Framework-neutral engine, editor controller, adapters, and agents        | Workspace audit events and collaboration service deployment          |

When a DOCX capability is useful outside one product, its implementation and
tests should originate in Folio. Hosts consume a released Folio API instead of
copying parsers, transforms, or serializers into application code.

The boundary does not prevent a host from storing Folio operation records in
its own audit log or presenting domain-specific UI. It prevents storage and
workflow concerns from leaking into the document engine.

## One document-operation contract

Every mutating surface should converge on a versioned, serializable document
operation contract. Toolbar commands, headless transforms, and agent tools
should not implement separate editing semantics.

An operation records at least:

- the contract version and operation kind;
- a stable document location;
- the intended mutation and review mode;
- optional author, timestamp, and source provenance;
- preconditions that make stale or ambiguous edits fail explicitly.

The review mode is one of:

- `direct`: change the document content immediately;
- `tracked`: encode the mutation as a WordprocessingML revision;
- `comment`: leave the content unchanged and attach a review comment.

Operations must produce structured results: applied changes, created object
identifiers, warnings, and typed failures. The same input document and valid
operation sequence must produce the same semantic output regardless of the
calling surface.

Callers must be able to inspect supported contract versions, operation kinds,
review modes, stories, and OOXML capabilities at runtime. A host should not
infer support from a package version or discover it only after a destructive
operation fails.

### Stable addressing

Paragraph indexes and raw ProseMirror positions are not durable document
addresses. Folio should expose stable handles for a document story and a range
within it. Stories include the main body, headers, footers, footnotes,
endnotes, comments, and text inside supported drawing objects.

An address may combine persisted OOXML identifiers with content fingerprints
and structural context. Resolution must either identify one range or return an
explicit stale or ambiguous result; it must never guess silently.

## Compatibility is a measured capability

OOXML support is not a single boolean. Folio tracks each feature independently
through these capability levels:

1. `read`: recognize and represent the source construct;
2. `render`: display it with documented fidelity;
3. `edit`: mutate it without corrupting adjacent content;
4. `create`: author the construct from a Folio API;
5. `preserve`: round-trip content that Folio does not interpret;
6. `unsupported`: detect and report a construct that cannot be handled safely.

Compatibility claims should identify the evidence behind them:

- the applicable ECMA-376 or ISO/IEC 29500 rule;
- Microsoft extension documentation where Word adds behavior;
- observed behavior from a reproducible Word interoperability fixture.

Tests and public APIs should also name the intended profile: OOXML
Transitional, OOXML Strict, or Word-compatible behavior. International text,
bidirectional layout, fonts, locale-dependent fields, and typography require
explicit fixtures rather than assumptions based on English documents.

Unknown elements, attributes, relationships, and package parts are preserved
by default. A transformation may discard them only through an explicit policy
and must report what was removed.

## Change acceptance rules

A new DOCX feature is complete only when:

- its Folio-owned behavior is available through an appropriate reusable API;
- interactive and headless paths share the same underlying operation;
- compatibility and preservation behavior are covered by fixtures;
- failures and unsupported constructs are structured and observable;
- a host can adopt the capability without copying Folio internals;
- published package changes include migration notes and a changeset.

These rules apply to vertical slices. A feature should land end to end for a
specific DOCX capability rather than adding a broad parser, model, or UI layer
with no usable workflow.
