# Plan: Controller and Engine Boundaries

Date: 2026-08-29

## Goal

Complete the next safe slice of Folio's seam architecture: route document I/O
through the framework-neutral editor controller, enforce the controller and
engine dependency direction, and make the Rust/WASM projection boundary
executable rather than documentary.

## Design Decisions

- **Migrate vertically:** Move load and serialization end to end through
  `FolioEditor` in React and Vue before moving another editor capability.
- **Preserve adapter contracts:** Existing React and Vue ref methods retain
  their signatures and host callbacks; only their internal route changes.
- **Use a named save mode:** The controller uses an extensible mode
  discriminator while adapters continue accepting their established
  `selective` option.
- **Make the controller a composition root:** Core modules below the controller
  may not import controller implementation types. Shared snapshots move into
  the pure model.
- **Keep engine inputs portable:** Pure layout code must not import ProseMirror,
  DOCX parsing, render, or controller modules.
- **Keep Rust bounded:** `@stll/docx-core/projection` may initialize and call the
  generated kernel only; no TypeScript parser or archive fallback may enter the
  boundary.
- **Avoid a package split:** Stabilize and enforce internal seams before adding
  published packages or moving interactive editing into Rust.

## Scope

**In scope:**

- Add document load and serialization methods to `FolioEditor`.
- Route React and Vue imperative document I/O through the controller.
- Move shared note-story, remote-selection, wrap, and tab-measurement types or
  helpers out of downstream implementation modules.
- Enforce controller-root, pure-engine, and Rust projection-boundary imports
  with lint rules and rule tests.
- Add release notes for affected published packages.

**Out of scope:**

- Changing public React or Vue ref signatures.
- Moving page layout, ProseMirror, or interactive editing into Rust.
- A big-bang migration of every adapter command or overlay.
- Breaking old public core subpath imports while adapters migrate.
- Changing DOCX parsing, serialization, layout, or rendering behavior.

## Implementation

- `packages/core/src/controller/folioEditor.ts` — own document I/O delegates and
  expose a typed serialization mode.
- React and Vue editor bindings — inject their existing load/serialization
  implementations and route imperative APIs through `FolioEditor`.
- `packages/core/src/types/` and `packages/core/src/layout-engine/measure/` —
  host shared pure types and tab measurement without upstream reach-ins.
- `.oxlint-plugins/folio-layer-boundaries.ts` — enforce controller, engine, and
  projection boundaries.
- Architecture lint tests and fixtures — prove forbidden edges fail and allowed
  seams remain usable.
- `.changeset/` — describe the additive controller capability.

## Test Cases

- Controller load, parsed-document load, serialization modes, missing-document
  results, and event behavior delegate exactly once.
- React and Vue existing ref APIs load and save through the controller without
  changing callback or dirty-state behavior.
- Engine code importing ProseMirror, DOCX, controller, painter, or bridge code
  fails lint.
- Non-controller core code importing controller implementation types fails
  lint.
- Projection code importing an XML/ZIP parser or a second local parser fails
  lint; the generated kernel import remains accepted.
- Focused unit tests, adapter parity checks, typecheck, lint, formatting, and
  builds pass.

## Open Questions

- Which controller slice should follow document I/O: active-story commands,
  overlay snapshots, or editor mount/destroy ownership?
- When the adapter exception set is small enough, should deep-import enforcement
  become a single controller-only rule or remain capability-specific?
