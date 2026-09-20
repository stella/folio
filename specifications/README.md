# Specification sources

`sources.json` records the exact external inputs used for specification-aware
tooling. Source archives and repositories are fetched into
`.cache/specifications/`; they are not committed.

Run `bun run specifications:check` to validate the manifest and verify any
cached inputs. Run `bun run specifications:fetch` to populate and verify the
cache.

`strict-names/` records the names ECMA-376 Part 1 spells by writing direction
where Part 4 spells them by physical side. The repository vendors no Strict
schema, so the direction of each pair is cited there by hand;
`bun run generate:strict-names` finds every slot it applies to in the committed
Transitional graph and fails on a rename the graph does not declare.

`reserved-values/` records what each model field decides about the reserved
value of the OOXML slot it carries. It is a tsconfig project of its own, checked
by `bun run typecheck` and budgeted like any package; nothing ships from it. See
`docs/reserved-values.md`.

The licensing fields are repository policy, not a replacement for the linked
notices. `implementation-facts-only` permits generated structural facts and
implementation metadata, but not copied prose or source archives. A
`needs-review` source remains reference-only until its licensing evidence is
complete.
