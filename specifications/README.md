# Specification sources

`sources.json` records the exact external inputs used for specification-aware
tooling. Source archives and repositories are fetched into
`.cache/specifications/`; they are not committed.

Run `bun run specifications:check` to validate the manifest and verify any
cached inputs. Run `bun run specifications:fetch` to populate and verify the
cache.

The licensing fields are repository policy, not a replacement for the linked
notices. `implementation-facts-only` permits generated structural facts and
implementation metadata, but not copied prose or source archives. A
`needs-review` source remains reference-only until its licensing evidence is
complete.
