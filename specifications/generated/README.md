# Generated specification data

`docx-transitional-schema.gen.json` is generated from the pinned ECMA-376 Part 2
and Part 4 schema archives. It contains structural facts only: namespaces,
documents, declarations, inheritance, compositors, scoped children,
attributes, and simple-type constraints.

Run `bun run specifications:fetch` and `bun run specifications:generate` to
regenerate it. Use `bun run specifications:generate:check` to verify a local
cache against the committed graph.
