# Saved OOXML validity oracle

Consumer save/reopen invariants and browser fuzz saves run `validateDocxPackage` before reopening. Swarm and long consumer flows inherit the same save boundary. The pending cross-host, host-API and canonical stacks call the same existing validator at their saved-output checkpoints; this change strengthens those calls when the stacks land. Metamorphic save helpers also validate each output.

The validator bounds archive inflation, checks all XML parts for syntax, checks WordprocessingML attributes against committed ECMA-376 graph facts, and checks OPC content types, relationship targets and relationship references. Required inherited attributes, enumerations, supported lexical patterns and integer facets are generated offline. Unknown extension vocabularies remain available for round trips. This is a targeted validity profile, not a complete XSD content-model validator.

`ST_DecimalNumber` derives from unbounded `xs:integer` in the committed schema. The signed 32-bit limit for WordprocessingML annotation IDs is a separate interoperability policy. It accepts note separator IDs -1 and 0 and catches timestamp-sized comment IDs. Schema integer facts remain authoritative for other numeric attributes; the generator applies the annotation policy explicitly at `w:id` slots. Product serializers should use the same ID policy when their owning changes land; this oracle does not rewrite output or hide a finding.

Regenerate facts with `bun scripts/generate-fuzz-schema.ts write`. CI checks exact derivation with `bun scripts/generate-fuzz-schema.ts check`. Mutation tests cover annotation families, enum and numeric values, declaration coverage and relationship integrity.
