# Reusable workflow contracts

`reusable/` contains complete public callee YAML, keyed by repository and commit SHA
and normalized to repository formatting.
The offline source guard enumerates reusable calls and checks caller permissions
against callee workflow and job declarations, including nested calls.

After changing a reusable workflow pin, refresh its source:

```sh
bun scripts/refresh-reusable-workflows.ts
```

Only this maintainer command fetches source. CI reads committed fixtures and fails
when a pin is missing or a caller grants insufficient permissions.

The refresh command applies repository formatting and verifies that the complete parsed YAML remains unchanged.
