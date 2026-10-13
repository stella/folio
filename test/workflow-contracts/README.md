# Reusable workflow contracts

`reusable/` contains exact public callee YAML, keyed by repository and commit SHA.
The offline source guard enumerates reusable calls and checks caller permissions
against callee workflow and job declarations, including nested calls.

After changing a reusable workflow pin, refresh its source:

```sh
bun scripts/refresh-reusable-workflows.ts
```

Only this maintainer command fetches source. CI reads committed fixtures and fails
when a pin is missing or a caller grants insufficient permissions.
