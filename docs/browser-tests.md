# Browser test hosts

The interaction, parity, rendering, and browser-fuzz projects use built React and
Vue playground previews. Playwright owns their processes and strict ports; a
running server on either port is an error. Vite development servers are not test
hosts because dependency discovery can reload a page during a test.

Run the usual command:

```sh
bun run test:interactions
bun run test:e2e:parity
```

Each host checks its build inputs and outputs before startup. Unchanged builds
are reused; changed, added, or deleted inputs and missing or modified output
files trigger a rebuild. The cache conservatively includes repository inputs.
The startup log reports cache-check and build duration. CI's later Playwright
invocations reuse the first build when those files remain unchanged.

For a persistent local preview, build each host, then start it in a separate
terminal:

```sh
bun scripts/playground-build.ts packages/playground
bun --filter @stll/playground preview
```

```sh
bun scripts/playground-build.ts packages/playground-vue
bun --filter @stll/playground-vue preview
```

Point tests at those previews explicitly:

```sh
FOLIO_PLAYGROUND_SERVER_MODE=existing-preview bun run test:interactions
FOLIO_PLAYGROUND_SERVER_MODE=existing-preview bun run test:e2e:parity
```

This mode leaves server ownership and rebuilds with the caller. Global setup
checks both hosts for bundled assets and exact repository fixture bytes before
any test starts; it rejects a development server. Rebuild the hosts after editing
source. The default mode is `managed-preview`; unknown mode names fail.

Override React's port with `FOLIO_PLAYGROUND_PORT` and Vue's with
`FOLIO_PLAYGROUND_VUE_PORT` when running tests. For a manually started Vue
preview, set its `FOLIO_PLAYGROUND_PORT` to the same Vue port.
