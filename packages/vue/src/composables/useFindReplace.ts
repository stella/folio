/**
 * Vue port of the find/replace UI-state container.
 *
 * PORT-BLOCKED (missing `utils/findReplace` UI-util): upstream builds this
 * composable's entire state on three symbols from
 * `@eigenpal/docx-editor-core/utils/findReplace` — the `FindMatch` and
 * `FindOptions` data contracts and the `createDefaultFindOptions()` factory.
 * That util is absent from our fork: it is neither in `@stll/folio-core`
 * (whose search surface is the lower-level `prosemirror/findReplaceSelection`
 * `FindMatchPosition` + `managers/FindReplaceManager`, a different shape) nor
 * in our ported Vue `../utils/*`. Every field of the reactive state
 * (`options: createDefaultFindOptions()`, `matches: FindMatch[]`, the
 * `currentMatch`/`setMatches`/`goToMatch` surface) depends on those types, so
 * there is no working subset to port without fabricating the data contract.
 *
 * Unblock: port upstream `packages/core/src/utils/findReplace.ts` (pure, no
 * runtime deps) into the Vue package as `../utils/findReplace.ts`, reconciling
 * its `FindMatch`/`FindOptions` with core's `FindMatchPosition` /
 * `FindReplaceManager`, then restore the full composable — its body is a direct
 * `reactive()` translation of the React hook with no other blockers.
 */

export {};
