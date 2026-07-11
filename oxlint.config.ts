import { library } from "@stll/oxlint-config";

// Standalone oxlint config for @stll/folio.
//
// `library()` (from @stll/oxlint-config) is the shared base for publishable
// @stll library packages: it wires the stella-lowercase + no-raw-colors
// plugins, type-aware linting, denyWarnings, and the strict library rule set.
// On top of it we register folio's render-pipeline layer-boundary plugin and
// the matching per-directory overrides (paths are relative to the repo root,
// where `bun run lint` runs).

export default library({
  options: {
    // Folio's source carries `eslint-disable` / `oxlint-disable` directives
    // calibrated for the full monorepo ruleset (ultracite's core + react
    // presets plus ~40 custom stella plugins). This standalone config uses the
    // curated `library()` rule set, a strict subset, so some of those
    // directives suppress rules that are not registered here and would be
    // reported as "unused". Keep the directives intact (they document intent
    // and stay aligned with upstream) rather than mass-editing source, and turn
    // off the unused-directive check, which is not coherent against the subset.
    reportUnusedDisableDirectives: "off",
    // Type-aware (tsgolint) lint rules are intentionally off here. In the
    // monorepo folio's tsconfig extends a workspace config the type-aware pass
    // never resolved, so those rules were effectively dormant over folio and
    // its source was never held to them. Running them against the inlined
    // standalone tsconfig would surface hundreds of never-enforced findings.
    // Type *safety* is fully covered by `bun run typecheck` (tsc --noEmit);
    // lint here enforces folio's architecture boundaries (the layer-boundary
    // plugin below) plus the AST-level hygiene rules folio already conforms to.
    typeAware: false,
  },
  rules: {
    // AST rules that oxlint delegates to the (dormant) type-aware pass in the
    // monorepo, so folio's source was never held to them. Folio's fork style
    // uses non-null assertions; keep parity with the rule set it conforms to.
    "typescript/no-non-null-assertion": "off",
    "no-useless-assignment": "off",
    // Stylistic unicorn rules the full stella config disables (they are "error"
    // in ultracite's base preset, which `library()` does not include). Folio's
    // source predates them; they catch no bugs, only style. Mirror the
    // monorepo's posture so the standalone lint matches what folio conforms to.
    "unicorn/no-useless-spread": "off",
    "unicorn/prefer-string-starts-ends-with": "off",
    "unicorn/prefer-string-replace-all": "off",
    "unicorn/switch-case-braces": "off",
    "unicorn/prefer-ternary": "off",
    "unicorn/no-array-sort": "off",
    "unicorn/throw-new-error": "off",
    "unicorn/consistent-function-scoping": "off",
    "unicorn/no-await-expression-member": "off",
    "unicorn/prefer-spread": "off",
    "unicorn/no-immediate-mutation": "off",
    "unicorn/filename-case": "off",
    "unicorn/escape-case": "off",
    "unicorn/no-hex-escape": "off",
    "unicorn/number-literal-case": "off",
    "unicorn/prefer-response-static-json": "off",
    // The shared config enables compiler analysis globally, including for Vue.
    // Folio does not currently run the React Compiler in any package build.
    "react/react-compiler": "off",
  },
  jsPlugins: [
    "./.oxlint-plugins/folio-layer-boundaries.ts",
    "./.oxlint-plugins/folio-asset-urls.ts",
  ],
  ignorePatterns: [
    // Module-augmentation files must use `interface` for declaration merging;
    // oxlint's --fix would rewrite it to `type` and break the augmentation.
    "packages/react/types/**/*.d.ts",
    // Machine-generated typed-message catalog (scripts/i18n-typegen.ts). Its
    // shape mirrors en.json byte-for-byte so `i18n-typegen --check` can diff it;
    // linting/`--fix` would rewrite it and break that drift check.
    "**/*.gen.ts",
  ],
  overrides: [
    {
      // Custom oxlint plugin sources traverse AST nodes the runtime delivers
      // as untyped (effectively `any`); strict any-flow rules add noise here.
      files: [".oxlint-plugins/**/*.{ts,tsx}"],
      rules: {
        "typescript/no-unsafe-assignment": "off",
        "typescript/no-unsafe-member-access": "off",
        "typescript/no-unsafe-call": "off",
        "typescript/no-unsafe-return": "off",
        "typescript/no-unsafe-argument": "off",
        "typescript/strict-boolean-expressions": "off",
        "require-unicode-regexp": "off",
        "no-nested-ternary": "off",
      },
    },
    {
      // Build/profiling/validation scripts: operational glue that prints to the
      // terminal and consumes untyped dynamic data. Covers the shared root
      // scripts (prepare-publish, validate-dist) and each package's own scripts.
      files: ["scripts/**/*.{ts,tsx}", "packages/*/scripts/**/*.{ts,tsx}"],
      rules: {
        "no-console": "off",
        "typescript/no-unsafe-assignment": "off",
        "typescript/no-unsafe-member-access": "off",
        "typescript/no-unsafe-call": "off",
        "typescript/no-unsafe-return": "off",
        "typescript/no-unsafe-argument": "off",
        "typescript/strict-boolean-expressions": "off",
        "typescript/no-redundant-type-constituents": "off",
      },
    },
    {
      // Folio render-pipeline layer boundaries. The painter is downstream of
      // the engine and bridge and must not import upstream concerns; the bridge
      // and engine must not pull from the painter. See
      // `.oxlint-plugins/folio-layer-boundaries.ts` and the matching test at
      // `src/core/__tests__/layer-boundaries.test.ts`.
      files: [
        "packages/core/src/layout-bridge/**/*.{ts,tsx}",
        "packages/core/src/layout-engine/**/*.{ts,tsx}",
        "packages/core/src/layout-painter/**/*.{ts,tsx}",
      ],
      rules: {
        "folio-layer-boundaries/no-upstream-import": "error",
      },
    },
    {
      // Worker/asset URL targets must survive the package build. A
      // `new URL("<x>.ts", import.meta.url)` in shipped source resolves to a
      // file the dist build never emits (it renames `.ts` -> `.js`), aborting a
      // downstream bundler with UNRESOLVED_ENTRY. See
      // `.oxlint-plugins/folio-asset-urls.ts` and the matching test at
      // `packages/core/src/__tests__/asset-url-extensions.test.ts`.
      files: ["packages/*/src/**/*.{ts,tsx}"],
      rules: {
        "folio-asset-urls/no-source-extension-url": "error",
      },
    },
    {
      // Folio core is the headless, framework-neutral core. Forbid React,
      // react-dom, and @stll/ui (type-only imports included) anywhere under
      // core/, so adapters can all sit on one shared core. See the matching
      // test at `src/core/__tests__/react-free-core.test.ts`.
      files: ["packages/core/src/**/*.{ts,tsx}"],
      rules: {
        "folio-layer-boundaries/no-react-in-core": "error",
      },
    },
    {
      // Render-storm guards for the React package. folio-react ships WITHOUT
      // the React Compiler (the tsdown build has no compiler pass, and
      // consumers import the prebuilt dist), so referential identity is
      // load-bearing: manual memoization is what keeps the editor's render
      // pipeline cheap, and an inline value passed as a context `value` (or to
      // a memoized child) silently defeats every bailout downstream. Consumers
      // that run the React Compiler over their own app code do not need these
      // rules; they are scoped to this package's source only.
      //
      files: ["packages/react/src/**/*.{ts,tsx}"],
      plugins: ["react", "react-perf"],
      rules: {
        "react/jsx-no-constructed-context-values": "error",
        // Intrinsic elements do not have React memoization boundaries, so a
        // fresh DOM prop cannot invalidate a child-component bailout. Keep the
        // rules strict at every component boundary without forcing no-op
        // memoization around native event handlers and style objects.
        "react-perf/jsx-no-jsx-as-prop": ["error", { nativeAllowList: "all" }],
        "react-perf/jsx-no-new-array-as-prop": ["error", { nativeAllowList: "all" }],
        "react-perf/jsx-no-new-function-as-prop": ["error", { nativeAllowList: "all" }],
        "react-perf/jsx-no-new-object-as-prop": ["error", { nativeAllowList: "all" }],
      },
    },
    {
      // Folio model seam. The model type layer is pure data: forbid it from
      // importing ProseMirror, DOM render, React, @stll/ui, or engine behavior.
      // See the matching test at `src/core/__tests__/model-purity.test.ts`.
      files: [
        "packages/core/src/types/**/*.ts",
        "packages/core/src/layout-engine/types.ts",
        "packages/core/src/layout-engine/measure/measureTypes.ts",
      ],
      rules: {
        "folio-layer-boundaries/model-is-pure-data": "error",
      },
    },
  ],
});
