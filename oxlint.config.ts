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
    // The React package runs the real compiler and ratchets its intentional
    // bailouts in scripts/react-compiler-bailouts.json; this rule cannot model
    // that gradual baseline and would turn every known bailout into an error.
    "react/react-compiler": "off",
  },
  jsPlugins: [
    "./.oxlint-plugins/folio-layer-boundaries.ts",
    "./.oxlint-plugins/folio-asset-urls.ts",
    "./.oxlint-plugins/folio-base64.ts",
    "./.oxlint-plugins/folio-fragment-ownership.ts",
    "./.oxlint-plugins/folio-identity-attributes.ts",
    "./.oxlint-plugins/folio-painted-text.ts",
    "./.oxlint-plugins/folio-verbatim-capture.ts",
    "./.oxlint-plugins/folio-ref-mirrors.ts",
    "./.oxlint-plugins/no-untranslated-jsx-literal.ts",
    "./.oxlint-plugins/folio-model-types.ts",
    "./.oxlint-plugins/folio-reserved-values.ts",
    "./.oxlint-plugins/folio-container-children.ts",
    "./.oxlint-plugins/folio-union-dispatch.ts",
    "./.oxlint-plugins/folio-xml-escaping.ts",
  ],
  ignorePatterns: [
    // Module-augmentation files must use `interface` for declaration merging;
    // oxlint's --fix would rewrite it to `type` and break the augmentation.
    "packages/react/types/**/*.d.ts",
    // Machine-generated typed-message catalog (scripts/i18n-typegen.ts). Its
    // shape mirrors en.json byte-for-byte so `i18n-typegen --check` can diff it;
    // linting/`--fix` would rewrite it and break that drift check.
    "**/*.gen.ts",
    // wasm-bindgen output is verified byte-for-byte by wasm:check. Formatting
    // or lint fixes would make the committed artifact differ from its source.
    "packages/docx-core/src/generated/**",
    "packages/core/src/generated/**",
    // Lint-rule fixtures contain deliberate violations; the repo-wide run must
    // skip them. scripts/no-untranslated-jsx-literal.test.ts lints them
    // explicitly with `--no-ignore` to assert the rule's behaviour.
    "test/__fixtures__/**",
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
      files: [
        "packages/core/src/layout-bridge/dom/clickToPositionDom.ts",
        "packages/core/src/layout-bridge/headerFooterLayout.ts",
        "packages/core/src/paged-layout/rangeProjection.ts",
        "packages/core/src/prosemirror/utils/visualLineNavigation.ts",
        "packages/core/src/render-dom/HeaderFooterSelectionOverlay.ts",
        "packages/core/src/render-dom/RenderedDomContext.ts",
        "packages/react/src/paged-editor/PagedEditor.tsx",
        "packages/react/src/paged-editor/useVisualLineNavigation.ts",
        "test/__fixtures__/painted-text.*.ts",
      ],
      rules: {
        "folio-painted-text/no-direct-text-shape": "error",
      },
    },
    {
      // The paginator owns page-fragment commits because insertion may consume
      // a pending section transition. The fixtures verify this custom rule;
      // repo-wide lint ignores their deliberate violation.
      files: [
        "packages/core/src/layout-engine/**/*.{ts,tsx}",
        "test/__fixtures__/fragment-ownership.*.ts",
      ],
      rules: {
        "folio-fragment-ownership/no-direct-page-fragment-push": "error",
      },
    },
    {
      // `docx/verbatimCapture` owns the conversion a Strict fragment needs
      // before a Transitional part replays it. The fixtures verify this custom
      // rule; repo-wide lint ignores their deliberate violation.
      files: ["packages/core/src/docx/**/*.ts", "test/__fixtures__/verbatim-capture.*.ts"],
      rules: {
        "folio-verbatim-capture/no-direct-element-to-xml": "error",
      },
    },
    {
      // One way to walk a container's children. The shrink-only baseline holds
      // the parsers that predate the dispatcher at their current count.
      // The fixtures verify this custom rule; repo-wide lint ignores their
      // deliberate violation.
      files: ["packages/core/src/docx/**/*.ts", "test/__fixtures__/container-children.*.ts"],
      rules: {
        "folio-container-children/no-hand-rolled-child-dispatch": "error",
      },
    },
    {
      // A test asserts on a name it already knows; it parses no package.
      files: ["packages/core/src/docx/**/*.test.ts"],
      rules: {
        "folio-container-children/no-hand-rolled-child-dispatch": "off",
      },
    },
    {
      // A test may serialize markup directly; it writes no package part.
      files: ["packages/core/src/docx/**/*.test.ts"],
      rules: {
        "folio-verbatim-capture/no-direct-element-to-xml": "off",
      },
    },
    {
      // `docx-core/serialize/xmlEscape` owns XML escaping. These directories
      // write the parts of a `.docx`, so a second escaper here decides whether
      // Word opens the package; the owner's own file is exempt inside the rule.
      // The fixtures verify this custom rule; repo-wide lint ignores their
      // deliberate violation.
      files: [
        "packages/core/src/docx/**/*.ts",
        "packages/core/src/internal/**/*.ts",
        "packages/docx-core/src/**/*.ts",
        "test/__fixtures__/xml-escaping.*.ts",
      ],
      rules: {
        "folio-xml-escaping/no-hand-rolled-xml-escape": "error",
      },
    },
    {
      // A model content union grows, and a chain of `else if` over its tag
      // absorbs the new member silently. These directories dispatch on those
      // unions and have been converted to `switch` plus a `never` default; the
      // glob may grow as the rest follow it, and may not shrink. The fixtures
      // verify this custom rule; repo-wide lint ignores their deliberate
      // violation.
      files: [
        "packages/core/src/prosemirror/conversion/**/*.ts",
        "packages/core/src/docx/serializer/**/*.ts",
        "packages/core/src/markdown/**/*.ts",
        "packages/docx-core/src/serialize/**/*.ts",
        "test/__fixtures__/union-dispatch.*.ts",
      ],
      rules: {
        "folio-union-dispatch/exhaustive-model-union-dispatch": "error",
      },
    },
    {
      // A test narrows a union to reach the value it asserts on; it projects
      // nothing, and a member it does not name fails the assertion below.
      files: [
        "packages/core/src/prosemirror/conversion/**/*.test.ts",
        "packages/core/src/docx/serializer/**/*.test.ts",
        "packages/core/src/markdown/**/*.test.ts",
        "packages/docx-core/src/serialize/**/*.test.ts",
      ],
      rules: {
        "folio-union-dispatch/exhaustive-model-union-dispatch": "off",
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
      // `utils/base64` owns bytes-to-base64 for the package. `btoa` needs a
      // binary string, and the usual ways of building one are either wrong in
      // browsers (`TextDecoder("latin1")` is windows-1252) or allocate a copy
      // of the input. See `.oxlint-plugins/folio-base64.ts` and the matching
      // test at `scripts/base64-owner-lint.test.ts`, which lints the
      // `test/__fixtures__` files covered here.
      files: ["packages/*/src/**/*.{ts,tsx}", "test/__fixtures__/base64-owner.*.ts"],
      rules: {
        "folio-base64/no-hand-rolled-base64": "error",
      },
    },
    {
      // Every user-facing string in the React editor must go through
      // use-intl (`useTranslations`) so it exists in the locale catalogs;
      // a raw JSX text literal ships English verbatim to all 18 locales.
      // See `.oxlint-plugins/no-untranslated-jsx-literal.ts` and the
      // matching test at `scripts/no-untranslated-jsx-literal.test.ts`,
      // which lints the `test/__fixtures__` files covered here.
      files: ["packages/react/src/**/*.tsx", "test/__fixtures__/**/*.tsx"],
      rules: {
        "no-untranslated-jsx-literal/no-untranslated-jsx-literal": [
          "error",
          {
            // "stella" is the product brand mark (e.g. the autocomplete caret
            // badge); like i18n-check's ALLOWED_IDENTICAL list, it reads the
            // same in every locale and never belongs in the catalogs.
            allowedText: ["stella"],
          },
        ],
      },
    },
    {
      // Folio core is the headless, framework-neutral core. Forbid React,
      // react-dom, and @stll/ui (type-only imports included) anywhere under
      // core/, so adapters can all sit on one shared core. See the matching
      // test at `src/core/__tests__/react-free-core.test.ts`.
      files: [
        "packages/core/src/**/*.{ts,tsx}",
        "test/__fixtures__/packages/core/src/**/*.{ts,tsx}",
      ],
      rules: {
        "folio-layer-boundaries/controller-and-engine-seams": "error",
        "folio-layer-boundaries/no-react-in-core": "error",
      },
    },
    {
      // A paint backend consumes the display list and nothing else. The coarse
      // half of that boundary (no edge to the layout engine at all) is a
      // dependency-cruiser rule; this one checks the shape of the remaining
      // edge, which dependency-cruiser cannot see: the display-list types must
      // arrive as types, so that data crosses the seam and behaviour does not.
      files: ["packages/core/src/pdf/**/*.ts", "packages/core/src/display-list/dom/**/*.ts"],
      rules: {
        "folio-layer-boundaries/paint-backend-seam": "error",
      },
    },
    {
      // Each artifact compiled from Rust has one boundary module around it: the
      // DOCX kernel behind docx-core's projection, the text shaper behind
      // core's. Prevent a TypeScript reimplementation of what the crate does,
      // or a second entry point to the artifact, from appearing silently.
      files: [
        "packages/docx-core/src/**/*.{ts,tsx}",
        "packages/core/src/**/*.{ts,tsx}",
        "test/__fixtures__/packages/docx-core/src/**/*.{ts,tsx}",
        "test/__fixtures__/packages/core/src/**/*.{ts,tsx}",
      ],
      rules: {
        "folio-layer-boundaries/rust-projection-boundary": "error",
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
      // A ref reassigned from render scope mirrors state; an imperative write
      // elsewhere is overwritten by the next render and defeats identity
      // checks against the ref. See `.oxlint-plugins/folio-ref-mirrors.ts`
      // and the matching test at `scripts/ref-mirror-lint.test.ts`.
      files: ["packages/react/src/**/*.{ts,tsx}", "test/__fixtures__/ref-mirror.*.ts"],
      rules: {
        "folio-ref-mirrors/no-write-to-render-mirrored-ref": "error",
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
    {
      // A field attached to a shared model type through a local intersection,
      // or read back with an `in` check, is invisible to every other
      // projection of the model (issue #845). See
      // `.oxlint-plugins/folio-model-types.ts` and the matching test at
      // `scripts/model-types-lint.test.ts`.
      files: ["packages/*/src/**/*.{ts,tsx}", "test/__fixtures__/model-types.*.ts"],
      rules: {
        "folio-model-types/no-model-intersection-widening": "error",
        "folio-model-types/no-in-check-on-model": "error",
      },
    },
    {
      // An OOXML reserved value read a second time is how the sentinel gets
      // handled in one place and missed in the next. Over published source the
      // rule runs from `oxlint.reserved-values.config.ts` against a shrink-only
      // baseline (`bun run check:reserved-values`), because the repository
      // predates the registry; here it covers only the fixtures its wiring test
      // lints. See `.oxlint-plugins/folio-reserved-values.ts` and the matching
      // test at `scripts/reserved-values-lint.test.ts`.
      files: ["test/__fixtures__/reserved-values.*.ts"],
      rules: {
        "folio-reserved-values/no-bare-reserved-compare": "error",
      },
    },
    {
      // An identity-bearing attribute resolved by prefix reads a foreign
      // attribute with the same local name as the real one. Over published
      // source the rule runs from `oxlint.identity-attributes.config.ts`
      // against a shrink-only baseline (`bun run check:identity-attributes`),
      // because the repository predates the derived set; here it covers only
      // the fixtures its wiring test lints. See
      // `.oxlint-plugins/folio-identity-attributes.ts` and the matching test at
      // `scripts/identity-attributes-lint.test.ts`.
      files: ["test/__fixtures__/identity-attributes.*.ts"],
      rules: {
        "folio-identity-attributes/no-prefix-resolved-identity-read": "error",
      },
    },
  ],
});
