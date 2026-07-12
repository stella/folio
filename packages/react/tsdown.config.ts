import babel from "@rolldown/plugin-babel";
import { reactCompilerPreset } from "@vitejs/plugin-react";
import type { Plugin } from "rolldown";
import { defineConfig } from "tsdown";

// CSS lives in a single bundled stylesheet (`dist/editor.css`, built separately
// by `scripts/build-css.ts`). The component-level `import "*.css"` statements are
// side-effect imports only; strip them from the JS so the dist bundle ships no
// inlined CSS and the consumer loads styling once via `@stll/folio-react/editor.css`.
const CSS_STUB_ID = "\0folio-stripped-style";
const stripCssImports = (): Plugin => ({
  name: "folio-strip-css-imports",
  resolveId(id) {
    if (id.endsWith(".css")) {
      return { id: CSS_STUB_ID, moduleSideEffects: false };
    }
    return null;
  },
  load(id) {
    if (id === CSS_STUB_ID) {
      return { code: "", moduleSideEffects: false };
    }
    return null;
  },
});

const compileReact = (): Plugin =>
  babel({
    // Folio publishes TSX and supports React 18 consumers. Parse the original
    // source before tsdown transforms it, then emit against the standalone
    // React 18 compiler runtime rather than React 19's built-in runtime.
    parserOpts: { plugins: ["typescript", "jsx"] },
    presets: [reactCompilerPreset({ target: "18" })],
  });

const entry = {
  index: "src/index.ts",
  "compat/eigenpal": "src/compat/eigenpal.tsx",
  dialogs: "src/dialogs.ts",
  // Bundled UI translations, exported at `@stll/folio-react/messages`. The 13
  // locale JSONs are inlined into this chunk; prepare-publish maps it to
  // `./dist/messages.js`.
  messages: "src/i18n/messages.ts",
};

const shared = {
  entry,
  // ESM-only: folio's own pervasive deps (better-result, marked) are ESM-only
  // packages with no `require` condition, so a CJS build would emit `require()`
  // calls that throw ERR_REQUIRE_ESM on consumers without require(ESM). ESM-only
  // also sidesteps the dual-package hazard entirely.
  format: ["esm"] as const,
  platform: "neutral" as const,
  outDir: "dist",
  plugins: [stripCssImports()],
  // Externalize every runtime dependency and peer so none is bundled into the
  // JS: consumers install React, ProseMirror, the @stll closure (including
  // @stll/folio-core), fonts, etc. themselves. tsdown auto-externalizes
  // `dependencies` + `peerDependencies` already; these explicit patterns also
  // cover subpath imports (`react/jsx-runtime`, `@stll/folio-core/prosemirror/...`,
  // `@fontsource/*`, `prosemirror-view/...`).
  deps: {
    // Force `.css` specifiers to be "bundled" so tsdown's auto-externalization
    // of dependency subpaths (e.g. `prosemirror-view/style/prosemirror.css`)
    // does not leave them as imports; the strip plugin then turns them into an
    // empty, side-effect-free module that rolldown drops.
    alwaysBundle: (id: string) => id.endsWith(".css"),
    neverBundle: [
      /^react(?:$|\/)/u,
      /^react-dom(?:$|\/)/u,
      /^use-intl(?:$|\/)/u,
      // The whole @stll closure stays external, including @stll/folio-core and
      // its subpaths (`@stll/folio-core/prosemirror/plugins/...`).
      /^@stll\//u,
      /^@base-ui\/react(?:$|\/)/u,
      /^@fontsource\//u,
      // Bare package names only: a `prosemirror-*/style/*.css` subpath must fall
      // through to the strip plugin, not be externalized as a CSS import.
      /^prosemirror-[a-z-]+$/u,
      /^y-prosemirror(?:$|\/)/u,
      "yjs",
      "lucide-react",
      "better-result",
    ],
  },
};

// JS and declarations are emitted in two separate rolldown passes. Generating
// both in one build makes a type-only re-export module collapse to an empty JS
// chunk that the declaration pass claims as a `.d.ts`; the JS bundle then emits
// a runtime-helper import pointing at that `.d.ts`, which fails to load.
// Splitting the passes keeps each graph's chunks self-referential (JS imports
// `.js`, declarations import `.d.ts`).
//
// tsdown runs array configs concurrently, so neither pass may `clean` (it would
// race the other's output). The `build` script clears `dist` up front instead.
export default defineConfig([
  { ...shared, plugins: [stripCssImports(), compileReact()], dts: false, clean: false },
  { ...shared, dts: { emitDtsOnly: true }, clean: false },
]);
