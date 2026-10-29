#!/usr/bin/env bun
// Build the single bundled stylesheet shipped as `@stll/folio-react/editor.css`.
//
// The JS build (tsdown) strips every component-level `import "*.css"`, so all
// the styling a consumer needs has to live in one file. This script composes
// it: folio's own stylesheets and the ProseMirror base CSS are inlined, while
// the `@stll/folio-fonts` `@import` is preserved as a bare `@import` statement
// (inlining it would break its relative `url()` font references — the
// consumer installs `@stll/folio-fonts`, which is a folio dep, and their
// bundler resolves the import).
//
// Run after `tsdown` (the package `build` script does both). Output:
// `dist/editor.css`.

import { panic } from "better-result";
import { $ } from "bun";
import { bundleAsync, transform } from "lightningcss";
import { rm } from "node:fs/promises";
import path from "node:path";

import { scopeEditorCss } from "../../../scripts/lib/scoped-editor-css";

const pkgDir = path.resolve(import.meta.dir, "..");
const srcDir = path.join(pkgDir, "src");
const stylesDir = path.join(srcDir, "styles");
const fontsCss = path.join(stylesDir, "fonts.css");
const outFile = path.join(pkgDir, "dist", "editor.css");

// ProseMirror's base stylesheet — the components that import it are stripped
// from the JS, so its rules have to be in the bundle. No relative `url()`s, so
// inlining is safe (unlike the font stylesheet).
const prosemirrorViewCss = Bun.resolveSync(
  "prosemirror-view/style/prosemirror.css",
  pkgDir,
);

// Virtual entry: the consumer-facing cascade order. ProseMirror base first,
// then folio's prosemirror layer, then the editor chrome, then AI decorations.
// `editor.css` pulls in fonts.css / autocomplete.css via its
// own relative `@import`s.
const ENTRY = path.join(stylesDir, "__folio_bundle_entry__.css");
const entrySource = [
  `@import ${JSON.stringify(prosemirrorViewCss)};`,
  `@import "./prosemirror-layer.css";`,
  `@import "./editor.css";`,
  `@import "./ai-suggestions.css";`,
].join("\n");

// The `@stll/folio-fonts` bare import inside fonts.css: capture and drop it so
// the bundler does not try to resolve (and inline) it, then re-emit it hoisted
// to the top of the output as a bare `@import`.
const FONT_IMPORT =
  /^[^\S\n]*@import\s+["'](?<spec>@stll\/folio-fonts\/[^"']+)["'][^\S\n]*;[^\S\n]*$/gmu;
const fontImports = new Set<string>();

const bundled = await bundleAsync({
  filename: ENTRY,
  minify: true,
  resolver: {
    resolve(specifier, originatingFile) {
      if (specifier === ENTRY) {
        return ENTRY;
      }
      if (path.isAbsolute(specifier)) {
        return specifier;
      }
      if (specifier.startsWith(".")) {
        return path.resolve(path.dirname(originatingFile), specifier);
      }
      return Bun.resolveSync(specifier, pkgDir);
    },
    async read(file) {
      if (file === ENTRY) {
        return entrySource;
      }
      const css = scopeEditorCss(await Bun.file(file).text());
      if (path.resolve(file) !== fontsCss) {
        return css;
      }
      return css.replace(FONT_IMPORT, (_match, spec: string) => {
        fontImports.add(spec);
        return "";
      });
    },
  },
});

if (fontImports.size === 0) {
  panic(
    "build-css: expected to capture the @stll/folio-fonts import from fonts.css, found none",
  );
}

const fontImportBlock = Array.from(
  fontImports,
  (spec) => `@import "${spec}";`,
).join("\n");
const bundledBody = new TextDecoder().decode(bundled.code);
const composed = `${fontImportBlock}\n${bundledBody}`;

// Final validation + minification pass over the combined stylesheet. `transform`
// (unlike `bundle`) preserves `@import` rules, so the hoisted font import
// survives at the top where the spec requires it.
const { code } = transform({
  filename: "editor.css",
  code: Buffer.from(composed),
  minify: true,
});

await Bun.write(outFile, code);
console.log(
  `built dist/editor.css (${(code.length / 1024).toFixed(1)} kB, ${fontImports.size} font import preserved)`,
);

// ---------------------------------------------------------------------------
// `dist/standalone.css` — the self-sufficient stylesheet.
//
// Same document + chrome CSS as `editor.css`, PLUS a pre-compiled copy of every
// Tailwind utility folio's own components use, scoped under `.folio-root` (see
// `src/styles/standalone.css` + `../tailwind.config.js`). A consumer imports
// this one file and needs no Tailwind pipeline of their own. `editor.css` and
// its own-Tailwind contract are untouched.
//
// Compilation runs the local `@tailwindcss/cli`: it scans the package source
// (`@source` in the entry) and emits utilities in a canonical, deterministic
// order, so the output is reproducible from the source alone.
const standaloneEntry = path.join(stylesDir, "standalone.css");
const standaloneOut = path.join(pkgDir, "dist", "standalone.css");
// Resolve the CLI entry from the installed package (its `exports` map does not
// expose the dist path directly, so resolve via package.json + its `bin`).
const tailwindCliPkgJson = Bun.resolveSync(
  "@tailwindcss/cli/package.json",
  pkgDir,
);
const tailwindCliBin = (await Bun.file(tailwindCliPkgJson).json()).bin;
const tailwindCli = path.resolve(
  path.dirname(tailwindCliPkgJson),
  typeof tailwindCliBin === "string"
    ? tailwindCliBin
    : tailwindCliBin.tailwindcss,
);
const utilitiesTmp = path.join(pkgDir, "dist", ".standalone-utilities.css");

const twResult =
  await $`bun ${tailwindCli} --input ${standaloneEntry} --output ${utilitiesTmp} --minify`
    .cwd(pkgDir)
    .nothrow();
if (twResult.exitCode !== 0) {
  panic(
    `build-css: tailwindcss failed to compile standalone utilities\n${twResult.stderr}`,
  );
}
const utilitiesCss = await Bun.file(utilitiesTmp).text();
await rm(utilitiesTmp, { force: true });

// Order: the hoisted font import (must lead the file), then the bundled
// document/chrome CSS, then the compiled utilities layer last so a component's
// utility className wins over the chrome defaults it decorates. The compiled
// utilities carry no `@import`, so the font import stays at the top.
const standaloneComposed = `${fontImportBlock}\n${bundledBody}\n${utilitiesCss}`;
const { code: standaloneCode } = transform({
  filename: "standalone.css",
  code: Buffer.from(standaloneComposed),
  minify: true,
});

await Bun.write(standaloneOut, standaloneCode);
console.log(
  `built dist/standalone.css (${(standaloneCode.length / 1024).toFixed(1)} kB, self-sufficient: editor.css + scoped Tailwind utilities)`,
);
