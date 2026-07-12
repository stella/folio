#!/usr/bin/env bun
// Clean-room validation that the *published* shape of a folio package works.
//
// Usage: `bun scripts/validate-dist.ts <core|react|agents|vue>`
//
// For the named package it builds, transforms its package.json to the dist
// shape exactly like the publish workflow (`prepare-publish.ts`), packs a
// tarball with `bun pm pack` (which rewrites `workspace:` / `catalog:`
// protocols to concrete versions), then installs that tarball into a throwaway
// project OUTSIDE the monorepo and runs the package's checks against it.
//
// vue already ships dist-shaped exports in-repo, so the prepare-publish
// transform is skipped for it (it only understands core/react/agents's
// `./src/*.ts` string exports). Its type check runs under `bundler` resolution
// only — the resolution Vite/Nuxt consumers use — because vite-plugin-dts
// emits extensionless relative re-exports that a node16/nodenext consumer
// rejects (a pre-existing vue build gap). nuxt is a Nuxt module (a different
// dist shape: `module.mjs` + `types.d.mts`) and is validated by its own
// `nuxt-module-build` step, not here.
//
//   core  — 4 checks:
//     1. Runtime  — ESM `import` of `.`, `/markdown`, `/server`, and a `./*`
//                   wildcard subpath loads with the expected exports.
//     2. Types    — a `.ts` consumer importing every surface typechecks under
//                   both `moduleResolution: node16` and `bundler`.
//     3. External — React is never bundled; prosemirror/jszip stay external.
//
//   react — 6 checks (the five below plus the messages subpath declaration):
//     1. Runtime  — ESM `import` of `@stll/folio-react` and
//                   `@stll/folio-react/messages` resolve, including its
//                   transitive `@stll/folio-core` imports (installed from a
//                   sibling tarball, npm-style). `getFolioMessages` returns a
//                   `{ folio }` catalog for a bundled and a fallback locale.
//     2. Types    — node16 + bundler.
//     3. CSS      — `dist/editor.css` exists, parses, carries the bundled rules,
//                   and preserves `@fontsource` `@import`s (not inlined).
//     4. External — React / react-dom / React Compiler runtime / ProseMirror
//                   AND `@stll/folio-core` are imported as externals, never
//                   bundled into the JS. Requiring the compiler runtime import
//                   also proves the published JS passed through the compiler.
//     5. Messages — the published `messages.d.ts` exports `FolioMessages` /
//                   `getFolioMessages` and never imports a `./messages/*.json`
//                   source file (dist inlines the JSON, ships no JSON).
//
//   agents — 3 checks (no CSS, no messages subpath — a headless tool-schema
//   library over @stll/folio-core):
//     1. Runtime  — ESM `import` of `@stll/folio-agents` resolves, including
//                   its transitive `@stll/folio-core` import (installed from a
//                   sibling tarball, npm-style), and exposes
//                   `getFolioToolDefinitions`, `executeFolioToolCall`,
//                   `toAnthropicTools`, `toOpenAITools`, `createReviewerBridge`,
//                   `createEditorRefBridge`, and `FOLIO_AGENT_TOOLS`.
//     2. Types    — node16 + bundler.
//     3. External — `@stll/folio-core` is imported as an external, never
//                   bundled into the JS.
//
// Exits non-zero on any failure. Run via `bun run validate-dist:<pkg>`.

import { panic } from "better-result";
import { $ } from "bun";
import { transform } from "lightningcss";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { scanDistUrlTargets } from "./dist-url-targets";
import {
  REACT_PEER_MAJORS,
  reactPeerInstallArgs,
  type ReactPeerMajor,
} from "./packaged-consumer-lib";
import { findUncoveredUtilities } from "./standalone-css-coverage";

const repoRoot = path.resolve(import.meta.dir, "..");
const prepareScript = path.join(repoRoot, "scripts", "prepare-publish.ts");
// Resolve the classic TypeScript compiler explicitly rather than via
// `.bin/tsc`: the native TS7 compiler is installed under the aliased
// `@typescript/native` package and also claims a `tsc` bin, so `.bin/tsc`
// is an ambiguous collision. The clean-room check simulates a consumer on
// the stable JS compiler, so pin it to the `typescript` (6.x) package.
const tscBin = path.join(repoRoot, "node_modules", "typescript", "bin", "tsc");

const target = process.argv[2];
if (target !== "core" && target !== "react" && target !== "agents" && target !== "vue") {
  panic("usage: bun scripts/validate-dist.ts <core|react|agents|vue>");
}

type CheckResult = { name: string; ok: boolean; detail: string };
const results: CheckResult[] = [];
const record = (name: string, ok: boolean, detail: string): void => {
  results.push({ name, ok, detail });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? ` — ${detail}` : ""}`);
};

// Build a package, transform its package.json to the dist shape (reversibly),
// pack a tarball into `destDir`, and return the tarball path.
//
// `prepare` runs `prepare-publish.ts` (source `./src/*.ts` exports -> dist).
// Skip it for packages that already ship dist-shaped exports in-repo
// (@stll/folio-vue): prepare-publish only understands `./src/*` string exports
// and panics on their `{ types, import }` objects. `bun pm pack` still rewrites
// their `workspace:` deps to concrete versions, so no transform is needed.
const buildAndPack = async (pkgDir: string, destDir: string, prepare: boolean): Promise<string> => {
  const name = path.basename(pkgDir);
  console.log(`→ building @stll/folio-${name}`);
  await $`bun run build`.cwd(pkgDir).quiet();

  const pkgJsonPath = path.join(pkgDir, "package.json");
  const original = await readFile(pkgJsonPath, "utf-8");
  try {
    if (prepare) {
      console.log(`→ transforming @stll/folio-${name} package.json to dist shape`);
      await $`bun ${prepareScript} ${pkgDir}`.quiet();
    }
    console.log(`→ packing @stll/folio-${name} tarball`);
    await $`bun pm pack --destination ${destDir}`.cwd(pkgDir).quiet();
  } finally {
    // Always restore the in-repo source-shape package.json.
    await writeFile(pkgJsonPath, original);
  }
  const tgz = (await readdir(destDir)).find((f) => f.endsWith(".tgz"));
  return tgz
    ? path.join(destDir, tgz)
    : panic(`validate-dist: bun pm pack produced no tarball for ${name}`);
};

const dirs: Record<string, string> = {
  core: path.join(repoRoot, "packages", "core"),
  react: path.join(repoRoot, "packages", "react"),
  agents: path.join(repoRoot, "packages", "agents"),
  vue: path.join(repoRoot, "packages", "vue"),
};
const coreDir = dirs.core;

const packDir = await mkdtemp(path.join(tmpdir(), "folio-pack-"));
const corePackDir = await mkdtemp(path.join(tmpdir(), "folio-core-pack-"));
const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-consumer-"));

// core/react/agents carry source-shaped exports that prepare-publish
// rewrites; vue already ships dist-shaped exports (skip the transform).
const needsPrepare = target !== "vue";
const pkgDir = dirs[target];
const pkgName = `@stll/folio-${target}`;
const tarball = await buildAndPack(pkgDir, packDir, needsPrepare);

// react, agents, and vue all depend on @stll/folio-core; pack it too so the
// clean room resolves it npm-style (no workspace leak). An `overrides` entry
// pins the transitive dependency to this exact tarball, since the package is
// not yet on a registry.
let coreTarball: string | null = null;
if (target === "react" || target === "agents" || target === "vue") {
  coreTarball = await buildAndPack(coreDir, corePackDir, true);
}

console.log(`→ installing tarball into ${consumerDir}`);
const consumerPkg: Record<string, unknown> = {
  name: "folio-dist-consumer",
  version: "0.0.0",
  private: true,
  type: "module",
};
if (coreTarball) {
  consumerPkg.overrides = { "@stll/folio-core": coreTarball };
}
await writeFile(
  path.join(consumerDir, "package.json"),
  `${JSON.stringify(consumerPkg, null, 2)}\n`,
);

const installArgs: string[] = [tarball];
if (target === "react") {
  if (!coreTarball) panic("validate-dist: react needs a @stll/folio-core tarball");
  installArgs.push(coreTarball, ...reactPeerInstallArgs("19"), "use-intl@^4");
}
if (target === "agents") {
  if (!coreTarball) panic("validate-dist: agents needs a @stll/folio-core tarball");
  installArgs.push(coreTarball);
}
if (target === "vue") {
  if (!coreTarball) panic("validate-dist: vue needs a @stll/folio-core tarball");
  installArgs.push(
    coreTarball,
    "vue@^3",
    "prosemirror-history@^1",
    "prosemirror-model@^1",
    "prosemirror-state@^1",
    "prosemirror-tables@^1",
    "prosemirror-view@^1",
  );
}
await $`bun add ${installArgs}`.cwd(consumerDir).quiet();

const installedDir = path.join(consumerDir, "node_modules", "@stll", `folio-${target}`);
const installedDist = path.join(installedDir, "dist");

// --- runtime expectations ---------------------------------------------------
const runtimeExpect: Record<string, Record<string, string[]>> = {
  core: {
    "@stll/folio-core": ["createEmptyDocument", "createDocx", "deriveBlockId"],
    "@stll/folio-core/markdown": ["toMarkdown", "fromMarkdown", "toMarkdownResult"],
    "@stll/folio-core/server": [
      "deriveBlockId",
      "createEmptyDocument",
      "createDocx",
      "FolioDocxReviewer",
      "applyFolioAIEditsToBuffer",
    ],
    "@stll/folio-core/types/block-id": ["deriveBlockId", "isFolioBlockId"],
  },
  react: {
    "@stll/folio-react": ["DocxEditor", "FolioUIProvider", "FormattingBar", "createDocx"],
    "@stll/folio-react/dialogs": [
      "WatermarkDialog",
      "HyperlinkDialog",
      "InsertTableDialog",
      "InsertImageDialog",
      "PasteSpecialDialog",
      "SplitCellDialog",
    ],
    "@stll/folio-react/messages": ["getFolioMessages", "FOLIO_LOCALES", "isFolioLocale"],
  },
  agents: {
    "@stll/folio-agents": [
      "getFolioToolDefinitions",
      "executeFolioToolCall",
      "toAnthropicTools",
      "toOpenAITools",
      "createReviewerBridge",
      "createEditorRefBridge",
      "FOLIO_AGENT_TOOLS",
      "parseSuggestChangesInput",
      "parseAddCommentInput",
    ],
  },
  vue: {
    "@stll/folio-vue": ["DocxEditor", "createDocx", "useWheelZoom", "i18nPlugin"],
    "@stll/folio-vue/composables": ["useDocxEditor", "useZoom", "useWheelZoom"],
    "@stll/folio-vue/dialogs": [
      "WatermarkDialog",
      "HyperlinkDialog",
      "InsertTableDialog",
      "InsertImageDialog",
      "PasteSpecialDialog",
      "SplitCellDialog",
    ],
  },
};

// The messages subpath must return a live `{ folio }` catalog for a bundled and
// a fallback locale, not merely export a name.
const messagesRuntimeCheck =
  target === "react"
    ? `
try {
  const m = await import("@stll/folio-react/messages");
  for (const loc of ["en", "xx"]) {
    const msgs = m.getFolioMessages(loc);
    if (!msgs || typeof msgs !== "object" || !("folio" in msgs)) {
      failed = true;
      console.error("getFolioMessages(" + JSON.stringify(loc) + ") did not resolve to a folio catalog");
    }
  }
} catch (err) { failed = true; console.error("messages subpath threw: " + (err?.message ?? err)); }
`
    : "";

// --- Check 1: runtime ESM import -------------------------------------------
const runtimeScript = `
const expect = ${JSON.stringify(runtimeExpect[target])};
let failed = false;
for (const [spec, names] of Object.entries(expect)) {
  try {
    const mod = await import(spec);
    const missing = names.filter((n) => !(n in mod));
    if (missing.length) { failed = true; console.error("missing from " + spec + ": " + missing.join(", ")); }
  } catch (err) { failed = true; console.error("import threw for " + spec + ": " + (err?.message ?? err)); }
}
${messagesRuntimeCheck}
process.exit(failed ? 1 : 0);
`;
const runtimeFile = path.join(consumerDir, "runtime-check.mjs");
await writeFile(runtimeFile, runtimeScript);
const runtime = await $`node ${runtimeFile}`.cwd(consumerDir).nothrow().quiet();
const runtimeSubpaths = Object.keys(runtimeExpect[target] ?? {}).length;
record(
  `runtime: ESM import of ${runtimeSubpaths} subpath(s)`,
  runtime.exitCode === 0,
  runtime.exitCode === 0
    ? "all subpaths load with expected exports"
    : runtime.stderr.toString().trim() || "non-zero exit",
);

// --- Check 2: types resolve under node16 AND bundler ------------------------
const consumerTsByTarget: Record<string, string> = {
  core: `
import { createEmptyDocument, createDocx, type Document } from "@stll/folio-core";
import { fromMarkdown, toMarkdown, type MarkdownOptions } from "@stll/folio-core/markdown";
import { deriveBlockId, type FolioBlockId } from "@stll/folio-core/server";
import { isFolioBlockId } from "@stll/folio-core/types/block-id";

export const used = [createEmptyDocument, createDocx, fromMarkdown, toMarkdown, deriveBlockId, isFolioBlockId];
export type Surface = [Document, MarkdownOptions, FolioBlockId];
`,
  react: `
import { DocxEditor, FolioUIProvider, createDocx, type DocxEditorProps } from "@stll/folio-react";
import { getFolioMessages, type FolioMessages } from "@stll/folio-react/messages";

export const used = [DocxEditor, FolioUIProvider, createDocx, getFolioMessages];
export type Surface = [DocxEditorProps, FolioMessages];
`,
  agents: `
import {
  getFolioToolDefinitions,
  executeFolioToolCall,
  toAnthropicTools,
  toOpenAITools,
  createReviewerBridge,
  createEditorRefBridge,
  FOLIO_AGENT_TOOLS,
} from "@stll/folio-agents";

export const used = [
  getFolioToolDefinitions,
  executeFolioToolCall,
  toAnthropicTools,
  toOpenAITools,
  createReviewerBridge,
  createEditorRefBridge,
  FOLIO_AGENT_TOOLS,
];
`,
  vue: `
import { DocxEditor, createDocx, type DocxEditorProps } from "@stll/folio-vue";
import { useDocxEditor, useZoom } from "@stll/folio-vue/composables";

export const used = [DocxEditor, createDocx, useDocxEditor, useZoom];
export type Surface = [DocxEditorProps];
`,
};
const consumerTs =
  consumerTsByTarget[target] ?? panic(`validate-dist: no consumer.ts for ${target}`);
await writeFile(path.join(consumerDir, "consumer.ts"), consumerTs);

const baseCompilerOptions = {
  target: "es2022",
  jsx: "react-jsx",
  strict: true,
  skipLibCheck: true,
  noEmit: true,
  lib: ["es2022", "dom", "dom.iterable"],
};
const tsconfigs: Record<string, { module: string; moduleResolution: string }> = {
  node16: { module: "node16", moduleResolution: "node16" },
  bundler: { module: "preserve", moduleResolution: "bundler" },
};
// vue's declarations come from vite-plugin-dts, which emits extensionless
// relative re-exports (`export * from './useZoom'`). Those resolve under
// `bundler` (what Vite/Nuxt consumers use) but not under node16/nodenext ESM,
// which requires explicit `.js` specifiers. Vue/Nuxt consumers use bundler
// resolution, so check that here; node16-clean declarations are a pre-existing
// vue build gap tracked separately. core/react/agents (tsdown) stay checked in both.
const typeCheckModes =
  target === "vue"
    ? Object.entries(tsconfigs).filter(([mode]) => mode === "bundler")
    : Object.entries(tsconfigs);
const typeCheckReactMajors: (ReactPeerMajor | null)[] =
  target === "react" ? [...REACT_PEER_MAJORS] : [null];
for (const reactMajor of typeCheckReactMajors) {
  if (reactMajor !== null) {
    await $`bun add ${reactPeerInstallArgs(reactMajor)}`.cwd(consumerDir).quiet();
  }

  const typeChecks = await Promise.all(
    typeCheckModes.map(async ([mode, opts]) => {
      const reactSuffix = reactMajor === null ? "" : `.react-${reactMajor}`;
      const file = path.join(consumerDir, `tsconfig.${mode}${reactSuffix}.json`);
      await writeFile(
        file,
        `${JSON.stringify(
          {
            compilerOptions: { ...baseCompilerOptions, ...opts },
            files: ["consumer.ts"],
          },
          null,
          2,
        )}\n`,
      );
      const tc = await $`${tscBin} -p ${file}`.cwd(consumerDir).nothrow().quiet();
      return { mode, tc };
    }),
  );
  for (const { mode, tc } of typeChecks) {
    const reactLabel = reactMajor === null ? "" : `, React ${reactMajor} types`;
    record(
      `types: tsc --noEmit (moduleResolution: ${mode}${reactLabel})`,
      tc.exitCode === 0,
      tc.exitCode === 0
        ? "consumer typechecks against published .d.ts"
        : `${tc.stdout.toString().trim()}${tc.stderr.toString().trim()}`.slice(0, 400),
    );
  }
}

// --- Check 3 (react only): bundled stylesheet -------------------------------
if (target === "react") {
  const cssPath = path.join(installedDist, "editor.css");
  if (!existsSync(cssPath)) {
    record("css: dist/editor.css present", false, "missing from tarball");
  } else {
    const css = await readFile(cssPath, "utf-8");
    const fontImports = css.match(/@import\s+["']@fontsource\/[^"']+["']/gu) ?? [];
    const requiredRules = [
      ".folio-root",
      ".ProseMirror",
      ".prosemirror-editor-wrapper",
      ".folio-ai-host",
      ".folio-default-button",
    ];
    const missingRules = requiredRules.filter((r) => !css.includes(r));
    // @fontsource must stay a bare @import, never inlined as font data.
    const fontsourceInlined =
      /url\([^)]*@fontsource/u.test(css) || /url\(["']?data:font/u.test(css);

    let parses = true;
    let parseDetail = "";
    try {
      transform({
        filename: "editor.css",
        code: Buffer.from(css),
        minify: false,
      });
    } catch (error) {
      parses = false;
      parseDetail = error instanceof Error ? error.message : String(error);
    }

    const ok =
      css.length > 10_000 &&
      parses &&
      missingRules.length === 0 &&
      fontImports.length >= 20 &&
      !fontsourceInlined;
    const detail = ok
      ? `${(css.length / 1024).toFixed(1)} kB, ${fontImports.length} @fontsource imports preserved, all bundled rules present`
      : [
          css.length <= 10_000 && `too small (${css.length}B)`,
          !parses && `invalid CSS: ${parseDetail}`,
          missingRules.length > 0 && `missing rules: ${missingRules.join(", ")}`,
          fontImports.length < 20 && `only ${fontImports.length} @fontsource imports`,
          fontsourceInlined && "@fontsource appears inlined as font data",
        ]
          .filter(Boolean)
          .join("; ");
    record("css: bundled, valid, @fontsource preserved", ok, detail);
  }

  // --- Check 3b: self-sufficient standalone.css -----------------------------
  // Same document + chrome CSS as editor.css PLUS pre-compiled Tailwind
  // utilities scoped under `.folio-root`, so a consumer needs no Tailwind of
  // their own. Must be fully compiled (no un-expanded `@tailwind`/`@config`/
  // `@source` directives), carry scoped utilities + token fallbacks, and still
  // preserve the @fontsource `@import`s.
  const standalonePath = path.join(installedDist, "standalone.css");
  if (!existsSync(standalonePath)) {
    record("css: dist/standalone.css present", false, "missing from tarball");
  } else {
    const css = await readFile(standalonePath, "utf-8");
    const fontImports = css.match(/@import\s+["']@fontsource\/[^"']+["']/gu) ?? [];
    // Scoped utilities (`.folio-root .<utility>`) plus low-specificity token
    // fallbacks are what make the sheet self-sufficient.
    const hasScopedUtilities = /\.folio-root \.[a-z]/u.test(css);
    const hasTokenFallback = /:where\(\.folio-root\)/u.test(css) && css.includes("--popover:");
    const hasDocRules = css.includes(".ProseMirror") && css.includes(".folio-default-button");
    // No Tailwind source directive may survive into the shipped file.
    const uncompiled = /@tailwind\b|@config\b|@source\b|@apply\b/u.test(css);
    const fontsourceInlined =
      /url\([^)]*@fontsource/u.test(css) || /url\(["']?data:font/u.test(css);

    let parses = true;
    let parseDetail = "";
    try {
      transform({ filename: "standalone.css", code: Buffer.from(css), minify: false });
    } catch (error) {
      parses = false;
      parseDetail = error instanceof Error ? error.message : String(error);
    }

    const ok =
      css.length > 10_000 &&
      parses &&
      hasScopedUtilities &&
      hasTokenFallback &&
      hasDocRules &&
      !uncompiled &&
      fontImports.length >= 20 &&
      !fontsourceInlined;
    const detail = ok
      ? `${(css.length / 1024).toFixed(1)} kB, scoped utilities + token fallbacks + doc rules, ${fontImports.length} @fontsource imports preserved`
      : [
          css.length <= 10_000 && `too small (${css.length}B)`,
          !parses && `invalid CSS: ${parseDetail}`,
          !hasScopedUtilities && "no `.folio-root .<utility>` rules",
          !hasTokenFallback && "no `:where(.folio-root)` token fallbacks",
          !hasDocRules && "missing bundled document/chrome rules",
          uncompiled && "un-expanded Tailwind directive present",
          fontImports.length < 20 && `only ${fontImports.length} @fontsource imports`,
          fontsourceInlined && "@fontsource appears inlined as font data",
        ]
          .filter(Boolean)
          .join("; ");
    record("css: standalone.css self-sufficient (scoped utilities + tokens)", ok, detail);
  }
}

// --- Check 3 (vue only): bundled stylesheet ---------------------------------
// The vue adapter ships a single `dist/folio-vue.css` (the `./editor.css`
// export). Assert it is present, parses, and carries the editor's document +
// chrome rules — a minimal but real check that the packed CSS is usable.
if (target === "vue") {
  const cssPath = path.join(installedDist, "folio-vue.css");
  if (!existsSync(cssPath)) {
    record("css: dist/folio-vue.css present", false, "missing from tarball");
  } else {
    const css = await readFile(cssPath, "utf-8");
    const requiredRules = [".ep-root", ".basic-toolbar"];
    const missingRules = requiredRules.filter((r) => !css.includes(r));

    let parses = true;
    let parseDetail = "";
    try {
      transform({ filename: "folio-vue.css", code: Buffer.from(css), minify: false });
    } catch (error) {
      parses = false;
      parseDetail = error instanceof Error ? error.message : String(error);
    }

    const ok = css.length > 5_000 && parses && missingRules.length === 0;
    const detail = ok
      ? `${(css.length / 1024).toFixed(1)} kB, editor rules present`
      : [
          css.length <= 5_000 && `too small (${css.length}B)`,
          !parses && `invalid CSS: ${parseDetail}`,
          missingRules.length > 0 && `missing rules: ${missingRules.join(", ")}`,
        ]
          .filter(Boolean)
          .join("; ");
    record("css: bundled folio-vue.css present and valid", ok, detail);
  }
}

// --- Check 4: externals not bundled into the JS -----------------------------
// Recurse: @stll/folio-core ships a source-mirrored dist tree, so the dep
// imports live in nested modules, not only at the dist root.
const jsFiles = (await readdir(installedDist, { recursive: true })).filter((f) =>
  f.endsWith(".js"),
);
const jsContents = await Promise.all(
  jsFiles.map(async (f) => ({
    file: f,
    code: await readFile(path.join(installedDist, f), "utf-8"),
  })),
);
const allJs = jsContents.map((c) => c.code).join("\n");
// Tell-tale internals that only exist if React's source were bundled.
const reactSentinels = [
  "react-stack-bottom-frame",
  "__SECRET_INTERNALS_DO_NOT_USE",
  "__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED",
];
const leaked = reactSentinels.filter((s) => allJs.includes(s));
// Each external must appear only as an import specifier, never bundled.
const externalsByTarget: Record<string, string[]> = {
  core: ["prosemirror-state", "prosemirror-model", "jszip"],
  react: [
    "react",
    "react-dom",
    "react/jsx-runtime",
    "react-compiler-runtime",
    "@stll/folio-core",
    "prosemirror-view",
  ],
  agents: ["@stll/folio-core"],
  vue: ["vue", "@stll/folio-core", "prosemirror-history", "prosemirror-state"],
};
const expectedExternals = externalsByTarget[target] ?? [];
const notExternalized = expectedExternals.filter(
  (p) => !new RegExp(`["']${p.replace(/\//gu, "\\/")}(?:/[^"']*)?["']`, "u").test(allJs),
);
const dataFontInlined = /["']data:font|["']data:application\/font/u.test(allJs);
const externalOk = leaked.length === 0 && notExternalized.length === 0 && !dataFontInlined;
const externalLabels: Record<string, string> = {
  core: "external: React never bundled; deps stay external",
  react: "external: React / compiler runtime / ProseMirror / @stll/folio-core not bundled",
  agents: "external: @stll/folio-core not bundled into JS",
  vue: "external: Vue / ProseMirror / @stll/folio-core not bundled into JS",
};
record(
  externalLabels[target] ?? "external: declared externals not bundled into JS",
  externalOk,
  externalOk
    ? `${expectedExternals.length} externals imported by specifier; no bundled internals`
    : [
        leaked.length > 0 && `bundled React internals found: ${leaked.join(", ")}`,
        notExternalized.length > 0 && `not imported as external: ${notExternalized.join(", ")}`,
        dataFontInlined && "font data inlined in JS",
      ]
        .filter(Boolean)
        .join("; "),
);

// --- Check (react only): standalone.css covers every utility the JS uses ----
// Every Tailwind utility a shipped component references must have a generated
// rule in the packed standalone.css. Fails if a class in the dist JS was not
// picked up by the standalone compile (a `@source` glob that missed a new
// component dir, a stale build), which would leave a consumer on the standalone
// path with an unstyled control and no error.
if (target === "react") {
  const standalonePath = path.join(installedDist, "standalone.css");
  if (!existsSync(standalonePath)) {
    record("css: standalone.css utility coverage", false, "standalone.css missing from tarball");
  } else {
    const standaloneCss = await readFile(standalonePath, "utf-8");
    const { candidates, uncovered } = findUncoveredUtilities({
      jsFiles: jsContents,
      standaloneCss,
    });
    record(
      "css: standalone.css covers every utility used in dist JS",
      uncovered.length === 0,
      uncovered.length === 0
        ? `${candidates} utility class(es) referenced, all present in standalone.css`
        : `missing from standalone.css: ${uncovered.join(", ")}`,
    );
  }
}

// --- Check: `new URL(..., import.meta.url)` targets ship in dist ------------
// Every asset a published module references via `new URL("<rel>",
// import.meta.url)` (worker entries, wasm, fonts) must be a file that actually
// ships in the tarball. A downstream bundler resolves these specifiers
// literally against the emitted module (Vite/rolldown's
// `vite:worker-import-meta-url` treats the worker entry as a build entry), so a
// dangling target — e.g. a specifier left pointing at a `.ts` source that the
// package build never emits — aborts the consumer build with UNRESOLVED_ENTRY.
// Scan logic (skip absolute URLs, strip `?query`/`#hash`) lives in
// `dist-url-targets.ts` with its own unit tests.
const urlTargets = scanDistUrlTargets(jsContents, installedDist);
record(
  "assets: new URL(..., import.meta.url) targets exist in dist",
  urlTargets.dangling.length === 0,
  urlTargets.dangling.length === 0
    ? `${urlTargets.total} URL target(s) resolve inside dist`
    : `missing target(s): ${urlTargets.dangling.join("; ")}`,
);

// --- Check 5 (react only): messages subpath declaration is self-contained ---
if (target === "react") {
  const dtsPath = path.join(installedDist, "messages.d.ts");
  if (!existsSync(dtsPath)) {
    record("messages: dist/messages.d.ts present", false, "missing from tarball");
  } else {
    const dts = await readFile(dtsPath, "utf-8");
    // dist ships no locale JSON (tsdown inlines it into the JS), so a
    // `./messages/*.json` import in the published declaration is unresolvable
    // for a TS consumer of `@stll/folio-react/messages`.
    const jsonImports = dts.match(/["'][^"']*messages\/[^"']*\.json["']/gu) ?? [];
    // The i18n surface now lives in `@stll/folio-core`; the react subpath either
    // declares the names itself or re-exports them wholesale from core (which
    // the clean room installs, so `export * from "@stll/folio-core/i18n/messages"`
    // resolves for a consumer). The runtime and tsc checks above already exercise
    // `getFolioMessages` / `FolioMessages` through this re-export; here we just
    // guard that the declaration exposes the surface with no source-JSON import.
    const declaresNames = dts.includes("FolioMessages") && dts.includes("getFolioMessages");
    const reexportsCore = /export\s*\*\s*from\s*["']@stll\/folio-core\/i18n\/messages["']/u.test(
      dts,
    );
    const declaresSurface = declaresNames || reexportsCore;
    const ok = jsonImports.length === 0 && declaresSurface;
    const detail = ok
      ? "getFolioMessages / FolioMessages surface exposed; no source-JSON import"
      : [
          jsonImports.length > 0 && `imports source JSON: ${jsonImports.join(", ")}`,
          !declaresSurface && "neither declares nor re-exports FolioMessages / getFolioMessages",
        ]
          .filter(Boolean)
          .join("; ");
    record("messages: .d.ts free of source-JSON imports", ok, detail);
  }
}

// --- Summary ----------------------------------------------------------------
await rm(packDir, { recursive: true, force: true });
await rm(corePackDir, { recursive: true, force: true });
await rm(consumerDir, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
console.log(
  `\n${failed.length === 0 ? "✓" : "✗"} ${pkgName}: ${results.length - failed.length}/${results.length} checks passed`,
);
if (failed.length > 0) {
  process.exit(1);
}
