#!/usr/bin/env bun
// Clean-room validation that the *published* shape of a folio package works.
//
// Usage: `bun scripts/validate-dist.ts <core|react>`
//
// For the named package it builds, transforms its package.json to the dist
// shape exactly like the publish workflow (`prepare-publish.ts`), packs a
// tarball with `bun pm pack` (which rewrites `workspace:` / `catalog:`
// protocols to concrete versions), then installs that tarball into a throwaway
// project OUTSIDE the monorepo and runs the package's checks against it.
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
//     4. External — React / react-dom / ProseMirror AND `@stll/folio-core` are
//                   imported as externals, never bundled into the JS.
//     5. Messages — the published `messages.d.ts` exports `FolioMessages` /
//                   `getFolioMessages` and never imports a `./messages/*.json`
//                   source file (dist inlines the JSON, ships no JSON).
//
// Exits non-zero on any failure. Run via `bun run validate-dist:<pkg>`.

import { panic } from "better-result";
import { $ } from "bun";
import { transform } from "lightningcss";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..");
const prepareScript = path.join(repoRoot, "scripts", "prepare-publish.ts");
const tscBin = path.join(repoRoot, "node_modules", ".bin", "tsc");

const target = process.argv[2];
if (target !== "core" && target !== "react") {
  panic("usage: bun scripts/validate-dist.ts <core|react>");
}

type CheckResult = { name: string; ok: boolean; detail: string };
const results: CheckResult[] = [];
const record = (name: string, ok: boolean, detail: string): void => {
  results.push({ name, ok, detail });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? ` — ${detail}` : ""}`);
};

// Build a package, transform its package.json to the dist shape (reversibly),
// pack a tarball into `destDir`, and return the tarball path.
const buildAndPack = async (pkgDir: string, destDir: string): Promise<string> => {
  const name = path.basename(pkgDir);
  console.log(`→ building @stll/folio-${name}`);
  await $`bun run build`.cwd(pkgDir).quiet();

  const pkgJsonPath = path.join(pkgDir, "package.json");
  const original = await readFile(pkgJsonPath, "utf-8");
  try {
    console.log(`→ transforming @stll/folio-${name} package.json to dist shape`);
    await $`bun ${prepareScript} ${pkgDir}`.quiet();
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

const coreDir = path.join(repoRoot, "packages", "core");
const reactDir = path.join(repoRoot, "packages", "react");

const packDir = await mkdtemp(path.join(tmpdir(), "folio-pack-"));
const corePackDir = await mkdtemp(path.join(tmpdir(), "folio-core-pack-"));
const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-consumer-"));

const pkgDir = target === "core" ? coreDir : reactDir;
const pkgName = `@stll/folio-${target}`;
const tarball = await buildAndPack(pkgDir, packDir);

// react depends on @stll/folio-core; pack it too so the clean room resolves it
// npm-style (no workspace leak). An `overrides` entry pins react's transitive
// dependency to this exact tarball, since the package is not yet on a registry.
let coreTarball: string | null = null;
if (target === "react") {
  coreTarball = await buildAndPack(coreDir, corePackDir);
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

const installArgs =
  target === "core"
    ? [tarball]
    : [
        tarball,
        coreTarball as string,
        "react@^19",
        "react-dom@^19",
        "use-intl@^4",
        "@types/react@^19",
        "@types/react-dom@^19",
      ];
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
    "@stll/folio-react/messages": ["getFolioMessages", "FOLIO_LOCALES", "isFolioLocale"],
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
const consumerTs =
  target === "core"
    ? `
import { createEmptyDocument, createDocx, type Document } from "@stll/folio-core";
import { fromMarkdown, toMarkdown, type MarkdownOptions } from "@stll/folio-core/markdown";
import { deriveBlockId, type FolioBlockId } from "@stll/folio-core/server";
import { isFolioBlockId } from "@stll/folio-core/types/block-id";

export const used = [createEmptyDocument, createDocx, fromMarkdown, toMarkdown, deriveBlockId, isFolioBlockId];
export type Surface = [Document, MarkdownOptions, FolioBlockId];
`
    : `
import { DocxEditor, FolioUIProvider, createDocx, type DocxEditorProps } from "@stll/folio-react";
import { getFolioMessages, type FolioMessages } from "@stll/folio-react/messages";

export const used = [DocxEditor, FolioUIProvider, createDocx, getFolioMessages];
export type Surface = [DocxEditorProps, FolioMessages];
`;
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
const typeChecks = await Promise.all(
  Object.entries(tsconfigs).map(async ([mode, opts]) => {
    const file = path.join(consumerDir, `tsconfig.${mode}.json`);
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
  record(
    `types: tsc --noEmit (moduleResolution: ${mode})`,
    tc.exitCode === 0,
    tc.exitCode === 0
      ? "consumer typechecks against published .d.ts"
      : `${tc.stdout.toString().trim()}${tc.stderr.toString().trim()}`.slice(0, 400),
  );
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
const expectedExternals =
  target === "core"
    ? ["prosemirror-state", "prosemirror-model", "jszip"]
    : ["react", "react-dom", "react/jsx-runtime", "@stll/folio-core", "prosemirror-view"];
const notExternalized = expectedExternals.filter(
  (p) => !new RegExp(`["']${p.replace(/\//gu, "\\/")}(?:/[^"']*)?["']`, "u").test(allJs),
);
const dataFontInlined = /["']data:font|["']data:application\/font/u.test(allJs);
const externalOk = leaked.length === 0 && notExternalized.length === 0 && !dataFontInlined;
record(
  target === "core"
    ? "external: React never bundled; deps stay external"
    : "external: React / ProseMirror / @stll/folio-core not bundled into JS",
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

// --- Check: `new URL(..., import.meta.url)` targets ship in dist ------------
// Every asset a published module references via `new URL("<rel>",
// import.meta.url)` (worker entries, wasm, fonts) must be a file that actually
// ships in the tarball. A downstream bundler resolves these specifiers
// literally against the emitted module (Vite/rolldown's
// `vite:worker-import-meta-url` treats the worker entry as a build entry), so a
// dangling target — e.g. a specifier left pointing at a `.ts` source that the
// package build never emits — aborts the consumer build with UNRESOLVED_ENTRY.
const urlTargetRe = /new URL\(\s*["']([^"']+)["']\s*,\s*import\.meta\.url\s*\)/gu;
const danglingUrlTargets: string[] = [];
let urlTargetCount = 0;
for (const { file, code } of jsContents) {
  for (const match of code.matchAll(urlTargetRe)) {
    const spec = match[1];
    // Only file-relative specifiers resolve against the module; skip absolute
    // URLs (`https:`, `data:`, `blob:`) and protocol-relative ones.
    if (/^[a-z][a-z0-9+.-]*:/iu.test(spec) || spec.startsWith("//")) {
      continue;
    }
    urlTargetCount += 1;
    const resolved = path.resolve(path.dirname(path.join(installedDist, file)), spec);
    if (!existsSync(resolved)) {
      danglingUrlTargets.push(`${file} -> ${spec}`);
    }
  }
}
record(
  "assets: new URL(..., import.meta.url) targets exist in dist",
  danglingUrlTargets.length === 0,
  danglingUrlTargets.length === 0
    ? `${urlTargetCount} URL target(s) resolve inside dist`
    : `missing target(s): ${danglingUrlTargets.join("; ")}`,
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
    const declaresSurface = dts.includes("FolioMessages") && dts.includes("getFolioMessages");
    const ok = jsonImports.length === 0 && declaresSurface;
    const detail = ok
      ? "FolioMessages / getFolioMessages self-contained; no source-JSON import"
      : [
          jsonImports.length > 0 && `imports source JSON: ${jsonImports.join(", ")}`,
          !declaresSurface && "missing FolioMessages / getFolioMessages declarations",
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
