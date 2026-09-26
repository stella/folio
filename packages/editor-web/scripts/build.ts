#!/usr/bin/env bun
// Build the editor bundle a VS Code custom-editor webview loads:
//
//   dist/vscode/editor.js    one minified IIFE: the editor, React and the
//                            webview protocol. It loads no further chunks,
//                            workers or wasm at run time.
//   dist/vscode/editor.css   folio's self-sufficient stylesheet
//                            (`@stll/folio-react/standalone.css`) plus the
//                            VS Code theme mapping in `src/vscode.css`
//   dist/vscode/fonts/       the bundled document fonts, `.woff2` only
//   dist/vscode/editor.js.LEGAL.txt
//                            the licence notices of the bundled packages
//
// The webview's Content Security Policy admits exactly this shape:
//   default-src 'none'; script-src 'nonce-…' <cspSource>;
//   style-src <cspSource> 'unsafe-inline'; font-src <cspSource> data: blob:;
//   img-src <cspSource> data: blob:; worker-src 'none'; connect-src 'none'
// so nothing here may fetch, spawn a worker, or evaluate code from a string.
//
// Run `bun install` at the repository root first.

import { $ } from "bun";
import { build, type Plugin } from "esbuild";
import { readdir, readFile, rm, stat } from "node:fs/promises";
import path from "node:path";

const packageRoot = path.resolve(import.meta.dir, "..");
const reactRoot = path.dirname(Bun.resolveSync("@stll/folio-react/package.json", packageRoot));
const outdir = path.join(packageRoot, "dist", "vscode");

// esbuild matches plugin filters as Go regular expressions, which take no flags.
// oxlint-disable-next-line require-unicode-regexp -- esbuild filter (Go syntax, no flags)
const STYLESHEET_FILTER = /\.css$/;
// oxlint-disable-next-line require-unicode-regexp -- esbuild filter (Go syntax, no flags)
const ANY_PATH_FILTER = /.*/;
// oxlint-disable-next-line require-unicode-regexp -- esbuild filter (Go syntax, no flags)
const FONTSOURCE_STYLESHEET_FILTER = /[\\/]@fontsource[\\/].*\.css$/;

/**
 * The React sources import stylesheets for their bundler; `standalone.css`
 * already carries every one of them, so the script bundle drops them rather
 * than emit a second, partial stylesheet.
 */
const DROPPED_NAMESPACE = "folio-dropped-stylesheet";

const dropStylesheetImports: Plugin = {
  name: "folio-drop-stylesheet-imports",
  setup: (builder) => {
    builder.onResolve({ filter: STYLESHEET_FILTER }, ({ path: stylesheet }) => ({
      path: stylesheet,
      namespace: DROPPED_NAMESPACE,
    }));
    builder.onLoad({ filter: ANY_PATH_FILTER, namespace: DROPPED_NAMESPACE }, () => ({
      contents: "",
      loader: "js",
    }));
  },
};

/**
 * `@fontsource` faces list a `.woff` fallback after each `.woff2`. The
 * webview's Chromium reads `.woff2`, so the fallbacks would only double the
 * shipped fonts.
 */
const WOFF_FALLBACK = /,\s*url\([^)]*\.woff\)\s*format\(["']woff["']\)/gu;

const woff2Only: Plugin = {
  name: "folio-woff2-only",
  setup: (builder) => {
    builder.onLoad({ filter: FONTSOURCE_STYLESHEET_FILTER }, async (args) => {
      const source = await readFile(args.path, "utf8");
      const stripped = source.replace(WOFF_FALLBACK, "");
      if (stripped.includes(".woff)")) {
        throw new Error(`${args.path} lists a .woff face this build cannot drop.`);
      }
      return { contents: stripped, loader: "css" };
    });
  },
};

// `standalone.css` is compiled from folio-react's sources with Tailwind; build
// it fresh so the bundle never ships a stale copy.
await $`bun scripts/build-css.ts`.cwd(reactRoot).quiet();

await rm(outdir, { recursive: true, force: true });

await Promise.all([
  build({
    entryPoints: { editor: path.join(packageRoot, "src", "vscode.tsx") },
    outdir,
    bundle: true,
    format: "iife",
    platform: "browser",
    target: "es2022",
    minify: true,
    legalComments: "linked",
    jsx: "automatic",
    define: {
      "process.env.NODE_ENV": JSON.stringify("production"),
      // The measure worker is addressed relative to `import.meta.url`, which an
      // IIFE has none of; an undefined base makes that URL throw, and the
      // measurer falls back to the main thread.
      "import.meta.url": "undefined",
    },
    plugins: [dropStylesheetImports],
    logLevel: "warning",
  }),
  build({
    stdin: {
      contents: [
        `@import ${JSON.stringify(path.join(reactRoot, "dist", "standalone.css"))};`,
        `@import "./src/vscode.css";`,
      ].join("\n"),
      resolveDir: packageRoot,
      sourcefile: "editor.css",
      loader: "css",
    },
    outfile: path.join(outdir, "editor.css"),
    bundle: true,
    minify: true,
    loader: { ".woff2": "file" },
    assetNames: "fonts/[name]-[hash]",
    plugins: [woff2Only],
    logLevel: "warning",
  }),
]);

const sizeOf = async (file: string) => (await stat(file)).size;
const kib = (bytes: number) => `${(bytes / 1024).toFixed(1)} KiB`;
const fonts = await readdir(path.join(outdir, "fonts"));
const fontBytes = (
  await Promise.all(fonts.map((font) => sizeOf(path.join(outdir, "fonts", font))))
).reduce((total, size) => total + size, 0);

console.log(
  [
    `Built ${path.relative(packageRoot, outdir)}`,
    `  editor.js   ${kib(await sizeOf(path.join(outdir, "editor.js")))}`,
    `  editor.css  ${kib(await sizeOf(path.join(outdir, "editor.css")))}`,
    `  fonts/      ${kib(fontBytes)} in ${String(fonts.length)} .woff2 files`,
  ].join("\n"),
);
