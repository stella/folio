#!/usr/bin/env bun
// Build the extension into dist/:
//
//   dist/extension.js      the extension host bundle (CommonJS, `vscode` external)
//   dist/cli/folio.mjs     the folio CLI from packages/cli, bundled with its
//                          dependencies, which the extension runs to save
//                          and as the MCP server
//   dist/cli/text_shaper_bg.wasm, dist/cli/fonts/
//                          the files the CLI loads at run time
//   dist/editor/           the editor webview's bundle (editor.js, editor.css,
//                          fonts/, licence notices), copied from
//                          packages/editor-web/dist/vscode, which is built
//                          first when it is missing
//
// It also copies the repository's LICENSE next to the manifest for packaging.
// Run `bun install` at the repository root first: the CLI's dependencies
// resolve from there.

import { $ } from "bun";
import { build, type Plugin } from "esbuild";
import {
  copyFile,
  cp,
  mkdir,
  readdir,
  readFile,
  rm,
  stat,
} from "node:fs/promises";
import path from "node:path";

import { copyFuzzyWasm, fuzzyFromDisk } from "./portable-fuzzy";

const extensionRoot = path.resolve(import.meta.dir, "..");
const repoRoot = path.resolve(extensionRoot, "..", "..");
const cliRoot = path.join(repoRoot, "packages", "cli");
const editorWebRoot = path.join(repoRoot, "packages", "editor-web");
const editorWebDist = path.join(editorWebRoot, "dist", "vscode");
const dist = path.join(extensionRoot, "dist");
const cliDist = path.join(dist, "cli");

const SHAPER_WASM = "text_shaper_bg.wasm";
const SHAPER_DEFAULT_SOURCE =
  "module_or_path = new URL('./text_shaper_bg.wasm', import.meta.url);";

/**
 * The generated shaper glue fetches its `.wasm` by URL, and Node.js cannot
 * fetch a `file:` URL. The bundled CLI reads the file beside it instead.
 */
const shaperFromDisk: Plugin = {
  name: "folio-shaper-from-disk",
  setup: (builder) => {
    builder.onLoad(
      { filter: /[\\/]generated[\\/]text_shaper\.js$/ },
      async (args) => {
        const source = await readFile(args.path, "utf8");
        if (!source.includes(SHAPER_DEFAULT_SOURCE)) {
          throw new Error(
            `${args.path} no longer loads its wasm as this build expects.`,
          );
        }
        return {
          contents: source.replace(
            SHAPER_DEFAULT_SOURCE,
            `module_or_path = await import("node:fs/promises").then((fs) => fs.readFile(new URL("./${SHAPER_WASM}", import.meta.url)));`,
          ),
          loader: "js",
        };
      },
    );
  },
};

const fontsRoot = path.join(repoRoot, "packages", "fonts");
const FONTS_ROOT_SOURCE =
  'const FONTS_ROOT = new URL("../fonts/", import.meta.url);';

/**
 * `@stll/folio-fonts/node` finds its files one level above its module. Bundled
 * into `dist/cli/folio.mjs`, the module and the copied files sit side by side.
 */
const fontsBesideCli: Plugin = {
  name: "folio-fonts-beside-cli",
  setup: (builder) => {
    builder.onLoad(
      { filter: /[\\/]packages[\\/]fonts[\\/]src[\\/]node\.ts$/ },
      async (args) => {
        const source = await readFile(args.path, "utf8");
        if (!source.includes(FONTS_ROOT_SOURCE)) {
          throw new Error(
            `${args.path} no longer locates its fonts as this build expects.`,
          );
        }
        return {
          contents: source.replace(
            FONTS_ROOT_SOURCE,
            'const FONTS_ROOT = new URL("./fonts/", import.meta.url);',
          ),
          loader: "ts",
        };
      },
    );
  },
};

/**
 * Copy what the CLI's font reader loads (each family's `unicode.json`, its
 * `.woff` faces, and its licence) into `dist/cli/fonts`. The `.woff2` files
 * are for browsers; the editor webview ships its own copies.
 */
const copyFonts = async (): Promise<void> => {
  const source = path.join(fontsRoot, "fonts");
  const target = path.join(cliDist, "fonts");
  let copied = 0;
  for (const entry of await readdir(source, {
    recursive: true,
    withFileTypes: true,
  })) {
    if (!entry.isFile() || entry.name.endsWith(".woff2")) continue;
    const from = path.join(entry.parentPath, entry.name);
    const to = path.join(target, path.relative(source, from));
    await mkdir(path.dirname(to), { recursive: true });
    await copyFile(from, to);
    if (entry.name.endsWith(".woff")) copied += 1;
  }
  if (copied === 0) throw new Error(`${source} has no faces to copy.`);
};

const exists = async (file: string): Promise<boolean> => {
  try {
    await stat(file);
    return true;
  } catch {
    return false;
  }
};

/** The files the editor webview loads; the bundle is built when one is missing. */
const EDITOR_FILES = ["editor.js", "editor.css", "fonts"];

const copyEditor = async (): Promise<void> => {
  const present = await Promise.all(
    EDITOR_FILES.map((file) => exists(path.join(editorWebDist, file))),
  );
  if (!present.every(Boolean)) await $`bun scripts/build.ts`.cwd(editorWebRoot);
  await cp(editorWebDist, path.join(dist, "editor"), { recursive: true });
};

await rm(dist, { recursive: true, force: true });

await Promise.all([
  build({
    entryPoints: [path.join(extensionRoot, "src", "extension.ts")],
    outfile: path.join(dist, "extension.js"),
    bundle: true,
    platform: "node",
    format: "cjs",
    target: "node18",
    external: ["vscode"],
    logLevel: "warning",
  }),
  build({
    entryPoints: [path.join(cliRoot, "src", "bin.ts")],
    outfile: path.join(cliDist, "folio.mjs"),
    bundle: true,
    platform: "node",
    format: "esm",
    target: "node18",
    // PNG output only: an optional peer the CLI imports on demand.
    external: ["playwright-core"],
    // CommonJS dependencies call `require` for Node.js built-ins.
    banner: {
      // Electron 27 exposes Web Crypto through node:crypto, without the global
      // used by the browser-native document model. Initialize only this CLI process.
      js: "import { createRequire as __folioCreateRequire } from 'node:module'; import { webcrypto as __folioWebCrypto } from 'node:crypto'; const require = __folioCreateRequire(import.meta.url); globalThis.crypto ??= __folioWebCrypto;",
    },
    plugins: [shaperFromDisk, fuzzyFromDisk, fontsBesideCli],
    logLevel: "warning",
  }),
]);

await copyFile(
  path.join(repoRoot, "packages", "core", "src", "generated", SHAPER_WASM),
  path.join(cliDist, SHAPER_WASM),
);
await copyFuzzyWasm(cliDist);
await copyFonts();
await copyEditor();
await copyFile(
  path.join(repoRoot, "LICENSE"),
  path.join(extensionRoot, "LICENSE"),
);

console.log(`Built ${path.relative(repoRoot, dist)}`);
