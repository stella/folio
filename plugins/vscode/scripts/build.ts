#!/usr/bin/env bun
// Build the extension into dist/:
//
//   dist/extension.js      the extension host bundle (CommonJS, `vscode` external)
//   dist/webview/main.js   the preview webview's script
//   dist/cli/folio.mjs     the folio CLI from packages/cli, bundled with its
//                          dependencies, which the extension runs for the
//                          preview and the MCP server
//   dist/cli/text_shaper_bg.wasm, dist/cli/node_modules/@fontsource/*
//                          the files the CLI loads at run time
//
// It also copies the repository's LICENSE next to the manifest for packaging.
// Run `bun install` at the repository root first: the CLI's dependencies
// resolve from there.

import { build, type Plugin } from "esbuild";
import { copyFile, mkdir, readdir, readFile, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

const extensionRoot = path.resolve(import.meta.dir, "..");
const repoRoot = path.resolve(extensionRoot, "..", "..");
const cliRoot = path.join(repoRoot, "packages", "cli");
const dist = path.join(extensionRoot, "dist");
const cliDist = path.join(dist, "cli");

const SHAPER_WASM = "text_shaper_bg.wasm";
const SHAPER_DEFAULT_SOURCE = "module_or_path = new URL('./text_shaper_bg.wasm', import.meta.url);";

/**
 * The generated shaper glue fetches its `.wasm` by URL, and Node.js cannot
 * fetch a `file:` URL. The bundled CLI reads the file beside it instead.
 */
const shaperFromDisk: Plugin = {
  name: "folio-shaper-from-disk",
  setup: (builder) => {
    builder.onLoad({ filter: /[\\/]generated[\\/]text_shaper\.js$/ }, async (args) => {
      const source = await readFile(args.path, "utf8");
      if (!source.includes(SHAPER_DEFAULT_SOURCE)) {
        throw new Error(`${args.path} no longer loads its wasm as this build expects.`);
      }
      return {
        contents: source.replace(
          SHAPER_DEFAULT_SOURCE,
          `module_or_path = await import("node:fs/promises").then((fs) => fs.readFile(new URL("./${SHAPER_WASM}", import.meta.url)));`,
        ),
        loader: "js",
      };
    });
  },
};

/** The `@fontsource` packages the CLI depends on, and so reads faces from. */
const fontsourcePackages = async (): Promise<string[]> => {
  const manifest: unknown = JSON.parse(await readFile(path.join(cliRoot, "package.json"), "utf8"));
  const dependencies =
    typeof manifest === "object" && manifest !== null && "dependencies" in manifest
      ? manifest.dependencies
      : undefined;
  if (typeof dependencies !== "object" || dependencies === null) {
    throw new Error("packages/cli/package.json has no dependencies.");
  }
  return Object.keys(dependencies).filter((name) => name.startsWith("@fontsource/"));
};

/** Faces folio paints: regular and bold, upright and italic, as `.woff`. */
const FACE_FILE = /-(?:400|700)-(?:normal|italic)\.woff$/u;

/**
 * Copy what the CLI's font loader reads (`package.json` to resolve the
 * package, `unicode.json`, the `.woff` faces) and each font's licence into
 * `dist/cli/node_modules`, where the bundled CLI's `require.resolve` finds them.
 */
const copyFonts = async (): Promise<void> => {
  const cliRequire = createRequire(path.join(cliRoot, "package.json"));
  for (const name of await fontsourcePackages()) {
    const source = path.dirname(cliRequire.resolve(`${name}/package.json`));
    const target = path.join(cliDist, "node_modules", name);
    await mkdir(path.join(target, "files"), { recursive: true });
    for (const file of ["package.json", "unicode.json", "LICENSE"]) {
      await copyFile(path.join(source, file), path.join(target, file));
    }
    const faces = (await readdir(path.join(source, "files"))).filter((file) =>
      FACE_FILE.test(file),
    );
    if (faces.length === 0) throw new Error(`${name} has no faces to copy.`);
    for (const face of faces) {
      await copyFile(path.join(source, "files", face), path.join(target, "files", face));
    }
  }
};

await rm(dist, { recursive: true, force: true });

await Promise.all([
  build({
    entryPoints: [path.join(extensionRoot, "src", "extension.ts")],
    outfile: path.join(dist, "extension.js"),
    bundle: true,
    platform: "node",
    format: "cjs",
    target: "node22",
    external: ["vscode"],
    logLevel: "warning",
  }),
  build({
    entryPoints: [path.join(extensionRoot, "src", "webview", "main.ts")],
    outfile: path.join(dist, "webview", "main.js"),
    bundle: true,
    platform: "browser",
    format: "iife",
    target: "es2022",
    logLevel: "warning",
  }),
  build({
    entryPoints: [path.join(cliRoot, "src", "bin.ts")],
    outfile: path.join(cliDist, "folio.mjs"),
    bundle: true,
    platform: "node",
    format: "esm",
    target: "node22",
    // PNG output only: an optional peer the CLI imports on demand.
    external: ["playwright-core"],
    // CommonJS dependencies call `require` for Node.js built-ins.
    banner: {
      js: "import { createRequire as __folioCreateRequire } from 'node:module'; const require = __folioCreateRequire(import.meta.url);",
    },
    plugins: [shaperFromDisk],
    logLevel: "warning",
  }),
]);

await copyFile(
  path.join(repoRoot, "packages", "core", "src", "generated", SHAPER_WASM),
  path.join(cliDist, SHAPER_WASM),
);
await copyFonts();
await copyFile(path.join(repoRoot, "LICENSE"), path.join(extensionRoot, "LICENSE"));

console.log(`Built ${path.relative(repoRoot, dist)}`);
