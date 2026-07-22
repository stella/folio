#!/usr/bin/env bun

// Invokes the native TypeScript 7 compiler (installed as the aliased
// `@typescript/native` package) directly, bypassing the `tsc` bin shim.
// Official 7.0 guidance: keep classic TypeScript 6.x side-by-side for tools
// that still need the JS Compiler API (named-exports.mjs, vue-tsc, nuxt,
// tsdown peers). The 7.0 npm package does not export that API (expected in
// 7.1). Resolving the binary by path keeps the two compilers unambiguous.
// See: https://devblogs.microsoft.com/typescript/announcing-typescript-7-0/

import { spawnSync } from "node:child_process";
import path from "node:path";

const tsc = path.join(import.meta.dir, "../node_modules/@typescript/native/bin/tsc");

const result = spawnSync(process.execPath, [tsc, ...process.argv.slice(2)], {
  stdio: "inherit",
});

if (result.error) {
  console.error(result.error);
}

process.exit(result.status ?? 1);
