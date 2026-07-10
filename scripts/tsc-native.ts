#!/usr/bin/env bun

// Invokes the native TypeScript 7 compiler (installed as the aliased
// `@typescript/native` package) directly, bypassing the `tsc` bin shim.
// The classic `typescript` package is pinned to 6.x for tooling that still
// needs the JS compiler API (tsdown, vue-tsc, nuxt build), and its `tsc` bin
// would otherwise shadow the native one. Resolving the binary by path keeps
// the two compilers unambiguous.

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
