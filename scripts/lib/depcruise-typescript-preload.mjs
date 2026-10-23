import { register } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

// Fixes CommonJS `require("typescript")` process-wide (see the patch file for
// why); the ESM side is registered as a loader hook right below.
await import("./depcruise-typescript-require-patch.cjs");

register(
  pathToFileURL(
    path.join(import.meta.dirname, "depcruise-typescript-loader.mjs"),
  ).href,
  pathToFileURL(import.meta.dirname + "/").href,
);
