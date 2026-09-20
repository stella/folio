"use strict";

const path = require("node:path");
const Module = require("node:module");

/**
 * Companion to depcruise-typescript-loader.mjs: that file fixes ESM
 * `import("typescript")`, this fixes the plain CommonJS `require()` call
 * dependency-cruiser's tryImport() makes first (to read typescript's
 * package.json and check its version) before it ever reaches the ESM import.
 * CommonJS resolution is a separate algorithm from the ESM loader hooks and
 * isn't affected by them, so it needs its own redirect.
 */
const REPOSITORY_ROOT = path.resolve(__dirname, "../..");
const TYPESCRIPT_ROOT = path.join(REPOSITORY_ROOT, "node_modules/typescript");
const originalResolveFilename = Module._resolveFilename;

Module._resolveFilename = function resolveFilename(request, ...rest) {
  if (request === "typescript" || request.startsWith("typescript/")) {
    return originalResolveFilename.call(
      this,
      path.join(TYPESCRIPT_ROOT, request.slice("typescript".length)),
      ...rest,
    );
  }
  return originalResolveFilename.call(this, request, ...rest);
};
