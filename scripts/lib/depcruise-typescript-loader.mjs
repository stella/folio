import { pathToFileURL } from "node:url";
import path from "node:path";

/**
 * dependency-cruiser's tsc-based extractor (needed to tag an import
 * "type-only", which the no-circular rule in .dependency-cruiser.cjs relies
 * on to scope itself to runtime cycles) dynamically imports "typescript" from
 * its own package. Under this repo's isolated bun install, dependency-cruiser
 * resolves through a global-store symlink whose real path lives outside the
 * repository, so Node's realpath-based module resolution can never walk back
 * up to this repo's node_modules/typescript. This loader hook redirects the
 * ESM side of that lookup to a real, local one-line CommonJS shim that
 * re-exports the repo's installed typescript; redirecting straight to
 * typescript's own entry file here produces an empty default export (the CJS
 * named-export synthesis Node performs for the ESM interop misbehaves when
 * the file arrives through a resolve hook), so the shim keeps that machinery
 * on its normal path. depcruise-typescript-require-patch.cjs handles the
 * CommonJS side of the same lookup (dependency-cruiser's version check).
 */
const REPOSITORY_ROOT = path.resolve(import.meta.dirname, "../..");
const SHIM_URL = pathToFileURL(
  path.join(REPOSITORY_ROOT, "scripts/lib/depcruise-typescript-shim.cjs"),
).href;

export function resolve(specifier, context, nextResolve) {
  if (specifier === "typescript") {
    return { shortCircuit: true, url: SHIM_URL };
  }
  return nextResolve(specifier, context);
}
