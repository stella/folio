#!/usr/bin/env bun
// CI gate: each bun.lock records the `trustedDependencies` its package.json
// declares.
//
// An empty list is the repository's stance that no dependency may run an
// install script. Some lockfile writers (Dependabot's among them) omit an
// empty list, and Bun then falls back to its built-in trusted packages, so
// the lock no longer states the policy the manifest does. The scoped-install
// settle step reports the same drift as an opaque "changed trustedDependencies".

import { readFileSync } from "node:fs";
import { join } from "node:path";

type Manifest = { trustedDependencies?: unknown };

/** Where a manifest and its lock live, relative to the repository root. */
const PROJECTS = [".", "plugins/vscode"] as const;

const canonical = (value: unknown): string =>
  JSON.stringify(Array.isArray(value) ? [...value].toSorted() : value);

/** Why a lock does not record its manifest's trusted dependencies, or nothing. */
export const trustViolation = (
  project: string,
  manifest: Manifest,
  lock: Manifest,
): string | undefined =>
  canonical(manifest.trustedDependencies) ===
  canonical(lock.trustedDependencies)
    ? undefined
    : `${join(project, "bun.lock")} records trustedDependencies ${canonical(lock.trustedDependencies)} but ${join(project, "package.json")} declares ${canonical(manifest.trustedDependencies)}. Run \`bun install\` in ${project} with the pinned Bun and commit the lockfile.`;

if (import.meta.main) {
  const root = join(import.meta.dir, "..");
  const violations = PROJECTS.flatMap((project) => {
    const manifest: Manifest = JSON.parse(
      readFileSync(join(root, project, "package.json"), "utf8"),
    );
    const lock: Manifest = Bun.JSONC.parse(
      readFileSync(join(root, project, "bun.lock"), "utf8"),
    );
    const violation = trustViolation(project, manifest, lock);
    return violation === undefined ? [] : [violation];
  });
  if (violations.length > 0) {
    console.error(violations.join("\n"));
    process.exit(1);
  }
}
