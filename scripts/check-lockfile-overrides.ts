#!/usr/bin/env bun
// CI gate: every root `overrides` pin must satisfy the range each locked
// package and workspace declares for it, unless the break is listed below.
//
// Bun applies an override even when it falls outside a dependent's declared
// range, and `bun install --frozen-lockfile` does not report it. Raising one
// package (prosemirror-view, postcss) can therefore leave an older override
// (prosemirror-model, nanoid) below the new dependent's minimum.

import { join } from "node:path";

type Ranges = Record<string, string>;

type Manifest = {
  dependencies?: Ranges;
  devDependencies?: Ranges;
  optionalDependencies?: Ranges;
  peerDependencies?: Ranges;
};

type Lockfile = {
  overrides?: Ranges;
  workspaces: Record<string, Manifest & { name?: string }>;
  /** `[name@version, registry, manifest, integrity]`; workspace links omit the manifest. */
  packages: Record<string, readonly [string, ...unknown[]]>;
};

/** A deliberate override outside a dependent's range: the dependent's name and the range it declares. */
type IntentionalBreak = { dependent: string; range: string; reason: string };

const INTENTIONAL_BREAKS: Readonly<Record<string, IntentionalBreak>> = {
  archiver: {
    dependent: "nitropack",
    range: "^7.0.1",
    reason: "archiver 8 resolves a dependency advisory",
  },
  "simple-git": {
    dependent: "@nuxt/devtools",
    range: "^3.36.0",
    reason: "simple-git 4 resolves dependency advisories",
  },
  tinypool: {
    dependent: "oxfmt",
    range: "2.1.0",
    reason: "tinypool 2.2.0 resolves dependency advisories",
  },
  valibot: {
    dependent: "@stll/conditions",
    range: "1.4.2",
    reason: "one valibot across workspaces (#1279)",
  },
};

const DEPENDENCY_FIELDS = [
  "dependencies",
  "devDependencies",
  "optionalDependencies",
  "peerDependencies",
] as const;

type Declaration = { dependent: string; range: string };

const isManifest = (value: unknown): value is Manifest =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Strips the version from `name@version`, keeping a scope's leading `@`. */
const packageName = (spec: string): string => spec.slice(0, spec.lastIndexOf("@"));

const declarationsOf = (lock: Lockfile, name: string): Declaration[] => {
  const declarations: Declaration[] = [];
  const collect = (dependent: string, manifest: Manifest) => {
    for (const field of DEPENDENCY_FIELDS) {
      const range = manifest[field]?.[name];
      if (range !== undefined) declarations.push({ dependent, range });
    }
  };
  for (const [path, workspace] of Object.entries(lock.workspaces)) {
    collect(workspace.name ?? (path || "root"), workspace);
  }
  for (const [spec, ...rest] of Object.values(lock.packages)) {
    const manifest = rest.at(1);
    if (isManifest(manifest)) collect(packageName(spec), manifest);
  }
  return declarations;
};

const isSemverRange = (range: string): boolean =>
  !/^(?:workspace|npm|file|link|git|github|https?):/.test(range);

/** Every override outside a declared range that no entry permits, plus every stale entry. */
export const overrideViolations = (
  lock: Lockfile,
  intentional: Readonly<Record<string, IntentionalBreak>>,
): string[] => {
  const violations: string[] = [];
  const overrides = lock.overrides ?? {};
  for (const [name, version] of Object.entries(overrides)) {
    const allowed = intentional[name];
    let allowedUsed = false;
    for (const { dependent, range } of declarationsOf(lock, name)) {
      if (!isSemverRange(range) || Bun.semver.satisfies(version, range)) continue;
      if (allowed?.dependent === dependent && allowed.range === range) {
        allowedUsed = true;
        continue;
      }
      violations.push(`${dependent} declares ${name} ${range}; the override pins ${version}.`);
    }
    if (allowed && !allowedUsed) {
      violations.push(
        `${name}: ${allowed.dependent} no longer declares ${allowed.range} outside the override; remove the intentional break.`,
      );
    }
  }
  for (const name of Object.keys(intentional)) {
    if (!(name in overrides)) {
      violations.push(`${name}: intentional break listed without an override; remove it.`);
    }
  }
  return violations;
};

if (import.meta.main) {
  const lock: Lockfile = Bun.JSONC.parse(
    await Bun.file(join(import.meta.dirname, "..", "bun.lock")).text(),
  );
  const violations = overrideViolations(lock, INTENTIONAL_BREAKS);
  for (const violation of violations) console.error(violation);
  if (violations.length > 0) process.exit(1);
  console.log("bun.lock override check: every override satisfies its dependents. OK.");
}
