#!/usr/bin/env bun
/**
 * Guard `bunfig.toml`'s `minimumReleaseAgeExcludes`:
 *
 * 1. Every `@stll/*` package the lockfile resolves FROM NPM must be excluded.
 *    folio's own published packages (`@stll/folio-core`, `@stll/folio-react`,
 *    `@stll/folio-vue`, `@stll/folio-nuxt`, `@stll/folio-agents`,
 *    `@stll/docx-core`) resolve through the `workspace:` protocol and never
 *    reach the registry, so they are exempt. The packages that DO need an
 *    exclusion are the first-party ones folio CONSUMES from npm, which is why
 *    the set is read out of the lockfile rather than hard-coded: the moment a
 *    workspace member is dropped, or a new `@stll/*` dependency is added, the
 *    requirement follows on its own.
 * 2. A temporary third-party exclusion annotated with
 *    `# quarantine-expires: <timestamp>` must be removed once Bun's release-age
 *    gate can take over again.
 *
 * An expiry is a wall-clock instant, so failing at that instant turns every
 * open branch red at once for a change nobody made. Removal is automatic
 * instead: `quarantine-prune.yml` runs `--prune` hourly and opens the PR that
 * deletes the entry. An expired entry is therefore only a warning while that
 * PR waits to be merged, and fails once it has been ignored past the window.
 * It has to fail eventually: a stale exclude is not inert, because the next
 * version of that package published would skip the quarantine unnoticed.
 *
 * The 5-day quarantine is a supply-chain control for third-party code.
 * First-party packages publish continuously, so they are excluded by name. The
 * failure this guards is not the exclusion itself but a PARTIAL one: a napi
 * package ships its platform binaries as separate `optionalDependencies`
 * published in the same minute, each recorded as its own lockfile entry. List
 * the parent and forget the children and the parent installs while every binary
 * is quarantined — and because they are optional, bun skips them without a
 * word. The package then resolves, imports cleanly, and throws on first use, on
 * every platform. folio ships no napi package today (`crates/docx-kernel` is
 * compiled to WebAssembly and committed inside `@stll/docx-core`), so the rule
 * reads the lockfile rather than a list: platform subpackages appear there as
 * ordinary entries and become required the day one arrives.
 *
 * Bun matches these entries EXACTLY. A `@stll/*` glob is accepted and silently
 * ignored, which is why this cannot be solved by pattern and needs a guard.
 *
 * Reads `bunfig.toml` and `bun.lock` and nothing else, so it runs without an
 * install — `quarantine-prune.yml` depends on that.
 *
 * Run: `bun run check:quarantine-excludes`
 */

import { join } from "node:path";

const ROOT = join(import.meta.dirname, "..");
const LOCKFILE = "bun.lock";
const BUNFIG = "bunfig.toml";
const SCRIPT_PATH = "scripts/check-quarantine-excludes.ts";
const EXPIRY_MARKER = "quarantine-expires:";
const EXACT_UTC_TIMESTAMP = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/u;

/** Packages resolved from the workspace itself never hit the registry. */
const WORKSPACE_PROTOCOL = "workspace:";

const readExcludeBlock = (bunfig: string): string => {
  const start = bunfig.indexOf("minimumReleaseAgeExcludes");
  if (start === -1) return "";
  const end = bunfig.indexOf("]", start);
  return bunfig.slice(start, end === -1 ? undefined : end);
};

const readExcludes = (bunfig: string): Set<string> =>
  new Set(
    [...readExcludeBlock(bunfig).matchAll(/"(?<name>@?[^"]+)"/gu)].flatMap((match) => {
      const name = match.groups?.["name"];
      return name === undefined ? [] : [name];
    }),
  );

type TemporaryExclude = {
  name: string;
  expiresAt: string;
};

type TemporaryExcludesResult = {
  entries: TemporaryExclude[];
  errors: string[];
};

type ParsedTemporaryExclude =
  | { kind: "error"; error: string }
  | { kind: "entry"; name: string; expiresAt: string };

/**
 * One annotated line. Both the guard and the prune read it through here, so
 * neither can disagree with the other about what a line means.
 */
const parseTemporaryExcludeLine = (line: string): ParsedTemporaryExclude => {
  const commentStart = line.indexOf("#");
  const declaration = line.slice(0, commentStart).trim();
  const comment = line.slice(commentStart + 1).trim();
  const quotedName = declaration.endsWith(",") ? declaration.slice(0, -1).trim() : declaration;
  const isQuotedName =
    quotedName.startsWith('"') &&
    quotedName.endsWith('"') &&
    !quotedName.slice(1, -1).includes('"');
  if (!comment.startsWith(EXPIRY_MARKER) || !isQuotedName) {
    return {
      kind: "error",
      error: `${BUNFIG} has a malformed temporary quarantine annotation: ${line.trim()}`,
    };
  }

  const name = quotedName.slice(1, -1);
  const expiresAt = comment.slice(EXPIRY_MARKER.length).trim();
  const expiresAtMs = Date.parse(expiresAt);
  if (
    !EXACT_UTC_TIMESTAMP.test(expiresAt) ||
    Number.isNaN(expiresAtMs) ||
    new Date(expiresAtMs).toISOString() !== expiresAt
  ) {
    return {
      kind: "error",
      error: `${BUNFIG} temporary quarantine exclude "${name}" has an invalid UTC expiry: ${expiresAt}`,
    };
  }

  return { kind: "entry", name, expiresAt };
};

const readTemporaryExcludes = (bunfig: string): TemporaryExcludesResult => {
  const entries: TemporaryExclude[] = [];
  const errors: string[] = [];

  for (const line of readExcludeBlock(bunfig).split("\n")) {
    if (!line.includes(EXPIRY_MARKER)) continue;

    const parsed = parseTemporaryExcludeLine(line);
    if (parsed.kind === "error") {
      errors.push(parsed.error);
      continue;
    }

    entries.push({ name: parsed.name, expiresAt: parsed.expiresAt });
  }

  return { entries, errors };
};

/**
 * `@stll` packages the lockfile pulls from the registry. Workspace members
 * resolve locally and are never subject to the release-age gate.
 */
const readRegistryFirstPartyPackages = (lockfile: string): Set<string> =>
  new Set(
    // Scope plus exactly one segment. Bun also keys nested resolutions by path
    // ("@stll/template-conditions/better-result"); those are not package names.
    [...lockfile.matchAll(/"(?<name>@stll\/[^"/]+)":\s*\[/gu)].flatMap((match) => {
      const name = match.groups?.["name"];
      if (name === undefined) return [];
      // The entry's own resolution follows the name; workspace members carry
      // the workspace protocol instead of a registry tarball. Read to the end
      // of that line and no further: a fixed-width window spills into the next
      // entry, and one workspace neighbour then hides a registry-resolved
      // package from the coverage check entirely.
      const lineEnd = lockfile.indexOf("\n", match.index);
      const entry = lockfile.slice(match.index, lineEnd === -1 ? undefined : lineEnd);
      return entry.includes(WORKSPACE_PROTOCOL) ? [] : [name];
    }),
  );

const HOUR_MS = 60 * 60 * 1000;
/** How long an expired entry is tolerated while its removal PR waits. */
const NOTICE_WINDOW_MS = 24 * HOUR_MS;

export type QuarantineExcludeCheckResult = {
  errors: string[];
  warnings: string[];
  excludeCount: number;
  firstPartyCount: number;
  activeTemporaryCount: number;
};

export const checkQuarantineExcludes = ({
  bunfig,
  lockfile,
  now = new Date(),
}: {
  bunfig: string;
  lockfile: string;
  now?: Date;
}): QuarantineExcludeCheckResult => {
  const excludes = readExcludes(bunfig);
  const firstPartyPackages = readRegistryFirstPartyPackages(lockfile);
  const missing = [...firstPartyPackages].filter((name) => !excludes.has(name)).sort();
  const temporary = readTemporaryExcludes(bunfig);
  const errors = [...temporary.errors];

  if (missing.length > 0) {
    errors.push(
      `${BUNFIG} minimumReleaseAgeExcludes is missing ${missing.length} first-party package(s):\n${missing
        .map((name) => `  "${name}",`)
        .join("\n")}\n\nAdd them. Until then, a fresh publish of any of these installs ` +
        `partially: the quarantine blocks it, and an optionalDependency that is ` +
        `blocked is skipped silently, so the failure surfaces at runtime rather ` +
        `than at install.`,
    );
  }

  const nowMs = now.getTime();
  const warnings: string[] = [];
  for (const { name, expiresAt } of temporary.entries) {
    const expiresAtMs = Date.parse(expiresAt);

    if (nowMs >= expiresAtMs + NOTICE_WINDOW_MS) {
      errors.push(
        `${BUNFIG} temporary quarantine exclude "${name}" expired at ${expiresAt} ` +
          `and is still here. Remove it (\`bun ${SCRIPT_PATH} --prune\`): the ` +
          `release-age gate admits the package on its own now, and leaving the ` +
          `entry lets the next version of it publish straight past the quarantine.`,
      );
      continue;
    }

    if (nowMs >= expiresAtMs) {
      warnings.push(
        `${BUNFIG} temporary quarantine exclude "${name}" expired at ${expiresAt}. ` +
          `Its removal PR opens automatically; merge it, or run ` +
          `\`bun ${SCRIPT_PATH} --prune\` to do it by hand.`,
      );
    }
  }

  return {
    errors,
    warnings,
    excludeCount: excludes.size,
    firstPartyCount: firstPartyPackages.size,
    activeTemporaryCount: temporary.entries.length,
  };
};

export type PruneResult = {
  bunfig: string;
  pruned: string[];
};

/**
 * Drops the entries whose expiry has passed. Only the entry line goes: the
 * comments above the block explain why the remaining entries are there.
 */
export const pruneExpiredExcludes = ({
  bunfig,
  now = new Date(),
}: {
  bunfig: string;
  now?: Date;
}): PruneResult => {
  const blockStart = bunfig.indexOf("minimumReleaseAgeExcludes");
  if (blockStart === -1) return { bunfig, pruned: [] };
  const blockEnd = bunfig.indexOf("]", blockStart);
  const pruned: string[] = [];

  // Each line is judged by its own annotation. Matching by name instead would
  // delete every entry sharing that name, including one whose expiry was
  // renewed and has not passed.
  const block = bunfig
    .slice(blockStart, blockEnd)
    .split("\n")
    .filter((line) => {
      if (!line.includes(EXPIRY_MARKER)) return true;
      const parsed = parseTemporaryExcludeLine(line);
      // A malformed annotation is the guard's to report, not this to delete.
      if (parsed.kind !== "entry") return true;
      if (now.getTime() < Date.parse(parsed.expiresAt)) return true;
      pruned.push(parsed.name);
      return false;
    })
    .join("\n");

  if (pruned.length === 0) return { bunfig, pruned: [] };

  return { bunfig: bunfig.slice(0, blockStart) + block + bunfig.slice(blockEnd), pruned };
};

const prune = async (): Promise<void> => {
  const bunfigPath = join(ROOT, BUNFIG);
  const { bunfig, pruned } = pruneExpiredExcludes({
    bunfig: await Bun.file(bunfigPath).text(),
  });

  if (pruned.length === 0) {
    console.log(`${BUNFIG}: no expired quarantine excludes to remove.`);
    return;
  }

  await Bun.write(bunfigPath, bunfig);
  console.log(
    `${BUNFIG}: removed ${pruned.length} expired quarantine exclude(s): ${pruned.join(", ")}.`,
  );
};

const check = async (): Promise<void> => {
  const result = checkQuarantineExcludes({
    bunfig: await Bun.file(join(ROOT, BUNFIG)).text(),
    lockfile: await Bun.file(join(ROOT, LOCKFILE)).text(),
  });

  if (result.warnings.length > 0) {
    console.warn(`${result.warnings.join("\n")}\n`);
  }

  if (result.errors.length > 0) {
    console.error(result.errors.join("\n\n"));
    process.exit(1);
  }

  console.log(
    `${BUNFIG}: ${result.excludeCount} quarantine excludes cover all ` +
      `${result.firstPartyCount} registry-backed first-party packages; ` +
      `${result.activeTemporaryCount} temporary exclude(s) remain active.`,
  );
};

if (import.meta.main) {
  await (process.argv.includes("--prune") ? prune() : check());
}
