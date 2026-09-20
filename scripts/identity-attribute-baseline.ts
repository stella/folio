#!/usr/bin/env bun

/**
 * Shrink-only guard for prefix-resolved reads of an identity-bearing attribute.
 *
 * `folio-identity-attributes/no-prefix-resolved-identity-read` says where such
 * a read belongs, but the repository predates the derived identity set and
 * still carries the reads the rule rejects. Turning it on repo-wide would only
 * produce a file of suppressions, so `bun run lint` leaves it off and this
 * guard holds the count per file against a committed baseline that may only go
 * down: new code cannot add one, and a file that loses its last prefix-resolved
 * read leaves the baseline for good.
 *
 * Modes:
 *   bun scripts/identity-attribute-baseline.ts                  report per file
 *   bun scripts/identity-attribute-baseline.ts --write-baseline regenerate baseline
 *   bun scripts/identity-attribute-baseline.ts --check          CI gate
 */

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BASELINE_PATH = path.join(REPO_ROOT, "scripts/identity-attribute-baseline.json");
const BASELINE_RELATIVE_PATH = "scripts/identity-attribute-baseline.json";
const LINT_CONFIG = "oxlint.identity-attributes.config.ts";
const RULE_CODE = "folio-identity-attributes(no-prefix-resolved-identity-read)";
const WRITE_COMMAND = "bun scripts/identity-attribute-baseline.ts --write-baseline";

type Diagnostic = { code?: unknown; filename?: unknown };

/**
 * oxlint's JSON report goes to a file, not a pipe: a piped stdout truncates
 * around 768 KiB and the report is bigger than that when other rules fire.
 */
const lintReport = (): string => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-identity-attributes-"));
  const reportPath = path.join(directory, "report.json");
  try {
    const result = Bun.spawnSync(
      ["bun", "--bun", "oxlint", "-c", LINT_CONFIG, "--format", "json", "packages"],
      { cwd: REPO_ROOT, stdout: Bun.file(reportPath), stderr: "pipe" },
    );
    const report = readFileSync(reportPath, "utf8");
    if (report.trim() === "") {
      throw new Error(`oxlint produced no report: ${result.stderr.toString()}`);
    }
    return report;
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

const diagnosticsOf = (parsed: unknown): Diagnostic[] => {
  if (typeof parsed !== "object" || parsed === null || !("diagnostics" in parsed)) {
    return [];
  }
  const { diagnostics } = parsed;
  return Array.isArray(diagnostics) ? (diagnostics as Diagnostic[]) : [];
};

const countsByFile = (report: string): Record<string, number> => {
  const diagnostics = diagnosticsOf(JSON.parse(report));
  const counts = new Map<string, number>();
  for (const { code, filename } of diagnostics) {
    if (typeof code !== "string" || !code.includes(RULE_CODE)) {
      continue;
    }
    if (typeof filename !== "string") {
      continue;
    }
    const file = path.relative(REPO_ROOT, path.resolve(REPO_ROOT, filename)).replaceAll("\\", "/");
    counts.set(file, (counts.get(file) ?? 0) + 1);
  }
  return Object.fromEntries([...counts.entries()].toSorted(([a], [b]) => a.localeCompare(b)));
};

const total = (counts: Record<string, number>): number =>
  Object.values(counts).reduce((sum, count) => sum + count, 0);

const modeFromArgv = (): "check" | "report" | "write" => {
  if (process.argv.includes("--write-baseline")) {
    return "write";
  }
  return process.argv.includes("--check") ? "check" : "report";
};

const mode = modeFromArgv();

const current = countsByFile(lintReport());

if (mode === "write") {
  writeFileSync(BASELINE_PATH, `${JSON.stringify(current, null, 2)}\n`);
  console.log(
    `Wrote ${total(current)} prefix-resolved identity reads across ${Object.keys(current).length} files to ${BASELINE_RELATIVE_PATH}`,
  );
  process.exit(0);
}

if (mode === "report") {
  for (const [file, count] of Object.entries(current)) {
    console.log(`${count}\t${file}`);
  }
  console.log(
    `\n${total(current)} prefix-resolved identity reads across ${Object.keys(current).length} files`,
  );
  process.exit(0);
}

const baseline = JSON.parse(readFileSync(BASELINE_PATH, "utf8")) as Record<string, number>;
const failures: string[] = [];

for (const [file, count] of Object.entries(current)) {
  const allowed = baseline[file];
  if (allowed === undefined) {
    failures.push(`${file}: ${count} new prefix-resolved identity read(s) in a clean file`);
    continue;
  }
  if (count > allowed) {
    failures.push(`${file}: prefix-resolved identity reads rose from ${allowed} to ${count}`);
  }
}

if (failures.length > 0) {
  console.error("Prefix-resolved identity reads increased:");
  for (const failure of failures) {
    console.error(`  ${failure}`);
  }
  console.error(
    "\nResolve the attribute with `getAttributeByNamespaceUri`, or, if you removed reads " +
      `elsewhere, regenerate the baseline with \`${WRITE_COMMAND}\`.`,
  );
  process.exit(1);
}

const removed = total(baseline) - total(current);
if (removed > 0) {
  console.error(
    `Prefix-resolved identity reads dropped by ${removed}. ` +
      `Lock the win in with \`${WRITE_COMMAND}\`.`,
  );
  process.exit(1);
}

console.log(
  `Identity-attribute guard: ${total(current)} prefix-resolved reads across ${Object.keys(current).length} files, none added.`,
);
