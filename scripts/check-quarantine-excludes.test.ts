import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import packageJson from "../package.json";
import { checkQuarantineExcludes, pruneExpiredExcludes } from "./check-quarantine-excludes";

const SCRIPT_PATH = "scripts/check-quarantine-excludes.ts";

// Shaped like folio's real lockfile: a first-party package pulled from the
// registry, next to a workspace member that never reaches it.
const lockfile = `
"packages": {
  "@stll/conditions": ["@stll/conditions@0.1.0", "", { "dependencies": { "valibot": "1.4.1" } }, "sha512-test"],
  "@stll/folio-core": ["@stll/folio-core@workspace:packages/core"],
}
`;

const createBunfig = (temporaryLine: string) => `
[install]
minimumReleaseAge = 432_000
minimumReleaseAgeExcludes = [
  "@stll/conditions",
  ${temporaryLine}
]
`;

describe("quarantine exclude guard", () => {
  test("accepts a temporary exclusion before its exact expiry", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      lockfile,
      now: new Date("2026-08-08T14:16:00.209Z"),
    });

    expect(result.errors).toEqual([]);
    expect(result.activeTemporaryCount).toBe(1);
  });

  // Nothing to say before it expires: the removal arrives on its own.
  test("stays silent until the expiry", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      lockfile,
      now: new Date("2026-08-08T14:16:00.209Z"),
    });

    expect(result.errors).toEqual([]);
    expect(result.warnings).toEqual([]);
  });

  // An expiry is a wall-clock instant. Failing at it turns every open branch
  // red at once for a change nobody made, so the instant only warns.
  test("warns rather than fails at the exact expiry", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      lockfile,
      now: new Date("2026-08-08T14:16:00.210Z"),
    });

    expect(result.errors).toEqual([]);
    expect(result.warnings.at(0)).toContain('temporary quarantine exclude "dompurify" expired at');
  });

  test("still only warns just inside the notice window", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      lockfile,
      now: new Date("2026-08-09T14:16:00.209Z"),
    });

    expect(result.errors).toEqual([]);
    expect(result.warnings).toHaveLength(1);
  });

  test("fails once an expired exclusion is ignored past the notice window", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      lockfile,
      now: new Date("2026-08-09T14:16:00.210Z"),
    });

    expect(result.warnings).toEqual([]);
    expect(result.errors.at(0)).toContain(
      'temporary quarantine exclude "dompurify" expired at 2026-08-08T14:16:00.210Z and is still here',
    );
  });

  test("rejects malformed expiry annotations", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08'),
      lockfile,
    });

    expect(result.errors).toContain(
      'bunfig.toml temporary quarantine exclude "dompurify" has an invalid UTC expiry: 2026-08-08',
    );
  });

  // folio's own published packages resolve through `workspace:` and never hit
  // the registry, so the release-age gate cannot apply to them.
  test("exempts workspace members and requires registry-resolved packages", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify",'),
      lockfile,
    });

    expect(result.errors).toEqual([]);
    expect(result.firstPartyCount).toBe(1);
  });

  // A fixed-width lookahead read past the entry, so a workspace member on the
  // next line marked its registry-resolved neighbour as a workspace member
  // too, and the coverage check silently stopped requiring it.
  test("counts a registry package followed by a workspace member", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify",'),
      lockfile: `
"packages": {
  "@stll/conditions": ["@stll/conditions@0.1.0", "", {}, "sha512-test"],
  "@stll/docx-utils": ["@stll/docx-utils@0.1.0", "", { "dependencies": { "jszip": "3.10.1" } }, "sha512-test"],
  "@stll/folio-core": ["@stll/folio-core@workspace:packages/core"],
}
`,
    });

    expect(result.firstPartyCount).toBe(2);
    expect(result.errors.at(0)).toContain('"@stll/docx-utils",');
  });

  // A napi package's platform binaries are optionalDependencies with their own
  // lockfile entries. Excluding only the parent installs it with every binding
  // silently skipped, so each child has to be required in its own right.
  test("requires the platform subpackages of a napi package, not just the parent", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"@stll/stdnum",'),
      lockfile: `
"packages": {
  "@stll/conditions": ["@stll/conditions@0.1.0", "", {}, "sha512-test"],
  "@stll/stdnum": ["@stll/stdnum@2.3.1", "", { "optionalDependencies": { "@stll/stdnum-darwin-arm64": "2.3.1", "@stll/stdnum-linux-x64-gnu": "2.3.1" } }, "sha512-test"],
  "@stll/stdnum-darwin-arm64": ["@stll/stdnum-darwin-arm64@2.3.1", "", { "os": "darwin", "cpu": "arm64" }, "sha512-test"],
  "@stll/stdnum-linux-x64-gnu": ["@stll/stdnum-linux-x64-gnu@2.3.1", "", { "os": "linux", "cpu": "x64" }, "sha512-test"],
}
`,
    });

    expect(result.errors.at(0)).toContain('"@stll/stdnum-darwin-arm64",');
    expect(result.errors.at(0)).toContain('"@stll/stdnum-linux-x64-gnu",');
  });

  // Bun keys nested resolutions by path; those are not package names.
  test("ignores nested resolution keys", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify",'),
      lockfile: `
"packages": {
  "@stll/conditions": ["@stll/conditions@0.1.0", "", {}, "sha512-test"],
  "@stll/template-conditions/better-result": ["better-result@2.9.2", "", {}, "sha512-test"],
}
`,
    });

    expect(result.firstPartyCount).toBe(1);
    expect(result.errors).toEqual([]);
  });

  test("retains the first-party package coverage guard", () => {
    const result = checkQuarantineExcludes({
      bunfig: createBunfig('"dompurify",'),
      lockfile: `${lockfile}\n"@stll/missing": ["@stll/missing@1.0.0"]`,
    });

    expect(result.errors.at(0)).toContain('"@stll/missing",');
  });
});

describe("quarantine exclude prune", () => {
  test("prunes only the expired entries", () => {
    const bunfig = `
[install]
minimumReleaseAge = 432_000
minimumReleaseAgeExcludes = [
  "@stll/conditions",
  # Security patches for dependency-audit findings.
  "dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z
  "nanoid", # quarantine-expires: 2026-08-09T10:39:22.487Z
  "@stll/oxlint-config",
]
`;
    const result = pruneExpiredExcludes({
      bunfig,
      now: new Date("2026-08-08T15:00:00.000Z"),
    });

    expect(result.pruned).toEqual(["dompurify"]);
    expect(result.bunfig).not.toContain("dompurify");
    expect(result.bunfig).toContain('"nanoid", # quarantine-expires:');
    expect(result.bunfig).toContain('"@stll/conditions"');
    expect(result.bunfig).toContain('"@stll/oxlint-config"');
    // The pruned file is what the guard should then accept.
    expect(
      checkQuarantineExcludes({
        bunfig: result.bunfig,
        lockfile,
        now: new Date("2026-08-08T15:00:00.000Z"),
      }).errors,
    ).toEqual([]);
  });

  // The prune deletes at the instant the guard starts warning, so the removal
  // PR is already open for the whole notice window.
  test("prunes at the exact expiry", () => {
    const result = pruneExpiredExcludes({
      bunfig: createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z'),
      now: new Date("2026-08-08T14:16:00.210Z"),
    });

    expect(result.pruned).toEqual(["dompurify"]);
  });

  // Matching by name would take the renewed entry with the expired one, which
  // silently drops a package back out of the quarantine before its time.
  test("keeps a renewed entry when an expired one shares its name", () => {
    const bunfig = `
[install]
minimumReleaseAge = 432_000
minimumReleaseAgeExcludes = [
  "@stll/conditions",
  "nanoid", # quarantine-expires: 2026-08-09T10:39:22.487Z
  "nanoid", # quarantine-expires: 2026-09-01T10:39:22.487Z
]
`;
    const result = pruneExpiredExcludes({
      bunfig,
      now: new Date("2026-08-10T00:00:00.000Z"),
    });

    expect(result.pruned).toEqual(["nanoid"]);
    expect(result.bunfig).toContain('"nanoid", # quarantine-expires: 2026-09-01T10:39:22.487Z');
    expect(result.bunfig).not.toContain("2026-08-09T10:39:22.487Z");
  });

  test("leaves a malformed annotation for the guard to report", () => {
    const bunfig = createBunfig('"dompurify", # quarantine-expires: nope');
    const result = pruneExpiredExcludes({ bunfig, now: new Date("2026-09-01T00:00:00.000Z") });

    expect(result.pruned).toEqual([]);
    expect(result.bunfig).toBe(bunfig);
  });

  test("leaves a file with nothing expired untouched", () => {
    const bunfig = createBunfig('"dompurify", # quarantine-expires: 2026-08-08T14:16:00.210Z');
    const result = pruneExpiredExcludes({ bunfig, now: new Date("2026-08-06T00:00:00.000Z") });

    expect(result.pruned).toEqual([]);
    expect(result.bunfig).toBe(bunfig);
  });
});

// Renaming the script silently unwires CI and the prune job, and neither
// failure shows up as a test failure anywhere else.
describe("quarantine exclude wiring", () => {
  test("the package script runs this guard", () => {
    expect(packageJson.scripts["check:quarantine-excludes"]).toBe(`bun ${SCRIPT_PATH}`);
  });

  test("CI runs the package script", async () => {
    const ci = await Bun.file(join(import.meta.dirname, "../.github/workflows/ci.yml")).text();

    expect(ci).toContain("bun run check:quarantine-excludes");
  });

  test("the prune workflow runs this guard in both modes", async () => {
    const workflow = await Bun.file(
      join(import.meta.dirname, "../.github/workflows/quarantine-prune.yml"),
    ).text();

    expect(workflow).toContain(`bun ${SCRIPT_PATH} --prune`);
    expect(workflow).toContain(`bun ${SCRIPT_PATH}\n`);
  });
});
