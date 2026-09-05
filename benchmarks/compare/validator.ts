/**
 * OOXML schema validation through the Open XML SDK projector the differential
 * harness already builds.
 *
 * A generated redline can be well-formed XML, round-trip correctly, and still
 * make Word report unreadable content: an unbalanced `w:moveFrom`/`w:moveTo`
 * pair or a revision id reused across parts is a schema fault, not a text one.
 * The validator is the only thing that catches that class, so it is a first
 * class assertion here rather than a nice-to-have.
 *
 * It auto-skips when .NET or the built projector is missing, the same way the
 * differential test does, so the benchmark stays runnable on a bare machine.
 */

import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const PROJECT_DIR = path.join(import.meta.dir, "../../packages/core/scripts/differential/dotnet");

const resolveProjector = (): string | null => {
  for (const configuration of ["Release", "Debug"]) {
    const dll = path.join(PROJECT_DIR, "bin", configuration, "net8.0/OpenXmlProjector.dll");
    if (existsSync(dll)) {
      return dll;
    }
  }
  return null;
};

type ValidationReport = { schemaVersion: number; errors: string[] };

/** Schema errors for one package, empty when it validates. */
export type PackageValidator = (buffer: ArrayBuffer) => string[];

/**
 * A validator, or `null` when this machine has no .NET toolchain. The caller
 * reports the skip; silently passing an unvalidated package would make the
 * suite claim a guarantee it did not check.
 */
export const resolvePackageValidator = (): PackageValidator | null => {
  const dotnet = spawnSync("dotnet", ["--version"], { encoding: "utf8" });
  if (dotnet.error || dotnet.status !== 0) {
    return null;
  }
  const projector = resolveProjector();
  if (projector === null) {
    return null;
  }

  return (buffer: ArrayBuffer): string[] => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-compare-validate-"));
    const file = path.join(directory, "candidate.docx");
    try {
      writeFileSync(file, new Uint8Array(buffer));
      const result = spawnSync("dotnet", [projector, "validate", file], {
        encoding: "utf8",
        maxBuffer: 32 * 1024 * 1024,
      });
      if (result.error || result.status !== 0) {
        return [`validator failed: ${result.stderr?.trim() ?? "unknown"}`];
      }
      const report = JSON.parse(result.stdout) as ValidationReport;
      return report.errors;
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  };
};

export const PACKAGE_VALIDATOR_HINT =
  "dotnet build packages/core/scripts/differential/dotnet -c Release";
