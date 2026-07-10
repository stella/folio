/**
 * Differential testing harness — folio parser vs. external OOXML references.
 *
 * Parses a single DOCX both with folio (in-process) and a reference parser
 * (python-docx or Open XML SDK subprocess), projects both into a normalised
 * structural shape, and prints any divergences. Exits 0 on equivalence, 1 on
 * any divergence, 2 on infrastructure failure (reference missing, fixture
 * missing, parse error).
 *
 * Usage:
 *   bun packages/core/scripts/differential/diff.ts <docx-path> [reference]
 *
 * References: python-docx (default), open-xml-sdk
 *
 * See README.md in this directory for setup and rationale.
 */

import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { parseDocx } from "../../src/docx/parser";
import {
  diffProjections,
  projectFolioDocument,
  type Divergence,
  type StructuralProjection,
} from "./projection";

const HERE = import.meta.dirname;
const PYTHON_SCRIPT = path.join(HERE, "python_docx_project.py");
const OPENXML_PROJECT = path.join(HERE, "dotnet");
const OPENXML_DLL = path.join(OPENXML_PROJECT, "bin/Release/net8.0/OpenXmlProjector.dll");

const EXIT_OK = 0;
const EXIT_DIVERGED = 1;
const EXIT_INFRA = 2;

export const DIFFERENTIAL_REFERENCES = ["python-docx", "open-xml-sdk"] as const;

export type DifferentialReference = (typeof DIFFERENTIAL_REFERENCES)[number];

export type DifferentialResult =
  | { ok: true; folio: StructuralProjection; reference: unknown }
  | {
      ok: false;
      reason: "diverged";
      folio: StructuralProjection;
      reference: unknown;
      divergences: Divergence[];
    }
  | { ok: false; reason: "infra"; message: string };

const isDifferentialReference = (value: string): value is DifferentialReference =>
  (DIFFERENTIAL_REFERENCES as readonly string[]).includes(value);

const resolveOpenXmlProjector = (): string | null => {
  if (existsSync(OPENXML_DLL)) {
    return OPENXML_DLL;
  }
  const debugDll = path.join(OPENXML_PROJECT, "bin/Debug/net8.0/OpenXmlProjector.dll");
  if (existsSync(debugDll)) {
    return debugDll;
  }
  return null;
};

const runReferenceProjection = (
  reference: DifferentialReference,
  resolved: string,
  options: { pythonBin?: string },
): { ok: true; reference: unknown } | { ok: false; message: string } => {
  if (reference === "python-docx") {
    const pythonBin = options.pythonBin ?? "python3";
    const pythonResult = spawnSync(pythonBin, [PYTHON_SCRIPT, resolved], {
      encoding: "utf-8",
      timeout: 30_000,
    });
    if (pythonResult.error || pythonResult.status !== 0) {
      const stderr = pythonResult.stderr ? pythonResult.stderr.trim() : "";
      const spawnError = pythonResult.error?.message ?? "";
      return {
        ok: false,
        message: [
          `python-docx projection failed (exit ${pythonResult.status ?? "?"}).`,
          spawnError ? `error: ${spawnError}` : "",
          stderr ? `stderr: ${stderr}` : "",
          "Hint: pip install python-docx",
        ]
          .filter(Boolean)
          .join("\n"),
      };
    }

    try {
      return { ok: true, reference: JSON.parse(pythonResult.stdout) };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        ok: false,
        message: `Failed to parse python-docx projector JSON: ${message}`,
      };
    }
  }

  const projector = resolveOpenXmlProjector();
  if (!projector) {
    return {
      ok: false,
      message: [
        "Open XML SDK projector is not built.",
        `Expected: ${OPENXML_DLL}`,
        "Hint: dotnet build packages/core/scripts/differential/dotnet -c Release",
      ].join("\n"),
    };
  }

  const dotnetResult = spawnSync("dotnet", [projector, resolved], {
    encoding: "utf-8",
    timeout: 30_000,
  });
  if (dotnetResult.error || dotnetResult.status !== 0) {
    const stderr = dotnetResult.stderr ? dotnetResult.stderr.trim() : "";
    const spawnError = dotnetResult.error?.message ?? "";
    return {
      ok: false,
      message: [
        `Open XML SDK projection failed (exit ${dotnetResult.status ?? "?"}).`,
        spawnError ? `error: ${spawnError}` : "",
        stderr ? `stderr: ${stderr}` : "",
        "Hint: dotnet build packages/core/scripts/differential/dotnet -c Release",
      ]
        .filter(Boolean)
        .join("\n"),
    };
  }

  try {
    return { ok: true, reference: JSON.parse(dotnetResult.stdout) };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      ok: false,
      message: `Failed to parse Open XML SDK projector JSON: ${message}`,
    };
  }
};

/**
 * Run the differential comparison without exiting the process. Returns
 * a structured result so callers (the smoke test, future corpus runners)
 * can decide how to fail.
 */
export async function runDifferential(
  docxPath: string,
  options: {
    pythonBin?: string;
    reference?: DifferentialReference;
  } = {},
): Promise<DifferentialResult> {
  const reference = options.reference ?? "python-docx";
  const resolved = path.resolve(docxPath);
  if (!existsSync(resolved)) {
    return {
      ok: false,
      reason: "infra",
      message: `Fixture not found: ${resolved}`,
    };
  }

  const buffer = readFileSync(resolved);

  let folioProjection: StructuralProjection;
  try {
    const doc = await parseDocx(buffer);
    folioProjection = projectFolioDocument(doc);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      ok: false,
      reason: "infra",
      message: `folio parseDocx failed: ${message}`,
    };
  }

  const referenceResult = runReferenceProjection(reference, resolved, options);
  if (!referenceResult.ok) {
    return {
      ok: false,
      reason: "infra",
      message: referenceResult.message,
    };
  }

  const divergences = diffProjections(folioProjection, referenceResult.reference);
  if (divergences.length === 0) {
    return { ok: true, folio: folioProjection, reference: referenceResult.reference };
  }
  return {
    ok: false,
    reason: "diverged",
    folio: folioProjection,
    reference: referenceResult.reference,
    divergences,
  };
}

export const isPythonDocxAvailable = (): boolean => {
  const result = spawnSync("python3", ["-c", "import docx"], {
    encoding: "utf-8",
    timeout: 5000,
  });
  return !result.error && result.status === 0;
};

export const isOpenXmlSdkAvailable = (): boolean => {
  const dotnet = spawnSync("dotnet", ["--version"], {
    encoding: "utf-8",
    timeout: 5000,
  });
  if (dotnet.error || dotnet.status !== 0) {
    return false;
  }
  return resolveOpenXmlProjector() !== null;
};

const formatDivergences = (divergences: readonly Divergence[]): string =>
  divergences
    .map(
      (d) =>
        `  ${d.path}: folio=${JSON.stringify(d.folio)} reference=${JSON.stringify(d.reference)}`,
    )
    .join("\n");

const isMain = import.meta.path === Bun.main;
if (isMain) {
  const docxPath = process.argv[2];
  const referenceArg = process.argv[3] ?? "python-docx";
  if (!docxPath) {
    console.error("usage: bun diff.ts <docx-path> [python-docx|open-xml-sdk]");
    process.exit(EXIT_INFRA);
  }
  if (!isDifferentialReference(referenceArg)) {
    console.error(`unknown reference: ${referenceArg}`);
    console.error("usage: bun diff.ts <docx-path> [python-docx|open-xml-sdk]");
    process.exit(EXIT_INFRA);
  }
  const result = await runDifferential(docxPath, { reference: referenceArg });
  if (result.ok) {
    console.log(`OK ${docxPath} (${referenceArg})`);
    console.log(JSON.stringify(result.folio, null, 2));
    process.exit(EXIT_OK);
  }
  if (result.reason === "infra") {
    console.error(result.message);
    process.exit(EXIT_INFRA);
  }
  console.error(`DIVERGED ${docxPath} (${referenceArg})`);
  console.error(formatDivergences(result.divergences));
  console.error("\nfolio projection:");
  console.error(JSON.stringify(result.folio, null, 2));
  console.error("\nreference projection:");
  console.error(JSON.stringify(result.reference, null, 2));
  process.exit(EXIT_DIVERGED);
}
