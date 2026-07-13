import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, test } from "bun:test";

const DIVERGENCES_SCRIPT = path.resolve(import.meta.dir, "../divergences.ts");

const VALID_REPORT = {
  generatedAt: "2026-01-01T00:00:00.000Z",
  reference: {
    id: "libreoffice",
    displayName: "LibreOffice Writer",
  },
  results: [],
  clusters: [],
};

type RunResult = {
  exitCode: number;
  stdout: string;
  stderr: string;
};

const runDivergences = async (report: unknown): Promise<RunResult> => {
  const dir = await mkdtemp(path.join(tmpdir(), "folio-divergences-test-"));
  const reportPath = path.join(dir, "report.json");
  await Bun.write(reportPath, JSON.stringify(report));

  try {
    const proc = Bun.spawn([process.execPath, DIVERGENCES_SCRIPT, "--report", reportPath], {
      stdout: "pipe",
      stderr: "pipe",
    });
    const [exitCode, stdout, stderr] = await Promise.all([
      proc.exited,
      new Response(proc.stdout).text(),
      new Response(proc.stderr).text(),
    ]);
    return { exitCode, stdout, stderr };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
};

describe("parity divergence report validation", () => {
  test("accepts explicit valid reference metadata", async () => {
    const result = await runDivergences(VALID_REPORT);

    expect(result.exitCode).toBe(0);
    expect(JSON.parse(result.stdout).reference).toEqual(VALID_REPORT.reference);
  });

  test("rejects reports without reference metadata", async () => {
    const { reference: _, ...report } = VALID_REPORT;
    const result = await runDivergences(report);

    expect(result.exitCode).toBe(2);
    expect(result.stderr).toContain("predates explicit reference-renderer metadata");
  });

  test("rejects unknown reference renderer ids", async () => {
    const result = await runDivergences({
      ...VALID_REPORT,
      reference: { id: "unknown", displayName: "Unknown" },
    });

    expect(result.exitCode).toBe(2);
    expect(result.stderr).toContain("predates explicit reference-renderer metadata");
  });
});
