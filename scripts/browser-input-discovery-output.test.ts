import { expect, test } from "bun:test";
import { chmodSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Run the actual guard against a failed reporter process. A helper-only test
// would miss the boundary discarding stdout before the error is surfaced.
test("browser input discovery retains failures from every reporter output stream", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-discovery-output-"));
  try {
    const executable = path.join(directory, "bunx");
    writeFileSync(
      executable,
      [
        "#!/usr/bin/env bun",
        'process.stdout.write(process.env["FOLIO_DISCOVERY_TEST_STDOUT"] ?? "");',
        'process.stderr.write(process.env["FOLIO_DISCOVERY_TEST_STDERR"] ?? "");',
        "process.exit(1);",
      ].join("\n"),
    );
    chmodSync(executable, 0o755);
    const json = JSON.stringify({ suites: [], errors: [{ message: "Discovery module failed" }] });
    const exercised = new Set<string>();
    for (const stream of ["stdout", "stderr", "both"] as const) {
      const stdout = stream === "stderr" ? "" : json;
      const stderr = stream === "stdout" ? "" : "Discovery loader failed";
      const result = Bun.spawnSync(
        [process.execPath, path.join(import.meta.dir, "check-browser-input-coverage.ts")],
        {
          env: {
            ...process.env,
            PATH: `${directory}:${process.env["PATH"] ?? ""}`,
            FOLIO_DISCOVERY_TEST_STDOUT: stdout,
            FOLIO_DISCOVERY_TEST_STDERR: stderr,
          },
          stdout: "pipe",
          stderr: "pipe",
        },
      );
      expect(result.exitCode).not.toBe(0);
      const failure = result.stderr.toString();
      expect(failure).toContain("Playwright test discovery failed");
      if (stdout.length > 0) expect(failure).toContain(json);
      if (stderr.length > 0) expect(failure).toContain(stderr);
      exercised.add(stream);
    }
    expect([...exercised].sort()).toEqual(["both", "stderr", "stdout"]);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
