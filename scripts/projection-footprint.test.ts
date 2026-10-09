import { expect, test } from "bun:test";
import binaryen from "binaryen";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const namedFixture = (source: string) => {
  const module = binaryen.parseText(source);
  try {
    if (!module.validate()) throw new TypeError("Footprint fixture must be valid WebAssembly.");
    binaryen.setDebugInfo(true);
    return module.emitBinary();
  } finally {
    binaryen.setDebugInfo(false);
    module.dispose();
  }
};

test("attribution normalizes Rust hashes and measures emitted function-body growth", () => {
  const directory = mkdtempSync(join(tmpdir(), "folio-footprint-attribution-"));
  try {
    const base = join(directory, "base.wasm");
    const head = join(directory, "head.wasm");
    const output = join(directory, "report.json");
    writeFileSync(
      base,
      namedFixture(`(module
        (func $probe::h1111111111111111 (export "probe") (param i32) (result i32)
          (local.get 0)))`),
    );
    writeFileSync(
      head,
      namedFixture(`(module
        (func $probe::h2222222222222222 (export "probe") (param i32) (result i32)
          (i32.add (local.get 0) (i32.const 1))))`),
    );
    const result = Bun.spawnSync(
      [process.execPath, join(import.meta.dir, "projection-footprint.ts"), base, head, output],
      { stdout: "pipe", stderr: "pipe" },
    );
    expect(result.exitCode, result.stderr.toString()).toBe(0);
    const report: unknown = JSON.parse(readFileSync(output, "utf8"));
    if (
      !isRecord(report) ||
      !isRecord(report["base"]) ||
      !isRecord(report["head"]) ||
      !Array.isArray(report["functions"])
    )
      throw new TypeError("Footprint analyzer must emit a structured report.");
    expect(report["functions"]).toEqual([{ name: "probe", baseBytes: 4, headBytes: 7, delta: 3 }]);
    expect(report["base"]["codeBytes"]).toBe(4);
    expect(report["head"]["codeBytes"]).toBe(7);
    const baseBytes = report["base"]["codeBytes"];
    const headBytes = report["head"]["codeBytes"];
    if (typeof baseBytes !== "number" || typeof headBytes !== "number")
      throw new TypeError("Footprint code sizes must be numeric.");
    expect(headBytes - baseBytes).toBe(3);
  } finally {
    rmSync(directory, { recursive: true });
  }
});
