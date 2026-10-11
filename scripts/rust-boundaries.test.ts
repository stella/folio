import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import ts from "typescript";
import { panic } from "better-result";

const REPOSITORY_ROOT = path.join(import.meta.dir, "..");

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

test("local and required kernel lint use one command for every tested crate", async () => {
  const workflow: unknown = Bun.YAML.parse(
    await readFile(path.join(REPOSITORY_ROOT, ".github/workflows/ci.yml"), "utf8"),
  );
  const manifest: unknown = JSON.parse(
    await readFile(path.join(REPOSITORY_ROOT, "package.json"), "utf8"),
  );
  if (!isRecord(workflow) || !isRecord(workflow["jobs"]) || !isRecord(manifest)) {
    return panic("Expected the workflow and package manifest");
  }
  const job = workflow["jobs"]["docx-kernel"];
  const scripts = manifest["scripts"];
  if (!isRecord(job) || !Array.isArray(job["steps"]) || !isRecord(scripts)) {
    return panic("Expected kernel steps and package scripts");
  }
  const command = scripts["kernel:check"];
  if (typeof command !== "string") return panic("Missing local kernel check command");
  const steps = job["steps"].filter(isRecord);
  expect(steps.filter((step) => step["run"] === "bun run kernel:check")).toHaveLength(1);
  const ciCommands = steps.flatMap((step) =>
    typeof step["run"] === "string" ? step["run"].split("\n") : [],
  );
  expect(ciCommands.some((line) => /cargo (?:clippy|fmt)|bun run rust:hawk/u.test(line))).toBe(
    false,
  );
  const crates = ciCommands.flatMap((line) => {
    const crate = line.match(/^cargo test -p ([\w-]+)$/u)?.at(1);
    return crate === undefined ? [] : [crate];
  });
  expect(crates.length).toBeGreaterThan(0);
  const expected = ["cargo fmt --all -- --check"];
  for (const crate of crates) {
    expected.push(
      `cargo clippy -p ${crate} --all-targets -- -D warnings`,
      `cargo clippy -p ${crate} --lib --features wasm --target wasm32-unknown-unknown -- -D warnings`,
    );
  }
  expected.push("bun run rust:hawk");
  expect(command.split(" && ")).toEqual(expected);
});

/**
 * Every artifact compiled from Rust, and the module that owns the boundary to
 * it. The lint rule keeps a second entry point from appearing; this keeps the
 * boundary module itself thin, which is the other half of the same contract:
 * what the crate does has one implementation, and the TypeScript beside it
 * initializes the artifact and translates its errors.
 */
const RUST_BOUNDARIES = [
  {
    label: "DOCX projection",
    module: "packages/docx-core/src/projection.ts",
    imports: ["better-result", "./generated/docx_kernel.js"],
    // A second OOXML or archive reader here would be a second answer to the
    // question the kernel exists to answer.
    forbidden: ["jszip", "fast-xml-parser", "DOMParser"],
  },
  {
    label: "text shaping",
    module: "packages/core/src/shaping/shaper.ts",
    imports: ["better-result", "../generated/text_shaper.js"],
    // Shaping in TypeScript beside the boundary is how the measurer and the
    // painter would come to disagree about which glyphs a word is.
    forbidden: ["harfbuzz", "opentype", "fontkit"],
  },
] as const;

describe("Rust boundary modules stay thin", () => {
  for (const { label, module, imports, forbidden } of RUST_BOUNDARIES) {
    test(`the ${label} boundary imports only its artifact`, async () => {
      const source = await readFile(path.join(REPOSITORY_ROOT, module), "utf8");
      const importedModules = ts
        .preProcessFile(source)
        .importedFiles.map(({ fileName }) => fileName);

      expect(importedModules).toEqual([...imports]);
      for (const name of forbidden) {
        expect(source).not.toContain(name);
      }
    });
  }
});
