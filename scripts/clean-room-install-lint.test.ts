import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

setDefaultTimeout(30_000);

const lint = (files: string[]) => {
  const result = Bun.spawnSync(
    ["bun", "--bun", "oxlint", "-c", "oxlint.fixtures.config.ts", "--no-ignore", ...files],
    { cwd: path.resolve(import.meta.dir, "..") },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return {
    exitCode: result.exitCode,
    output,
    violations: output.split("folio-clean-room-install(install-owner)").length - 1,
  };
};

test("the install ownership guard rejects shell and subprocess install spellings", () => {
  const result = lint(["test/__fixtures__/clean-room-install.invalid.ts"]);
  expect(result.violations, result.output).toBe(9);
  expect(result.exitCode).not.toBe(0);
});

test("consumer installs use the owner and other Bun commands remain available", () => {
  const result = lint([
    "test/__fixtures__/clean-room-install.valid.ts",
    "scripts/clean-room-install.ts",
  ]);
  expect(result.violations, result.output).toBe(0);
  expect(result.exitCode, result.output).toBe(0);
});

test("every script delegates installs to the release-policy owner", () => {
  const files = [...new Bun.Glob("**/*.{ts,js,mts,mjs,cts,cjs}").scanSync(import.meta.dir)].map(
    (file) => path.join("scripts", file),
  );
  expect(files.length).toBeGreaterThan(0);
  const result = lint(files);
  expect(result.violations, result.output).toBe(0);
  expect(result.exitCode, result.output).toBe(0);
});

test("shell installs stay in the repository-mounted canonical build", () => {
  const installers = [...new Bun.Glob("**/*.sh").scanSync(import.meta.dir)].filter((file) =>
    /\bbun\s+(?:install|add)\b/u.test(
      readFileSync(path.join(import.meta.dir, file), "utf8")
        .split("\n")
        .filter((line) => !line.trimStart().startsWith("#"))
        .join("\n"),
    ),
  );
  // This install uses the full repository mount, including its bunfig.toml.
  // New shell installers must delegate to the TypeScript install owner.
  expect(installers).toEqual(["regenerate-wasm-canonically.sh"]);
  const canonical = readFileSync(
    path.join(import.meta.dir, "regenerate-wasm-canonically.sh"),
    "utf8",
  );
  expect(canonical).toContain('--volume "${root}:/workspace"');
  const dockerfile = readFileSync(path.join(import.meta.dir, "canonical-wasm.Dockerfile"), "utf8");
  expect(dockerfile).toContain("WORKDIR /workspace");
});
