import { expect, test } from "bun:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

test("publish preparation preserves private exceptions to wildcard exports", async () => {
  const packageRoot = await mkdtemp(path.join(tmpdir(), "folio-prepare-publish-"));
  try {
    await mkdir(path.join(packageRoot, "dist"));
    await writeFile(path.join(packageRoot, "dist/index.js"), "export {};\n");
    await writeFile(path.join(packageRoot, "dist/index.d.ts"), "export {};\n");
    await writeFile(path.join(packageRoot, "dist/documentClone.js"), "export {};\n");
    await writeFile(path.join(packageRoot, "dist/documentClone.d.ts"), "export {};\n");
    await writeFile(
      path.join(packageRoot, "package.json"),
      `${JSON.stringify(
        {
          name: "@stll/fixture",
          version: "1.0.0",
          exports: {
            ".": "./src/index.ts",
            "./document-clone": "./src/documentClone.ts",
            "./private": null,
            "./*": "./src/*.ts",
          },
        },
        null,
        2,
      )}\n`,
    );

    const result = Bun.spawnSync(
      [process.execPath, path.join(import.meta.dir, "prepare-publish.ts"), packageRoot],
      { stderr: "pipe", stdout: "pipe" },
    );
    expect(result.exitCode).toBe(0);
    const manifest = await Bun.file(path.join(packageRoot, "package.json")).json();

    expect(manifest.exports["./private"]).toBeNull();
    expect(manifest.exports["./document-clone"]).toEqual({
      types: "./dist/documentClone.d.ts",
      import: "./dist/documentClone.js",
    });
    expect(manifest.exports["./*"]).toEqual({
      types: "./dist/*.d.ts",
      import: "./dist/*.js",
    });
  } finally {
    await rm(packageRoot, { force: true, recursive: true });
  }
});
