import { expect, test } from "bun:test";
import { mkdir, mkdtemp, realpath, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";

test("publish preparation maps source, asset and private exports to dist", async () => {
  const packageRoot = await realpath(
    await mkdtemp(path.join(tmpdir(), "folio-prepare-publish-")),
  );
  try {
    await mkdir(path.join(packageRoot, "dist"));
    await writeFile(path.join(packageRoot, "dist/index.js"), "export {};\n");
    await writeFile(path.join(packageRoot, "dist/index.d.ts"), "export {};\n");
    await writeFile(
      path.join(packageRoot, "dist/documentClone.js"),
      "export {};\n",
    );
    await writeFile(
      path.join(packageRoot, "dist/documentClone.d.ts"),
      "export {};\n",
    );
    await mkdir(path.join(packageRoot, "dist/generated"));
    await writeFile(
      path.join(packageRoot, "dist/generated/artifact_bg.wasm"),
      "",
    );
    await writeFile(
      path.join(packageRoot, "package.json"),
      `${JSON.stringify(
        {
          name: "@stll/fixture",
          version: "1.0.0",
          type: "module",
          exports: {
            ".": "./src/index.ts",
            "./document-clone": "./src/documentClone.ts",
            "./artifact/wasm": "./src/generated/artifact_bg.wasm",
            "./private": null,
            "./*": "./src/*.ts",
          },
        },
        null,
        2,
      )}\n`,
    );

    const result = Bun.spawnSync(
      [
        process.execPath,
        path.join(import.meta.dir, "prepare-publish.ts"),
        packageRoot,
      ],
      { stderr: "pipe", stdout: "pipe" },
    );
    expect(result.exitCode).toBe(0);
    const manifest = await Bun.file(
      path.join(packageRoot, "package.json"),
    ).json();

    expect(manifest.exports["./private"]).toBeNull();
    expect(manifest.exports["./document-clone"]).toEqual({
      types: "./dist/documentClone.d.ts",
      import: "./dist/documentClone.js",
      default: "./dist/documentClone.js",
    });
    const require = createRequire(path.join(packageRoot, "consumer.cjs"));
    // Self-resolution applies the require condition just like the Vue CJS consumer.
    expect(require.resolve("@stll/fixture/document-clone")).toBe(
      path.join(packageRoot, "dist/documentClone.js"),
    );
    expect(require.resolve("@stll/fixture/index")).toBe(
      path.join(packageRoot, "dist/index.js"),
    );
    expect(manifest.exports["./artifact/wasm"]).toBe(
      "./dist/generated/artifact_bg.wasm",
    );
    expect(manifest.exports["./*"]).toEqual({
      types: "./dist/*.d.ts",
      import: "./dist/*.js",
      default: "./dist/*.js",
    });
  } finally {
    await rm(packageRoot, { force: true, recursive: true });
  }
});
