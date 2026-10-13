// CI proof: the editor (without MCP declarations) compiles at the manifest's
// floor, while the immediately preceding API lacks workspace.save.
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import manifest from "../package.json";

const root = path.resolve(import.meta.dir, "..");
const floor = manifest.devDependencies["@types/vscode"];
assert.equal(manifest.engines.vscode, `^${floor}`);
const previous = "1.85.0";
assert.equal(floor, "1.86.0", "Update the adjacent-version proof when changing the floor");
const directory = await mkdtemp(path.join(root, ".engine-floor-"));

try {
  const response = await fetch(
    `https://raw.githubusercontent.com/microsoft/vscode/${previous}/src/vscode-dts/vscode.d.ts`,
  );
  assert.ok(response.ok, `Download ${previous} typings: ${response.status}`);
  const previousTypes = path.join(directory, "previous.d.ts");
  await writeFile(previousTypes, await response.text());
  for (const version of [floor, previous]) {
    const config = path.join(directory, `tsconfig-${version}.json`);
    await writeFile(
      config,
      JSON.stringify({
        extends: path.join(root, "tsconfig.json"),
        compilerOptions: {
          types: ["node", "bun-types"],
          paths: {
            vscode: [
              version === floor
                ? path.join(root, "node_modules/@types/vscode/index.d.ts")
                : previousTypes,
            ],
          },
        },
        include: [path.join(root, "src/editor.ts")],
      }),
    );
    const result = spawnSync(
      process.execPath,
      ["check", "--no-pretty", "--all", `--project=${config}`],
      {
        cwd: root,
        encoding: "utf8",
      },
    );
    assert.equal(result.error, undefined);
    const diagnostics = result.stdout + result.stderr;
    if (version === floor) {
      assert.equal(result.status, 0, diagnostics);
      console.log(`Editor without MCP compiles against @types/vscode ${version}.`);
      continue;
    }
    assert.equal(result.status, 1, diagnostics);
    assert.match(diagnostics, /Property 'save' does not exist on type 'typeof workspace'/u);
    assert.equal((diagnostics.match(/error TS\d+:/gu) ?? []).length, 1, diagnostics);
    console.log(`VS Code ${version} rejected only workspace.save; confirmed API floor ${floor}.`);
  }
} finally {
  await rm(directory, { recursive: true, force: true });
}
