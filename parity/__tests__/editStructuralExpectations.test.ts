import { Glob } from "bun";
import { describe, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import path from "node:path";

import { EDIT_OPERATION_SCRIPTS } from "../fixtures/edit-operation-scripts";
import { readEditStructure } from "../fixtures/edit-structure";
import { runEditOperationScript } from "../fixtures/run-edit-operation-script";

const FIXTURES_DIR = path.join(import.meta.dir, "../fixtures");
const expectationFiles = [...new Glob("*.expectation.json").scanSync({ cwd: FIXTURES_DIR })].sort();
const pinnedScripts = EDIT_OPERATION_SCRIPTS.filter(({ expectation }) => expectation === "pinned");

describe("saved edit structural expectations", () => {
  test("every reviewed expectation matches the saved operation result", async () => {
    expect(pinnedScripts.length).toBeGreaterThan(0);
    expect(expectationFiles).toEqual(
      pinnedScripts.map(({ id }) => `${id}.expectation.json`).toSorted(),
    );
    const exercised = new Set<string>();
    for (const file of expectationFiles) {
      const fixture = await Bun.file(path.join(FIXTURES_DIR, file)).json();
      const script = EDIT_OPERATION_SCRIPTS.find(({ id }) => id === fixture.script);
      if (!script) throw new Error(`${file}: unknown script`);
      expect(file).toBe(`${script.id}.expectation.json`);
      exercised.add(script.id);
      const seed = await Bun.file(path.join(FIXTURES_DIR, script.seed)).arrayBuffer();
      expect(createHash("sha256").update(new Uint8Array(seed)).digest("hex")).toBe(
        fixture.seedSha256,
      );
      const saved = await runEditOperationScript(seed, script);
      expect(await readEditStructure(saved)).toEqual(fixture.structure);
    }
    expect([...exercised].toSorted()).toEqual(pinnedScripts.map(({ id }) => id).toSorted());
  });
});
