import { describe, expect, test } from "bun:test";
import path from "node:path";

import JSZip from "jszip";

import { EDIT_OPERATION_SCRIPTS } from "../fixtures/edit-operation-scripts";
import { readEditStructure } from "../fixtures/edit-structure";
import { runEditOperationScript } from "../fixtures/run-edit-operation-script";

describe("saved synthetic edit scripts", () => {
  test("every script saves a readable package with its structural carriers", async () => {
    const names = EDIT_OPERATION_SCRIPTS.map(({ id }) => id);
    expect(new Set(names).size).toBe(names.length);

    for (const script of EDIT_OPERATION_SCRIPTS) {
      const seed = await Bun.file(
        path.join(import.meta.dir, "../fixtures", script.seed),
      ).arrayBuffer();
      const saved = await runEditOperationScript(seed, script);
      const zip = await JSZip.loadAsync(saved);
      expect(zip.file("word/document.xml"), script.id).not.toBeNull();
      const structure = await readEditStructure(saved);
      expect(structure.blocks.length, script.id).toBeGreaterThan(0);
      if (script.action === "replyToComment") {
        expect(structure.comments.some(({ replies }) => replies.includes(script.text))).toBeTrue();
      }
      if (script.action === "deleteTableRow") {
        expect(structure.cells.length).toBeGreaterThan(0);
      }
      if (script.action === "replaceInBlock") {
        expect(structure.notes).toHaveLength(2);
      }
    }
  });
});
