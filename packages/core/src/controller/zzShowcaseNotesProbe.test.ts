import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { parseShapeDocument } from "../__tests__/editorHarness";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createCanonicalSession } from "./canonicalSession";

// Scratch probe, not for merge: how the showcase's notes reach the canonical save.
test("probe showcase notes through a canonical save", async () => {
  const bytes = new Uint8Array(
    readFileSync(
      path.resolve(
        import.meta.dir,
        "../../../playground/public/folio-showcase.docx",
      ),
    ),
  );
  const opened = await parseShapeDocument(
    (await prepareCanonicalDocxInput(bytes)).unwrap(),
  );
  const notes = (opened.package.footnotes ?? []).map((note) => ({
    id: note.id,
    type: (note as { noteType?: unknown }).noteType ?? null,
    keys: Object.keys(note),
    paraIds: note.content.map((block) =>
      block.type === "paragraph" ? (block.paraId ?? null) : block.type,
    ),
  }));
  console.error("PROBE notes " + JSON.stringify(notes));
  const session = createCanonicalSession(opened).unwrap();
  const snapshot = session.captureSaveSnapshot();
  console.error(
    "PROBE held " + JSON.stringify([...snapshot.heldUnreferencedNotes]),
  );
  console.error(
    "PROBE snapshot notes " +
      JSON.stringify(
        (snapshot.document.package.footnotes ?? []).map((note) => ({
          id: note.id,
          paraIds: note.content.map((block) =>
            block.type === "paragraph" ? (block.paraId ?? null) : block.type,
          ),
        })),
      ),
  );
  for (const mode of ["selective", "full"] as const) {
    try {
      const saved = await serializeCanonicalSave({
        snapshot: {
          ...snapshot,
          structure: mode === "full" ? "changed" : snapshot.structure,
        },
        featureFlags: { selectiveSave: mode === "selective" },
      });
      console.error(`PROBE ${mode} ok ${saved.buffer.byteLength}`);
    } catch (error) {
      console.error(`PROBE ${mode} threw ${String((error as Error).message)}`);
    }
  }
  expect(true).toBe(true);
});
