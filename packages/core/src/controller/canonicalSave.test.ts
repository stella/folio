import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "@stll/docx-core/ops";
import { createCanonicalSession } from "./canonicalSession";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { schema } from "../prosemirror/schema";
import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createSimpleDocument } from "../docx/serializer/documentSerializer";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import type { SaveDiagnostic } from "../docx/saveDiagnostics";

const IDS = ["11111111", "22222222", "33333333"] as const;
const SOURCE_BLOCKS = IDS.map(
  (id, index) => `<w:p w14:paraId='${id}'><w:r>\n<w:t>text${index}</w:t>\t</w:r></w:p>`,
);
const openSource = async () => {
  const zip = await JSZip.loadAsync(await createDocx(createSimpleDocument([{ text: "seed" }])));
  zip.file(
    "word/document.xml",
    `<?xml version='1.0'?>\n<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'><w:body>\n${SOURCE_BLOCKS.join("\n<!-- gap -->\n")}\n</w:body></w:document>`,
  );
  return parseDocx(await zip.generateAsync({ type: "arraybuffer" }), { preloadFonts: false });
};

test("generated canonical histories save the model and preserve every block outside cumulative touched ids", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(
        fc.record({
          id: fc.constantFrom(IDS[0], IDS[1]),
          text: fc.constantFrom("x", "é", "漢", "😀"),
          undo: fc.boolean(),
        }),
        { minLength: 1, maxLength: 8 },
      ),
      async (edits) => {
        const session = createCanonicalSession(await openSource()).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        expect(session.captureSaveSnapshot().changedBlockIds).toEqual([]);
        for (const { id, text, undo } of edits) {
          const prepared = session
            .prepareOperations(state, [
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { story: OP_STORIES.MAIN, blockId: id, offset: 0 },
                text,
              },
            ])
            .unwrap();
          const staged = session.captureSaveSnapshot();
          expect(staged.version).toBe(session.version);
          state = state.apply(prepared.transaction);
          prepared.publish().unwrap();
          expect(session.captureSaveSnapshot().changedBlockIds).toContain(id);
          if (undo) {
            const inverse = session.prepareUndo(state).unwrap();
            state = state.apply(inverse.transaction);
            inverse.publish().unwrap();
            expect(session.captureSaveSnapshot().changedBlockIds).toContain(id);
          }
        }
        const transientId = "0abcdefa";
        const split = session
          .prepareOperations(state, [
            {
              type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
              at: { story: OP_STORIES.MAIN, blockId: IDS[2], offset: 2 },
              newBlockId: transientId,
              newHalf: "second",
            },
          ])
          .unwrap();
        state = state.apply(split.transaction);
        split.publish().unwrap();
        expect(session.captureSaveSnapshot().changedBlockIds).toContain(transientId);

        const join = session
          .prepareOperations(state, [
            {
              type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
              story: OP_STORIES.MAIN,
              blockId: IDS[2],
              nextBlockId: transientId,
              survivor: "first",
            },
          ])
          .unwrap();
        state = state.apply(join.transaction);
        join.publish().unwrap();
        const snapshot = session.captureSaveSnapshot();
        expect(snapshot.changedBlockIds).not.toContain(transientId);
        expect(session.document.package.document.content).toHaveLength(IDS.length);
        const selectiveDiagnostics: SaveDiagnostic[] = [];
        const selectiveSave = await serializeCanonicalSave({
          snapshot,
          featureFlags: { selectiveSave: true },
          options: {
            mode: FOLIO_DOCX_SERIALIZATION_MODE.preferSelective,
            onDiagnostic: (diagnostic) => selectiveDiagnostics.push(diagnostic),
          },
        });
        expect(selectiveDiagnostics.some(({ type }) => type === "selectiveSaveRefused")).toBe(
          false,
        );
        expect(
          describePackageDifferences(
            snapshot.document,
            await parseDocx(selectiveSave.buffer, { preloadFonts: false }),
          ),
        ).toEqual({ messages: [], omitted: 0 });
        for (const [index, id] of IDS.entries()) {
          if (!snapshot.changedBlockIds.includes(id))
            expect(await canonicalSaveParagraphXml(selectiveSave.buffer, id)).toBe(
              SOURCE_BLOCKS[index],
            );
        }
        for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
          const saved = await serializeCanonicalSave({ snapshot, options: { mode } });
          const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
          expect(describePackageDifferences(snapshot.document, reopened)).toEqual({
            messages: [],
            omitted: 0,
          });
          const xml = await (
            await JSZip.loadAsync(saved.buffer)
          )
            .file("word/document.xml")
            ?.async("text");
          for (const [index, id] of IDS.entries()) {
            if (!snapshot.changedBlockIds.includes(id)) expect(xml).toContain(SOURCE_BLOCKS[index]);
          }
        }
        const source = snapshot.document.package.document.content.at(0);
        if (source?.type === "paragraph") source.content = [];
        expect(session.captureSaveSnapshot().document.package.document.content.at(0)).not.toEqual(
          source,
        );
      },
    ),
    { seed: 20261004, numRuns: 12 },
  );
});

test("selective refusal emits a typed diagnostic and trusted full replay preserves untouched bytes", async () => {
  const session = createCanonicalSession(await openSource()).unwrap();
  const diagnostics: SaveDiagnostic[] = [];
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
    featureFlags: { selectiveSaveMaxBytes: 1 },
    options: { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) },
  });
  expect(diagnostics).toEqual([{ type: "selectiveSaveRefused", part: "word/document.xml" }]);
  expect(saved.diagnostics).toEqual(diagnostics);
  const xml = await (await JSZip.loadAsync(saved.buffer)).file("word/document.xml")?.async("text");
  for (const block of SOURCE_BLOCKS) expect(xml).toContain(block);
});

test("captured canonical saves retain the exact version when a newer commit arrives", async () => {
  const session = createCanonicalSession(await openSource()).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const snapshot = session.captureSaveSnapshot();
  const pending = serializeCanonicalSave({ snapshot });
  const edit = session
    .prepareOperations(state, [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story: OP_STORIES.MAIN, blockId: IDS[0], offset: 0 },
        text: "new",
      },
    ])
    .unwrap();
  state = state.apply(edit.transaction);
  edit.publish().unwrap();
  const saved = await pending;
  expect(saved.version).toBe(0);
  expect(session.captureSaveSnapshot().version).toBe(1);
  expect(
    describePackageDifferences(
      snapshot.document,
      await parseDocx(saved.buffer, { preloadFonts: false }),
    ),
  ).toEqual({ messages: [], omitted: 0 });
  session.beginComposition();
  expect(() => session.captureSaveSnapshot()).toThrow("Composition must finish");
});
