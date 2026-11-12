/**
 * A suggested column deletion writes the same table markup from the canonical
 * session as from the editor's own command: the grid stays while the deletion
 * is pending and each deleted cell is marked. Accepting removes the column and
 * its grid entry; rejecting restores the table exactly.
 */
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createHarnessState,
  harnessManager,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import {
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (id: string, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cell("10000001", "A")}${cell("10000002", "B")}${cell("10000003", "C")}</w:tr>` +
  `<w:tr>${cell("10000004", "D")}${cell("10000005", "E")}${cell("10000006", "F")}</w:tr></w:tbl>`;

const source = async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${TABLE}<w:p w14:paraId="20000001"/><w:sectPr/></w:body></w:document>`,
  );
  return new Uint8Array(await zip.generate());
};

/**
 * The saved table, with the revision stamps the two paths choose apart and the
 * namespace declarations a splice restates; everything else must be equal.
 */
const tableXml = async (bytes: Uint8Array | ArrayBuffer) => {
  const xml =
    (await (
      await ZipArchive.read(
        bytes instanceof Uint8Array ? bytes.slice().buffer : bytes,
      )
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  const table = /<w:tbl[ >].*<\/w:tbl>/su.exec(xml)?.[0] ?? "";
  // A spliced table restates the part's namespaces on its own start tag.
  return table
    .replace(/^<w:tbl[^>]*>/u, "<w:tbl>")
    .replaceAll(/ w:id="\d+"/gu, ' w:id="#"')
    .replaceAll(/ w:author="[^"]*"/gu, ' w:author="#"')
    .replaceAll(/ w:date="[^"]*"/gu, ' w:date="#"');
};

const caretIn = (state: EditorState, text: string) => {
  let at = -1;
  state.doc.descendants((node, position) => {
    if (node.isTextblock && node.textContent === text) at = position + 1;
    return at < 0;
  });
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at)),
  );
};

const canonicalSession = async () =>
  createCanonicalSession(
    await parseDocx(
      (await prepareCanonicalDocxInput(await source())).unwrap(),
      {
        preloadFonts: false,
        detectVariables: false,
      },
    ),
  ).unwrap();

const deleteColumn = (session: CanonicalSession, text: string) => {
  const state = caretIn(
    EditorState.create({ schema, doc: session.projection.doc }),
    text,
  );
  const command = harnessManager().getCommand("deleteColumn");
  if (!command) throw new Error("Missing command: deleteColumn");
  const intents = getCanonicalCommandIntents(command(), state);
  if (intents === undefined) throw new Error("deleteColumn has no lowering.");
  const commit = prepareCanonicalCommands(session, state, intents).unwrap();
  expect(commit.publish().isOk()).toBe(true);
  return state.apply(commit.transaction);
};

const canonicalSave = async (session: CanonicalSession) =>
  (
    await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    })
  ).buffer;

describe("a suggested column deletion in the canonical session", () => {
  test("writes the table the editor's own command writes", async () => {
    const base = await parseShapeDocument(await source());
    const editor = harnessManager().getCommand("deleteColumn");
    if (!editor) throw new Error("Missing command: deleteColumn");
    const tracked = caretIn(createHarnessState(base, "suggesting"), "B");
    let edited = tracked;
    expect(
      editor()(tracked, (transaction) => {
        edited = tracked.apply(transaction);
      }),
    ).toBe(true);
    const expected = await tableXml(
      (await saveHarnessState(edited, base)).bytes,
    );

    const session = await canonicalSession();
    session.setMode({ type: "suggesting", author: "Reviewer" });
    deleteColumn(session, "B");
    const saved = await tableXml(await canonicalSave(session));
    expect(saved).toBe(expected);
    // The grid keeps all three columns while the deletion is pending.
    expect(saved.match(/<w:gridCol /gu)).toHaveLength(3);
    expect(saved.match(/<w:cellDel /gu)).toHaveLength(2);
  });

  for (const decision of ["accept", "reject"] as const)
    test(`${decision}ed, saves and reopens as the ${decision === "accept" ? "direct deletion" : "source"}`, async () => {
      const direct = await canonicalSession();
      deleteColumn(direct, "B");
      const untouched = await canonicalSession();
      const expected = await tableXml(
        await canonicalSave(decision === "accept" ? direct : untouched),
      );

      const session = await canonicalSession();
      session.setMode({ type: "suggesting", author: "Reviewer" });
      const state = deleteColumn(session, "B");
      const ids = [
        ...JSON.stringify(session.document.package.document.content).matchAll(
          /"tableCellDeletion","info":\{"id":(\d+)/gu,
        ),
      ].map(([, id]) => Number(id));
      expect(ids).toHaveLength(2);
      const resolved = session
        .prepareResolve(state, { revisionIds: ids, resolution: decision })
        .unwrap();
      expect(resolved.publish().isOk()).toBe(true);
      const saved = await canonicalSave(session);
      expect(await tableXml(saved)).toBe(expected);
      // The file reopens to the same table.
      const reopened = createCanonicalSession(
        await parseDocx(new Uint8Array(saved), {
          preloadFonts: false,
          detectVariables: false,
        }),
      ).unwrap();
      expect(await tableXml(await canonicalSave(reopened))).toBe(expected);
    });

  for (const decision of ["accept", "reject"] as const)
    test(`two suggested column deletions ${decision}ed together save as the ${decision === "accept" ? "direct deletions" : "source"}`, async () => {
      const direct = await canonicalSession();
      deleteColumn(direct, "B");
      deleteColumn(direct, "C");
      const untouched = await canonicalSession();
      const expected = await tableXml(
        await canonicalSave(decision === "accept" ? direct : untouched),
      );
      const session = await canonicalSession();
      session.setMode({ type: "suggesting", author: "Reviewer" });
      deleteColumn(session, "B");
      const state = deleteColumn(session, "C");
      const ids = [
        ...JSON.stringify(session.document.package.document.content).matchAll(
          /"tableCellDeletion","info":\{"id":(\d+)/gu,
        ),
      ].map(([, id]) => Number(id));
      expect(ids).toHaveLength(4);
      const resolved = session
        .prepareResolve(state, { revisionIds: ids, resolution: decision })
        .unwrap();
      expect(resolved.publish().isOk()).toBe(true);
      expect(await tableXml(await canonicalSave(session))).toBe(expected);
    });
});
