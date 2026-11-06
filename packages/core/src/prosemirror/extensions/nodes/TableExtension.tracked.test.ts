/**
 * With Track Changes on, the editor's row and column commands make tracked
 * changes: accepting gives what the edit would have given directly, and
 * rejecting restores the table.
 */

import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createHarnessState,
  harnessManager,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
} from "../../../__tests__/editorHarness";
import { createDocx } from "../../../docx/rezip";
import { createEmptyDocument } from "../../../utils/createDocument";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (text: string) =>
  `<w:tc><w:tcPr/><w:p w14:paraId="1000000${text}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const TABLES = {
  "a plain table": `<w:tr>${cell("1")}${cell("2")}</w:tr><w:tr>${cell("3")}${cell("4")}</w:tr>`,
  "a table with a short row": `<w:tr>${cell("1")}${cell("2")}</w:tr><w:tr>${cell("3")}</w:tr>`,
  "a table with a declared w:gridAfter": `<w:tr>${cell("1")}${cell("2")}</w:tr><w:tr><w:trPr><w:gridAfter w:val="1"/></w:trPr>${cell("3")}</w:tr>`,
} as const;

const load = async (rows: string) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>${rows}</w:tbl><w:p w14:paraId="10000009"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return parseShapeDocument(new Uint8Array(await zip.generate()));
};

const run = (state: EditorState, command: string): EditorState => {
  let at = -1;
  state.doc.descendants((node, position) => {
    if (node.isTextblock && node.textContent === "1") at = position + 1;
    return at < 0;
  });
  const selected = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at)),
  );
  const factory = harnessManager().getCommand(command);
  if (!factory) throw new Error(`Missing command: ${command}`);
  let next = selected;
  const handled = factory()(selected, (tr: Transaction) => {
    next = selected.apply(tr);
  });
  expect(handled).toBe(true);
  return next;
};

/** Saved rows as the text of their cells, and no revision left over. */
const savedTable = async (
  state: EditorState,
  base: Awaited<ReturnType<typeof load>>,
) => {
  const { bytes } = await saveHarnessState(state, base);
  const xml =
    (await (
      await ZipArchive.read(bytes.slice().buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return {
    rows: [...xml.matchAll(/<w:tr[ >].*?<\/w:tr>/gsu)].map(([row]) =>
      [...row.matchAll(/<w:tc>.*?<\/w:tc>/gsu)].map(([saved]) =>
        [...saved.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
          .map(([, text]) => text)
          .join(""),
      ),
    ),
    revisions: /<w:(cellIns|cellDel|ins|del)[ />]/u.test(xml),
  };
};

describe.each(Object.entries(TABLES))("on %s", (_name, rows) => {
  describe.each([
    "addRowBelow",
    "addRowAbove",
    "addColumnLeft",
    "addColumnRight",
    "deleteRow",
    "deleteColumn",
  ])("%s with Track Changes on", (command) => {
    test("leaves a revision, not a finished change", async () => {
      const base = await load(rows);
      const tracked = run(createHarnessState(base, "suggesting"), command);
      expect((await savedTable(tracked, base)).revisions).toBe(true);
    });

    test("rejected, restores the table", async () => {
      const base = await load(rows);
      const source = await savedTable(
        createHarnessState(base, "editing"),
        base,
      );
      const tracked = run(createHarnessState(base, "suggesting"), command);
      const rejected = resolveAllChanges(tracked, "reject");
      expect(await savedTable(rejected, base)).toEqual(source);
    });

    test("accepted, gives what the direct command gives", async () => {
      const base = await load(rows);
      const direct = await savedTable(
        run(createHarnessState(base, "editing"), command),
        base,
      );
      const tracked = run(createHarnessState(base, "suggesting"), command);
      const accepted = resolveAllChanges(tracked, "accept");
      expect(await savedTable(accepted, base)).toEqual(direct);
    });
  });
});
