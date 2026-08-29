import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseDocx } from "../parser";
import { createDocx } from "../rezip";
import { createEmptyDocument } from "../../utils/createDocument";
import { extractDocxText } from "./extractDocxText";
import {
  FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME,
  FOLIO_YJS_UPDATE_MAX_BYTES,
  FolioYjsDocxMaterializationError,
  materializeYjsDocx,
} from "./materializeYjsDocx";

const createCollaborativeUpdate = async (sourceDocx: ArrayBuffer, text: string) => {
  const sourceDocument = await parseDocx(sourceDocx, { preloadFonts: false });
  const initialState = EditorState.create({ doc: toProseDoc(sourceDocument) });
  const bodyEnd = initialState.doc.content.size - 1;
  const nextState = initialState.apply(
    initialState.tr.replaceWith(1, bodyEnd, initialState.schema.text(text)),
  );
  const ydoc = new Y.Doc();
  prosemirrorToYXmlFragment(
    nextState.doc,
    ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
  );
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

describe("materializeYjsDocx", () => {
  test("replaces body content while preserving the source DOCX package", async () => {
    const sourceDocx = await createDocx(createEmptyDocument({ initialText: "Original text" }));
    const yjsUpdate = await createCollaborativeUpdate(sourceDocx, "Collaborative text");

    const output = await materializeYjsDocx({ sourceDocx, yjsUpdate });
    const extracted = await extractDocxText(output);
    const [sourceZip, outputZip] = await Promise.all([
      JSZip.loadAsync(sourceDocx),
      JSZip.loadAsync(output),
    ]);

    expect(extracted.paragraphs.map(({ text }) => text)).toContain("Collaborative text");
    expect(extracted.paragraphs.map(({ text }) => text)).not.toContain("Original text");
    expect(await outputZip.file("word/styles.xml")?.async("text")).toBe(
      await sourceZip.file("word/styles.xml")?.async("text"),
    );
  });

  test("rejects a state update without Folio's document fragment", async () => {
    const sourceDocx = await createDocx(createEmptyDocument());
    const ydoc = new Y.Doc();
    ydoc.getMap("unrelated").set("value", true);
    const yjsUpdate = Y.encodeStateAsUpdate(ydoc);
    ydoc.destroy();

    await expect(materializeYjsDocx({ sourceDocx, yjsUpdate })).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "missing_document",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });

  test("rejects an oversized update before parsing the source DOCX", async () => {
    await expect(
      materializeYjsDocx({
        sourceDocx: new Uint8Array([1, 2, 3]),
        yjsUpdate: new Uint8Array(FOLIO_YJS_UPDATE_MAX_BYTES + 1),
      }),
    ).rejects.toMatchObject({
      _tag: "FolioYjsDocxMaterializationError",
      code: "update_too_large",
    } satisfies Partial<FolioYjsDocxMaterializationError>);
  });
});
